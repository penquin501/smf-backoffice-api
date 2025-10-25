# dbd_scraper_improved.py
# เวอร์ชันปรับปรุง: เพิ่ม debugging, retry logic, และ screenshot เมื่อเกิดข้อผิดพลาด

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import cv2
import numpy as np

import argparse
import re
import time
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# ------------------------- Selenium helpers ------------------------- #
def make_driver(download_dir: Path, headless: bool = False) -> webdriver.Chrome:
    opts = ChromeOptions()
    prefs = {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    opts.add_experimental_option("prefs", prefs)
    
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.add_argument("--log-level=3")

    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
    else:
        opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    return driver


def save_debug_screenshot(driver, name: str, out_dir: Path):
    """บันทึก screenshot และ page source สำหรับ debug"""
    try:
        ss_path = out_dir / f"debug_{name}_{int(time.time())}.png"
        driver.save_screenshot(str(ss_path))
        print(f"💾 Screenshot saved: {ss_path}")
        
        html_path = out_dir / f"debug_{name}_{int(time.time())}.html"
        html_path.write_text(driver.page_source, encoding='utf-8')
        print(f"💾 HTML saved: {html_path}")
    except Exception as e:
        print(f"⚠️  ไม่สามารถบันทึก debug files: {e}")


def wait_click(driver, locator, timeout=20):
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.3)
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    return el


def wait_send_keys(driver, locator, text, timeout=20):
    el = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    el.clear()
    el.send_keys(text)
    return el


def try_close_popups(driver, verbose=False):
    """ปิดป๊อปอัปทุกชนิด"""
    candidates = [
        (By.XPATH, "//button[normalize-space()='ปิด' or normalize-space()='Close']"),
        (By.XPATH, "//span[normalize-space()='ปิด']/ancestor::button"),
        (By.XPATH, "//button[@aria-label='Close' or @aria-label='ปิด']"),
        (By.XPATH, "//*[self::button or self::a][contains(.,'ยอมรับ') or contains(.,'ตกลง') or contains(.,'รับทราบ')]"),
        (By.XPATH, "//*[self::button or self::div][contains(@class,'close') or contains(@class,'modal-close')]"),
        (By.XPATH, "//div[contains(@class,'modal')]//button"),
        (By.XPATH, "//*[@role='dialog']//button"),
    ]
    closed_count = 0
    for by, xp in candidates:
        try:
            elems = driver.find_elements(by, xp)
            for e in elems:
                if e.is_displayed():
                    try:
                        driver.execute_script("arguments[0].click();", e)
                        closed_count += 1
                        time.sleep(0.3)
                    except:
                        pass
        except Exception:
            pass
    if verbose and closed_count > 0:
        print(f"✓ ปิดป๊อปอัป {closed_count} อัน")


def wait_for_downloads(folder: Path, before_set: set, timeout=90) -> Path:
    end = time.time() + timeout
    while time.time() < end:
        after = set(folder.glob("*"))
        new = [p for p in after - before_set if p.exists() and not p.name.endswith(".crdownload")]
        new = [p for p in new if p.stat().st_size > 0]
        if new:
            return sorted(new, key=lambda p: p.stat().st_mtime)[-1]
        time.sleep(0.5)
    raise TimeoutError("รอโหลดไฟล์ไม่ทันเวลา ลองใหม่อีกที")


# ------------------------- Scrape flow ------------------------- #
def search_by_juristic_id(driver, juristic_id: str, out_dir: Path):
    """ค้นหาเลขนิติบุคคล"""
    print(f"🔍 กำลังค้นหาเลขนิติบุคคล: {juristic_id}")
    driver.get("https://datawarehouse.dbd.go.th/index")
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2.0)
    try_close_popups(driver, verbose=True)

    # หาช่องค้นหา
    candidate_inputs = [
        (By.XPATH, "//input[@type='text' and (contains(@placeholder,'เลข') or contains(@placeholder,'ค้นหา') or contains(@placeholder,'010'))]"),
        (By.XPATH, "//input[contains(@class,'search') or contains(@name,'search')]"),
        (By.XPATH, "//div[contains(@class,'search')]//input"),
        (By.XPATH, "//input[@type='text']"),
    ]
    input_el = None
    for loc in candidate_inputs:
        try:
            input_el = WebDriverWait(driver, 8).until(EC.visibility_of_element_located(loc))
            print(f"✓ พบช่องค้นหา")
            break
        except Exception:
            continue
    
    if input_el is None:
        save_debug_screenshot(driver, "search_input_not_found", out_dir)
        raise RuntimeError("หา input สำหรับค้นหาไม่เจอ ตรวจสอบ screenshot ในโฟลเดอร์ output")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_el)
    input_el.clear()
    input_el.send_keys(juristic_id)
    time.sleep(0.5)

    # ปุ่มค้นหา
    search_button_found = False
    try:
        wait_click(driver, (By.XPATH, "//button[.//i[contains(@class,'search')] or contains(.,'ค้นหา') or .//span[contains(@class,'icon')]]"), 10)
        search_button_found = True
    except Exception:
        try:
            input_el.send_keys(u"\ue007")
            search_button_found = True
        except:
            pass
    
    if not search_button_found:
        save_debug_screenshot(driver, "search_button_not_found", out_dir)
        raise RuntimeError("ไม่สามารถกดปุ่มค้นหาได้")

    print("⏳ รอผลการค้นหา...")
    time.sleep(3.0)
    try_close_popups(driver, verbose=True)
    
    # รอให้หน้า profile ขึ้น
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(.,'ข้อมูลนิติบุคคล') or contains(.,'บริษัท')][not(self::script)]"))
        )
        print("✓ พบหน้าข้อมูลนิติบุคคล")
    except Exception:
        save_debug_screenshot(driver, "company_page_not_found", out_dir)
        raise RuntimeError("ไม่พบหน้าข้อมูลบริษัท อาจเป็นเพราะเลขนิติบุคคลไม่ถูกต้อง หรือระบบมีปัญหา")
    
    time.sleep(2.0)
    try_close_popups(driver, verbose=True)


def go_financial_tab(driver, out_dir: Path, max_retries=3):
    """ไปที่แท็บข้อมูลงบการเงิน"""
    print("📊 กำลังเปิดแท็บข้อมูลงบการเงิน...")
    
    for attempt in range(max_retries):
        try:
            try_close_popups(driver, verbose=True)
            time.sleep(1.0)
            
            # Scroll หน้าเว็บลงมาหน่อย เผื่อแท็บอยู่ข้างล่าง
            driver.execute_script("window.scrollTo(0, 300);")
            time.sleep(0.5)

            # ขั้นตอนที่ 1: เปิด dropdown "ข้อมูลงบการเงิน" ก่อน
            print("🔽 กำลังเปิด dropdown...")
            dropdown_opened = False
            dropdown_candidates = [
                (By.XPATH, "//span[contains(.,'ข้อมูลงบการเงิน')]/parent::*"),
                (By.XPATH, "//*[@class='dropdown' and contains(.,'ข้อมูลงบการเงิน')]"),
                (By.XPATH, "//*[contains(@class,'dropdown') and contains(.,'งบการเงิน')]"),
                (By.XPATH, "//li[@class='dropdown' or contains(@class,'dropdown')]//span[contains(.,'งบการเงิน')]"),
            ]
            
            for loc in dropdown_candidates:
                try:
                    dropdown = driver.find_element(loc[0], loc[1])
                    if dropdown.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dropdown)
                        time.sleep(0.3)
                        try:
                            dropdown.click()
                        except:
                            driver.execute_script("arguments[0].click();", dropdown)
                        print("✓ เปิด dropdown สำเร็จ")
                        dropdown_opened = True
                        time.sleep(1.0)
                        break
                except:
                    continue
            
            if not dropdown_opened:
                print("⚠️  ไม่พบ dropdown หรือไม่จำเป็นต้องเปิด กำลังค้นหาแท็บโดยตรง...")
            
            # ขั้นตอนที่ 2: คลิกที่ลิงก์ "งบการเงิน" ใน dropdown menu
            candidates = [
                (By.XPATH, "//a[@href='#tab22' or @lang='tab22']"),  # ตาม HTML ที่เห็น
                (By.XPATH, "//a[contains(@class,'tabinfo') and contains(.,'งบการเงิน')]"),
                (By.XPATH, "//ul[@class='dropdown-menu']//a[contains(.,'งบการเงิน')]"),
                (By.XPATH, "//a[contains(normalize-space(),'งบการเงิน')]"),
                (By.XPATH, "//button[contains(normalize-space(),'งบการเงิน')]"),
                (By.XPATH, "//*[contains(@class,'tab') and contains(.,'งบการเงิน')]"),
                (By.XPATH, "//*[@role='tab' and contains(.,'งบการเงิน')]"),
                (By.XPATH, "//div[contains(.,'งบการเงิน') and contains(@class,'clickable')]"),
                (By.XPATH, "//*[contains(text(),'งบการเงิน')]"),
            ]
            
            tab = None
            for i, loc in enumerate(candidates):
                try:
                    elements = driver.find_elements(loc[0], loc[1])
                    for el in elements:
                        if el.is_displayed():
                            print(f"✓ พบแท็บงบการเงิน (candidate #{i+1})")
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                            time.sleep(0.5)
                            try:
                                el.click()
                            except:
                                driver.execute_script("arguments[0].click();", el)
                            tab = el
                            break
                    if tab:
                        break
                except Exception as e:
                    continue

            if tab is None:
                if attempt < max_retries - 1:
                    print(f"⚠️  ไม่พบแท็บงบการเงิน (ครั้งที่ {attempt+1}/{max_retries}) กำลังลองใหม่...")
                    time.sleep(2.0)
                    continue
                else:
                    save_debug_screenshot(driver, "financial_tab_not_found", out_dir)
                    raise RuntimeError("หาแท็บ 'ข้อมูลงบการเงิน' ไม่เจอ แม้หลังจากลอง " + str(max_retries) + " ครั้ง")

            # รอให้เนื้อหาแท็บโหลด - เพิ่มเวลารอให้มากขึ้น
            print("⏳ รอโหลดเนื้อหางบการเงิน...")
            time.sleep(5.0)  # เพิ่มจาก 2 เป็น 5 วินาที
            
            # ลอง scroll ภายในพื้นที่เนื้อหา
            driver.execute_script("window.scrollTo(0, 500);")
            time.sleep(1.0)
            
            try_close_popups(driver, verbose=True)
            
            # ตรวจสอบว่าเนื้อหาโหลดแล้ว - ใช้ XPath ที่หลวมกว่า
            content_found = False
            try:
                # ลองหลาย pattern
                content_patterns = [
                    "//*[contains(.,'งบแสดงฐานะการเงิน') or contains(.,'ฐานะการเงิน')]",
                    "//*[contains(.,'งบกำไรขาดทุน') or contains(.,'กำไรขาดทุน')]",
                    "//*[contains(.,'อัตราส่วนทางการเงิน') or contains(.,'อัตราส่วน')]",
                    "//*[contains(.,'สินทรัพย์')]",
                    "//*[contains(.,'หนี้สิน')]",
                    "//*[contains(.,'รายได้')]",
                    "//table[contains(@class,'financial') or contains(@class,'report')]",
                    "//div[contains(@class,'financial-data') or contains(@class,'report-data')]",
                ]
                
                for pattern in content_patterns:
                    try:
                        elements = driver.find_elements(By.XPATH, pattern)
                        visible_elements = [el for el in elements if el.is_displayed() and el.text.strip()]
                        if visible_elements:
                            print(f"✓ พบเนื้อหางบการเงิน: {pattern[:50]}...")
                            content_found = True
                            break
                    except:
                        continue
                
                if content_found:
                    print("✓ เนื้อหางบการเงินโหลดสำเร็จ")
                    time.sleep(2.0)
                    return
                else:
                    raise Exception("ไม่พบเนื้อหางบการเงิน")
                    
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"⚠️  เนื้อหาไม่โหลด ({e}) กำลังลองใหม่...")
                    # ลอง refresh หน้า
                    driver.refresh()
                    time.sleep(3.0)
                    continue
                else:
                    save_debug_screenshot(driver, "financial_content_not_loaded", out_dir)
                    # ตรวจสอบว่ามีข้อความ error บนหน้าเว็บหรือไม่
                    page_text = driver.find_element(By.TAG_NAME, "body").text
                    print(f"📄 เนื้อหาบนหน้า (100 ตัวอักษรแรก): {page_text[:100]}")
                    raise RuntimeError("แท็บงบการเงินเปิดแล้ว แต่เนื้อหาไม่โหลด - อาจเป็นเพราะบริษัทนี้ไม่มีข้อมูลงบการเงินในระบบ")
                    
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"⚠️  เกิดข้อผิดพลาด: {e} กำลังลองใหม่...")
                time.sleep(2.0)
            else:
                raise


def click_print_excel(driver, out_dir: Path):
    """คลิกปุ่มพิมพ์ Excel"""
    try_close_popups(driver, verbose=True)

    # คลิกที่ dropdown "พิมพ์ข้อมูล"
    wait_click(driver, (By.XPATH, "//a[contains(.,'พิมพ์ข้อมูล')]"), 20)
    time.sleep(0.5)

    # คลิกที่ปุ่ม Excel (สังเกตว่าจริงๆ เป็น <a id="finXLS">)
    try:
        wait_click(driver, (By.XPATH, "//a[@id='finXLS' or contains(.,'Excel')]"), 10)
    except Exception as e:
        save_debug_screenshot(driver, "excel_button_not_found", out_dir)
        raise RuntimeError(f"หา Excel download link ไม่เจอ: {e}")

    time.sleep(1.0)
    
def download_company_info_pdf(driver, juristic_id: str, out_dir: Path) -> Path:
    """ดาวน์โหลด PDF ข้อมูลนิติบุคคลจากแท็บแรก - ใช้วิธี Print to PDF"""
    print("📄 กำลังดาวน์โหลด PDF ข้อมูลนิติบุคคล...")
    
    try:
        # ปิดป๊อปอัปก่อน
        try_close_popups(driver, verbose=True)
        time.sleep(1.0)
        
        # ตรวจสอบว่าอยู่ที่แท็บข้อมูลนิติบุคคลแล้ว
        try:
            company_tab_candidates = [
                (By.XPATH, "//a[contains(.,'ข้อมูลนิติบุคคล') or @href='#tab11']"),
                (By.XPATH, "//li[contains(@class,'active')]//a[contains(.,'ข้อมูลนิติบุคคล')]"),
            ]
            
            for loc in company_tab_candidates:
                try:
                    tab = driver.find_element(loc[0], loc[1])
                    if tab.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", tab)
                        time.sleep(0.3)
                        tab.click()
                        print("✓ เปิดแท็บข้อมูลนิติบุคคล")
                        time.sleep(2.0)
                        break
                except:
                    continue
        except Exception as e:
            print(f"⚠️  อาจอยู่ที่แท็บข้อมูลนิติบุคคลอยู่แล้ว: {e}")
        
        # Scroll เพื่อให้เห็นปุ่มพิมพ์ข้อมูล
        driver.execute_script("window.scrollTo(0, 200);")
        time.sleep(0.5)
        
        # บันทึก window handles ปัจจุบัน
        original_window = driver.current_window_handle
        original_windows = driver.window_handles
        
        # คลิกที่ปุ่ม "พิมพ์ข้อมูล" (ตรงมุมขวาบน)
        print_button_candidates = [
            (By.XPATH, "//button[contains(.,'พิมพ์ข้อมูล')]"),
            (By.XPATH, "//a[contains(.,'พิมพ์ข้อมูล')]"),
            (By.XPATH, "//*[contains(@class,'print') and contains(.,'พิมพ์')]"),
            (By.XPATH, "//button[.//i[contains(@class,'print')]]"),
            (By.CSS_SELECTOR, "button.btn-print"),
            (By.CSS_SELECTOR, "a.btn-print"),
        ]
        
        print_btn = None
        for loc in print_button_candidates:
            try:
                elements = driver.find_elements(loc[0], loc[1])
                for el in elements:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        time.sleep(0.5)
                        print("✓ พบปุ่มพิมพ์ข้อมูล")
                        try:
                            el.click()
                        except:
                            driver.execute_script("arguments[0].click();", el)
                        print_btn = el
                        break
                if print_btn:
                    break
            except:
                continue
        
        if print_btn is None:
            save_debug_screenshot(driver, "print_button_not_found", out_dir)
            raise RuntimeError("ไม่พบปุ่ม 'พิมพ์ข้อมูล'")
        
        # รอให้เปิดแท็บใหม่หรือดาวน์โหลด
        print("⏳ รอดาวน์โหลดไฟล์ PDF...")
        time.sleep(3.0)
        
        # ตรวจสอบว่ามีแท็บใหม่เปิดขึ้นมาหรือไม่
        new_windows = [w for w in driver.window_handles if w not in original_windows]
        
        if new_windows:
            # กรณีเปิดแท็บใหม่
            print("✓ ตรวจพบแท็บใหม่ กำลังดาวน์โหลดจากแท็บนั้น...")
            driver.switch_to.window(new_windows[0])
            time.sleep(2.0)
            
            # ใช้ Chrome DevTools Protocol สั่งพิมพ์เป็น PDF
            pdf_path = out_dir / f"{juristic_id}_company_info.pdf"
            
            try:
                result = driver.execute_cdp_cmd("Page.printToPDF", {
                    "landscape": False,
                    "printBackground": True,
                    "preferCSSPageSize": True,
                })
                
                import base64
                with open(pdf_path, 'wb') as f:
                    f.write(base64.b64decode(result['data']))
                
                print(f"✓ บันทึก PDF สำเร็จ: {pdf_path.name}")
                
                # ปิดแท็บและกลับไปแท็บเดิม
                driver.close()
                driver.switch_to.window(original_window)
                
                return pdf_path
            except Exception as e:
                print(f"⚠️  ไม่สามารถใช้ printToPDF: {e}")
                driver.close()
                driver.switch_to.window(original_window)
                raise
        else:
            # กรณีไม่มีแท็บใหม่ ลองรอไฟล์ดาวน์โหลดปกติ
            print("⏳ รอไฟล์ดาวน์โหลดแบบปกติ...")
            before = set(out_dir.glob("*"))
            
            # รอให้ไฟล์ดาวน์โหลด - เพิ่มเวลารอ
            timeout = 60
            end_time = time.time() + timeout
            downloaded = None
            
            while time.time() < end_time:
                time.sleep(1.0)
                
                # หาไฟล์ PDF ที่ดาวน์โหลดมาใหม่
                current_files = set(out_dir.glob("*.pdf"))
                new_pdfs = [f for f in current_files - before if f.exists()]
                
                # กรองไฟล์ที่กำลังดาวน์โหลดอยู่
                complete_pdfs = [f for f in new_pdfs if not f.name.endswith('.crdownload')]
                
                if complete_pdfs:
                    # เลือกไฟล์ที่ดาวน์โหลดล่าสุด
                    downloaded = sorted(complete_pdfs, key=lambda p: p.stat().st_mtime)[-1]
                    
                    # ตรวจสอบว่าไฟล์มีขนาด
                    if downloaded.stat().st_size > 1000:  # มากกว่า 1KB
                        print(f"✓ พบไฟล์ PDF: {downloaded.name} ({downloaded.stat().st_size} bytes)")
                        break
                
                # แสดงความคืบหน้า
                remaining = int(end_time - time.time())
                if remaining % 10 == 0:
                    print(f"   รอไฟล์... (เหลือ {remaining} วินาที)")
            
            if downloaded is None:
                # ถ้ายังไม่เจอ ลองหาไฟล์ PDF ทั้งหมดที่ดาวน์โหลดล่าสุด
                print("⚠️  ไม่เจอไฟล์ใหม่ กำลังตรวจสอบไฟล์ทั้งหมด...")
                all_pdfs = list(out_dir.glob("*.pdf"))
                if all_pdfs:
                    # เรียงตาม modified time
                    latest_pdf = sorted(all_pdfs, key=lambda p: p.stat().st_mtime)[-1]
                    # ตรวจสอบว่าไฟล์นี้ถูกแก้ไขใน 5 นาทีที่ผ่านมาหรือไม่
                    if time.time() - latest_pdf.stat().st_mtime < 300:
                        downloaded = latest_pdf
                        print(f"✓ พบไฟล์ PDF ล่าสุด: {downloaded.name}")
            
            if downloaded is None:
                print("❌ ไม่พบไฟล์ PDF ที่ดาวน์โหลด")
                # แสดงไฟล์ทั้งหมดในโฟลเดอร์
                all_files = list(out_dir.glob("*"))
                print(f"📁 ไฟล์ในโฟลเดอร์ ({len(all_files)} ไฟล์):")
                for f in sorted(all_files, key=lambda p: p.stat().st_mtime, reverse=True)[:5]:
                    print(f"   - {f.name} ({f.stat().st_size} bytes)")
                raise TimeoutError("ไม่พบไฟล์ PDF ที่ดาวน์โหลด")
            
            # เปลี่ยนชื่อไฟล์ให้สื่อความหมาย
            new_name = out_dir / f"{juristic_id}_company_info.pdf"
            
            # ถ้าไฟล์ชื่อเดียวกันมีอยู่แล้ว ให้ลบทิ้ง
            if new_name.exists():
                print(f"⚠️  พบไฟล์เก่า กำลังลบ: {new_name.name}")
                new_name.unlink()
            
            try:
                downloaded.rename(new_name)
                downloaded = new_name
                print(f"✓ เปลี่ยนชื่อเป็น: {downloaded.name}")
            except Exception as e:
                print(f"⚠️  ไม่สามารถเปลี่ยนชื่อไฟล์: {e}")
                print(f"   ใช้ชื่อเดิม: {downloaded.name}")
            
            print(f"✓ ดาวน์โหลด PDF สำเร็จ: {downloaded.name}")
            return downloaded
        
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการดาวน์โหลด PDF: {e}")
        save_debug_screenshot(driver, "pdf_download_error", out_dir)
        raise
    
def preprocess_image_for_ocr(image):
    """ปรับแต่งภาพก่อนทำ OCR"""
    # แปลงเป็น numpy array
    img_array = np.array(image)
    
    # แปลงเป็น grayscale
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else:
        gray = img_array
    
    # เพิ่มความคมชัด
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    
    # ลด noise
    gray = cv2.medianBlur(gray, 3)
    
    return Image.fromarray(gray)


def extract_text_from_pdf(pdf_path: Path, out_dir: Path) -> pd.DataFrame:
    """แปลง PDF เป็นข้อความด้วย OCR และจัดเก็บใน DataFrame"""
    print(f"📖 กำลังทำ OCR ไฟล์: {pdf_path.name}")
    
    try:
        # แปลง PDF เป็นรูปภาพ
        print("   🖼️  แปลง PDF เป็นรูปภาพ...")
        images = convert_from_path(pdf_path, dpi=300)
        print(f"   ✓ พบ {len(images)} หน้า")
        
        # เก็บข้อมูลแต่ละหน้า
        all_data = []
        
        for page_num, image in enumerate(images, 1):
            print(f"   🔍 กำลัง OCR หน้า {page_num}/{len(images)}...")
            
            # ปรับแต่งภาพก่อน OCR
            processed_img = preprocess_image_for_ocr(image)
            
            # ทำ OCR (รองรับภาษาไทยและอังกฤษ)
            try:
                text = pytesseract.image_to_string(
                    processed_img, 
                    lang='tha+eng',  # ภาษาไทย + อังกฤษ
                    config='--psm 6'  # Assume uniform block of text
                )
            except Exception as e:
                print(f"   ⚠️  ไม่สามารถ OCR ด้วยภาษาไทย: {e}")
                print(f"   ℹ️  ลอง OCR ด้วยภาษาอังกฤษอย่างเดียว...")
                text = pytesseract.image_to_string(
                    processed_img, 
                    lang='eng',
                    config='--psm 6'
                )
            
            # เก็บข้อมูลแต่ละบรรทัด
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            for line_num, line in enumerate(lines, 1):
                all_data.append({
                    'หน้า': page_num,
                    'บรรทัด': line_num,
                    'ข้อความ': line
                })
            
            print(f"   ✓ หน้า {page_num}: พบ {len(lines)} บรรทัด")
        
        # สร้าง DataFrame
        df = pd.DataFrame(all_data)
        print(f"✓ OCR เสร็จสมบูรณ์: รวม {len(df)} บรรทัด")
        
        return df
        
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการ OCR: {e}")
        save_debug_screenshot(None, "ocr_error", out_dir)
        raise


def extract_structured_data(df: pd.DataFrame, juristic_id: str) -> pd.DataFrame:
    """แยกข้อมูลสำคัญจาก OCR results - ปรับให้เหมาะกับรูปแบบ DBD"""
    print("🔍 กำลังแยกข้อมูลสำคัญ...")
    
    structured_data = []
    all_text = '\n'.join(df['ข้อความ'].tolist())
    
    # รูปแบบข้อมูลที่ต้องการหา - ปรับให้ตรงกับ PDF จริง
    patterns = {
        'ชื่อบริษัท': [
            r'บริษัท\s+([^\n]+?)\s+จำกัด',
            r'ชื่อ\s*[:：]\s*([^\n]+)',
        ],
        'เลขทะเบียนนิติบุคคล': [
            r'เลขทะเบียนนิติบุคคล\s*[:：]\s*([0-9]+)',
            r'เลข.*?นิติบุคคล.*?[:：]\s*([0-9]+)',
        ],
        'ประเภทนิติบุคคล': [
            r'ประเภทนิติบุคคล\s*[:：]\s*([^\n]+)',
        ],
        'วันที่จดทะเบียนจัดตั้ง': [
            r'วันที่จดทะเบียนจัดตั้ง\s*[:：]\s*([0-9/]+)',
            r'วันจดทะเบียน.*?[:：]\s*([0-9/]+)',
        ],
        'สถานะนิติบุคคล': [
            r'สถานะนิติบุคคล\s*[:：]\s*([^\n]+)',
        ],
        'ทุนจดทะเบียน': [
            r'ทุนจดทะเบียน.*?[:：]\s*([0-9,\.]+)',
        ],
        'ที่ตั้ง': [
            r'ที่ตั้ง\s*[:：]\s*([^\n]+(?:\n(?![^\n]*[:：])[^\n]+)*)',
        ],
        'หมวดธุรกิจตอนจดทะเบียน': [
            r'หมวดธุรกิจตอนจดทะเบียน\s*[:：]\s*([^\n]+)',
        ],
        'หมวดธุรกิจจากงบการเงิน': [
            r'หมวดธุรกิจ\s*\(มาจากงบการเงินปีล่าสุด\)\s*[:：]\s*([^\n]+)',
        ],
        'วัตถุประสงค์ตอนจดทะเบียน': [
            r'วัตถุประสงค์ตอนจดทะเบียน\s*[:：]\s*([^\n]+)',
        ],
        'วัตถุประสงค์จากงบการเงิน': [
            r'วัตถุประสงค์\s*\(มาจากงบการเงินปีล่าสุด\)\s*[:：]\s*([^\n]+)',
        ],
        'ปีที่ส่งงบการเงิน': [
            r'ปีที่ส่งงบการเงิน\s*[:：]\s*([^\n]+)',
        ],
    }
    
    # หาข้อมูลตาม pattern
    for field_name, pattern_list in patterns.items():
        for pattern in pattern_list:
            match = re.search(pattern, all_text, re.IGNORECASE | re.MULTILINE)
            if match:
                value = match.group(1).strip()
                # ตัด whitespace ที่เกินมา
                value = re.sub(r'\s+', ' ', value)
                
                structured_data.append({
                    'juristic_id': juristic_id,
                    'ฟิลด์': field_name,
                    'ค่า': value,
                })
                break  # ถ้าเจอแล้วไม่ต้องลอง pattern อื่น
    
    # แยกข้อมูลกรรมการ
    print("   🔍 กำลังแยกข้อมูลกรรมการ...")
    director_match = re.search(r'กรรมการ\s*[:：]\s*(.*?)(?=คณะกรรมการลงชื่อผูกพัน|ข้อควรทราบ|$)', 
                               all_text, re.DOTALL)
    if director_match:
        directors_text = director_match.group(1)
        # แยกแต่ละคน (รูปแบบ: เลขที่.ชื่อ)
        directors = re.findall(r'\d+\.\s*([^\n]+)', directors_text)
        if directors:
            directors_list = '\n'.join([f"{i+1}. {d.strip()}" for i, d in enumerate(directors)])
            structured_data.append({
                'juristic_id': juristic_id,
                'ฟิลด์': 'กรรมการ',
                'ค่า': directors_list,
            })
            print(f"   ✓ พบกรรมการ {len(directors)} คน")
    
    # แยกข้อมูลคณะกรรมการลงชื่อผูกพัน
    signing_match = re.search(r'คณะกรรมการลงชื่อผูกพัน\s*[:：]\s*(.*?)(?=รวมเป็น|ข้อควรทราบ|$)', 
                              all_text, re.DOTALL)
    if signing_match:
        signing_text = signing_match.group(1).strip()
        signing_text = re.sub(r'\s+', ' ', signing_text)
        structured_data.append({
            'juristic_id': juristic_id,
            'ฟิลด์': 'คณะกรรมการลงชื่อผูกพัน',
            'ค่า': signing_text,
        })
    
    if structured_data:
        result_df = pd.DataFrame(structured_data)
        print(f"✓ พบข้อมูลสำคัญ {len(result_df)} รายการ")
        
        # เรียงลำดับฟิลด์ให้เป็นระเบียบ
        field_order = [
            'ชื่อบริษัท', 'เลขทะเบียนนิติบุคคล', 'ประเภทนิติบุคคล', 
            'วันที่จดทะเบียนจัดตั้ง', 'สถานะนิติบุคคล', 'ทุนจดทะเบียน',
            'ที่ตั้ง', 'หมวดธุรกิจตอนจดทะเบียน', 'หมวดธุรกิจจากงบการเงิน',
            'วัตถุประสงค์ตอนจดทะเบียน', 'วัตถุประสงค์จากงบการเงิน',
            'ปีที่ส่งงบการเงิน', 'กรรมการ', 'คณะกรรมการลงชื่อผูกพัน'
        ]
        result_df['ฟิลด์'] = pd.Categorical(result_df['ฟิลด์'], categories=field_order, ordered=True)
        result_df = result_df.sort_values('ฟิลด์').reset_index(drop=True)
        
        return result_df
    else:
        print("⚠️  ไม่พบข้อมูลที่ตรงกับรูปแบบ - จะส่งข้อความทั้งหมดแทน")
        df['juristic_id'] = juristic_id
        df['ฟิลด์'] = 'ข้อความทั่วไป'
        return df[['juristic_id', 'หน้า', 'ฟิลด์', 'ข้อความ']]


def ocr_pdf_to_excel(pdf_path: Path, juristic_id: str, out_dir: Path) -> Path:
    """ทำ OCR และบันทึกเป็น Excel"""
    print("="*60)
    print("🔤 เริ่มกระบวนการ OCR")
    print("="*60)
    
    try:
        # ตรวจสอบว่าติดตั้ง Tesseract หรือยัง
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            print("❌ ไม่พบ Tesseract OCR!")
            print("📥 กรุณาติดตั้ง:")
            print("   Windows: https://github.com/UB-Mannheim/tesseract/wiki")
            print("   Mac: brew install tesseract tesseract-lang")
            print("   Linux: sudo apt-get install tesseract-ocr tesseract-ocr-tha")
            raise RuntimeError("ต้องติดตั้ง Tesseract OCR ก่อน")
        
        # OCR ไฟล์ PDF
        ocr_df = extract_text_from_pdf(pdf_path, out_dir)
        
        # บันทึกข้อความทั้งหมด
        raw_output = out_dir / f"{juristic_id}_company_info_ocr_raw.xlsx"
        ocr_df.to_excel(raw_output, index=False)
        print(f"💾 บันทึกข้อความทั้งหมด: {raw_output.name}")
        
        # แยกข้อมูลสำคัญ
        structured_df = extract_structured_data(ocr_df, juristic_id)
        
        # บันทึกข้อมูลที่จัดโครงสร้างแล้ว
        structured_output = out_dir / f"{juristic_id}_company_info_ocr_structured.xlsx"
        structured_df.to_excel(structured_output, index=False)
        print(f"💾 บันทึกข้อมูลสำคัญ: {structured_output.name}")
        
        print("="*60)
        print("✅ OCR เสร็จสมบูรณ์!")
        print(f"📁 ไฟล์ข้อความทั้งหมด: {raw_output}")
        print(f"📁 ไฟล์ข้อมูลสำคัญ: {structured_output}")
        print("="*60)
        
        return structured_output
        
    except Exception as e:
        print(f"❌ OCR ล้มเหลว: {e}")
        raise

def switch_report(driver, name: str, out_dir: Path):
    """สลับไปรายงานอื่น"""
    print(f"📄 กำลังสลับไปรายงาน: {name}")
    try_close_popups(driver, verbose=True)
    time.sleep(0.5)
    
    try:
        _ = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, f"//*[self::a or self::button or self::div][contains(.,'{name}')]"))
        )
    except Exception:
        pass
    
    wait_click(driver, (By.XPATH, f"//*[self::a or self::button or self::div][contains(.,'{name}')]"), 20)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, f"//*[contains(.,'{name}')]")))
    time.sleep(1.0)


def scrape_and_download_all(driver, download_dir: Path) -> Dict[str, Path]:
    """ดาวน์โหลดงบการเงินทั้ง 3 แบบ"""
    files = {}

    print("📥 กำลังดาวน์โหลดงบแสดงฐานะการเงิน...")
    before = set(download_dir.glob("*"))
    click_print_excel(driver, download_dir)
    files["งบแสดงฐานะการเงิน"] = wait_for_downloads(download_dir, before, timeout=120)
    print(f"✓ ดาวน์โหลดสำเร็จ: {files['งบแสดงฐานะการเงิน'].name}")

    print("📥 กำลังดาวน์โหลดงบกำไรขาดทุน...")
    switch_report(driver, "งบกำไรขาดทุน", download_dir)
    before = set(download_dir.glob("*"))
    click_print_excel(driver, download_dir)
    files["งบกำไรขาดทุน"] = wait_for_downloads(download_dir, before, timeout=120)
    print(f"✓ ดาวน์โหลดสำเร็จ: {files['งบกำไรขาดทุน'].name}")

    print("📥 กำลังดาวน์โหลดอัตราส่วนทางการเงิน...")
    switch_report(driver, "อัตราส่วนทางการเงิน", download_dir)
    before = set(download_dir.glob("*"))
    click_print_excel(driver, download_dir)
    files["อัตราส่วนทางการเงิน"] = wait_for_downloads(download_dir, before, timeout=120)
    print(f"✓ ดาวน์โหลดสำเร็จ: {files['อัตราส่วนทางการเงิน'].name}")

    return files


# ------------------------- Tidy & Merge ------------------------- #
def buddhist_to_gregorian(year):
    try:
        y = int(re.findall(r"\d{4}", str(year))[0])
        return y - 543 if y > 2400 else y
    except Exception:
        return None


def tidy_excel(path: Path, report_type: str, juristic_id: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, header=0)
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)

    year_cols = []
    for c in df.columns:
        if re.search(r"(25\d{2}|20\d{2})", str(c)):
            year_cols.append(c)
    if not year_cols:
        year_cols = [c for c in df.columns if re.search(r"\d{4}", str(c))]

    name_col = df.columns[0]
    long_df = df.melt(id_vars=[name_col], value_vars=year_cols, var_name="ปี", value_name="ค่า")
    long_df.rename(columns={name_col: "รายการ"}, inplace=True)
    long_df["ปี"] = long_df["ปี"].apply(buddhist_to_gregorian)
    long_df["report_type"] = report_type
    long_df["juristic_id"] = juristic_id
    long_df = long_df.dropna(subset=["ค่า"]).reset_index(drop=True)
    return long_df


def merge_to_single_excel(download_map: Dict[str, Path], out_path: Path, juristic_id: str) -> pd.DataFrame:
    print("🔄 กำลังรวมไฟล์...")
    frames = []
    for rep, p in download_map.items():
        frames.append(tidy_excel(Path(p), rep, juristic_id))
    merged = pd.concat(frames, ignore_index=True)

    def tag_cat(x: str) -> str:
        s = str(x)
        if re.search("สินทรัพย์|ทรัพย์สิน", s):
            return "สินทรัพย์"
        if re.search("หนี้สิน|ส่วนของผู้ถือหุ้น", s):
            return "หนี้สิน/ส่วนของผู้ถือหุ้น"
        if re.search("รายได้|ค่าใช้จ่าย|ต้นทุน|กำไร|ขาดทุน", s):
            return "รายได้/ค่าใช้จ่าย"
        return "อื่นๆ"

    merged["หมวด"] = merged["รายการ"].apply(tag_cat)
    cols = ["juristic_id", "report_type", "ปี", "หมวด", "รายการ", "ค่า"]
    merged = merged[cols].sort_values(["report_type", "ปี", "หมวด", "รายการ"]).reset_index(drop=True)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    merged.to_excel(out_path, index=False)
    print(f"✓ รวมไฟล์สำเร็จ: {out_path}")
    return merged


# ------------------------- Main ------------------------- #
def main():
    parser = argparse.ArgumentParser(description="Scrape DBD Datawarehouse financials (improved version)")
    parser.add_argument("--juristic-id", required=True, help="เลขนิติบุคคล เช่น 0105560001219")
    parser.add_argument("--out-dir", default="./downloads", help="โฟลเดอร์เก็บไฟล์")
    parser.add_argument("--headless", action="store_true", help="รันแบบไม่โชว์หน้าต่าง")
    parser.add_argument("--skip-pdf", action="store_true", help="ข้ามการดาวน์โหลด PDF ข้อมูลนิติบุคคล")
    parser.add_argument("--skip-ocr", action="store_true", help="ข้ามการทำ OCR (ใช้เมื่อไม่มี Tesseract)")
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print("="*60)
    print("🚀 DBD Financial Scraper (Improved Version)")
    print("="*60)

    driver = make_driver(out_dir, headless=args.headless)
    try:
        # ค้นหาบริษัท
        search_by_juristic_id(driver, args.juristic_id, out_dir)
        
        # ดาวน์โหลด PDF ข้อมูลนิติบุคคล (ถ้าไม่ได้ skip)
        pdf_file = None
        if not args.skip_pdf:
            try:
                pdf_file = download_company_info_pdf(driver, args.juristic_id, out_dir)
                
                # ทำ OCR (ถ้าไม่ได้ skip)
                if pdf_file and not args.skip_ocr:
                    try:
                        ocr_result = ocr_pdf_to_excel(pdf_file, args.juristic_id, out_dir)
                    except Exception as e:
                        print(f"⚠️  OCR ล้มเหลว: {e}")
                        print("💡 ใช้ --skip-ocr เพื่อข้ามขั้นตอนนี้")
                        
            except Exception as e:
                print(f"⚠️  ไม่สามารถดาวน์โหลด PDF: {e}")
                print("📝 จะทำการดาวน์โหลดงบการเงินต่อ...")
        
        # ไปที่แท็บงบการเงิน
        go_financial_tab(driver, out_dir)
        
        # ดาวน์โหลดงบการเงิน
        downloaded = scrape_and_download_all(driver, out_dir)

        # เปลี่ยนชื่อไฟล์งบการเงิน
        rename_map = {}
        for rep, p in downloaded.items():
            suffix = {
                "งบแสดงฐานะการเงิน": "balance",
                "งบกำไรขาดทุน": "income",
                "อัตราส่วนทางการเงิน": "ratios",
            }[rep]
            newp = out_dir / f"{args.juristic_id}_{suffix}.xlsx"
            try:
                Path(p).rename(newp)
                rename_map[rep] = newp
            except Exception:
                rename_map[rep] = Path(p)

        # รวมไฟล์งบการเงิน
        merged_path = out_dir / f"{args.juristic_id}_dbd_merged.xlsx"
        merge_to_single_excel(rename_map, merged_path, args.juristic_id)
        
        print("="*60)
        print(f"✅ เสร็จสมบูรณ์!")
        if pdf_file:
            print(f"📄 ไฟล์ PDF: {pdf_file}")
            if not args.skip_ocr:
                print(f"🔤 ไฟล์ OCR: {args.juristic_id}_company_info_ocr_*.xlsx")
        print(f"📁 ไฟล์งบการเงินรวม: {merged_path}")
        print("="*60)
        
    except Exception as e:
        print("\n" + "="*60)
        print(f"❌ เกิดข้อผิดพลาด: {e}")
        print("="*60)
        save_debug_screenshot(driver, "final_error", out_dir)
        raise
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()