<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\BolBsRaw;

class PublicApiController extends Controller
{
    public function bol_bs_store(Request $request)
    {
        // Mapping: Request key -> Database field
        $fieldMap = [
            'company_id' => 'company_id',
            'company_name' => 'company_name',
            'year' => 'year',
            'Assets' => 'assets',
            'Cash and deposits at financial institutions' => 'cash_and_deposits_at_financial_institutions',
            'Accounts receivable' => 'accounts_receivable',
            'Accounts and notes receivable - net' => 'accounts_and_notes_receivable_net',
            'Total short-term loans consolidation' => 'total_short_term_loans_consolidation',
            'Inventories-net' => 'inventories_net',
            'Accrued income' => 'accrued_income',
            'Prepaid expenses' => 'prepaid_expenses',
            'Other current assets' => 'other_current_assets',
            'Others - Total current assets' => 'others_total_current_assets',
            'Total current assets' => 'total_current_assets',
            'Total long-term loans and investments' => 'total_long_term_loans_and_investments',
            'Property, plant and equipment - net' => 'property_plant_and_equipment_net',
            'Other non-current assets' => 'other_non_current_assets',
            'Others - Total non-current assets' => 'others_total_non_current_assets',
            'Total non-current assets' => 'total_non_current_assets',
            'Total assets' => 'total_assets',
            "Liabilities and shareholders' equity" => 'liabilities_and_shareholders_equity',
            'Liabilities' => 'liabilities',
            'Bank overdrafts and short-term loans from financial institutions' => 'bank_overdrafts_and_short_term_loans_from_financial_institutions',
            'Accounts payable' => 'accounts_payable',
            'Total accounts payable and notes payable' => 'total_accounts_payable_and_notes_payable',
            'Current portion of long-term loans' => 'current_portion_of_long_term_loans',
            'Total short-term loans' => 'total_short_term_loans',
            'Accrued expenses' => 'accrued_expenses',
            'Unearned revenues' => 'unearned_revenues',
            'Other current liabilities' => 'other_current_liabilities',
            'Others - Total current liabilities' => 'others_total_current_liabilities',
            'Total current liabilities' => 'total_current_liabilities',
            'Total long-term loans' => 'total_long_term_loans',
            'Other non-current liabilities' => 'other_non_current_liabilities',
            'Others - Total non-current liabilities' => 'others_total_non_current_liabilities',
            'Total non-current liabilities' => 'total_non_current_liabilities',
            'Total Liabilities' => 'total_liabilities',
            "Shareholder's equity" => 'shareholders_equity',
            'Authorized preferred stocks' => 'authorized_preferred_stocks',
            'Authorized common stocks' => 'authorized_common_stocks',
            'Issued and paid-up preferred stocks' => 'issued_and_paid_up_preferred_stocks',
            'Issued and paid-up common stocks' => 'issued_and_paid_up_common_stocks',
            'Appraisal surplus on property, plant and equipment' => 'appraisal_surplus_on_property_plant_and_equipment',
            'Accumulated retained earnings' => 'accumulated_retained_earnings',
            'Others' => 'others',
            'Total shareholders\' equity' => 'total_shareholders_equity',
            'Total liabilities and shareholders\' equity' => 'total_liabilities_and_shareholders_equity',
            'Additional information for shareholders\' equity' => 'additional_information_for_shareholders_equity',
            'Common stocks' => 'common_stocks',
            'No.of shares - Authorized' => 'no_of_shares_authorized',
            'Par value (Baht) - Authorized' => 'par_value_baht_authorized',
            'No.of shares - Issued and paid-up' => 'no_of_shares_issued_and_paid_up',
            'Par value (Baht) - Issued and Paid-up' => 'par_value_baht_issued_and_paid_up',
        ];

        $data = [];
        foreach ($request->all() as $key => $value) {
            $mappedItem = [];
            foreach ($value as $k => $v) {
                if (isset($fieldMap[$k])) {
                    $mappedItem[$fieldMap[$k]] = $v;
                }
            }
            $data[] = $mappedItem;
        }
        // dd($data);
        $checked = [];
        foreach ($data as $kk => $item) {
            $validated = validator($item, [
                'company_id' => 'required|string',
                'company_name' => 'nullable|string',
                'year' => 'required|integer',
                'assets' => 'nullable|numeric',
                'cash_and_deposits_at_financial_institutions' => 'nullable|numeric',
                'accounts_receivable' => 'nullable|numeric',
                'accounts_and_notes_receivable_net' => 'nullable|numeric',
                'total_short_term_loans_consolidation' => 'nullable|numeric',
                'inventories_net' => 'nullable|numeric',
                'accrued_income' => 'nullable|numeric',
                'prepaid_expenses' => 'nullable|numeric',
                'other_current_assets' => 'nullable|numeric',
                'others_total_current_assets' => 'nullable|numeric',
                'total_current_assets' => 'nullable|numeric',
                'total_long_term_loans_and_investments' => 'nullable|numeric',
                'property_plant_and_equipment_net' => 'nullable|numeric',
                'other_non_current_assets' => 'nullable|numeric',
                'others_total_non_current_assets' => 'nullable|numeric',
                'total_non_current_assets' => 'nullable|numeric',
                'total_assets' => 'nullable|numeric',
                'liabilities_and_shareholders_equity' => 'nullable|numeric',
                'liabilities' => 'nullable|numeric',
                'bank_overdrafts_and_short_term_loans_from_financial_institutions' => 'nullable|numeric',
                'accounts_payable' => 'nullable|numeric',
                'total_accounts_payable_and_notes_payable' => 'nullable|numeric',
                'current_portion_of_long_term_loans' => 'nullable|numeric',
                'total_short_term_loans' => 'nullable|numeric',
                'accrued_expenses' => 'nullable|numeric',
                'unearned_revenues' => 'nullable|numeric',
                'other_current_liabilities' => 'nullable|numeric',
                'others_total_current_liabilities' => 'nullable|numeric',
                'total_current_liabilities' => 'nullable|numeric',
                'total_long_term_loans' => 'nullable|numeric',
                'other_non_current_liabilities' => 'nullable|numeric',
                'others_total_non_current_liabilities' => 'nullable|numeric',
                'total_non_current_liabilities' => 'nullable|numeric',
                'total_liabilities' => 'nullable|numeric',
                'shareholders_equity' => 'nullable|numeric',
                'authorized_preferred_stocks' => 'nullable|numeric',
                'authorized_common_stocks' => 'nullable|numeric',
                'issued_and_paid_up_preferred_stocks' => 'nullable|numeric',
                'issued_and_paid_up_common_stocks' => 'nullable|numeric',
                'appraisal_surplus_on_property_plant_and_equipment' => 'nullable|numeric',
                'accumulated_retained_earnings' => 'nullable|numeric',
                'others' => 'nullable|numeric',
                'total_shareholders_equity' => 'nullable|numeric',
                'total_liabilities_and_shareholders_equity' => 'nullable|numeric',
                'additional_information_for_shareholders_equity' => 'nullable|numeric',
                'common_stocks' => 'nullable|numeric',
                'no_of_shares_authorized' => 'nullable|integer',
                'par_value_baht_authorized' => 'nullable|numeric',
                'no_of_shares_issued_and_paid_up' => 'nullable|integer',
                'par_value_baht_issued_and_paid_up' => 'nullable|numeric',
            ])->validate();

            $checked[] = $validated;
        }

        $output = [];

        foreach ($checked as $item) {
            $bolBs = BolBsRaw::updateOrCreate(
                [
                    'company_id' => $item['company_id'],
                    'year' => $item['year']
                ],
                $item
            );

            $output[] = $bolBs;
        }

        return response()->json([
            'success' => true,
            'data' => $output,
            'message' => 'All data saved successfully.',
        ]);
    }
}
