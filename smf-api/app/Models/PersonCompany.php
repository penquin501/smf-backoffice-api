<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

/**
 * Class PersonCompany
 *
 * @property int $id
 * @property string|null $registered_no         เลขทะเบียนนิติบุคคล
 * @property string|null $citizen_id            เลขบัตรประชาชน / Tax ID
 * @property string|null $first_name            ชื่อ
 * @property string|null $last_name             นามสกุล
 * @property string|null $phone                 เบอร์โทรศัพท์
 * @property bool|null $is_owner                เป็นเจ้าของหรือไม่ (1=ใช่)
 * @property int|null $director_no              ลำดับกรรมการ
 * @property string|null $boj5_doc_no           เลขเอกสาร บอจ.5 (ถ้ามี)
 */
class PersonCompany extends Model
{
    use HasFactory;

    protected $table = 'person_company';

    protected $fillable = [
        'registered_no',
        'citizen_id',
        'prefix',
        'first_name',
        'last_name',
        'phone',
        'is_owner',
        'director_no',
        'boj5_doc_no',
    ];

    protected $casts = [
        'is_owner' => 'boolean',
        'director_no' => 'integer',
    ];

    /**
     * Relationship: เชื่อมกับ CompanyEntity (ผ่าน registered_no)
     */
    public function company()
    {
        return $this->belongsTo(CompanyEntity::class, 'registered_no', 'registered_no');
    }

    // /**
    //  * Accessor: รวมชื่อเต็ม
    //  */
    // public function getFullNameAttribute(): string
    // {
    //     return trim(($this->first_name ?? '') . ' ' . ($this->last_name ?? ''));
    // }
}
