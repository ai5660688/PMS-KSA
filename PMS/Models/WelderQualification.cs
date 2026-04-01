using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models
{
    [Table("Welder_List_tbl")]
    public class WelderQualification
    {
        // Primary key: JCC_No (unique qualification identifier)
        [Key]
        [Required, StringLength(18)]
        public string JCC_No { get; set; } = string.Empty;

        // FK to Welder (many qualifications per welder)
        [ForeignKey(nameof(Welder))]
        [Column("Welder_ID_WL")]
        public int Welder_ID_WL { get; set; }

        public DateTime? Test_Date { get; set; }
        [StringLength(30)] public string? Welding_Process { get; set; }
        // DB column is nvarchar(100) (was 50) => adjust annotation to prevent validation truncation issues
        [StringLength(100)] public string? Material_P_No { get; set; }
        [StringLength(75)] public string? Code_Reference { get; set; }
        [StringLength(20)] public string? Consumable_Root_F_No { get; set; }
        [StringLength(40)] public string? Consumable_Root_Spec { get; set; }
        [StringLength(100)] public string? Consumable_Filling_Cap_F_No { get; set; }
        [StringLength(100)] public string? Consumable_Filling_Cap_Spec { get; set; }
        [StringLength(100)] public string? Position_Progression { get; set; }
        [StringLength(60)] public string? Max_Thickness { get; set; }
        [StringLength(60)] public string? Diameter_Range { get; set; }
        public DateTime? Date_Issued { get; set; }
        [StringLength(40)] public string? Remarks { get; set; }
        [StringLength(17)]
        [Column("Qualifiaction_Cert_Ref_No")]
        public string? Qualification_Cert_Ref_No { get; set; }
        [StringLength(30)] public string? WQT_Agency { get; set; }
        [StringLength(40)] public string? Received_from_Aramco { get; set; }
        public DateTime? DATE_OF_LAST_CONTINUITY { get; set; }
        [StringLength(20)] public string? RECORDING_THE_CONTINUITY_RECORD { get; set; }
        [Column("Batch_No")]
        // Keep nullable to avoid changing existing logic that uses .HasValue
        public int? Batch_No { get; set; }

        // Audit
        public DateTime? Welder_List_Updated_Date { get; set; }
        public int? Welder_List_Updated_By { get; set; }

        // NEW: Qualification file metadata
        [StringLength(100)] public string? JCC_FileName { get; set; }
        public int? JCC_FileSize { get; set; }
        public DateTime? JCC_UploadDate { get; set; }
        public int? JCC_Upload_By { get; set; }
        [StringLength(105)] public string? JCC_BlobName { get; set; }

        public Welder? Welder { get; set; }
    }
}