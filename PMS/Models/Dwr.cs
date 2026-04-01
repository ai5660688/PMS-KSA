using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("DWR_tbl")]
public class Dwr
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.None)]
    public int Joint_ID_DWR { get; set; }

    [MaxLength(9)] public string? ROOT_A { get; set; }
    [MaxLength(9)] public string? ROOT_B { get; set; }
    [MaxLength(9)] public string? FILL_A { get; set; }
    [MaxLength(9)] public string? FILL_B { get; set; }
    [MaxLength(9)] public string? CAP_A { get; set; }
    [MaxLength(9)] public string? CAP_B { get; set; }

    [Column(TypeName = "datetime")] public DateTime? DATE_WELDED { get; set; }
    [Column(TypeName = "datetime")] public DateTime? ACTUAL_DATE_WELDED { get; set; }

    public int? WPS_ID_DWR { get; set; }

    [MaxLength(8)] public string? POST_VISUAL_INSPECTION_QR_NO { get; set; }

    public double? PREHEAT_TEMP_C { get; set; }

    [MaxLength(2)] public string? Open_Closed { get; set; }
    [MaxLength(8)] public string? IP_or_T { get; set; }

    public bool Weld_Confirmed { get; set; }

    [MaxLength(50)] public string? DWR_REMARKS { get; set; }

    public bool PID_SELECTION { get; set; }

    [Column(TypeName = "datetime")] public DateTime? SELECTION_DATE { get; set; }

    public bool HT_SELECTION { get; set; }

    [Column(TypeName = "datetime")] public DateTime? HT_SELECTION_DATE { get; set; }

    public int? RFI_ID_DWR { get; set; }

    [Column(TypeName = "datetime")] public DateTime? DWR_Updated_Date { get; set; }

    public int? DWR_Updated_By { get; set; }

    public int? DWR_U_C_ID { get; set; }
}
