using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("DFR_tbl")]
public class Dfr
{
    [Key]
    public int Joint_ID { get; set; }

    [MaxLength(4)] public string? LOCATION { get; set; }
    [MaxLength(10)] public string? LAYOUT_NUMBER { get; set; }
    [MaxLength(6)] public string? WELD_NUMBER { get; set; }
    [MaxLength(8)] public string? J_Add { get; set; }
    [MaxLength(4)] public string? WELD_TYPE { get; set; }
    [MaxLength(5)] public string? SHEET { get; set; }
    [MaxLength(8)] public string? DFR_REV { get; set; }
    [MaxLength(9)] public string? SPOOL_NUMBER { get; set; }
    public double? DIAMETER { get; set; }
    [MaxLength(8)] public string? SCHEDULE { get; set; }
    [MaxLength(20)] public string? MATERIAL_A { get; set; }
    [MaxLength(20)] public string? MATERIAL_B { get; set; }
    [MaxLength(30)] public string? GRADE_A { get; set; }
    [MaxLength(30)] public string? GRADE_B { get; set; }
    [MaxLength(20)] public string? HEAT_NUMBER_A { get; set; }
    [MaxLength(20)] public string? HEAT_NUMBER_B { get; set; }
    public int? WPS_ID_DFR { get; set; }
    [Column(TypeName="datetime")] public DateTime? FITUP_DATE { get; set; }
    [MaxLength(8)] public string? FITUP_INSPECTION_QR_NUMBER { get; set; }
    [MaxLength(9)] public string? TACK_WELDER { get; set; }
    [MaxLength(35)] public string? TP_NUMBER { get; set; }
    public double? Special_RT { get; set; }
    public double? OL_DIAMETER { get; set; }
    [MaxLength(8)] public string? OL_SCHEDULE { get; set; }
    public double? OL_Thick { get; set; }
    public bool Deleted { get; set; }
    public bool Cancelled { get; set; }
    [Column(TypeName="datetime")] public DateTime? DFR_Hold_Date { get; set; }
    [Column(TypeName="datetime")] public DateTime? DFR_Hold_Release_Date { get; set; }
    [MaxLength(100)] public string? DFR_Hold_Reason { get; set; }
    public bool Fitup_Confirmed { get; set; }
    [MaxLength(35)] public string? DFR_REMARKS { get; set; }
    public int? RFI_ID_DFR { get; set; }
    public int? Sch_ID_DFR { get; set; }
    public int? Spool_ID_DFR { get; set; }
    public int? Line_Sheet_ID_DFR { get; set; }
    public int Project_No { get; set; }
    [Column(TypeName="datetime")] public DateTime? DFR_Updated_Date { get; set; }
    public int? DFR_Updated_By { get; set; }
    [MaxLength(8)] public string? HOLD_DFR { get; set; }
    [Column(TypeName="datetime")] public DateTime? HOLD_DFR_Date_D { get; set; }
    public int? U_C_ID_DFR { get; set; }
    public int? MATR_ID_DFR_1 { get; set; }
    public int? MATR_ID_DFR_2 { get; set; }

    [MaxLength(35)] public string? DFR_Test_Package_No_WS { get; set; }
    [MaxLength(35)] public string? DFR_Test_Package_No { get; set; }
}
