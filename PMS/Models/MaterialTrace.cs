using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("MATERIAL_TRACE_tbl")]
public class MaterialTrace
{
    [Key]
    public int N { get; set; }

    [MaxLength(10)] public string? Marian_MRR { get; set; }
    [MaxLength(10)] public string? Material_Type { get; set; }
    [MaxLength(30)] public string? MATR_GRADE { get; set; }
    [MaxLength(20)] public string? Material_Des { get; set; }

    // SQL float -> double?
    public double? Dia_In1 { get; set; }
    [MaxLength(8)] public string? SC1 { get; set; }
    public double? Dia_In2 { get; set; }
    [MaxLength(8)] public string? SC2 { get; set; }

    [MaxLength(150)] public string? Description { get; set; }
    public int? ACTUAL_PO_QTY { get; set; }
    public int? RECEIVED_Qty { get; set; }
    [MaxLength(35)] public string? IRC { get; set; }
    [MaxLength(70)] public string? Manufacturer_Supplier { get; set; }
    public int? PO { get; set; }
    [MaxLength(70)] public string? SHIPMENT_NO { get; set; }
    public DateTime? ATS { get; set; }
    public int? IDENTIFICATION_CODE { get; set; }
    public int? POS_NO { get; set; }
    public int? SUB_PO { get; set; }
    [MaxLength(70)] public string? HEAT_NO { get; set; }
    public int? EPM_NO { get; set; }
    [MaxLength(70)] public string? RFI_NO { get; set; }
    public DateTime? RFI_DATE { get; set; }
    [MaxLength(70)] public string? RELEASE_NOTE { get; set; }
    [MaxLength(7)] public string? RFI_STATUS { get; set; }
    [MaxLength(70)] public string? Remarks { get; set; }

    public int MATR_Project_No { get; set; }
    public DateTime? MATR_Updated_Date { get; set; }
    public int? MATR_Updated_By { get; set; }
}
