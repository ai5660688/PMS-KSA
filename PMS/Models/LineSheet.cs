using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Line_Sheet_tbl")]
public class LineSheet
{
    [Key]
    public int Line_Sheet_ID { get; set; }

    public int? Project_No { get; set; }

    [MaxLength(10)] public string? LS_LAYOUT_NO { get; set; }
    [MaxLength(5)] public string? LS_SHEET { get; set; }
    [MaxLength(8)] public string? LS_REV { get; set; }
    public double? LS_DIAMETER { get; set; }
    [MaxLength(35)] public string? LS_REMARKS { get; set; }
    [MaxLength(10)] public string? LS_Scope { get; set; }
    public int? Line_ID_LS { get; set; }

    // Hold tracking columns
    [Column(TypeName="datetime")] public DateTime? LS_Hold_Date { get; set; }
    [Column(TypeName="datetime")] public DateTime? LS_Hold_Release_Date { get; set; }
    [MaxLength(100)] public string? LS_Hold_Reason { get; set; }

    // Audit and DWG fields
    [Column(TypeName="datetime")] public DateTime? LS_Updated_Date { get; set; }
    public int? LS_Updated_By { get; set; }

    [MaxLength(100)] public string? DWG_S_FileName { get; set; }
    public int? DWG_S_FileSize { get; set; }
    [Column(TypeName = "datetime2")] public DateTime? DWG_S_UploadDate { get; set; }
    [MaxLength(105)] public string? DWG_S_BlobName { get; set; }
}
