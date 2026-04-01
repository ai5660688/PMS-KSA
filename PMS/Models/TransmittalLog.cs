using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Transmittal_Log_tbl")]
public class TransmittalLog
{
    [Key]
    [Column("ID")]
    public int ID { get; set; }

    [MaxLength(30)] public string? Transmittal { get; set; }

    [Column("Date", TypeName = "datetime")]
    public DateTime? Date { get; set; }

    [MaxLength(30)] public string? Doc_No { get; set; }

    [MaxLength(20)] public string? eGesDoc_No { get; set; }

    [MaxLength(150)] public string? Documents_Title { get; set; }

    [MaxLength(3)] public string? REV { get; set; }

    [MaxLength(50)] public string? Remarks { get; set; }

    [MaxLength(30)] public string? Transmittal_No { get; set; }

    [MaxLength(30)] public string? TR_Comment { get; set; }

    [Column("Date_TR", TypeName = "datetime")]
    public DateTime? Date_TR { get; set; }

    [MaxLength(3)] public string? TR_REV { get; set; }

    [MaxLength(30)] public string? SAPMT_Comment { get; set; }

    [Column("Date_SAPMT", TypeName = "datetime")]
    public DateTime? Date_SAPMT { get; set; }

    [MaxLength(3)] public string? SAPMT_REV { get; set; }

    [MaxLength(50)] public string? Categorize { get; set; }

    [MaxLength(25)] public string? Discipline { get; set; }

    public int? TR_Project_No { get; set; }

    [Column("Transmittal_Updated_Date", TypeName = "datetime")]
    public DateTime? Transmittal_Updated_Date { get; set; }

    public int? Transmittal_Updated_By { get; set; }

    [MaxLength(100)] public string? TR_FileName { get; set; }

    public int? TR_FileSize { get; set; }

    [Column(TypeName = "datetime2(7)")]
    public DateTime? TR_UploadDate { get; set; }

    [MaxLength(105)] public string? TR_BlobName { get; set; }
}
