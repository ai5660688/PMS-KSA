using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("RFI_tbl")]
public class Rfi
{
    [Key]
    public int RFI_ID { get; set; }

    [Column("RFI_Project_No")]
    public int? RFI_Project_No { get; set; }

    [MaxLength(25)] public string? DISCIPLINE { get; set; }

    [MaxLength(30)]
    [Column("Sub_Contractor")]
    public string? Sub_Contractor { get; set; }

    [Column("SUB_DISCIPLINE")]
    [MaxLength(25)]
    public string? SubDiscipline { get; set; }

    [MaxLength(30)]
    public string? RFI_CONTRACTOR { get; set; }

    [MaxLength(10)] public string? SubCon_RFI_No { get; set; }

    public int? EPMNo { get; set; }

    [MaxLength(25)] public string? SATIP { get; set; }
    [MaxLength(25)] public string? SAIC { get; set; }

    public double? ACTIVITY { get; set; }

    [MaxLength(25)] public string? RFI_LOCATION { get; set; }
    [MaxLength(15)] public string? UNIT { get; set; }
    [MaxLength(150)] public string? ELEMENT { get; set; }

    public string? RFI_DESCRIPTION { get; set; }

    [Column("DATE", TypeName = "datetime")] public DateTime? Date { get; set; }
    [Column("TIME")] public DateTime? Time { get; set; }

    [MaxLength(4)] public string? COMPANY_INSPECTION_LEVEL { get; set; }
    [MaxLength(4)] public string? CONTRACTOR_INSPECTION_LEVEL { get; set; }
    [MaxLength(4)] public string? SUBCON_INSPECTION_LEVEL { get; set; }

    [MaxLength(30)] public string? TR_QC { get; set; }
    [MaxLength(30)] public string? SUB_CON_QC { get; set; }
    [MaxLength(30)] public string? PID { get; set; }
    [MaxLength(30)] public string? PMT { get; set; }

    [MaxLength(50)] public string? INSPECTION_STATUS { get; set; }
    [MaxLength(50)] public string? REFRENCE_DRAWING_No { get; set; }
    [MaxLength(8)] public string? SCAN_COPY { get; set; }
    [MaxLength(100)] public string? REMARKS { get; set; }
    [MaxLength(500)] public string? QR_Code_Link { get; set; }

    [Column("RFI_Updated_Date", TypeName = "datetime")]
    public DateTime? RFI_Updated_Date { get; set; }

    public int? RFI_Updated_By { get; set; }

    [MaxLength(100)] public string? RFI_FileName { get; set; }
    public int? RFI_FileSize { get; set; }

    [Column(TypeName = "datetime2(7)")]
    public DateTime? RFI_UploadDate { get; set; }

    [MaxLength(105)] public string? RFI_BlobName { get; set; }
}