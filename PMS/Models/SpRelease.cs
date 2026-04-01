using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System;

namespace PMS.Models;

[Table("SP_Release_tbl")]
public class SpRelease
{
    [Key]
    public int Spool_ID { get; set; }

    public int SP_Project_No { get; set; }

    [MaxLength(1)]
    public string? SP_TYPE { get; set; }

    // Diameter for the spool (set from DFR_tbl.DIAMETER when created from Daily Fit-up)
    public double? SP_DIA { get; set; }

    [MaxLength(10)]
    public string? SP_LAYOUT_NUMBER { get; set; }

    [MaxLength(5)]
    public string? SP_SHEET { get; set; }

    [MaxLength(30)]
    public string? SP_Material { get; set; }

    // Added: spool number string as present in database (if available)
    [MaxLength(50)]
    public string? SP_SPOOL_NUMBER { get; set; }

    // New: link back to Line_Sheet table when known
    public int? Line_Sheet_ID_SP { get; set; }

    // New: core dates used for pruning logic
    [Column(TypeName = "datetime")] public DateTime? SP_Date { get; set; }
    [Column(TypeName = "datetime")] public DateTime? SP_Date_SUPERSEDED { get; set; }

    // Hold tracking columns (used for highlighting in Daily Fit-up)
    [Column(TypeName = "datetime")] public DateTime? SP_Hold_Date { get; set; }
    [Column(TypeName = "datetime")] public DateTime? SP_Hold_Release_Date { get; set; }

    // New: QR fields needed for supersede transfer logic
    [MaxLength(8)]
    public string? SP_QR_NUMBER { get; set; }

    [MaxLength(8)]
    public string? SP_QR_SUPERSEDED { get; set; }

    // Added: status fields used by supersede logic
    [MaxLength(50)]
    public string? SP_STATUS { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? SP_STATUS_DATE { get; set; }

    // Audit columns (matching DFR_tbl pattern)
    [Column(TypeName = "datetime")]
    public DateTime? SP_Updated_Date { get; set; }
    public int? SP_Updated_By { get; set; }
}
