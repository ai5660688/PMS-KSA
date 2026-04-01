using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Other_NDE_tbl")]
public class OtherNde
{
    [Key]
    public int Joint_ID_NDE { get; set; }

    [MaxLength(4)] public string? OTHER_NDE_TYPE { get; set; }
    [MaxLength(8)] public string? OTHER_NDE_REPORT_NUMBER { get; set; }
    [Column(TypeName = "datetime")] public DateTime? OTHER_NDE_DATE { get; set; }

    [MaxLength(8)] public string? Root_PT { get; set; }
    [Column(TypeName = "datetime")] public DateTime? Root_PT_DATE { get; set; }

    [MaxLength(8)] public string? Bevel_PT { get; set; }
    [Column(TypeName = "datetime")] public DateTime? Bevel_PT_DATE { get; set; }

    [MaxLength(8)] public string? PMI_REPORT_NUMBER { get; set; }
    [Column(TypeName = "datetime")] public DateTime? PMI_DATE { get; set; }

    [MaxLength(50)] public string? Consumables_Used { get; set; }
    [MaxLength(30)] public string? Surface_Condition { get; set; }
    [MaxLength(30)] public string? Material_Specification { get; set; }

    [Column(TypeName = "datetime")] public DateTime? Calibration_Validity { get; set; }

    [MaxLength(35)] public string? Penetrant_Type { get; set; }
    [MaxLength(35)] public string? Developer_Type { get; set; }
    [MaxLength(35)] public string? Penetrant_Remover { get; set; }
    [MaxLength(35)] public string? Stage_of_Examination { get; set; }

    [MaxLength(50)] public string? Other_NDE_REMARKS { get; set; }

    [Column(TypeName = "datetime")] public DateTime? Other_NDE_Updated_Date { get; set; }

    public int? Other_NDE_Updated_By { get; set; }
}
