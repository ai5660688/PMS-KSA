using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Projects_tbl")]
public class Project
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.None)]
    public int Project_ID { get; set; }

    [MaxLength(50)] public string? Project_Name { get; set; }
    [MaxLength(30)] public string? Client { get; set; }
    [MaxLength(10)] public string? Contractor_Project_No { get; set; }
    [MaxLength(30)] public string? Contractor { get; set; }
    [MaxLength(20)] public string? PO_No { get; set; }
    [MaxLength(30)] public string? WS_Location { get; set; }
    [MaxLength(30)] public string? FW_Location { get; set; }
    [MaxLength(20)] public string? BI_JO { get; set; }
    [MaxLength(20)] public string? Contract_No { get; set; }
    [MaxLength(5)] public string? Line_Sheet { get; set; }

    // Audit fields
    [Column(TypeName = "datetime")]
    public DateTime? PR_Updated_Date { get; set; }
    public int? PR_Updated_By { get; set; }

    // NEW: link for welders project mapping (nullable to remain compatible with existing data)
    public int? Welders_Project_ID { get; set; }

    // NEW: Additional fields
    [MaxLength(15)] public string? Project_Type { get; set; }
    [MaxLength(15)] public string? ARAMCO_NON { get; set; }
    public bool Default_P { get; set; }
}