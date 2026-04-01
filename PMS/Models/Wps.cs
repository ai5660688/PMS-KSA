using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("WPS_tbl")]
public class Wps
{
    [Key]
    public int WPS_ID { get; set; }

    [Required, MaxLength(20)]
    public string WPS { get; set; } = string.Empty;

    public int? Project_WPS { get; set; } // nullable in DB

    [MaxLength(11)]
    public string? Weld_Process { get; set; }

    [Required]
    public bool PWHT { get; set; }

    [MaxLength(150)]
    public string? Dia_Range { get; set; }

    [MaxLength(150)]
    public string? Thickness_Range { get; set; }

    public double? Thickness_Range_From { get; set; }
    public double? Thickness_Range_To { get; set; }

    // NEW columns
    public string? WPS_Service { get; set; }          // nvarchar(MAX)
    public string? WPS_Pipe_Class { get; set; }       // nvarchar(MAX)

    [MaxLength(20)]
    public string? WPS_P_NO { get; set; }

    [MaxLength(150)]
    public string? WPS_Material { get; set; }

    [MaxLength(20)]
    public string? Electrode { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? WPS_Updated_Date { get; set; }
    public int? WPS_Updated_By { get; set; }
}