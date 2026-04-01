using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("PMS_Updated_Confirmed_tbl")]
public class UpdatedConfirmed
{
    [Key]
    public int U_C_ID { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime Updated_Confirmed_Date { get; set; }

    [MaxLength(4)]
    public string? U_C_Location { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? Fitup_Updated_Date { get; set; }

    [Column(TypeName = "decimal(6,2)")]
    public decimal? Fitup_Dia { get; set; }

    public int? Fitup_Updated_By { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? Welding_Updated_Date { get; set; }

    [Column(TypeName = "decimal(6,2)")]
    public decimal? Welding_Total_Dia { get; set; }

    public int? Welding_Updated_By { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? Fitup_Confirmed_Date { get; set; }

    [Column(TypeName = "decimal(6,2)")]
    public decimal? Fitup_Confirmed_Dia { get; set; }

    public int? Fitup_Confirmed_By { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? Welding_Confirmed_Date { get; set; }

    [Column(TypeName = "decimal(6,2)")]
    public decimal? Welding_Confirmed_Dia { get; set; }

    public int? Welding_Confirmed_By { get; set; }

    public int U_C_Project_No { get; set; }
}
