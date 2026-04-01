using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("PWHT_HT_tbl")]
public class PwhtHt
{
    [Key]
    public int Joint_ID_PWHT_HT { get; set; }

    [MaxLength(8)] public string? PWHT_REPORT_NUMBER { get; set; }
    [Column(TypeName = "datetime")] public DateTime? PWHT_DATE { get; set; }
    [MaxLength(13)] public string? HEAT_RATE { get; set; }
    [MaxLength(13)] public string? COOL_RATE { get; set; }
    [MaxLength(20)] public string? SOAK_TEMP { get; set; }
    public short? SOAK_TIME { get; set; }
    [MaxLength(5)] public string? Chart_No { get; set; }
    [MaxLength(13)] public string? Loading_Temp_C { get; set; }
    [MaxLength(15)] public string? Rate_of_Heating_C_Hr { get; set; }
    [MaxLength(13)] public string? Soaking_Temp_C { get; set; }
    [MaxLength(30)] public string? Soaking_Time_Mins { get; set; }
    [MaxLength(15)] public string? Rate_of_Cooling_C_Hr { get; set; }
    [MaxLength(13)] public string? Unloading_Temp_C { get; set; }
    [MaxLength(15)] public string? Searial_No { get; set; }
    [Column(TypeName = "datetime")] public DateTime? Cal_Due_Date { get; set; }
    [MaxLength(50)] public string? PWHT_HT_REMARKS { get; set; }
    [MaxLength(8)] public string? HARDNESS_TESTING_REPORT_NUMBER { get; set; }
    [Column(TypeName = "datetime")] public DateTime? HARDNESS_DATE { get; set; }
    public double? HAZ_1 { get; set; }
    public double? WELD_2 { get; set; }
    public double? HAZ_3 { get; set; }
    [Column(TypeName = "datetime")] public DateTime? PWHT_HT_Updated_Date { get; set; }
    public int? PWHT_HT_Updated_By { get; set; }
}
