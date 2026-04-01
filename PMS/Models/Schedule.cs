using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Schedule_tbl")]
public class Schedule
{
    [Key]
    public int Sch_ID { get; set; }

    [MaxLength(8)]
    public string? SCH { get; set; }      // Schedule designation (e.g., STD, XS)

    public double? NPS { get; set; }      // Nominal pipe size / diameter

    public double? DN { get; set; }       // Added: nominal diameter (mm)
    public double? OD { get; set; }       // Added: outside diameter (mm)
    public double? OD_In { get; set; }    // Added: outside diameter (inches?)

    public double? THICK { get; set; }    // Nominal wall thickness

    // Added Olet specific thickness columns
    public double? OLET_THICK_SS { get; set; }
    public double? OLET_THICK_CS { get; set; }
    public double? OLET_THICK_SSS { get; set; }
}
