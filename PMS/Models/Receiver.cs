using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.EntityFrameworkCore;

namespace PMS.Models;

[Table("Receivers_tbl")]
[Keyless]
public class Receiver
{
    // Razor nvarchar(30) NOT NULL
    [Required]
    [MaxLength(30)]
    public string Razor { get; set; } = string.Empty;

    // Project_Receivers int NOT NULL
    [Required]
    public int Project_Receivers { get; set; }

    // Location_Receivers nvarchar(10) NULL
    [MaxLength(10)]
    public string? Location_Receivers { get; set; }

    // Receivers nvarchar(MAX) NULL
    public string? Receivers { get; set; }

    // Receivers_cc nvarchar(MAX) NULL (new)
    public string? Receivers_cc { get; set; }
}
