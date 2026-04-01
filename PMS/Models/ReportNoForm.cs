using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.EntityFrameworkCore;

namespace PMS.Models;

[Table("Report_No_Form_tbl")]
[Keyless]
public class ReportNoForm
{
    [Required]
    public int Project_Report { get; set; }

    [Required]
    [MaxLength(20)]
    public string Report_Type { get; set; } = string.Empty;

    [Required]
    [MaxLength(2)]
    public string Report_Location { get; set; } = string.Empty;

    [MaxLength(25)]
    public string? Report_No_Form { get; set; }

    [MaxLength(20)]
    public string? Remarks { get; set; }
}
