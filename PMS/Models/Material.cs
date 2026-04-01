using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Material_tbl")]
public class Material
{
    [Key]
    public int MAT_SN { get; set; }

    [Required]
    [MaxLength(20)]
    public string MAT { get; set; } = string.Empty;

    [Required]
    [MaxLength(30)]
    public string MAT_GRADE { get; set; } = string.Empty;

    [MaxLength(100)]
    public string? MAT_Notes { get; set; }
}
