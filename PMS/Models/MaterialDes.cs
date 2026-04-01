using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Material_Des_tbl")]
public class MaterialDes
{
    [Key]
    public int MATD_SN { get; set; }

    [Required]
    [MaxLength(30)]
    public string MATD_Description { get; set; } = string.Empty;
}
