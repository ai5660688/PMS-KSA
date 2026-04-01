using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("PLN_tbl")]
public class Pln
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public int PLN_ID { get; set; }

    public int PLN_Project_No { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? PLN_DATE { get; set; }

    public double? PLN_DIA { get; set; }
    public string? PLN_LOCATION { get; set; }
}
