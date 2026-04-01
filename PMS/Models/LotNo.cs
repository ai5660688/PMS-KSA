using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Lot_No_tbl")]
public class LotNo
{
    [Key]
    public int Lot_ID { get; set; }

    [Required, MaxLength(10)]
    public string Lot_No { get; set; } = string.Empty;

    [Column(TypeName = "datetime")]
    public DateTime? From_Date { get; set; }

    [Column(TypeName = "datetime")]
    public DateTime? To_Date { get; set; }

    public int? Lot_Project_No { get; set; }
}
