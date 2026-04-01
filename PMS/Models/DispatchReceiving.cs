using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Dispatch_Receiving_tbl")]
public class DispatchReceiving
{
 [Key]
 public int Spool_ID_DR { get; set; }

 [MaxLength(8)]
 public string? DR_QR_NUMBER { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? DR_Date { get; set; }

 public int? DR_RFI_ID { get; set; }

 [MaxLength(8)]
 public string? DR_QR_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? DR_Date_SUPERSEDED { get; set; }

 [MaxLength(8)]
 public string? Bundle_No { get; set; }

 [MaxLength(70)]
 public string? DR_REMARKS { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? DELIVERY_DATE { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? DELIVERY_DATE_SUPERSEDED { get; set; }

 [MaxLength(8)]
 public string? DN_No { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? ARRIVED_Date { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? ARRIVED_Date_SUPERSEDED { get; set; }

 [MaxLength(8)]
 public string? MRI_RFI { get; set; }

 [MaxLength(8)]
 public string? MRI_RFI_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? MRI_Date { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? MRI_Date_SUPERSEDED { get; set; }

 [MaxLength(35)]
 public string? Received_By { get; set; }

 [MaxLength(35)]
 public string? Received_By_SUPERSEDED { get; set; }

 [MaxLength(35)]
 public string? Storage_Location { get; set; }

 [MaxLength(70)]
 public string? MRI_REMARKS { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Issue_Date { get; set; }

 public int? Installation_RFI_ID { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Installation_Date { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? DR_Updated_Date { get; set; }

 public int? DR_Updated_By { get; set; }
}
