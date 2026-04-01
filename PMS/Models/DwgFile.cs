using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("DWG_File_tbl")]
public class DwgFile
{
 [Key]
 public int Id { get; set; }

 public int Project_ID { get; set; }

 [Required]
 [MaxLength(10)]
 public string Mode { get; set; } = null!;

 public int? Line_Sheet_ID { get; set; }

 public int? Line_ID { get; set; }

 [MaxLength(100)]
 public string? FileName { get; set; }

 public int FileSize { get; set; }

 [Column(TypeName = "datetime2(7)")]
 public DateTime UploadDate { get; set; }

 // NEW: Uploader UserID (same pattern as Welder_List_Updated_By)
 public int? UploadBy { get; set; }

 [Required]
 [MaxLength(105)]
 public string BlobName { get; set; } = null!;

 [MaxLength(100)]
 public string? ContentType { get; set; }

 public int RevisionOrder { get; set; }

 [MaxLength(20)]
 public string? RevisionTag { get; set; }
}
