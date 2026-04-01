using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace PMS.Models;

public class DrawingKeyOption
{
 public string Key { get; set; } = string.Empty; // e.g., LAYOUT/SHEET or LAYOUT only
 public string Display { get; set; } = string.Empty; // shown in dropdown
}

public class DrawingRevisionVm
{
 public int Id { get; set; }
 public int RevisionOrder { get; set; }
 public string? RevisionTag { get; set; }
 public string? RawRevisionTag { get; set; }
 public string FileName { get; set; } = string.Empty;
 public int FileSize { get; set; }
 public string UploadDate { get; set; } = string.Empty;
 // NEW: Uploader full name
 public string? UploadedBy { get; set; }
}

public class DrawingsViewModel
{
 public int SelectedProjectId { get; set; }

 public string Mode { get; set; } = "Sheet"; // Sheet or Line

 public List<DrawingKeyOption> Keys { get; set; } = new();

 // Selected key fields
 public string? Layout { get; set; }
 public string? Sheet { get; set; }

 public string KeyDisplay => Mode == "Sheet" ? ($"{Layout}-{Sheet}") : (Layout ?? string.Empty);

 // Current revisions
 public List<DrawingRevisionVm> Revisions { get; set; } = new();

 // Upload input
 [Display(Name = "Revision Tag")] public string? NewRevisionTag { get; set; }
}
