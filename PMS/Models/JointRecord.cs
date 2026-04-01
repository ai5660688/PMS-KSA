using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace PMS.Models;

public partial class JointRecordAnalysisRow
{
    // Layout/Line Information
    [Display(Name = "Layout No")]
    [Required(ErrorMessage = "Layout No is required")]
    [RegularExpression(@"^[A-Z0-9]{1,4}-\d{3,5}$", ErrorMessage = "Format: FLUID-LINENO (e.g., P-11645)")]
    public string? LayoutNo { get; set; }

    [Display(Name = "Line Revision")]
    public string? LRev { get; set; }

    [Display(Name = "Sheet Revision")]
    public string? LSRev { get; set; }

    [Display(Name = "Unit No")]
    public string? UnitNumber { get; set; }

    [Display(Name = "ISO")]
    [RegularExpression(@"^[A-Z0-9]{1,4}-\d{3,5}-[A-Z0-9]{3,12}$", ErrorMessage = "Format: FLUID-LINENO-LINECLASS")]
    public string? ISO { get; set; }

    [Display(Name = "Line Class")]
    public string? LineClass { get; set; }

    [Display(Name = "Material")]
    public string? Material { get; set; }

    [Display(Name = "Line No")]
    [Required(ErrorMessage = "Line No is required")]
    public string? LineNo { get; set; }

    [Display(Name = "Fluid")]
    [Required(ErrorMessage = "Fluid is required")]
    public string? Fluid { get; set; }

    // Sheet Information
    [Display(Name = "Sheet No")]
    [Required(ErrorMessage = "Sheet No is required")]
    public string? LSSheet { get; set; }

    [Display(Name = "ISO Dia")]
    public string? LSDiameter { get; set; }

    [Display(Name = "Scope")]
    public string? LSScope { get; set; }

    // Weld Information
    [Display(Name = "Location")]
    public string? Location { get; set; }

    [Display(Name = "Weld No")]
    [Required(ErrorMessage = "Weld No is required")]
    [RegularExpression(@"^[A-Z0-9]+$", ErrorMessage = "Weld No must be alphanumeric")]
    public string? WeldNumber { get; set; }

    [Display(Name = "J-Add")]
    [Required(ErrorMessage = "J-Add is required")]
    [RegularExpression(@"^(NEW|[A-Z][A-Z0-9]*)$", ErrorMessage = "Format: NEW or alphanumeric code (e.g., R1, CS1)")]
    public string? JAdd { get; set; }

    [Display(Name = "Weld Type")]
    [RegularExpression(@"^(BW|FW|SW|SOF|BR|LET|TH|SP|FJ)$", ErrorMessage = "Invalid weld type")]
    public string? WeldType { get; set; }

    // Spool Information
    [Display(Name = "Spool")]
    public string? SpoolNumber { get; set; }

    [Display(Name = "Dia In")]
    public string? Diameter { get; set; }

    [Display(Name = "Spool Dia")]
    public string? SpDia { get; set; }

    // Schedule Information
    [Display(Name = "Schedule")]
    public string? Schedule { get; set; }

    [Display(Name = "OL Dia")]
    public string? OLDiameter { get; set; }

    [Display(Name = "OL Schedule")]
    public string? OLSchedule { get; set; }

    [Display(Name = "OL Thick")]
    public string? OLThick { get; set; }

    // Material Information
    [Display(Name = "Material A")]
    public string? MaterialA { get; set; }

    [Display(Name = "Material B")]
    public string? MaterialB { get; set; }

    [Display(Name = "Grade A")]
    public string? GradeA { get; set; }

    [Display(Name = "Grade B")]
    public string? GradeB { get; set; }

    [Display(Name = "System Remarks")]
    public string? Remarks { get; set; }
    public string? SourceFile { get; set; }

    // Delete / Cancel status (from DFR_tbl)
    public bool Deleted { get; set; }
    public bool Cancelled { get; set; }
    public string? FitupDate { get; set; }

    // Validation method
    public List<string> Validate()
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(LayoutNo))
            errors.Add("Layout No is required");
        else if (!LayoutRegex().IsMatch(LayoutNo))
            errors.Add("Invalid Layout No format");

        if (string.IsNullOrWhiteSpace(WeldNumber))
            errors.Add("Weld No is required");
        else if (!WeldNumberRegex().IsMatch(WeldNumber))
            errors.Add("Weld No must be alphanumeric");

        if (string.IsNullOrWhiteSpace(JAdd))
            errors.Add("J-Add is required");
        else if (!JAddRegex().IsMatch(JAdd))
            errors.Add("Invalid J-Add format");

        // Weld Type is optional but validate format if provided
        if (!string.IsNullOrWhiteSpace(WeldType) && !WeldTypeRegex().IsMatch(WeldType))
            errors.Add("Invalid Weld Type");

        // Weld type specific validation
        if ((WeldType == "BR" || WeldType == "LET") &&
            (string.IsNullOrWhiteSpace(OLDiameter) || OLDiameter == "##"))
        {
            errors.Add("OL Diameter required for branch weld type");
        }

        return errors;
    }
}

partial class JointRecordAnalysisRow
{
    [GeneratedRegex(@"^[A-Z0-9]{1,4}-\d{3,5}$", RegexOptions.CultureInvariant)]
    private static partial Regex LayoutRegex();

    [GeneratedRegex(@"^[A-Z0-9]+$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex WeldNumberRegex();

    [GeneratedRegex(@"^(NEW|[A-Z][A-Z0-9]*)$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex JAddRegex();

    [GeneratedRegex(@"^(BW|FW|SW|SOF|BR|LET|TH|SP|FJ)$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
    private static partial Regex WeldTypeRegex();
}

public class JointRecordAnalysisResult
{
    public int ProjectId { get; set; }
    public string? AnalysisName { get; set; }
    public List<JointRecordAnalysisRow> Rows { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public List<string> Errors { get; set; } = new();
}

public class JointRecordViewModel
{
    public List<ProjectOption> Projects { get; set; } = new();
    public int SelectedProjectId { get; set; }
    public string Scope { get; set; } = "Joints";
    public string Mode { get; set; } = "upload";
    public string? FilterLayout { get; set; }
    public string? FilterSheet { get; set; }
    public List<JointRecordAnalysisRow> Rows { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public List<string> MaterialDescriptions { get; set; } = new();
    public List<string> Materials { get; set; } = new();
    public List<string> Locations { get; set; } = new();
    public List<string> WeldTypes { get; set; } = new();
    public List<string> Grades { get; set; } = new();
    public List<string> JAddOptions { get; set; } = new();
}

public class JointRecordCommitRequest
{
    public int ProjectId { get; set; }
    public string Scope { get; set; } = "Joints";
    public string Mode { get; set; } = "upload";
    public bool UpdateOnlyNonNull { get; set; }
    public List<JointRecordAnalysisRow> Rows { get; set; } = new();
}

public class JointRecordCommitResult
{
    public bool Success { get; set; }
    public int Created { get; set; }
    public int Updated { get; set; }
    public int Skipped { get; set; }
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public Dictionary<int, string> RowRemarks { get; set; } = new();
}

public class BulkAnalyzeRequest
{
    public int ProjectId { get; set; }
    public string Scope { get; set; } = "Joints";
    public List<int>? DrawingIds { get; set; }
}

public class BulkAnalyzeDrawingItem
{
    public int ProjectId { get; set; }
    public string Layout { get; set; } = string.Empty;
    public string? Sheet { get; set; }
    public string? RevisionTag { get; set; }
    public string? FileName { get; set; }
    public int DrawingId { get; set; }
}

public class BulkAnalyzeResult
{
    public int ProjectId { get; set; }
    public int TotalDrawings { get; set; }
    public int SuccessfulDrawings { get; set; }
    public int FailedDrawings { get; set; }
    public List<JointRecordAnalysisRow> Rows { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public List<string> Errors { get; set; } = new();
    public List<BulkAnalyzeDrawingItem> ProcessedDrawings { get; set; } = new();
}

public class DrawingMetadata
{
    public string? PlantNo { get; set; }
    public string? DrawingNo { get; set; }
    public string? SheetNo { get; set; }
    public string? LRev { get; set; }
    public string? LSRev { get; set; }
    public string? Iso { get; set; }
    public string? Fluid { get; set; }
    public string? LineNo { get; set; }
    public string? LineClass { get; set; }
    public string? Diameter { get; set; }
    public string? Schedule { get; set; }
    public string? Material { get; set; }
    public string? Location { get; set; }
    public List<WeldRecord> Welds { get; set; } = new();
    public List<string> Spools { get; set; } = new();
    public List<MaterialInfo> Materials { get; set; } = new();
    public string? DrawingTitle { get; set; }
    public string? Scale { get; set; }
    public string? DrawnBy { get; set; }
    public string? CheckedBy { get; set; }
    public string? ApprovedBy { get; set; }
    public string? Date { get; set; }
    public List<string> BillOfMaterials { get; set; } = new();
    public List<string> RevisionHistory { get; set; } = new();
    public List<string> Notes { get; set; } = new();
    public bool HasInspectionRequirements { get; set; }
    public List<CuttingListEntry> CuttingListEntries { get; set; } = new();
    /// <summary>
    /// Spool number (e.g. "01") → primary pipe SIZE from the cutting list.
    /// Built from the mode (most frequent) SIZE per spool in the cutting list table.
    /// </summary>
    public Dictionary<string, string> CuttingListSpoolSizes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>
    /// Ordered list of spool numbers (e.g. ["01","02","03"]) derived from the
    /// cutting list, spool labels, or WS group count. Used by the successive
    /// WS group algorithm to assign spools to rows in the correct order.
    /// </summary>
    public List<string> OrderedSpools { get; set; } = new();
}

public class WeldRecord
{
    public string? Number { get; set; }
    public string? Type { get; set; }
    public string? Size { get; set; }
    public string? Schedule { get; set; }
    public string? Thickness { get; set; }
    public string? ShopCode { get; set; }
    public string? ShopNo { get; set; }
    public string? SpoolNo { get; set; }
    public string? SpoolDia { get; set; }
    public string? Date { get; set; }
    public string? Remarks { get; set; }
    public string? JAdd { get; set; }
    public string? MaterialA { get; set; }
    public string? MaterialB { get; set; }
    public string? GradeA { get; set; }
    public string? GradeB { get; set; }
    public string? OLDiameter { get; set; }
    public string? OLSchedule { get; set; }
    public string? OLThick { get; set; }
    public string? Location { get; set; }
    public string? Process { get; set; }
    /// <summary>
    /// When true, Material A/B and Grade A/B were explicitly extracted from a
    /// structured weld list table (e.g. PdfSaudiWeldListRowRegex,
    /// PdfWeldTableRowWithLocationRegex, PdfWeldTableRowRegex). BOM inference
    /// must not override these values.
    /// </summary>
    public bool ExplicitTableData { get; set; }
    /// <summary>
    /// Line index in the OCR text where this weld was detected.
    /// Used for PT NO proximity-based Material A/B inference.
    /// </summary>
    public int SourceLineIndex { get; set; } = -1;
}

public class CuttingListEntry
{
    public string? PieceNo { get; set; }
    public string? SpoolNo { get; set; }
    public string? Size { get; set; }
    public string? Length { get; set; }
    public string? End1 { get; set; }
    public string? End2 { get; set; }
    public string? Ident { get; set; }
}

public class MaterialInfo
{
    public string? PartNo { get; set; }
    public string? Description { get; set; }
    public string? MaterialType { get; set; }
    public string? Grade { get; set; }
    public string? Size { get; set; }
    public string? Schedule { get; set; }
    public string? Thickness { get; set; }
    public string? Specification { get; set; }
    /// <summary>
    /// Quantity of this BOM item (e.g. 3 elbows, 2 flanges). Defaults to 1
    /// when the quantity could not be parsed from the BOM description text.
    /// Pipe items may have length in MM instead of a discrete count.
    /// </summary>
    public int Quantity { get; set; } = 1;
}