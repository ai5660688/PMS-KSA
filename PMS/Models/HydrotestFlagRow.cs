namespace PMS.Models;

/// <summary>
/// Flat row returned by the Hydrotest Flag joined query
/// (Projects → Line_Sheet → LINE_LIST ←LEFT JOIN→ DFR).
/// </summary>
public class HydrotestFlagRow
{
    /// <summary>LINE_LIST_tbl.Line_ID – used as the key when saving edits.</summary>
    public int LineId { get; set; }

    /// <summary>DFR_tbl.Joint_ID – null when no DFR row exists (LEFT JOIN).</summary>
    public int? JointId { get; set; }

    /// <summary>Project_ID + ' - ' + Project_Name.</summary>
    public string ProjectNo { get; set; } = string.Empty;

    public string? LayoutNo { get; set; }

    /// <summary>Line_Sheet_tbl.LS_SHEET.</summary>
    public string? SheetNo { get; set; }

    /// <summary>Computed: LOCATION-WELD_NUMBER [J_Add].</summary>
    public string? WeldNo { get; set; }

    // ── Editable fields (all live in LINE_LIST_tbl) ──
    public string? System { get; set; }
    public string? SubSystem { get; set; }
    public string? TestPackageNo { get; set; }
    public string? TestType { get; set; }
    public string? SpoolNo { get; set; }
    public string? PId { get; set; }
    public string? TestPackageNoWs { get; set; }
}
