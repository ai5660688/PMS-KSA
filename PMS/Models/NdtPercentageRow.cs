namespace PMS.Models;

/// <summary>
/// Flat row returned by the NDT Percentage joined query.
/// When Sheet / Weld filters are active the query joins DFR_tbl,
/// exposing Joint_ID, SheetNo, WeldNo and the editable Special_RT flag.
/// </summary>
public class NdtPercentageRow
{
    // ── Keys ──
    public int LineId { get; set; }
    public int? JointId { get; set; }

    // ── Read-only / display ──
    public string? LayoutNo { get; set; }
    public string? LineClass { get; set; }
    public string? Fluid { get; set; }
    public string? SheetNo { get; set; }
    public string? WeldNo { get; set; }

    // ── Editable LINE_LIST fields ──
    public string? MaterialDes { get; set; }
    public string? Material { get; set; }
    public string? DesignTemp { get; set; }
    public string? Category { get; set; }
    public double? RtShop { get; set; }
    public double? RtField { get; set; }
    public double? RtFieldShopSw { get; set; }
    public double? MtShop { get; set; }
    public double? MtField { get; set; }
    public double? PtShop { get; set; }
    public double? PtField { get; set; }
    public bool? PwhtYn { get; set; }
    public bool? Pwht20 { get; set; }
    public double? Ht { get; set; }
    public double? HtAfterPwht { get; set; }
    public bool? Pmi { get; set; }
    public string? DwgRemarks { get; set; }

    // ── Editable DFR field ──
    /// <summary>DFR_tbl.Special_RT — stored as float, displayed as Special Joint decimal input.</summary>
    public double? SpecialRt { get; set; }
}
