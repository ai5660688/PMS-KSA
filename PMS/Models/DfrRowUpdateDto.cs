using System;

namespace PMS.Models;

public class DfrRowUpdateDto
{
    public int JointId { get; set; }
    public string? JointNo { get; set; } // new full joint number composed client-side
    public string? Location { get; set; } // new: LOCATION
    public string? LayoutNumber { get; set; } // new: LAYOUT_NUMBER
    public string? WeldNumber { get; set; } // new: WELD_NUMBER (editable Joint No segment)
    public string? JAdd { get; set; } // new: J_Add (repair / change code)
    public string? WeldType { get; set; }
    public string? Rev { get; set; }
    public string? SpoolNo { get; set; }
    public double? DiaIn { get; set; }
    public string? Sch { get; set; }
    public string? MaterialA { get; set; }
    public string? MaterialB { get; set; }
    public string? GradeA { get; set; }
    public string? GradeB { get; set; }
    public string? HeatNumberA { get; set; }
    public string? HeatNumberB { get; set; }
    public int? WpsId { get; set; }
    public string? Wps { get; set; } // added plain WPS string; server will resolve to ID if possible
    public DateTime? FitupDate { get; set; }
    public string? FitupReport { get; set; }
    public string? TackWelder { get; set; }
    public double? OlDia { get; set; }
    public string? OlSch { get; set; }
    public double? OlThick { get; set; }
    public string? HoldDfr { get; set; } // HOLD_DFR text
    public DateTime? HoldDfrDate { get; set; } // HOLD_DFR_Date_D
    public string? Remarks { get; set; }
    public bool? Deleted { get; set; }
    public bool? Cancelled { get; set; }
    public bool? FitupConfirmed { get; set; }
    public string? Sheet { get; set; } // NEW: editable sheet value
    public int? RfiId { get; set; } // NEW: selected RFI_tbl.RFI_ID to save
    // New: user confirmation flag for supersede transfer (when FITUP_DATE > SP_Date)
    public bool? ConfirmSupersede { get; set; }

    // NOTE: DWR-specific fields moved to DwrRowUpdateDto.
}
