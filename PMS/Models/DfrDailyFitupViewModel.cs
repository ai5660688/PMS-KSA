using System;
using System.Collections.Generic;

namespace PMS.Models;

public class DfrDailyFitupViewModel
{
    public int SelectedProjectId { get; set; }
    public string? SelectedLocation { get; set; }
    public string? SearchLayout { get; set; }
    public string? SearchSheet { get; set; }
    public string? SearchFitupDate { get; set; }
    public string? SearchFitupReport { get; set; }
    public string? HeaderView { get; set; }

    public List<ProjectOption> Projects { get; set; } = new();
    public List<string> Locations { get; set; } = new() { "All", "Shop", "Field" };

    public List<string> LocationOptions { get; set; } = new();
    public List<string> JAddOptions { get; set; } = new();
    public List<string> WeldTypeOptions { get; set; } = new();
    public List<string> IpOrTOptions { get; set; } = new();
    public string? DefaultWeldType { get; set; }
    public List<string> DefaultWeldTypes { get; set; } = new();
    public List<string> LayoutOptions { get; set; } = new();
    public List<string> SheetOptions { get; set; } = new();

    public string? LineClass { get; set; }
    public string? LineMaterial { get; set; }
    public string? LineCategory { get; set; }
    public double? RtShop { get; set; }
    public double? RtField { get; set; }
    public bool? PwhtYN { get; set; }
    public bool? Pwht20mm { get; set; }
    public string? LsRev { get; set; }

    public string? FitupReportHeader { get; set; }
    public string? FitupReportShop { get; set; }
    public string? FitupReportField { get; set; }
    public string? FitupReportThreaded { get; set; }
    public string? FitupDateHeader { get; set; }
    public int? SelectedRfiId { get; set; }
    public int? SelectedWpsId { get; set; }
    public string? SelectedTacker { get; set; }
    // New separate header tacker selections for welding positions
    public string? SelectedTackerRootA { get; set; }
    public string? SelectedTackerRootB { get; set; }
    public string? SelectedTackerFillA { get; set; }
    public string? SelectedTackerFillB { get; set; }
    public string? SelectedTackerCapA { get; set; }
    public string? SelectedTackerCapB { get; set; }

    public List<RfiOption> RfiOptions { get; set; } = new();
    public List<WpsOption> WpsOptions { get; set; } = new();
    public List<WelderSymbolOption> TackerOptions { get; set; } = new();

    public List<string> MaterialDescriptions { get; set; } = new();
    public List<string> MaterialGrades { get; set; } = new();

    public List<DfrRowVm> Rows { get; set; } = new();

    public int RowLimit { get; set; }
    public int RowsReturned { get; set; }
    public bool IsTruncated { get; set; }
}

public class ProjectOption
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string? LineSheet { get; set; }
    public int? WeldersProjectId { get; set; }
}

public class RfiOption
{
    public int Id { get; set; }
    public string Display { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

public class WpsOption
{
    public int Id { get; set; }
    public string Wps { get; set; } = string.Empty;
    public string? ThicknessRange { get; set; }
    public bool? Pwht { get; set; }
}

public class DfrRowVm
{
    public int JointId { get; set; }
    public string? JointNo { get; set; }
    public string? WeldNumber { get; set; }
    public string? Location { get; set; }
    public string? LayoutNumber { get; set; }
    public string? JAdd { get; set; }
    public string? WeldType { get; set; }
    public string? Sheet { get; set; }
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
    public string? Wps { get; set; }
    public List<WpsOption> WpsCandidates { get; set; } = new();
    public string? FitupDate { get; set; }
    // DWR-provided actual weld date (ACTUAL_DATE_WELDED)
    public string? ActualDate { get; set; }
    public string? FitupReport { get; set; }
    public int? RfiId { get; set; } // NEW: selected RFI id
    public string? RfiNo { get; set; } // NEW: selected RFI No display
    public int? RFI_ID_DWR { get; set; }
    public int? WPS_ID_DWR { get; set; }
    public DateTime? DATE_WELDED { get; set; }
    public DateTime? ACTUAL_DATE_WELDED { get; set; }
    public string? TackWelder { get; set; }
    public double? OlDia { get; set; }
    public string? OlSch { get; set; }
    public double? OlThick { get; set; }
    public bool Deleted { get; set; }
    public bool Cancelled { get; set; }
    public bool FitupConfirmed { get; set; }
    // NEW: DWR weld confirmed state to drive DailyWelding UI
    public bool Weld_Confirmed { get; set; }
    public string? HoldDfr { get; set; }
    public string? HoldDfrDate { get; set; }
    public string? Remarks { get; set; }
    public string? UpdatedBy { get; set; }
    public string? UpdatedDate { get; set; }
    // DWR-supplied Open/Closed status and IP/T indicator
    public string? OpenClosed { get; set; }
    public string? IPOrT { get; set; }

    // Hold flags for UI highlighting
    public bool DfrOnHold { get; set; }
    public bool SpOnHold { get; set; }
    public bool SheetOnHold { get; set; }

    // Per-row metadata to support WPS candidates in Date/Report views
    public string? LineClass { get; set; }

    // DWR PREHEAT value (nullable) - added so DwrForm reflection can discover PREHEAT_TEMP_C on the row VM
    public double? PREHEAT_TEMP_C { get; set; }

    // --- Added: DWR welding crew and QR fields so DailyWelding can display saved values ---
    // ROOT/FILL/CAP from `DWR_tbl`
    public string? ROOT_A { get; set; }
    public string? ROOT_B { get; set; }
    public string? FILL_A { get; set; }
    public string? FILL_B { get; set; }
    public string? CAP_A { get; set; }
    public string? CAP_B { get; set; }
    // DWR QR number (POST_VISUAL_INSPECTION_QR_NO)
    public string? POST_VISUAL_INSPECTION_QR_NO { get; set; }
}