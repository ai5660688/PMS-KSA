using System;
using System.Collections.Generic;

namespace PMS.Models;

public class IpTSelectionViewModel
{
    public int SelectedProjectId { get; set; }
    public List<int> SelectedProjectIds { get; set; } = new();
    public int? SelectedLotId { get; set; }
    public string Location { get; set; } = "All";
    public List<ProjectOption> Projects { get; set; } = new();
    public List<LotOption> LotOptions { get; set; } = new();
    public List<string> WeldTypeOptions { get; set; } = new();
    public List<string> SelectedWeldTypes { get; set; } = new();
}

public class IpTSelectionRowDto
{
    public int JointId { get; set; }
    public string JAdd { get; set; } = string.Empty;
    public DateTime? ActualDateWelded { get; set; }
    public double? Diameter { get; set; }
    public string LineClass { get; set; } = string.Empty;
    public string WelderSymbol { get; set; } = string.Empty;
    public string? FinalRtResult { get; set; }
    public string? IpOrT { get; set; }
    public string? RepairWelder { get; set; }
    public string? RootA { get; set; }
    public string? RootB { get; set; }
    public string? FillA { get; set; }
    public string? FillB { get; set; }
    public string? CapA { get; set; }
    public string? CapB { get; set; }
    public string LotNo { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string WeldNumber { get; set; } = string.Empty;
    public string LayoutNumber { get; set; } = string.Empty;
    public string? Sheet { get; set; }
    public bool HasRepairCq { get; set; }
    public string? DwrRemarks { get; set; }
    public string TracerStatus { get; set; } = string.Empty;
    public DateTime? RequestedDate { get; set; }
}

public class IpTSelectionQuery
{
    public int ProjectId { get; set; }
    public List<int>? ProjectIds { get; set; }
    public int? LotId { get; set; }
    public string? Location { get; set; }
    public string? LineClass { get; set; }
    public string? Layout { get; set; }
    public string? Sheet { get; set; }
    public string? JointNo { get; set; }
    public string? RepairWelder { get; set; }
    public List<string>? WeldTypes { get; set; }
    public bool MatchLot { get; set; } = true;
    public bool MatchDate { get; set; } = true;
    public bool MatchDiameter { get; set; } = true;
    public bool MatchLineClass { get; set; } = true;
    public bool MatchWelder { get; set; } = true;
    public bool MatchProcess { get; set; } = true;
}

public class IpTDropdownRequest
{
    public int JointId { get; set; }
    public int ProjectId { get; set; }
    public string WelderSymbol { get; set; } = string.Empty;
    public List<string>? WeldTypes { get; set; }
    public bool MatchLot { get; set; } = true;
    public bool MatchDate { get; set; } = true;
    public bool MatchDiameter { get; set; } = true;
    public bool MatchLineClass { get; set; } = true;
    public bool MatchWelder { get; set; } = true;
    public bool MatchProcess { get; set; } = true;
}

public class IpTMatchedRowDto
{
    public int JointId { get; set; }
    public string JointNo { get; set; } = string.Empty;
    public string JAdd { get; set; } = string.Empty;
    public string LayoutNumber { get; set; } = string.Empty;
    public string? Sheet { get; set; }
    public DateTime? ActualDateWelded { get; set; }
    public double? Diameter { get; set; }
    public string LineClass { get; set; } = string.Empty;
    public string Root { get; set; } = string.Empty;
    public string FillCap { get; set; } = string.Empty;
    public string? CurrentIpOrT { get; set; }
    public string LotNo { get; set; } = string.Empty;
    public DateTime? RequestedDate { get; set; }
}

public class IpTUpdateRequest
{
    public int JointId { get; set; }
    public int ProjectId { get; set; }
    public string WelderSymbol { get; set; } = string.Empty;
    public string? IpOrTValue { get; set; }
    public List<string>? WeldTypes { get; set; }
    public bool MatchLot { get; set; } = true;
    public bool MatchDate { get; set; } = true;
    public bool MatchDiameter { get; set; } = true;
    public bool MatchLineClass { get; set; } = true;
    public bool MatchWelder { get; set; } = true;
    public bool MatchProcess { get; set; } = true;
}

public class IpTBatchUpdateRequest
{
    public int ProjectId { get; set; }
    public List<IpTRowAssignment> Assignments { get; set; } = new();
}

public class IpTRowAssignment
{
    public int JointId { get; set; }
    public string? IpOrTValue { get; set; }
}
