using System;
using System.Collections.Generic;

namespace PMS.Models;

public class HtSelectionViewModel
{
    public int SelectedProjectId { get; set; }
    public List<int> SelectedProjectIds { get; set; } = new();
    public int? SelectedLotId { get; set; }
    public string Location { get; set; } = "All";
    public string LotCategory { get; set; } = "Welding";
    public string HtMode { get; set; } = "WithoutPWHT";
    public List<ProjectOption> Projects { get; set; } = new();
    public List<string> WeldTypeOptions { get; set; } = new();
    public List<string> SelectedWeldTypes { get; set; } = new();
    public List<LotOption> LotOptions { get; set; } = new();
}

public class HtSelectionRowDto
{
    public int JointId { get; set; }
    public string LineClass { get; set; } = string.Empty;
    public string LayoutNumber { get; set; } = string.Empty;
    public string Sheet { get; set; } = string.Empty;
    public string SpoolNumber { get; set; } = string.Empty;
    public string JointNo { get; set; } = string.Empty;
    public double? Diameter { get; set; }
    public double? ActualThick { get; set; }
    public string Root { get; set; } = string.Empty;
    public string FillCap { get; set; } = string.Empty;
    public bool HtSelection { get; set; }
    public DateTime? HtSelectionDate { get; set; }
    public string LotNo { get; set; } = string.Empty;
    public DateTime? LotFrom { get; set; }
    public DateTime? LotTo { get; set; }
    public string Location { get; set; } = string.Empty;
    public string WeldType { get; set; } = string.Empty;
    public double HtFraction { get; set; }
}

public class HtSelectionStatDto
{
    public string LineClass { get; set; } = string.Empty;
    public string Root { get; set; } = string.Empty;
    public string FillCap { get; set; } = string.Empty;
    public int GroupCount { get; set; }
    public int SelectedCount { get; set; }
}
