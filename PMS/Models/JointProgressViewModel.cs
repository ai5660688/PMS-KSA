using System;
using System.Collections.Generic;

namespace PMS.Models;

public class JointProgressViewModel
{
    public int SelectedProjectId { get; set; }
    public List<int> SelectedProjectIds { get; set; } = new();
    public DateTime? StartDate { get; set; }
    public DateTime? EndDate { get; set; }
    public string Location { get; set; } = "All";
    public List<ProjectOption> Projects { get; set; } = new();
    public List<string> WeldTypeOptions { get; set; } = new();
    public List<string> SelectedWeldTypes { get; set; } = new();
    public string DateBasis { get; set; } = "Welding";
    public string Grouping { get; set; } = "Monthly";
}

public class JointProgressRowDto
{
    public DateTime WeldDate { get; set; }
    public string Location { get; set; } = string.Empty;
    public double WeldedDiaIn { get; set; }
    public double ThreadedDiaIn { get; set; }
    public double PlannedDiaIn { get; set; }
    public double TotalDiaIn => WeldedDiaIn + ThreadedDiaIn;
    public double CumulativeDiaIn { get; set; }
}

public class JointProgressChartPoint
{
    public string Date { get; set; } = string.Empty;
    public double PlannedDailyTotal { get; set; }
    public double DailyTotal { get; set; }
    public double CumulativeShop { get; set; }
    public double CumulativeField { get; set; }
    public double CumulativeTotal { get; set; }
}

public class JointProgressResponse
{
    public List<JointProgressRowDto> Rows { get; set; } = new();
    public List<JointProgressChartPoint> Chart { get; set; } = new();
    public double TotalWelded { get; set; }
    public double TotalThreaded { get; set; }
    public double TotalPlanned { get; set; }
    public double Total { get; set; }
}
