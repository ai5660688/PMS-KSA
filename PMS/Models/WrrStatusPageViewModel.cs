using System.Collections.Generic;

namespace PMS.Models;

public sealed class WrrStatusPageViewModel
{
    public IReadOnlyList<ProjectOption> Projects { get; init; } = [];
    public IReadOnlyList<int> SelectedProjectIds { get; init; } = [];
    public double ProjectAcceptableRate { get; init; } = 5.0;
    public double LinearAcceptableRate { get; init; } = 0.2;
    public string AsOfDate { get; init; } = string.Empty;
    public IReadOnlyList<WrrJointRow> JointRows { get; init; } = [];
    public IReadOnlyList<WrrLinearRow> LinearRows { get; init; } = [];
}

public sealed class WrrJointRow
{
    public string Week { get; init; } = string.Empty;
    public string From { get; init; } = string.Empty;
    public string To { get; init; } = string.Empty;
    public int RtDone { get; init; }
    public int RtReject { get; init; }
    public double WeeklyWrr { get; init; }
    public double CummWrr { get; init; }
    public double AcceptableRate { get; init; }
}

public sealed class WrrLinearRow
{
    public string Week { get; init; } = string.Empty;
    public string From { get; init; } = string.Empty;
    public string To { get; init; } = string.Empty;
    public double WeldLengthRtd { get; init; }
    public double LengthWeldDefect { get; init; }
    public double WeeklyWrr { get; init; }
    public double CummWrr { get; init; }
}
