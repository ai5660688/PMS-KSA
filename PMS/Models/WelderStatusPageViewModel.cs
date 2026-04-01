using System.Collections.Generic;
using System.Data;

namespace PMS.Models;

public sealed class WelderStatusPageViewModel
{
    public string? Week { get; init; }
    public IReadOnlyList<LotOption> AvailableWeeks { get; init; } = [];
    public IReadOnlyList<int> SelectedProjectIds { get; init; } = [];
    public IReadOnlyList<ProjectOption> Projects { get; init; } = [];
    public DataTable Results { get; init; } = new();
}
