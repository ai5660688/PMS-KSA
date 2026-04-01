using System.Collections.Generic;

namespace PMS.Models;

public sealed class TransmittalStatusGroup
{
    public string Category { get; set; } = string.Empty;
    public List<TransmittalStatusRow> Rows { get; set; } = [];
}

public sealed class TransmittalStatusRow
{
    public string Label { get; set; } = string.Empty;
    public int Value { get; set; }
    public string? Remarks { get; set; }
    public bool IsIndented { get; set; }
}
