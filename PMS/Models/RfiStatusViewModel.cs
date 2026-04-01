using System.Collections.Generic;

namespace PMS.Models;

public sealed class RfiStatusGroup
{
    public string Contractor { get; set; } = string.Empty;
    public string LocationType { get; set; } = string.Empty;
    public List<RfiStatusRow> Rows { get; set; } = [];
    public RfiStatusRow Total { get; set; } = new();
}

public sealed class RfiStatusRow
{
    public string Discipline { get; set; } = string.Empty;
    public int Raised { get; set; }
    public int Closed { get; set; }
    public int Cancelled { get; set; }
    public int UnderReviewInspDate { get; set; }
    public int UnderReviewLess72 { get; set; }
    public int UnderReviewMore72 { get; set; }
    public int UnderQcPidSign { get; set; }
    public int PmtSignBalance { get; set; }
    public int UnderScanningWithPmtSign { get; set; }
    public int TotalOpen { get; set; }
}
