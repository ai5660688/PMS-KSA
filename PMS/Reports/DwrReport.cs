using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using PMS.Models;

namespace PMS.Reports;

/// <summary>
/// Daily Welding Report – joins DWR data with DFR layout info.
/// Grouped by welding date with summary counts.
/// </summary>
public sealed class DwrReport : PmsReportBase
{
    private readonly IReadOnlyList<DwrReportRow> _rows;

    public DwrReport(IReadOnlyList<DwrReportRow> rows, string? projectName = null, string? logoPath = null)
        : base("Daily Welding Report (DWR)", projectName, logoPath)
    {
        _rows = rows;
    }

    protected override void ComposeContent(IContainer container)
    {
        container.PaddingVertical(6).Column(col =>
        {
            col.Item().Text($"Total Welds: {_rows.Count}")
                .FontSize(9).Bold().FontColor(PrimaryColor);

            col.Item().PaddingTop(6);

            var groups = _rows
                .OrderByDescending(r => r.DateWelded)
                .GroupBy(r => r.DateWelded?.ToString("yyyy-MM-dd") ?? "(No Date)");

            foreach (var group in groups)
            {
                // Group header
                col.Item().PaddingTop(8).Background(HeaderBg)
                    .Padding(4).Row(row =>
                    {
                        row.RelativeItem().Text($"Date Welded: {group.Key}")
                            .FontSize(10).Bold().FontColor(PrimaryColor);
                        row.ConstantItem(120).AlignRight()
                            .Text($"Welds: {group.Count()}")
                            .FontSize(9).FontColor(AccentColor);
                    });

                // Detail table
                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(c =>
                    {
                        c.ConstantColumn(70);  // Layout No
                        c.ConstantColumn(55);  // Weld No
                        c.ConstantColumn(65);  // Spool No
                        c.ConstantColumn(55);  // Root A
                        c.ConstantColumn(55);  // Root B
                        c.ConstantColumn(55);  // Fill A
                        c.ConstantColumn(55);  // Cap A
                        c.ConstantColumn(50);  // Preheat
                        c.ConstantColumn(55);  // Visual QR
                        c.RelativeColumn();     // Remarks
                    });

                    table.Header(header =>
                    {
                        foreach (var h in new[] {
                            "Layout No", "Weld No", "Spool No", "Root A", "Root B",
                            "Fill A", "Cap A", "Preheat °C", "Visual QR", "Remarks" })
                        {
                            header.Cell().BorderBottom(1).BorderColor(AccentColor)
                                .Padding(3).Text(h)
                                .FontSize(8).Bold().FontColor(PrimaryColor);
                        }
                    });

                    int rowIndex = 0;
                    foreach (var r in group)
                    {
                        var bg = rowIndex++ % 2 == 1 ? RowAltBg : "#ffffff";

                        void Cell(string? val) =>
                            table.Cell().Background(bg).Padding(3)
                                .Text(val ?? "").FontSize(8);

                        Cell(r.LayoutNumber);
                        Cell(r.WeldNumber);
                        Cell(r.SpoolNumber);
                        Cell(r.RootA);
                        Cell(r.RootB);
                        Cell(r.FillA);
                        Cell(r.CapA);
                        Cell(r.PreheatTempC?.ToString("F0"));
                        Cell(r.VisualQrNo);
                        Cell(r.Remarks);
                    }
                });
            }
        });
    }
}

/// <summary>
/// Flattened row combining DWR + DFR fields for the report.
/// </summary>
public class DwrReportRow
{
    public string? LayoutNumber { get; set; }
    public string? WeldNumber { get; set; }
    public string? SpoolNumber { get; set; }
    public DateTime? DateWelded { get; set; }
    public string? RootA { get; set; }
    public string? RootB { get; set; }
    public string? FillA { get; set; }
    public string? CapA { get; set; }
    public double? PreheatTempC { get; set; }
    public string? VisualQrNo { get; set; }
    public string? Remarks { get; set; }
}
