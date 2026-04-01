using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using PMS.Models;

namespace PMS.Reports;

/// <summary>
/// Daily Fit-up Report – grouped by Layout Number.
/// MS-Access-style banded report: group header → detail rows → group footer totals.
/// </summary>
public sealed class DfrReport : PmsReportBase
{
    private readonly IReadOnlyList<Dfr> _rows;

    public DfrReport(IReadOnlyList<Dfr> rows, string? projectName = null, string? logoPath = null)
        : base("Daily Fit-up Report (DFR)", projectName, logoPath)
    {
        _rows = rows;
    }

    protected override void ComposeContent(IContainer container)
    {
        container.PaddingVertical(6).Column(col =>
        {
            // Summary line
            col.Item().Text($"Total Joints: {_rows.Count}")
                .FontSize(9).Bold().FontColor(PrimaryColor);

            col.Item().PaddingTop(6);

            // Group by Layout Number
            var groups = _rows
                .OrderBy(r => r.LAYOUT_NUMBER)
                .ThenBy(r => r.WELD_NUMBER)
                .GroupBy(r => r.LAYOUT_NUMBER ?? "(No Layout)");

            foreach (var group in groups)
            {
                // ── Group header (like MS Access Group Header band) ──
                col.Item().PaddingTop(8).Background(HeaderBg)
                    .Padding(4).Row(row =>
                    {
                        row.RelativeItem().Text($"Layout: {group.Key}")
                            .FontSize(10).Bold().FontColor(PrimaryColor);
                        row.ConstantItem(120).AlignRight()
                            .Text($"Joints: {group.Count()}")
                            .FontSize(9).FontColor(AccentColor);
                    });

                // ── Detail table ──
                col.Item().Table(table =>
                {
                    table.ColumnsDefinition(c =>
                    {
                        c.ConstantColumn(55);  // Weld No
                        c.ConstantColumn(65);  // Spool No
                        c.ConstantColumn(40);  // Type
                        c.ConstantColumn(35);  // Sheet
                        c.ConstantColumn(45);  // Dia
                        c.ConstantColumn(45);  // Sch
                        c.RelativeColumn();     // Material A
                        c.RelativeColumn();     // Grade A
                        c.ConstantColumn(70);  // Heat No A
                        c.ConstantColumn(60);  // Tack Welder
                        c.ConstantColumn(65);  // Fitup Date
                        c.ConstantColumn(55);  // QR No
                    });

                    // Column headers
                    table.Header(header =>
                    {
                        foreach (var h in new[] {
                            "Weld No", "Spool No", "Type", "Sheet", "Dia",
                            "Sch", "Material A", "Grade A", "Heat No A",
                            "Tack Wldr", "Fitup Date", "QR No" })
                        {
                            header.Cell().BorderBottom(1).BorderColor(AccentColor)
                                .Padding(3).Text(h)
                                .FontSize(8).Bold().FontColor(PrimaryColor);
                        }
                    });

                    // Data rows
                    int rowIndex = 0;
                    foreach (var r in group)
                    {
                        var bg = rowIndex++ % 2 == 1 ? RowAltBg : "#ffffff";

                        void Cell(string? val) =>
                            table.Cell().Background(bg).Padding(3)
                                .Text(val ?? "").FontSize(8);

                        Cell(r.WELD_NUMBER);
                        Cell(r.SPOOL_NUMBER);
                        Cell(r.WELD_TYPE);
                        Cell(r.SHEET);
                        Cell(r.DIAMETER?.ToString("F1"));
                        Cell(r.SCHEDULE);
                        Cell(r.MATERIAL_A);
                        Cell(r.GRADE_A);
                        Cell(r.HEAT_NUMBER_A);
                        Cell(r.TACK_WELDER);
                        Cell(r.FITUP_DATE?.ToString("yyyy-MM-dd"));
                        Cell(r.FITUP_INSPECTION_QR_NUMBER);
                    }
                });
            }
        });
    }
}
