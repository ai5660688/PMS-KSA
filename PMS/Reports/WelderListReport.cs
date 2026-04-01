using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using PMS.Models;

namespace PMS.Reports;

/// <summary>
/// Welder List Report – tabular listing of all welders with status.
/// </summary>
public sealed class WelderListReport : PmsReportBase
{
    private readonly IReadOnlyList<Welder> _welders;

    public WelderListReport(IReadOnlyList<Welder> welders, string? projectName = null, string? logoPath = null)
        : base("Welder List Report", projectName, logoPath)
    {
        _welders = welders;
    }

    protected override void ComposeContent(IContainer container)
    {
        container.PaddingVertical(6).Column(col =>
        {
            col.Item().Text($"Total Welders: {_welders.Count}")
                .FontSize(9).Bold().FontColor(PrimaryColor);

            col.Item().PaddingTop(6);

            col.Item().Table(table =>
            {
                table.ColumnsDefinition(c =>
                {
                    c.ConstantColumn(30);  // #
                    c.ConstantColumn(80);  // Symbol
                    c.RelativeColumn(2);   // Name
                    c.ConstantColumn(65);  // Location
                    c.ConstantColumn(80);  // Mobile
                    c.RelativeColumn(2);   // Email
                    c.ConstantColumn(75);  // Mobilization
                    c.ConstantColumn(75);  // Demobilization
                    c.ConstantColumn(70);  // Status
                });

                table.Header(header =>
                {
                    foreach (var h in new[] {
                        "#", "Symbol", "Name", "Location", "Mobile",
                        "Email", "Mobilized", "Demobilized", "Status" })
                    {
                        header.Cell().BorderBottom(1).BorderColor(AccentColor)
                            .Background(HeaderBg).Padding(3)
                            .Text(h).FontSize(8).Bold().FontColor(PrimaryColor);
                    }
                });

                int rowIndex = 0;
                foreach (var w in _welders.OrderBy(w => w.Welder_Symbol))
                {
                    var bg = rowIndex % 2 == 1 ? RowAltBg : "#ffffff";
                    rowIndex++;

                    void Cell(string? val) =>
                        table.Cell().Background(bg).Padding(3)
                            .Text(val ?? "").FontSize(8);

                    Cell(rowIndex.ToString());
                    Cell(w.Welder_Symbol);
                    Cell(w.Name);
                    Cell(w.Welder_Location);
                    Cell(w.Mobile_No);
                    Cell(w.Email);
                    Cell(w.Mobilization_Date?.ToString("yyyy-MM-dd"));
                    Cell(w.Demobilization_Date?.ToString("yyyy-MM-dd"));
                    Cell(w.Status);
                }
            });
        });
    }
}
