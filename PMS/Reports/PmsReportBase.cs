using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace PMS.Reports;

/// <summary>
/// Base class for all PMS PDF reports. Provides consistent branding,
/// page layout (header / detail / footer), and PMS colour palette.
/// Subclasses only need to implement <see cref="ComposeContent"/>.
/// </summary>
public abstract class PmsReportBase : IDocument
{
    // ── PMS brand colours ──
    protected static readonly string PrimaryColor = "#176d8a";
    protected static readonly string AccentColor = "#1f99be";
    protected static readonly string HeaderBg = "#f0fbff";
    protected static readonly string RowAltBg = "#faf8f2";

    protected string ReportTitle { get; }
    protected string? ProjectName { get; }
    protected string? LogoPath { get; }

    protected PmsReportBase(string reportTitle, string? projectName = null, string? logoPath = null)
    {
        ReportTitle = reportTitle;
        ProjectName = projectName;
        LogoPath = logoPath;
    }

    public void Compose(IDocumentContainer container)
    {
        container.Page(page =>
        {
            page.Size(PageSizes.A4.Landscape());
            page.MarginVertical(20);
            page.MarginHorizontal(25);
            page.DefaultTextStyle(x => x.FontSize(9).FontFamily("Calibri"));

            page.Header().Element(ComposeHeader);
            page.Content().Element(ComposeContent);
            page.Footer().Element(ComposeFooter);
        });
    }

    // ── Header ──
    private void ComposeHeader(IContainer container)
    {
        container.Column(col =>
        {
            col.Item().Row(row =>
            {
                if (!string.IsNullOrEmpty(LogoPath) && File.Exists(LogoPath))
                {
                    row.ConstantItem(60).Height(30).Image(LogoPath, ImageScaling.FitArea);
                    row.ConstantItem(10); // spacer
                }

                row.RelativeItem().Column(inner =>
                {
                    inner.Item().Text(ReportTitle)
                        .FontSize(16).Bold().FontColor(PrimaryColor);

                    if (!string.IsNullOrWhiteSpace(ProjectName))
                        inner.Item().Text(ProjectName)
                            .FontSize(10).FontColor(AccentColor);
                });

                row.ConstantItem(140).AlignRight().Column(inner =>
                {
                    inner.Item().Text($"Date: {DateTime.Now:yyyy-MM-dd}")
                        .FontSize(8).FontColor(PrimaryColor);
                    inner.Item().Text($"Time: {DateTime.Now:HH:mm}")
                        .FontSize(8).FontColor(PrimaryColor);
                });
            });

            col.Item().PaddingVertical(4)
                .LineHorizontal(1).LineColor(AccentColor);
        });
    }

    // ── Content – implemented by each report ──
    protected abstract void ComposeContent(IContainer container);

    // ── Footer ──
    private void ComposeFooter(IContainer container)
    {
        container.Row(row =>
        {
            row.RelativeItem().AlignLeft()
                .Text("Piping Management System (PMS)")
                .FontSize(7).FontColor(PrimaryColor);

            row.RelativeItem().AlignCenter()
                .Text(t =>
                {
                    t.Span("Page ").FontSize(7).FontColor(PrimaryColor);
                    t.CurrentPageNumber().FontSize(7).FontColor(PrimaryColor);
                    t.Span(" of ").FontSize(7).FontColor(PrimaryColor);
                    t.TotalPages().FontSize(7).FontColor(PrimaryColor);
                });

            row.RelativeItem().AlignRight()
                .Text($"© {DateTime.Now.Year} PMS")
                .FontSize(7).FontColor(PrimaryColor);
        });
    }
}
