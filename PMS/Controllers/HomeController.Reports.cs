using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QuestPDF.Fluent;
using PMS.Reports;

namespace PMS.Controllers;

public partial class HomeController
{
    // ── DFR Report ──
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DfrReportPdf(int? projectNo)
    {
        var projectId = projectNo ?? HttpContext.Session.GetInt32("ProjectID");

        var query = _context.DFR_tbl.Where(d => !d.Deleted && !d.Cancelled);
        if (projectId.HasValue)
            query = query.Where(d => d.Project_No == projectId.Value);

        var rows = await query.OrderBy(d => d.LAYOUT_NUMBER)
            .ThenBy(d => d.WELD_NUMBER)
            .ToListAsync();

        var project = projectId.HasValue
            ? await _context.Projects_tbl.FindAsync(projectId.Value)
            : null;

        var logoPath = Path.Combine(_env.WebRootPath, "img", "PMS_logo.png");

        var report = new DfrReport(rows, project?.Project_Name, logoPath);
        var pdf = report.GeneratePdf();

        return File(pdf, "application/pdf", $"DFR_Report_{DateTime.Now:yyyyMMdd}.pdf");
    }

    // ── DWR Report ──
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DwrReportPdf(int? projectNo)
    {
        var projectId = projectNo ?? HttpContext.Session.GetInt32("ProjectID");

        var query = from dwr in _context.DWR_tbl
                    join dfr in _context.DFR_tbl on dwr.Joint_ID_DWR equals dfr.Joint_ID
                    where !dfr.Deleted && !dfr.Cancelled
                    select new { dwr, dfr };

        if (projectId.HasValue)
            query = query.Where(x => x.dfr.Project_No == projectId.Value);

        var data = await query.OrderByDescending(x => x.dwr.DATE_WELDED)
            .ThenBy(x => x.dfr.LAYOUT_NUMBER)
            .ToListAsync();

        var rows = data.Select(x => new DwrReportRow
        {
            LayoutNumber = x.dfr.LAYOUT_NUMBER,
            WeldNumber = x.dfr.WELD_NUMBER,
            SpoolNumber = x.dfr.SPOOL_NUMBER,
            DateWelded = x.dwr.DATE_WELDED,
            RootA = x.dwr.ROOT_A,
            RootB = x.dwr.ROOT_B,
            FillA = x.dwr.FILL_A,
            CapA = x.dwr.CAP_A,
            PreheatTempC = x.dwr.PREHEAT_TEMP_C,
            VisualQrNo = x.dwr.POST_VISUAL_INSPECTION_QR_NO,
            Remarks = x.dwr.DWR_REMARKS
        }).ToList();

        var project = projectId.HasValue
            ? await _context.Projects_tbl.FindAsync(projectId.Value)
            : null;

        var logoPath = Path.Combine(_env.WebRootPath, "img", "PMS_logo.png");

        var report = new DwrReport(rows, project?.Project_Name, logoPath);
        var pdf = report.GeneratePdf();

        return File(pdf, "application/pdf", $"DWR_Report_{DateTime.Now:yyyyMMdd}.pdf");
    }

    // ── Welder List Report ──
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> WelderListReportPdf()
    {
        var welders = await _context.Welders_tbl
            .OrderBy(w => w.Welder_Symbol)
            .ToListAsync();

        var logoPath = Path.Combine(_env.WebRootPath, "img", "PMS_logo.png");

        var report = new WelderListReport(welders, null, logoPath);
        var pdf = report.GeneratePdf();

        return File(pdf, "application/pdf", $"Welder_List_{DateTime.Now:yyyyMMdd}.pdf");
    }
}
