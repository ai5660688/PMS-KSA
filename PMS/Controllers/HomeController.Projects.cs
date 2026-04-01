using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    // GET: /Home/ProjectLog (Admin only)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ProjectLog()
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        var projects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .ToListAsync();

        return View(projects);
    }

    // POST: /Home/SaveProject
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveProject(Project input)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            var existing = await _context.Projects_tbl.FirstOrDefaultAsync(p => p.Project_ID == input.Project_ID);

            if (existing == null)
            {
                if (input.Default_P)
                {
                    var otherDefaults = await _context.Projects_tbl
                        .Where(p => p.Default_P)
                        .ToListAsync();
                    foreach (var other in otherDefaults)
                        other.Default_P = false;
                }

                input.PR_Updated_Date = AppClock.Now;
                input.PR_Updated_By = userId;
                _context.Projects_tbl.Add(input);
                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Added project {input.Project_ID}.";
            }
            else
            {
                existing.Project_Name = input.Project_Name;
                existing.Client = input.Client;
                existing.Contractor_Project_No = input.Contractor_Project_No;
                existing.Contractor = input.Contractor;
                existing.PO_No = input.PO_No;
                existing.WS_Location = input.WS_Location;
                existing.FW_Location = input.FW_Location;
                existing.BI_JO = input.BI_JO;
                existing.Contract_No = input.Contract_No;
                existing.Welders_Project_ID = input.Welders_Project_ID;
                existing.Line_Sheet = input.Line_Sheet;
                existing.Project_Type = input.Project_Type;
                existing.ARAMCO_NON = input.ARAMCO_NON;
                existing.Default_P = input.Default_P;

                if (existing.Default_P)
                {
                    var otherDefaults = await _context.Projects_tbl
                        .Where(p => p.Project_ID != existing.Project_ID && p.Default_P)
                        .ToListAsync();
                    foreach (var other in otherDefaults)
                        other.Default_P = false;
                }

                existing.PR_Updated_Date = AppClock.Now;
                existing.PR_Updated_By = userId;

                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Saved project {existing.Project_ID}.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveProject failed for {Id}", input.Project_ID);
            TempData["Msg"] = "Failed to save project.";
        }
        return RedirectToAction(nameof(ProjectLog));
    }

    // POST: /Home/DeleteProject
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteProject([FromForm] int id)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        try
        {
            var proj = await _context.Projects_tbl.FirstOrDefaultAsync(p => p.Project_ID == id);
            if (proj == null) return NotFound();

            _context.Projects_tbl.Remove(proj);
            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Deleted project {id}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteProject failed for {Id}", id);
            TempData["Msg"] = "Failed to delete project.";
        }
        return RedirectToAction(nameof(ProjectLog));
    }

    // GET: /Home/ExportProjects?q=search
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportProjects([FromQuery] string? q)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        var qry = _context.Projects_tbl.AsNoTracking();
        if (!string.IsNullOrWhiteSpace(q))
        {
            var term = q.Trim();
            if (int.TryParse(term, out var idVal))
            {
                qry = qry.Where(p => p.Project_ID == idVal || (p.Project_Name != null && EF.Functions.Like(p.Project_Name, $"%{term}%")));
            }
            else
            {
                qry = qry.Where(p => p.Project_Name != null && EF.Functions.Like(p.Project_Name, $"%{term}%"));
            }
        }

        var rows = await qry
            .OrderBy(p => p.Project_ID)
            .Select(p => new
            {
                p.Project_ID,
                p.Project_Name,
                p.Contractor_Project_No,
                p.Client,
                p.Contractor,
                p.PO_No,
                p.WS_Location,
                p.FW_Location,
                p.BI_JO,
                p.Contract_No,
                p.Welders_Project_ID,
                p.Line_Sheet,
                p.Project_Type,
                p.ARAMCO_NON,
                p.Default_P
            })
            .ToListAsync();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Project Log");

        string[] headers = { "ID", "Project Name", "Contractor Project", "Client", "Contractor", "PO No", "WS Loc", "FW Loc", "BI/JO", "Contract No", "Welders Project ID", "Line / Sheet", "Project Type", "ARAMCO/NON", "Default" };
        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        int r = 2;
        foreach (var p in rows)
        {
            ws.Cell(r, 1).Value = p.Project_ID;
            ws.Cell(r, 2).Value = p.Project_Name ?? string.Empty;
            ws.Cell(r, 3).Value = p.Contractor_Project_No ?? string.Empty;
            ws.Cell(r, 4).Value = p.Client ?? string.Empty;
            ws.Cell(r, 5).Value = p.Contractor ?? string.Empty;
            ws.Cell(r, 6).Value = p.PO_No ?? string.Empty;
            ws.Cell(r, 7).Value = p.WS_Location ?? string.Empty;
            ws.Cell(r, 8).Value = p.FW_Location ?? string.Empty;
            ws.Cell(r, 9).Value = p.BI_JO ?? string.Empty;
            ws.Cell(r, 10).Value = p.Contract_No ?? string.Empty;
            if (p.Welders_Project_ID.HasValue)
                ws.Cell(r, 11).SetValue(p.Welders_Project_ID.Value);
            else
                ws.Cell(r, 11).SetValue(string.Empty);
            ws.Cell(r, 12).Value = p.Line_Sheet ?? string.Empty;
            ws.Cell(r, 13).Value = p.Project_Type ?? string.Empty;
            ws.Cell(r, 14).Value = p.ARAMCO_NON ?? string.Empty;
            ws.Cell(r, 15).Value = p.Default_P ? "Yes" : "No";
            ws.Row(r).Height = 17;
            r++;
        }

        int lastRow = r - 1;
        var fullRange = ws.Range(1, 1, lastRow, headers.Length);
        var table = fullRange.CreateTable();
        table.Theme = XLTableTheme.TableStyleMedium2;
        table.ShowTotalsRow = false;

        ws.Row(1).Height = 30;
        ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"ProjectLog_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}
