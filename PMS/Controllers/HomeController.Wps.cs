using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    private static string NormalizeWeldProcessLabel(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var trimmed = value.Trim();
        return trimmed.ToUpperInvariant() switch
        {
            "SAW" => "GTAW + SAW",
            "SMAW" => "GTAW + SMAW",
            "FCAW-GS" => "GTAW + FCAW-GS",
            _ => trimmed
        };
    }

    // GET: /Home/WpsLog (Admin only)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> WpsLog()
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        var list = await _context.WPS_tbl
            .AsNoTracking()
            .OrderBy(w => w.WPS_ID)
            .ToListAsync();

        var projects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_Name)
            .Select(p => new { p.Project_ID, p.Project_Name, p.Client })
            .ToListAsync();

        var defaultProjectId = await GetDefaultProjectIdAsync();
        var fallbackProjectId = projects.Count == 0 ? (int?)null : projects.Max(p => p.Project_ID);
        var selectedProjectId = defaultProjectId ?? fallbackProjectId;

        ViewBag.Projects = projects
            .Select(p =>
            {
                var name = (p.Project_Name ?? string.Empty).Trim();
                var hasName = !string.IsNullOrWhiteSpace(name);
                return new SelectListItem
                {
                    Value = p.Project_ID.ToString(),
                    Text = hasName ? $"{p.Project_ID} - {name}" : p.Project_ID.ToString(),
                    Selected = (selectedProjectId.HasValue && p.Project_ID == selectedProjectId.Value)
                };
            })
            .ToList();

        ViewBag.DefaultProjectId = selectedProjectId;
        return View(list);
    }

    // POST: /Home/SaveWps
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveWps(Wps input)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            var normalizedWeldProcess = NormalizeWeldProcessLabel(input.Weld_Process);
            var existing = await _context.WPS_tbl.FirstOrDefaultAsync(x => x.WPS_ID == input.WPS_ID);

            if (existing == null)
            {
                input.Weld_Process = normalizedWeldProcess;
                input.WPS_Updated_Date = AppClock.Now;
                input.WPS_Updated_By = userId;
                _context.WPS_tbl.Add(input);
                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Added WPS '{input.WPS}'.";
            }
            else
            {
                existing.WPS = input.WPS;
                existing.Project_WPS = input.Project_WPS;
                existing.Weld_Process = normalizedWeldProcess;
                existing.PWHT = input.PWHT;
                existing.Dia_Range = input.Dia_Range;
                existing.Thickness_Range = input.Thickness_Range;
                existing.Thickness_Range_From = input.Thickness_Range_From;
                existing.Thickness_Range_To = input.Thickness_Range_To;
                existing.WPS_P_NO = input.WPS_P_NO;
                existing.WPS_Material = input.WPS_Material;
                existing.Electrode = input.Electrode;
                existing.WPS_Pipe_Class = input.WPS_Pipe_Class;
                existing.WPS_Service = input.WPS_Service;

                existing.WPS_Updated_Date = AppClock.Now;
                existing.WPS_Updated_By = userId;

                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Updated WPS '{existing.WPS}'.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveWps failed for {Id}", input.WPS_ID);
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            if (!string.IsNullOrWhiteSpace(input?.WPS) && reason?.Contains("Uniqe_WPS", StringComparison.OrdinalIgnoreCase) == true)
            {
                TempData["Msg"] = $"Failed to save WPS. WPS '{input.WPS}' already exists.";
            }
            else
            {
                TempData["Msg"] = string.IsNullOrWhiteSpace(reason)
                    ? "Failed to save WPS."
                    : $"Failed to save WPS. {reason}";
            }
        }
        return RedirectToAction(nameof(WpsLog));
    }

    // POST: /Home/DeleteWps
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteWps([FromForm] int id)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        try
        {
            var w = await _context.WPS_tbl.FirstOrDefaultAsync(x => x.WPS_ID == id);
            if (w == null) return NotFound();

            _context.WPS_tbl.Remove(w);
            await _context.SaveChangesAsync();
            var wpsName = string.IsNullOrWhiteSpace(w?.WPS) ? null : w.WPS.Trim();
            TempData["Msg"] = wpsName == null
                ? $"Deleted WPS ID {id}."
                : $"Deleted WPS '{wpsName}'.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteWps failed for {Id}", id);
            TempData["Msg"] = "Failed to delete WPS.";
        }
        return RedirectToAction(nameof(WpsLog));
    }

    // GET: /Home/ExportWps?q=search
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportWps([FromQuery] string? q)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        var qry = _context.WPS_tbl.AsNoTracking();
        if (!string.IsNullOrWhiteSpace(q))
        {
            var term = q.Trim();
            qry = qry.Where(w =>
                EF.Functions.Like(w.WPS!, $"%{term}%") ||
                (w.WPS_Material != null && EF.Functions.Like(w.WPS_Material, $"%{term}%")) ||
                (w.WPS_Pipe_Class != null && EF.Functions.Like(w.WPS_Pipe_Class, $"%{term}%")) ||
                (w.Project_WPS != null && EF.Functions.Like(w.Project_WPS.ToString()!, $"%{term}%"))
            );
        }

        var rows = await qry
            .OrderBy(w => w.WPS_ID)
            .Select(w => new
            {
                w.WPS,
                Project = w.Project_WPS,
                w.Weld_Process,
                PWHT = w.PWHT,
                w.Dia_Range,
                w.Thickness_Range,
                w.Thickness_Range_From,
                w.Thickness_Range_To,
                w.WPS_P_NO,
                w.WPS_Material,
                w.Electrode,
                w.WPS_Pipe_Class,
                w.WPS_Service
            })
            .ToListAsync();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("WPS Log");

        string[] headers = { "WPS", "Project", "Weld Process", "PWHT", "Dia Range", "Thic. Range", "Thic. From", "Thic. To", "P NO", "Material", "Electrode", "Pipe Class", "Service" };

        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        int r = 2;
        foreach (var w in rows)
        {
            ws.Cell(r, 1).Value = w.WPS ?? string.Empty;
            ws.Cell(r, 2).Value = w.Project;
            ws.Cell(r, 3).Value = NormalizeWeldProcessLabel(w.Weld_Process);
            ws.Cell(r, 4).Value = w.PWHT ? "Yes" : "No";
            ws.Cell(r, 5).Value = w.Dia_Range ?? string.Empty;
            ws.Cell(r, 6).Value = w.Thickness_Range ?? string.Empty;
            ws.Cell(r, 7).Value = w.Thickness_Range_From;
            ws.Cell(r, 8).Value = w.Thickness_Range_To;
            ws.Cell(r, 9).Value = w.WPS_P_NO ?? string.Empty;
            ws.Cell(r, 10).Value = w.WPS_Material ?? string.Empty;
            ws.Cell(r, 11).Value = w.Electrode ?? string.Empty;
            ws.Cell(r, 12).Value = w.WPS_Pipe_Class ?? string.Empty;
            ws.Cell(r, 13).Value = w.WPS_Service ?? string.Empty;
            ws.Row(r).Height = 17;
            r++;
        }

        int lastRow = r - 1;
        var fullRange = ws.Range(1, 1, lastRow, headers.Length);
        var table = fullRange.CreateTable();
        table.Theme = XLTableTheme.TableStyleMedium2;
        table.ShowTotalsRow = false;

        ws.Row(1).Height = 30;
        ws.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Column(7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"WpsLog_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}
