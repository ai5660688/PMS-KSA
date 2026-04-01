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
    // GET: /Home/PlnLog
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> PlnLog([FromQuery] int? projectId)
    {
        // Load all projects for both filter and form
        var allProjects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .Select(p => new { p.Project_ID, p.Project_Name })
            .ToListAsync();

        var defaultProjectId = await GetDefaultProjectIdAsync(projectId);

        // Filter PLN list by selected project
        var query = _context.PLN_tbl.AsNoTracking();
        if (defaultProjectId.HasValue)
            query = query.Where(p => p.PLN_Project_No == defaultProjectId.Value);

        var list = await query.OrderBy(p => p.PLN_ID).ToListAsync();

        // Project filter dropdown (all projects)
        ViewBag.FilterProjects = allProjects
            .Select(p =>
            {
                var name = (p.Project_Name ?? string.Empty).Trim();
                var hasName = !string.IsNullOrWhiteSpace(name);
                return new SelectListItem
                {
                    Value = p.Project_ID.ToString(),
                    Text = hasName ? $"{p.Project_ID} - {name}" : p.Project_ID.ToString(),
                    Selected = (defaultProjectId.HasValue && p.Project_ID == defaultProjectId.Value)
                };
            })
            .ToList();

        // Lookup map: Project_ID -> "ID - Name" for table display
        ViewBag.ProjectMap = allProjects.ToDictionary(
            p => p.Project_ID,
            p =>
            {
                var name = (p.Project_Name ?? string.Empty).Trim();
                return !string.IsNullOrWhiteSpace(name) ? $"{p.Project_ID} - {name}" : p.Project_ID.ToString();
            });

        ViewBag.DefaultProjectId = defaultProjectId;

        // Full name for header
        var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
        var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
        var composed = (fn + " " + ln).Trim();
        if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
        ViewBag.FullName = composed;

        return View(list);
    }

    // POST: /Home/SavePln
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SavePln(Pln input)
    {
        try
        {
            var existing = await _context.PLN_tbl.FirstOrDefaultAsync(x => x.PLN_ID == input.PLN_ID);

            if (existing == null)
            {
                _context.PLN_tbl.Add(input);
                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Added PLN record #{input.PLN_ID}.";
            }
            else
            {
                existing.PLN_Project_No = input.PLN_Project_No;
                existing.PLN_LOCATION = input.PLN_LOCATION;
                existing.PLN_DATE = input.PLN_DATE;
                existing.PLN_DIA = input.PLN_DIA;

                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Updated PLN record #{existing.PLN_ID}.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SavePln failed for {Id}", input.PLN_ID);
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            TempData["Msg"] = string.IsNullOrWhiteSpace(reason)
                ? "Failed to save PLN record."
                : $"Failed to save PLN record. {reason}";
        }
        return RedirectToAction(nameof(PlnLog));
    }

    // POST: /Home/DeletePln
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeletePln([FromForm] int id)
    {
        try
        {
            var p = await _context.PLN_tbl.FirstOrDefaultAsync(x => x.PLN_ID == id);
            if (p == null) return NotFound();

            _context.PLN_tbl.Remove(p);
            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Deleted PLN record #{id}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeletePln failed for {Id}", id);
            TempData["Msg"] = "Failed to delete PLN record.";
        }
        return RedirectToAction(nameof(PlnLog));
    }

    // GET: /Home/ExportPln?projectId=&q=search
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportPln([FromQuery] int? projectId, [FromQuery] string? q)
    {
        var qry = _context.PLN_tbl.AsNoTracking();
        if (projectId.HasValue)
            qry = qry.Where(p => p.PLN_Project_No == projectId.Value);

        if (!string.IsNullOrWhiteSpace(q))
        {
            var term = q.Trim();
            qry = qry.Where(p =>
                (p.PLN_LOCATION != null && EF.Functions.Like(p.PLN_LOCATION, $"%{term}%")) ||
                EF.Functions.Like(p.PLN_Project_No.ToString(), $"%{term}%")
            );
        }

        var rows = await qry
            .OrderBy(p => p.PLN_ID)
            .ToListAsync();

        // Build project name lookup
        var projectMap = await _context.Projects_tbl.AsNoTracking()
            .Select(p => new { p.Project_ID, p.Project_Name })
            .ToDictionaryAsync(
                p => p.Project_ID,
                p =>
                {
                    var name = (p.Project_Name ?? string.Empty).Trim();
                    return !string.IsNullOrWhiteSpace(name) ? $"{p.Project_ID} - {name}" : p.Project_ID.ToString();
                });

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Productivity Plan");

        string[] headers = { "ID", "Project No", "Location", "Date", "Dia" };

        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        int r = 2;
        foreach (var p in rows)
        {
            ws.Cell(r, 1).Value = p.PLN_ID;
            ws.Cell(r, 2).Value = projectMap.TryGetValue(p.PLN_Project_No, out var projLabel) ? projLabel : p.PLN_Project_No.ToString();
            ws.Cell(r, 3).Value = p.PLN_LOCATION ?? string.Empty;
            if (p.PLN_DATE.HasValue)
            {
                ws.Cell(r, 4).Value = p.PLN_DATE.Value;
                ws.Cell(r, 4).Style.NumberFormat.Format = "yyyy-MM-dd";
            }
            ws.Cell(r, 5).Value = p.PLN_DIA ?? 0;
            ws.Row(r).Height = 17;
            r++;
        }

        int lastRow = r - 1;
        if (lastRow >= 2)
        {
            var fullRange = ws.Range(1, 1, lastRow, headers.Length);
            var table = fullRange.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowTotalsRow = false;
        }

        ws.Row(1).Height = 30;
        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"ProductivityPlan_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    // GET: /Home/DownloadPlnTemplate
    [SessionAuthorization]
    [HttpGet]
    public IActionResult DownloadPlnTemplate()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Productivity Plan");

        string[] headers = { "Project No", "Location", "Date", "Dia" };
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#e6f6fa");
            cell.Style.Font.FontColor = XLColor.FromHtml("#176d8a");
        }

        ws.Column(1).Width = 15;  // Project No
        ws.Column(2).Width = 12;  // Location
        ws.Column(3).Width = 14;  // Date
        ws.Column(4).Width = 10;  // Dia
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return File(ms.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "ProductivityPlan_Template.xlsx");
    }

    // POST: /Home/ImportPln
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImportPln(IFormFile? excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            TempData["Msg"] = "Please select an Excel file to import.";
            return RedirectToAction(nameof(PlnLog));
        }

        try
        {
            using var stream = excelFile.OpenReadStream();
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();

            int imported = 0;
            int updated = 0;
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            // Detect header row: find the columns by name
            var headerRow = ws.Row(1);
            int colId = -1, colProjectNo = -1, colLocation = -1, colDate = -1, colDia = -1;
            for (int c = 1; c <= ws.LastColumnUsed()?.ColumnNumber(); c++)
            {
                var header = (headerRow.Cell(c).GetString() ?? string.Empty).Trim().ToUpperInvariant();
                if (header is "ID" or "PLN_ID") colId = c;
                else if (header is "PROJECT NO" or "PROJECT_NO" or "PLN_PROJECT_NO") colProjectNo = c;
                else if (header is "LOCATION" or "PLN_LOCATION") colLocation = c;
                else if (header is "DATE" or "PLN_DATE") colDate = c;
                else if (header is "DIA" or "PLN_DIA") colDia = c;
            }

            // Fallback: assume column order ID, Project No, Location, Date, Dia
            if (colProjectNo < 0) { colId = 1; colProjectNo = 2; colLocation = 3; colDate = 4; colDia = 5; }

            for (int row = 2; row <= lastRow; row++)
            {
                var projectNoRaw = ws.Cell(row, colProjectNo).GetString()?.Trim();
                if (string.IsNullOrWhiteSpace(projectNoRaw)) continue;
                if (!int.TryParse(projectNoRaw, out var projectNo)) continue;

                var location = colLocation > 0 ? ws.Cell(row, colLocation).GetString()?.Trim() : null;
                if (location?.Length > 4) location = location[..4];

                DateTime? date = null;
                if (colDate > 0)
                {
                    var dateCell = ws.Cell(row, colDate);
                    if (dateCell.DataType == XLDataType.DateTime)
                        date = dateCell.GetDateTime();
                    else if (DateTime.TryParse(dateCell.GetString(), out var parsed))
                        date = parsed;
                }

                double? dia = null;
                if (colDia > 0)
                {
                    var diaRaw = ws.Cell(row, colDia).GetString()?.Trim();
                    if (double.TryParse(diaRaw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var parsedDia))
                        dia = parsedDia;
                }

                // Check if row has an ID and it already exists (update)
                int? existingId = null;
                if (colId > 0)
                {
                    var idRaw = ws.Cell(row, colId).GetString()?.Trim();
                    if (int.TryParse(idRaw, out var parsedId) && parsedId > 0)
                        existingId = parsedId;
                }

                if (existingId.HasValue)
                {
                    var existing = await _context.PLN_tbl.FirstOrDefaultAsync(x => x.PLN_ID == existingId.Value);
                    if (existing != null)
                    {
                        existing.PLN_Project_No = projectNo;
                        existing.PLN_LOCATION = location;
                        existing.PLN_DATE = date;
                        existing.PLN_DIA = dia;
                        updated++;
                        continue;
                    }
                }

                // New record
                _context.PLN_tbl.Add(new Pln
                {
                    PLN_Project_No = projectNo,
                    PLN_LOCATION = location,
                    PLN_DATE = date,
                    PLN_DIA = dia
                });
                imported++;
            }

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Import complete. {imported} added, {updated} updated.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ImportPln failed");
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            TempData["Msg"] = $"Import failed. {reason}";
        }

        return RedirectToAction(nameof(PlnLog));
    }
}
