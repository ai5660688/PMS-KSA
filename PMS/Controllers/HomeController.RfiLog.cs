using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    // GET: /Home/RfiLog
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> RfiLog([FromQuery] int? projectId)
    {
        var allProjects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .Select(p => new { p.Project_ID, p.Project_Name })
            .ToListAsync();

        var defaultProjectId = await GetDefaultProjectIdAsync(projectId);

        var query = _context.RFI_tbl.AsNoTracking();
        if (defaultProjectId.HasValue)
            query = query.Where(r => r.RFI_Project_No == defaultProjectId.Value);

        var list = await query.OrderBy(r => r.RFI_ID).ToListAsync();

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

        ViewBag.ProjectMap = allProjects.ToDictionary(
            p => p.Project_ID,
            p =>
            {
                var name = (p.Project_Name ?? string.Empty).Trim();
                return !string.IsNullOrWhiteSpace(name) ? $"{p.Project_ID} - {name}" : p.Project_ID.ToString();
            });

        ViewBag.DefaultProjectId = defaultProjectId;

        var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
        var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
        var composed = (fn + " " + ln).Trim();
        if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
        ViewBag.FullName = composed;

        return View(list);
    }

    // POST: /Home/SaveRfi
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveRfi(Rfi input)
    {
        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            input.RFI_Updated_Date = AppClock.Now;
            input.RFI_Updated_By = userId;

            var existing = input.RFI_ID > 0
                ? await _context.RFI_tbl.FirstOrDefaultAsync(x => x.RFI_ID == input.RFI_ID)
                : null;

            if (existing == null)
            {
                _context.RFI_tbl.Add(input);
                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Added RFI record #{input.RFI_ID}.";
            }
            else
            {
                // Preserve file-related fields managed by upload endpoints
                var fileName = existing.RFI_FileName;
                var fileSize = existing.RFI_FileSize;
                var uploadDate = existing.RFI_UploadDate;
                var blobName = existing.RFI_BlobName;

                _context.Entry(existing).CurrentValues.SetValues(input);

                existing.RFI_FileName = fileName;
                existing.RFI_FileSize = fileSize;
                existing.RFI_UploadDate = uploadDate;
                existing.RFI_BlobName = blobName;

                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Updated RFI record #{existing.RFI_ID}.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveRfi failed for {Id}", input.RFI_ID);
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            TempData["Msg"] = string.IsNullOrWhiteSpace(reason)
                ? "Failed to save RFI record."
                : $"Failed to save RFI record. {reason}";
        }
        return RedirectToAction(nameof(RfiLog), new { projectId = input.RFI_Project_No });
    }

    // POST: /Home/DeleteRfi
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteRfi([FromForm] int id)
    {
        int? projectId = null;
        try
        {
            var r = await _context.RFI_tbl.FirstOrDefaultAsync(x => x.RFI_ID == id);
            if (r == null) return NotFound();

            projectId = r.RFI_Project_No;

            // Remove attached PDF from disk
            if (!string.IsNullOrWhiteSpace(r.RFI_BlobName))
            {
                try
                {
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files", Path.GetFileName(r.RFI_BlobName));
                    if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);
                }
                catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete PDF for RFI {Id}", r.RFI_ID); }
            }

            _context.RFI_tbl.Remove(r);
            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Deleted RFI record #{id}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteRfi failed for {Id}", id);
            TempData["Msg"] = "Failed to delete RFI record.";
        }
        return RedirectToAction(nameof(RfiLog), new { projectId });
    }

    // POST: /Home/UploadRfiFile
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UploadRfiFile([FromForm] int rfiId, IFormFile? pdfFile)
    {
        int? projectId = null;
        try
        {
            var rfi = await _context.RFI_tbl.FirstOrDefaultAsync(x => x.RFI_ID == rfiId);
            if (rfi == null) return NotFound();
            projectId = rfi.RFI_Project_No;

            if (pdfFile == null || pdfFile.Length == 0)
            {
                TempData["Msg"] = "Please select a PDF file to upload.";
                return RedirectToAction(nameof(RfiLog), new { projectId });
            }

            const long maxPdfSize = 50 * 1024 * 1024; // 50 MB
            if (pdfFile.Length > maxPdfSize)
            {
                TempData["Msg"] = "PDF file exceeds the 50 MB size limit.";
                return RedirectToAction(nameof(RfiLog), new { projectId });
            }

            if (!string.Equals(Path.GetExtension(pdfFile.FileName), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                TempData["Msg"] = "Only PDF files are allowed.";
                return RedirectToAction(nameof(RfiLog), new { projectId });
            }

            // Remove previous file if any
            if (!string.IsNullOrWhiteSpace(rfi.RFI_BlobName))
            {
                try
                {
                    var oldPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files", Path.GetFileName(rfi.RFI_BlobName));
                    if (System.IO.File.Exists(oldPath)) System.IO.File.Delete(oldPath);
                }
                catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete old PDF for RFI {Id}", rfi.RFI_ID); }
            }

            var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files");
            Directory.CreateDirectory(folder);

            var blobName = $"{Guid.NewGuid():N}.pdf";
            var savePath = Path.Combine(folder, blobName);
            using (var fs = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write))
            {
                await pdfFile.CopyToAsync(fs);
            }

            rfi.RFI_FileName = Path.GetFileName(pdfFile.FileName);
            rfi.RFI_FileSize = (int)pdfFile.Length;
            rfi.RFI_UploadDate = AppClock.Now;
            rfi.RFI_BlobName = blobName;
            rfi.RFI_Updated_Date = AppClock.Now;
            rfi.RFI_Updated_By = HttpContext.Session.GetInt32("UserID");

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"PDF uploaded for RFI #{rfiId}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "UploadRfiFile failed for RFI {Id}", rfiId);
            TempData["Msg"] = $"PDF upload failed. {ex.GetBaseException()?.Message ?? ex.Message}";
        }
        return RedirectToAction(nameof(RfiLog), new { projectId });
    }

    // GET: /Home/DownloadRfiFile?id=  (force download)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DownloadRfiFile([FromQuery] int id)
    {
        var rfi = await _context.RFI_tbl.AsNoTracking().FirstOrDefaultAsync(x => x.RFI_ID == id);
        if (rfi == null || string.IsNullOrWhiteSpace(rfi.RFI_BlobName)) return NotFound();

        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files", Path.GetFileName(rfi.RFI_BlobName));
        if (!System.IO.File.Exists(filePath)) return NotFound();

        var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
        var downloadName = rfi.RFI_FileName ?? $"RFI_{rfi.RFI_ID}.pdf";
        return File(bytes, "application/pdf", downloadName);
    }

    // GET: /Home/OpenRfiFile?id=  (open inline in browser)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> OpenRfiFile([FromQuery] int id)
    {
        var rfi = await _context.RFI_tbl.AsNoTracking().FirstOrDefaultAsync(x => x.RFI_ID == id);
        if (rfi == null || string.IsNullOrWhiteSpace(rfi.RFI_BlobName)) return NotFound();

        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files", Path.GetFileName(rfi.RFI_BlobName));
        if (!System.IO.File.Exists(filePath)) return NotFound();

        var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
        return File(bytes, "application/pdf");
    }

    // POST: /Home/DeleteRfiFile
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteRfiFile([FromForm] int rfiId)
    {
        int? projectId = null;
        try
        {
            var rfi = await _context.RFI_tbl.FirstOrDefaultAsync(x => x.RFI_ID == rfiId);
            if (rfi == null) return NotFound();
            projectId = rfi.RFI_Project_No;

            if (!string.IsNullOrWhiteSpace(rfi.RFI_BlobName))
            {
                try
                {
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files", Path.GetFileName(rfi.RFI_BlobName));
                    if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);
                }
                catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete PDF for RFI {Id}", rfi.RFI_ID); }
            }

            rfi.RFI_FileName = null;
            rfi.RFI_FileSize = null;
            rfi.RFI_UploadDate = null;
            rfi.RFI_BlobName = null;
            rfi.RFI_Updated_Date = AppClock.Now;
            rfi.RFI_Updated_By = HttpContext.Session.GetInt32("UserID");

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"PDF removed from RFI #{rfiId}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteRfiFile failed for RFI {Id}", rfiId);
            TempData["Msg"] = "Failed to remove PDF.";
        }
        return RedirectToAction(nameof(RfiLog), new { projectId });
    }

    // GET: /Home/ExportRfi?projectId=&q=search
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportRfi([FromQuery] int? projectId, [FromQuery] string? q)
    {
        var qry = _context.RFI_tbl.AsNoTracking();
        if (projectId.HasValue)
            qry = qry.Where(r => r.RFI_Project_No == projectId.Value);

        if (!string.IsNullOrWhiteSpace(q))
        {
            var term = q.Trim();
            qry = qry.Where(r =>
                (r.DISCIPLINE != null && EF.Functions.Like(r.DISCIPLINE, $"%{term}%")) ||
                (r.SubCon_RFI_No != null && EF.Functions.Like(r.SubCon_RFI_No, $"%{term}%")) ||
                (r.RFI_LOCATION != null && EF.Functions.Like(r.RFI_LOCATION, $"%{term}%")) ||
                (r.RFI_DESCRIPTION != null && EF.Functions.Like(r.RFI_DESCRIPTION, $"%{term}%")) ||
                (r.ELEMENT != null && EF.Functions.Like(r.ELEMENT, $"%{term}%")) ||
                (r.INSPECTION_STATUS != null && EF.Functions.Like(r.INSPECTION_STATUS, $"%{term}%")) ||
                (r.PID != null && EF.Functions.Like(r.PID, $"%{term}%")) ||
                (r.REMARKS != null && EF.Functions.Like(r.REMARKS, $"%{term}%"))
            );
        }

        var rows = await qry.OrderBy(r => r.RFI_ID).ToListAsync();

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
        var ws = wb.Worksheets.Add("RFI Log");

        string[] headers = {
            "RFI ID", "Project No", "Discipline", "Sub Contractor", "Sub Discipline",
            "RFI Contractor", "SubCon RFI No", "EPM No", "SATIP", "SAIC",
            "Activity", "Location", "Unit", "Element", "Description",
            "Date", "Time", "Company Insp Level", "Contractor Insp Level", "SubCon Insp Level",
            "TR QC", "Sub Con QC", "PID", "PMT", "Inspection Status",
            "Ref Drawing No", "Scan Copy", "Remarks", "QR Code Link"
        };

        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        int row = 2;
        foreach (var r in rows)
        {
            ws.Cell(row, 1).Value = r.RFI_ID;
            ws.Cell(row, 2).Value = projectMap.TryGetValue(r.RFI_Project_No ?? 0, out var projLabel) ? projLabel : (r.RFI_Project_No?.ToString() ?? "");
            ws.Cell(row, 3).Value = r.DISCIPLINE ?? "";
            ws.Cell(row, 4).Value = r.Sub_Contractor ?? "";
            ws.Cell(row, 5).Value = r.SubDiscipline ?? "";
            ws.Cell(row, 6).Value = r.RFI_CONTRACTOR ?? "";
            ws.Cell(row, 7).Value = r.SubCon_RFI_No ?? "";
            ws.Cell(row, 8).Value = r.EPMNo ?? 0;
            ws.Cell(row, 9).Value = r.SATIP ?? "";
            ws.Cell(row, 10).Value = r.SAIC ?? "";
            ws.Cell(row, 11).Value = r.ACTIVITY ?? 0;
            ws.Cell(row, 12).Value = r.RFI_LOCATION ?? "";
            ws.Cell(row, 13).Value = r.UNIT ?? "";
            ws.Cell(row, 14).Value = r.ELEMENT ?? "";
            ws.Cell(row, 15).Value = r.RFI_DESCRIPTION ?? "";
            if (r.Date.HasValue)
            {
                ws.Cell(row, 16).Value = r.Date.Value;
                ws.Cell(row, 16).Style.NumberFormat.Format = "yyyy-MM-dd";
            }
            if (r.Time.HasValue)
            {
                ws.Cell(row, 17).Value = r.Time.Value.ToString("HH:mm");
            }
            ws.Cell(row, 18).Value = r.COMPANY_INSPECTION_LEVEL ?? "";
            ws.Cell(row, 19).Value = r.CONTRACTOR_INSPECTION_LEVEL ?? "";
            ws.Cell(row, 20).Value = r.SUBCON_INSPECTION_LEVEL ?? "";
            ws.Cell(row, 21).Value = r.TR_QC ?? "";
            ws.Cell(row, 22).Value = r.SUB_CON_QC ?? "";
            ws.Cell(row, 23).Value = r.PID ?? "";
            ws.Cell(row, 24).Value = r.PMT ?? "";
            ws.Cell(row, 25).Value = r.INSPECTION_STATUS ?? "";
            ws.Cell(row, 26).Value = r.REFRENCE_DRAWING_No ?? "";
            ws.Cell(row, 27).Value = r.SCAN_COPY ?? "";
            ws.Cell(row, 28).Value = r.REMARKS ?? "";
            ws.Cell(row, 29).Value = r.QR_Code_Link ?? "";
            ws.Row(row).Height = 17;
            row++;
        }

        int lastRow = row - 1;
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
        var fileName = $"RFI_Log_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    // GET: /Home/DownloadRfiTemplate
    [SessionAuthorization]
    [HttpGet]
    public IActionResult DownloadRfiTemplate()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("RFI Log");

        string[] headers = {
            "Discipline", "Sub Contractor", "Sub Discipline",
            "RFI Contractor", "SubCon RFI No", "EPM No", "SATIP", "SAIC",
            "Activity", "Location", "Unit", "Element", "Description",
            "Date", "Time", "Company Insp Level", "Contractor Insp Level", "SubCon Insp Level",
            "TR QC", "Sub Con QC", "PID", "PMT", "Inspection Status",
            "Ref Drawing No", "Scan Copy", "Remarks", "QR Code Link", "Project No"
        };
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#e6f6fa");
            cell.Style.Font.FontColor = XLColor.FromHtml("#176d8a");
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return File(ms.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "RFI_Log_Template.xlsx");
    }

    // POST: /Home/ImportRfi
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImportRfi(IFormFile? excelFile, [FromForm] int? projectId)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            TempData["Msg"] = "Please select an Excel file to import.";
            return RedirectToAction(nameof(RfiLog));
        }

        try
        {
            using var stream = excelFile.OpenReadStream();
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();

            int imported = 0;
            int updated = 0;
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            var userId = HttpContext.Session.GetInt32("UserID");

            // Detect columns by header name
            var headerRow = ws.Row(1);
            int colId = -1, colProjectNo = -1, colDiscipline = -1, colSubContractor = -1, colSubDiscipline = -1,
                colRfiContractor = -1, colSubConRfiNo = -1, colEpmNo = -1, colSatip = -1, colSaic = -1,
                colActivity = -1, colLocation = -1, colUnit = -1, colElement = -1, colDescription = -1,
                colDate = -1, colTime = -1, colCompanyInsp = -1, colContractorInsp = -1, colSubConInsp = -1,
                colTrQc = -1, colSubConQc = -1, colPid = -1, colPmt = -1, colInspStatus = -1,
                colRefDrawing = -1, colScanCopy = -1, colRemarks = -1, colQrCode = -1;

            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            for (int c = 1; c <= lastCol; c++)
            {
                var h = (headerRow.Cell(c).GetString() ?? "").Trim().ToUpperInvariant();
                switch (h)
                {
                    case "RFI ID" or "RFI_ID" or "ID": colId = c; break;
                    case "PROJECT NO" or "PROJECT_NO" or "RFI_PROJECT_NO": colProjectNo = c; break;
                    case "DISCIPLINE": colDiscipline = c; break;
                    case "SUB CONTRACTOR" or "SUB_CONTRACTOR": colSubContractor = c; break;
                    case "SUB DISCIPLINE" or "SUB_DISCIPLINE": colSubDiscipline = c; break;
                    case "RFI CONTRACTOR" or "RFI_CONTRACTOR": colRfiContractor = c; break;
                    case "SUBCON RFI NO" or "SUBCON_RFI_NO": colSubConRfiNo = c; break;
                    case "EPM NO" or "EPMNO": colEpmNo = c; break;
                    case "SATIP": colSatip = c; break;
                    case "SAIC": colSaic = c; break;
                    case "ACTIVITY": colActivity = c; break;
                    case "LOCATION" or "RFI_LOCATION": colLocation = c; break;
                    case "UNIT": colUnit = c; break;
                    case "ELEMENT": colElement = c; break;
                    case "DESCRIPTION" or "RFI_DESCRIPTION": colDescription = c; break;
                    case "DATE": colDate = c; break;
                    case "TIME": colTime = c; break;
                    case "COMPANY INSP LEVEL" or "COMPANY_INSPECTION_LEVEL": colCompanyInsp = c; break;
                    case "CONTRACTOR INSP LEVEL" or "CONTRACTOR_INSPECTION_LEVEL": colContractorInsp = c; break;
                    case "SUBCON INSP LEVEL" or "SUBCON_INSPECTION_LEVEL": colSubConInsp = c; break;
                    case "TR QC" or "TR_QC": colTrQc = c; break;
                    case "SUB CON QC" or "SUB_CON_QC": colSubConQc = c; break;
                    case "PID": colPid = c; break;
                    case "PMT": colPmt = c; break;
                    case "INSPECTION STATUS" or "INSPECTION_STATUS": colInspStatus = c; break;
                    case "REF DRAWING NO" or "REFRENCE_DRAWING_NO": colRefDrawing = c; break;
                    case "SCAN COPY" or "SCAN_COPY": colScanCopy = c; break;
                    case "REMARKS": colRemarks = c; break;
                    case "QR CODE LINK" or "QR_CODE_LINK": colQrCode = c; break;
                }
            }

            for (int r = 2; r <= lastRow; r++)
            {
                string GetStr(int col) => col > 0 ? (ws.Cell(r, col).GetString()?.Trim() ?? "") : "";
                int? GetInt(int col)
                {
                    if (col <= 0) return null;
                    var raw = ws.Cell(r, col).GetString()?.Trim();
                    return int.TryParse(raw, out var v) ? v : null;
                }
                double? GetDouble(int col)
                {
                    if (col <= 0) return null;
                    var raw = ws.Cell(r, col).GetString()?.Trim();
                    return double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;
                }
                DateTime? GetDate(int col)
                {
                    if (col <= 0) return null;
                    var cell = ws.Cell(r, col);
                    if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
                    return DateTime.TryParse(cell.GetString(), out var parsed) ? parsed : null;
                }
                DateTime? GetTime(int col)
                {
                    if (col <= 0) return null;
                    var cell = ws.Cell(r, col);
                    if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
                    var raw = cell.GetString()?.Trim();
                    if (TimeSpan.TryParse(raw, out var ts)) return DateTime.MinValue.Add(ts);
                    if (DateTime.TryParse(raw, out var parsed)) return parsed;
                    return null;
                }

                // Skip entirely blank rows
                var discipline = GetStr(colDiscipline);
                var location = GetStr(colLocation);
                var description = GetStr(colDescription);
                if (string.IsNullOrWhiteSpace(discipline) && string.IsNullOrWhiteSpace(location) && string.IsNullOrWhiteSpace(description))
                    continue;

                int? existingId = GetInt(colId);
                Rfi? existing = null;
                if (existingId.HasValue && existingId.Value > 0)
                    existing = await _context.RFI_tbl.FirstOrDefaultAsync(x => x.RFI_ID == existingId.Value);

                var rfi = existing ?? new Rfi();
                rfi.RFI_Project_No = GetInt(colProjectNo) ?? projectId;
                rfi.DISCIPLINE = NullIfEmpty(GetStr(colDiscipline), 25);
                rfi.Sub_Contractor = NullIfEmpty(GetStr(colSubContractor), 255);
                rfi.SubDiscipline = NullIfEmpty(GetStr(colSubDiscipline), 25);
                rfi.RFI_CONTRACTOR = NullIfEmpty(GetStr(colRfiContractor), 255);
                rfi.SubCon_RFI_No = NullIfEmpty(GetStr(colSubConRfiNo), 10);
                rfi.EPMNo = GetInt(colEpmNo);
                rfi.SATIP = NullIfEmpty(GetStr(colSatip), 25);
                rfi.SAIC = NullIfEmpty(GetStr(colSaic), 25);
                rfi.ACTIVITY = GetDouble(colActivity);
                rfi.RFI_LOCATION = NullIfEmpty(GetStr(colLocation), 25);
                rfi.UNIT = NullIfEmpty(GetStr(colUnit), 15);
                rfi.ELEMENT = NullIfEmpty(GetStr(colElement), 150);
                rfi.RFI_DESCRIPTION = NullIfEmpty(GetStr(colDescription));
                rfi.Date = GetDate(colDate);
                rfi.Time = GetTime(colTime);
                rfi.COMPANY_INSPECTION_LEVEL = NullIfEmpty(GetStr(colCompanyInsp), 4);
                rfi.CONTRACTOR_INSPECTION_LEVEL = NullIfEmpty(GetStr(colContractorInsp), 4);
                rfi.SUBCON_INSPECTION_LEVEL = NullIfEmpty(GetStr(colSubConInsp), 4);
                rfi.TR_QC = NullIfEmpty(GetStr(colTrQc), 30);
                rfi.SUB_CON_QC = NullIfEmpty(GetStr(colSubConQc), 30);
                rfi.PID = NullIfEmpty(GetStr(colPid), 30);
                rfi.PMT = NullIfEmpty(GetStr(colPmt), 30);
                rfi.INSPECTION_STATUS = NullIfEmpty(GetStr(colInspStatus), 50);
                rfi.REFRENCE_DRAWING_No = NullIfEmpty(GetStr(colRefDrawing), 50);
                rfi.SCAN_COPY = NullIfEmpty(GetStr(colScanCopy), 8);
                rfi.REMARKS = NullIfEmpty(GetStr(colRemarks), 100);
                rfi.QR_Code_Link = NullIfEmpty(GetStr(colQrCode), 500);
                rfi.RFI_Updated_Date = AppClock.Now;
                rfi.RFI_Updated_By = userId;

                if (existing != null)
                    updated++;
                else
                {
                    _context.RFI_tbl.Add(rfi);
                    imported++;
                }
            }

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Import complete. {imported} added, {updated} updated.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ImportRfi failed");
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            TempData["Msg"] = $"Import failed. {reason}";
        }

        return RedirectToAction(nameof(RfiLog), new { projectId });
    }

    // POST: /Home/BulkUploadRfiPdfs
    // Matches each uploaded PDF filename against SubCon_RFI_No for the selected project.
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkUploadRfiPdfs([FromForm] int projectId, List<IFormFile>? pdfFiles)
    {
        if (pdfFiles == null || pdfFiles.Count == 0)
        {
            TempData["Msg"] = "Please select at least one PDF file.";
            return RedirectToAction(nameof(RfiLog), new { projectId });
        }

        try
        {
            // Load all RFIs for this project that have a SubCon RFI No
            var rfis = await _context.RFI_tbl
                .Where(r => r.RFI_Project_No == projectId && r.SubCon_RFI_No != null)
                .ToListAsync();

            var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files");
            Directory.CreateDirectory(folder);

            var userId = HttpContext.Session.GetInt32("UserID");
            int matched = 0;
            int skipped = 0;
            var unmatchedNames = new List<string>();

            foreach (var file in pdfFiles)
            {
                if (file.Length == 0) continue;

                if (!string.Equals(Path.GetExtension(file.FileName), ".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    skipped++;
                    continue;
                }

                var fileNameWithoutExt = Path.GetFileNameWithoutExtension(file.FileName) ?? "";

                // Find the RFI whose SubCon_RFI_No appears anywhere in the filename (case-insensitive)
                var rfi = rfis.FirstOrDefault(r =>
                    !string.IsNullOrWhiteSpace(r.SubCon_RFI_No) &&
                    fileNameWithoutExt.Contains(r.SubCon_RFI_No.Trim(), StringComparison.OrdinalIgnoreCase));

                if (rfi == null)
                {
                    unmatchedNames.Add(Path.GetFileName(file.FileName));
                    continue;
                }

                // Remove previous file if any
                if (!string.IsNullOrWhiteSpace(rfi.RFI_BlobName))
                {
                    try
                    {
                        var oldPath = Path.Combine(folder, Path.GetFileName(rfi.RFI_BlobName));
                        if (System.IO.File.Exists(oldPath)) System.IO.File.Delete(oldPath);
                    }
                    catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete old PDF for RFI {Id}", rfi.RFI_ID); }
                }

                var blobName = $"{Guid.NewGuid():N}.pdf";
                var savePath = Path.Combine(folder, blobName);
                using (var fs = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write))
                {
                    await file.CopyToAsync(fs);
                }

                rfi.RFI_FileName = Path.GetFileName(file.FileName);
                rfi.RFI_FileSize = (int)file.Length;
                rfi.RFI_UploadDate = AppClock.Now;
                rfi.RFI_BlobName = blobName;
                rfi.RFI_Updated_Date = AppClock.Now;
                rfi.RFI_Updated_By = userId;
                matched++;
            }

            await _context.SaveChangesAsync();

            var parts = new List<string>();
            if (matched > 0) parts.Add($"{matched} PDF(s) matched & uploaded");
            if (skipped > 0) parts.Add($"{skipped} non-PDF file(s) skipped");
            if (unmatchedNames.Count > 0) parts.Add($"{unmatchedNames.Count} file(s) had no SubCon RFI No match: {string.Join(", ", unmatchedNames)}");
            TempData["Msg"] = parts.Count > 0 ? string.Join(". ", parts) + "." : "No files processed.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "BulkUploadRfiPdfs failed for project {ProjectId}", projectId);
            TempData["Msg"] = $"Bulk PDF upload failed. {ex.GetBaseException()?.Message ?? ex.Message}";
        }

        return RedirectToAction(nameof(RfiLog), new { projectId });
    }

    // GET: /Home/BulkDownloadRfiPdfs?projectId=
    // Downloads all attached PDFs for the project as a ZIP archive.
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> BulkDownloadRfiPdfs([FromQuery] int projectId)
    {
        var rfis = await _context.RFI_tbl.AsNoTracking()
            .Where(r => r.RFI_Project_No == projectId && r.RFI_BlobName != null)
            .OrderBy(r => r.RFI_ID)
            .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.RFI_FileName, r.RFI_BlobName })
            .ToListAsync();

        if (rfis.Count == 0)
        {
            TempData["Msg"] = "No PDF files found for this project.";
            return RedirectToAction(nameof(RfiLog), new { projectId });
        }

        var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "rfi-files");

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rfi in rfis)
            {
                if (string.IsNullOrWhiteSpace(rfi.RFI_BlobName)) continue;
                var filePath = Path.Combine(folder, Path.GetFileName(rfi.RFI_BlobName));
                if (!System.IO.File.Exists(filePath)) continue;

                // Build a friendly download name: prefer original file name, fall back to RFI_ID
                var entryName = rfi.RFI_FileName ?? $"RFI_{rfi.RFI_ID}.pdf";
                // Ensure unique names inside the ZIP
                var baseName = Path.GetFileNameWithoutExtension(entryName);
                var ext = Path.GetExtension(entryName);
                var candidate = entryName;
                int counter = 1;
                while (!usedNames.Add(candidate))
                {
                    candidate = $"{baseName}_{counter}{ext}";
                    counter++;
                }

                var entry = archive.CreateEntry(candidate, CompressionLevel.Fastest);
                using var entryStream = entry.Open();
                using var fileStream = System.IO.File.OpenRead(filePath);
                await fileStream.CopyToAsync(entryStream);
            }
        }

        ms.Position = 0;
        var zipName = $"RFI_PDFs_Project_{projectId}_{AppClock.Now:yyyyMMdd_HHmmss}.zip";
        return File(ms.ToArray(), "application/zip", zipName);
    }

    // POST: /Home/BulkUpdateInspectionStatus
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkUpdateInspectionStatus([FromBody] BulkInspectionStatusRequest request)
    {
        if (request == null || request.Ids == null || request.Ids.Count == 0 || string.IsNullOrWhiteSpace(request.Status))
            return Json(new { success = false, message = "Invalid request." });

        var status = request.Status.Trim();
        if (status.Length > 50)
            return Json(new { success = false, message = "Status value too long (max 50)." });

        var userId = HttpContext.Session.GetInt32("UserID");
        var rfis = await _context.RFI_tbl
            .Where(r => request.Ids.Contains(r.RFI_ID))
            .ToListAsync();

        foreach (var rfi in rfis)
        {
            rfi.INSPECTION_STATUS = status;
            rfi.RFI_Updated_Date = AppClock.Now;
            rfi.RFI_Updated_By = userId;
        }

        await _context.SaveChangesAsync();
        return Json(new { success = true, count = rfis.Count });
    }

    public class BulkInspectionStatusRequest
    {
        public List<int> Ids { get; set; } = new();
        public string Status { get; set; } = "";
    }

    private static string? NullIfEmpty(string? s, int? maxLength = null)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        return maxLength.HasValue && s.Length > maxLength.Value ? s[..maxLength.Value] : s;
    }
}
