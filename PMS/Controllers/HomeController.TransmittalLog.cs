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
    // GET: /Home/TransmittalLog
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> TransmittalLog([FromQuery] int? projectId)
    {
        var allProjects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .Select(p => new { p.Project_ID, p.Project_Name })
            .ToListAsync();

        var defaultProjectId = await GetDefaultProjectIdAsync(projectId);

        var query = _context.Transmittal_Log_tbl.AsNoTracking();
        if (defaultProjectId.HasValue)
            query = query.Where(t => t.TR_Project_No == defaultProjectId.Value);

        var list = await query.OrderBy(t => t.ID).ToListAsync();

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

        // Resolve default discipline from the user's Access value
        var access = HttpContext.Session.GetString("Access") ?? string.Empty;
        ViewBag.DefaultDiscipline = ResolveDisciplineFromAccess(access);

        var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
        var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
        var composed = (fn + " " + ln).Trim();
        if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
        ViewBag.FullName = composed;

        return View(list);
    }

    // POST: /Home/SaveTransmittal
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveTransmittal(TransmittalLog input)
    {
        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            input.Transmittal_Updated_Date = AppClock.Now;
            input.Transmittal_Updated_By = userId;

            var existing = input.ID > 0
                ? await _context.Transmittal_Log_tbl.FirstOrDefaultAsync(x => x.ID == input.ID)
                : null;

            if (existing == null)
            {
                _context.Transmittal_Log_tbl.Add(input);
                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Added Transmittal record #{input.ID}.";
            }
            else
            {
                // Preserve file-related fields managed by upload endpoints
                var fileName = existing.TR_FileName;
                var fileSize = existing.TR_FileSize;
                var uploadDate = existing.TR_UploadDate;
                var blobName = existing.TR_BlobName;

                _context.Entry(existing).CurrentValues.SetValues(input);

                existing.TR_FileName = fileName;
                existing.TR_FileSize = fileSize;
                existing.TR_UploadDate = uploadDate;
                existing.TR_BlobName = blobName;

                await _context.SaveChangesAsync();
                TempData["Msg"] = $"Updated Transmittal record #{existing.ID}.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveTransmittal failed for {Id}", input.ID);
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            TempData["Msg"] = string.IsNullOrWhiteSpace(reason)
                ? "Failed to save Transmittal record."
                : $"Failed to save Transmittal record. {reason}";
        }
        return RedirectToAction(nameof(TransmittalLog), new { projectId = input.TR_Project_No });
    }

    // POST: /Home/DeleteTransmittal
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteTransmittal([FromForm] int id)
    {
        int? projectId = null;
        try
        {
            var t = await _context.Transmittal_Log_tbl.FirstOrDefaultAsync(x => x.ID == id);
            if (t == null) return NotFound();

            projectId = t.TR_Project_No;

            if (!string.IsNullOrWhiteSpace(t.TR_BlobName))
            {
                try
                {
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files", Path.GetFileName(t.TR_BlobName));
                    if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);
                }
                catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete PDF for Transmittal {Id}", t.ID); }
            }

            _context.Transmittal_Log_tbl.Remove(t);
            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Deleted Transmittal record #{id}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteTransmittal failed for {Id}", id);
            TempData["Msg"] = "Failed to delete Transmittal record.";
        }
        return RedirectToAction(nameof(TransmittalLog), new { projectId });
    }

    // POST: /Home/UploadTransmittalFile
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UploadTransmittalFile([FromForm] int trId, IFormFile? pdfFile)
    {
        int? projectId = null;
        try
        {
            var tr = await _context.Transmittal_Log_tbl.FirstOrDefaultAsync(x => x.ID == trId);
            if (tr == null) return NotFound();
            projectId = tr.TR_Project_No;

            if (pdfFile == null || pdfFile.Length == 0)
            {
                TempData["Msg"] = "Please select a PDF file to upload.";
                return RedirectToAction(nameof(TransmittalLog), new { projectId });
            }

            const long maxPdfSize = 50 * 1024 * 1024;
            if (pdfFile.Length > maxPdfSize)
            {
                TempData["Msg"] = "PDF file exceeds the 50 MB size limit.";
                return RedirectToAction(nameof(TransmittalLog), new { projectId });
            }

            if (!string.Equals(Path.GetExtension(pdfFile.FileName), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                TempData["Msg"] = "Only PDF files are allowed.";
                return RedirectToAction(nameof(TransmittalLog), new { projectId });
            }

            if (!string.IsNullOrWhiteSpace(tr.TR_BlobName))
            {
                try
                {
                    var oldPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files", Path.GetFileName(tr.TR_BlobName));
                    if (System.IO.File.Exists(oldPath)) System.IO.File.Delete(oldPath);
                }
                catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete old PDF for Transmittal {Id}", tr.ID); }
            }

            var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files");
            Directory.CreateDirectory(folder);

            var blobName = $"{Guid.NewGuid():N}.pdf";
            var savePath = Path.Combine(folder, blobName);
            using (var fs = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write))
            {
                await pdfFile.CopyToAsync(fs);
            }

            tr.TR_FileName = Path.GetFileName(pdfFile.FileName);
            tr.TR_FileSize = (int)pdfFile.Length;
            tr.TR_UploadDate = AppClock.Now;
            tr.TR_BlobName = blobName;
            tr.Transmittal_Updated_Date = AppClock.Now;
            tr.Transmittal_Updated_By = HttpContext.Session.GetInt32("UserID");

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"PDF uploaded for Transmittal #{trId}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "UploadTransmittalFile failed for Transmittal {Id}", trId);
            TempData["Msg"] = $"PDF upload failed. {ex.GetBaseException()?.Message ?? ex.Message}";
        }
        return RedirectToAction(nameof(TransmittalLog), new { projectId });
    }

    // GET: /Home/DownloadTransmittalFile?id=
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DownloadTransmittalFile([FromQuery] int id)
    {
        var tr = await _context.Transmittal_Log_tbl.AsNoTracking().FirstOrDefaultAsync(x => x.ID == id);
        if (tr == null || string.IsNullOrWhiteSpace(tr.TR_BlobName)) return NotFound();

        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files", Path.GetFileName(tr.TR_BlobName));
        if (!System.IO.File.Exists(filePath)) return NotFound();

        var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
        var downloadName = tr.TR_FileName ?? $"TR_{tr.ID}.pdf";
        return File(bytes, "application/pdf", downloadName);
    }

    // GET: /Home/OpenTransmittalFile?id=
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> OpenTransmittalFile([FromQuery] int id)
    {
        var tr = await _context.Transmittal_Log_tbl.AsNoTracking().FirstOrDefaultAsync(x => x.ID == id);
        if (tr == null || string.IsNullOrWhiteSpace(tr.TR_BlobName)) return NotFound();

        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files", Path.GetFileName(tr.TR_BlobName));
        if (!System.IO.File.Exists(filePath)) return NotFound();

        var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
        return File(bytes, "application/pdf");
    }

    // POST: /Home/DeleteTransmittalFile
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteTransmittalFile([FromForm] int trId)
    {
        int? projectId = null;
        try
        {
            var tr = await _context.Transmittal_Log_tbl.FirstOrDefaultAsync(x => x.ID == trId);
            if (tr == null) return NotFound();
            projectId = tr.TR_Project_No;

            if (!string.IsNullOrWhiteSpace(tr.TR_BlobName))
            {
                try
                {
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files", Path.GetFileName(tr.TR_BlobName));
                    if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);
                }
                catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete PDF for Transmittal {Id}", tr.ID); }
            }

            tr.TR_FileName = null;
            tr.TR_FileSize = null;
            tr.TR_UploadDate = null;
            tr.TR_BlobName = null;
            tr.Transmittal_Updated_Date = AppClock.Now;
            tr.Transmittal_Updated_By = HttpContext.Session.GetInt32("UserID");

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"PDF removed from Transmittal #{trId}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteTransmittalFile failed for Transmittal {Id}", trId);
            TempData["Msg"] = "Failed to remove PDF.";
        }
        return RedirectToAction(nameof(TransmittalLog), new { projectId });
    }

    // GET: /Home/ExportTransmittal?projectId=&q=
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportTransmittal([FromQuery] int? projectId, [FromQuery] string? q)
    {
        var qry = _context.Transmittal_Log_tbl.AsNoTracking();
        if (projectId.HasValue)
            qry = qry.Where(t => t.TR_Project_No == projectId.Value);

        if (!string.IsNullOrWhiteSpace(q))
        {
            var term = q.Trim();
            qry = qry.Where(t =>
                (t.Transmittal != null && EF.Functions.Like(t.Transmittal, $"%{term}%")) ||
                (t.Doc_No != null && EF.Functions.Like(t.Doc_No, $"%{term}%")) ||
                (t.eGesDoc_No != null && EF.Functions.Like(t.eGesDoc_No, $"%{term}%")) ||
                (t.Documents_Title != null && EF.Functions.Like(t.Documents_Title, $"%{term}%")) ||
                (t.Remarks != null && EF.Functions.Like(t.Remarks, $"%{term}%")) ||
                (t.Transmittal_No != null && EF.Functions.Like(t.Transmittal_No, $"%{term}%")) ||
                (t.Discipline != null && EF.Functions.Like(t.Discipline, $"%{term}%")) ||
                (t.Categorize != null && EF.Functions.Like(t.Categorize, $"%{term}%"))
            );
        }

        var rows = await qry.OrderBy(t => t.ID).ToListAsync();

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
        var ws = wb.Worksheets.Add("Transmittal Log");

        string[] headers = {
            "ID", "Project No", "Transmittal", "Date", "Doc No", "eGesDoc No",
            "Documents Title", "REV", "Remarks", "Transmittal No",
            "TR Comment", "Date TR", "TR REV",
            "SAPMT Comment", "Date SAPMT", "SAPMT REV",
            "Categorize", "Discipline"
        };

        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        int row = 2;
        foreach (var t in rows)
        {
            ws.Cell(row, 1).Value = t.ID;
            ws.Cell(row, 2).Value = projectMap.TryGetValue(t.TR_Project_No ?? 0, out var projLabel) ? projLabel : (t.TR_Project_No?.ToString() ?? "");
            ws.Cell(row, 3).Value = t.Transmittal ?? "";
            if (t.Date.HasValue) { ws.Cell(row, 4).Value = t.Date.Value; ws.Cell(row, 4).Style.NumberFormat.Format = "yyyy-MM-dd"; }
            ws.Cell(row, 5).Value = t.Doc_No ?? "";
            ws.Cell(row, 6).Value = t.eGesDoc_No ?? "";
            ws.Cell(row, 7).Value = t.Documents_Title ?? "";
            ws.Cell(row, 8).Value = t.REV ?? "";
            ws.Cell(row, 9).Value = t.Remarks ?? "";
            ws.Cell(row, 10).Value = t.Transmittal_No ?? "";
            ws.Cell(row, 11).Value = t.TR_Comment ?? "";
            if (t.Date_TR.HasValue) { ws.Cell(row, 12).Value = t.Date_TR.Value; ws.Cell(row, 12).Style.NumberFormat.Format = "yyyy-MM-dd"; }
            ws.Cell(row, 13).Value = t.TR_REV ?? "";
            ws.Cell(row, 14).Value = t.SAPMT_Comment ?? "";
            if (t.Date_SAPMT.HasValue) { ws.Cell(row, 15).Value = t.Date_SAPMT.Value; ws.Cell(row, 15).Style.NumberFormat.Format = "yyyy-MM-dd"; }
            ws.Cell(row, 16).Value = t.SAPMT_REV ?? "";
            ws.Cell(row, 17).Value = t.Categorize ?? "";
            ws.Cell(row, 18).Value = t.Discipline ?? "";
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
        var fileName = $"Transmittal_Log_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    // GET: /Home/DownloadTransmittalTemplate
    [SessionAuthorization]
    [HttpGet]
    public IActionResult DownloadTransmittalTemplate()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Transmittal Log");

        string[] headers = {
            "Transmittal", "Date", "Doc No", "eGesDoc No",
            "Documents Title", "REV", "Remarks", "Transmittal No",
            "TR Comment", "Date TR", "TR REV",
            "SAPMT Comment", "Date SAPMT", "SAPMT REV",
            "Categorize", "Discipline", "Project No"
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
            "Transmittal_Log_Template.xlsx");
    }

    // POST: /Home/ImportTransmittal
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImportTransmittal(IFormFile? excelFile, [FromForm] int? projectId)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            TempData["Msg"] = "Please select an Excel file to import.";
            return RedirectToAction(nameof(TransmittalLog));
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

            var headerRow = ws.Row(1);
            int colId = -1, colProjectNo = -1, colTransmittal = -1, colDate = -1, colDocNo = -1,
                colEGesDocNo = -1, colDocTitle = -1, colRev = -1, colRemarks = -1, colTransmittalNo = -1,
                colTrComment = -1, colDateTr = -1, colTrRev = -1, colSapmtComment = -1, colDateSapmt = -1,
                colSapmtRev = -1, colCategorize = -1, colDiscipline = -1;

            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            for (int c = 1; c <= lastCol; c++)
            {
                var h = (headerRow.Cell(c).GetString() ?? "").Trim().ToUpperInvariant();
                switch (h)
                {
                    case "ID": colId = c; break;
                    case "PROJECT NO" or "PROJECT_NO" or "TR_PROJECT_NO": colProjectNo = c; break;
                    case "TRANSMITTAL": colTransmittal = c; break;
                    case "DATE": colDate = c; break;
                    case "DOC NO" or "DOC_NO": colDocNo = c; break;
                    case "EGESDOC NO" or "EGESDOC_NO": colEGesDocNo = c; break;
                    case "DOCUMENTS TITLE" or "DOCUMENTS_TITLE": colDocTitle = c; break;
                    case "REV": colRev = c; break;
                    case "REMARKS": colRemarks = c; break;
                    case "TRANSMITTAL NO" or "TRANSMITTAL_NO": colTransmittalNo = c; break;
                    case "TR COMMENT" or "TR_COMMENT": colTrComment = c; break;
                    case "DATE TR" or "DATE_TR": colDateTr = c; break;
                    case "TR REV" or "TR_REV": colTrRev = c; break;
                    case "SAPMT COMMENT" or "SAPMT_COMMENT": colSapmtComment = c; break;
                    case "DATE SAPMT" or "DATE_SAPMT": colDateSapmt = c; break;
                    case "SAPMT REV" or "SAPMT_REV": colSapmtRev = c; break;
                    case "CATEGORIZE": colCategorize = c; break;
                    case "DISCIPLINE": colDiscipline = c; break;
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
                DateTime? GetDate(int col)
                {
                    if (col <= 0) return null;
                    var cell = ws.Cell(r, col);
                    if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
                    return DateTime.TryParse(cell.GetString(), out var parsed) ? parsed : null;
                }

                var transmittal = GetStr(colTransmittal);
                var docNo = GetStr(colDocNo);
                var docTitle = GetStr(colDocTitle);
                if (string.IsNullOrWhiteSpace(transmittal) && string.IsNullOrWhiteSpace(docNo) && string.IsNullOrWhiteSpace(docTitle))
                    continue;

                int? existingId = GetInt(colId);
                TransmittalLog? existing = null;
                if (existingId.HasValue && existingId.Value > 0)
                    existing = await _context.Transmittal_Log_tbl.FirstOrDefaultAsync(x => x.ID == existingId.Value);

                var tr = existing ?? new TransmittalLog();
                tr.TR_Project_No = GetInt(colProjectNo) ?? projectId;
                tr.Transmittal = TrNullIfEmpty(GetStr(colTransmittal), 30);
                tr.Date = GetDate(colDate);
                tr.Doc_No = TrNullIfEmpty(GetStr(colDocNo), 30);
                tr.eGesDoc_No = TrNullIfEmpty(GetStr(colEGesDocNo), 20);
                tr.Documents_Title = TrNullIfEmpty(GetStr(colDocTitle), 150);
                tr.REV = TrNullIfEmpty(GetStr(colRev), 3);
                tr.Remarks = TrNullIfEmpty(GetStr(colRemarks), 50);
                tr.Transmittal_No = TrNullIfEmpty(GetStr(colTransmittalNo), 30);
                tr.TR_Comment = TrNullIfEmpty(GetStr(colTrComment), 30);
                tr.Date_TR = GetDate(colDateTr);
                tr.TR_REV = TrNullIfEmpty(GetStr(colTrRev), 3);
                tr.SAPMT_Comment = TrNullIfEmpty(GetStr(colSapmtComment), 30);
                tr.Date_SAPMT = GetDate(colDateSapmt);
                tr.SAPMT_REV = TrNullIfEmpty(GetStr(colSapmtRev), 3);
                tr.Categorize = TrNullIfEmpty(GetStr(colCategorize), 50);
                tr.Discipline = TrNullIfEmpty(GetStr(colDiscipline), 25);
                tr.Transmittal_Updated_Date = AppClock.Now;
                tr.Transmittal_Updated_By = userId;

                if (existing != null)
                    updated++;
                else
                {
                    _context.Transmittal_Log_tbl.Add(tr);
                    imported++;
                }
            }

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Import complete. {imported} added, {updated} updated.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ImportTransmittal failed");
            var reason = ex.GetBaseException()?.Message ?? ex.Message;
            TempData["Msg"] = $"Import failed. {reason}";
        }

        return RedirectToAction(nameof(TransmittalLog), new { projectId });
    }

    // POST: /Home/BulkUploadTransmittalPdfs
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkUploadTransmittalPdfs([FromForm] int projectId, List<IFormFile>? pdfFiles)
    {
        if (pdfFiles == null || pdfFiles.Count == 0)
        {
            TempData["Msg"] = "Please select at least one PDF file.";
            return RedirectToAction(nameof(TransmittalLog), new { projectId });
        }

        try
        {
            var transmittals = await _context.Transmittal_Log_tbl
                .Where(t => t.TR_Project_No == projectId && t.Doc_No != null)
                .ToListAsync();

            var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files");
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

                var tr = transmittals.FirstOrDefault(t =>
                    !string.IsNullOrWhiteSpace(t.Doc_No) &&
                    fileNameWithoutExt.Contains(t.Doc_No.Trim(), StringComparison.OrdinalIgnoreCase));

                if (tr == null)
                {
                    unmatchedNames.Add(Path.GetFileName(file.FileName));
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(tr.TR_BlobName))
                {
                    try
                    {
                        var oldPath = Path.Combine(folder, Path.GetFileName(tr.TR_BlobName));
                        if (System.IO.File.Exists(oldPath)) System.IO.File.Delete(oldPath);
                    }
                    catch (Exception fileEx) { _logger.LogWarning(fileEx, "Failed to delete old PDF for Transmittal {Id}", tr.ID); }
                }

                var blobName = $"{Guid.NewGuid():N}.pdf";
                var savePath = Path.Combine(folder, blobName);
                using (var fs = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write))
                {
                    await file.CopyToAsync(fs);
                }

                tr.TR_FileName = Path.GetFileName(file.FileName);
                tr.TR_FileSize = (int)file.Length;
                tr.TR_UploadDate = AppClock.Now;
                tr.TR_BlobName = blobName;
                tr.Transmittal_Updated_Date = AppClock.Now;
                tr.Transmittal_Updated_By = userId;
                matched++;
            }

            await _context.SaveChangesAsync();

            var parts = new List<string>();
            if (matched > 0) parts.Add($"{matched} PDF(s) matched & uploaded");
            if (skipped > 0) parts.Add($"{skipped} non-PDF file(s) skipped");
            if (unmatchedNames.Count > 0) parts.Add($"{unmatchedNames.Count} file(s) had no Doc No match: {string.Join(", ", unmatchedNames)}");
            TempData["Msg"] = parts.Count > 0 ? string.Join(". ", parts) + "." : "No files processed.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "BulkUploadTransmittalPdfs failed for project {ProjectId}", projectId);
            TempData["Msg"] = $"Bulk PDF upload failed. {ex.GetBaseException()?.Message ?? ex.Message}";
        }

        return RedirectToAction(nameof(TransmittalLog), new { projectId });
    }

    // GET: /Home/BulkDownloadTransmittalPdfs?projectId=
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> BulkDownloadTransmittalPdfs([FromQuery] int projectId)
    {
        var transmittals = await _context.Transmittal_Log_tbl.AsNoTracking()
            .Where(t => t.TR_Project_No == projectId && t.TR_BlobName != null)
            .OrderBy(t => t.ID)
            .Select(t => new { t.ID, t.Doc_No, t.TR_FileName, t.TR_BlobName })
            .ToListAsync();

        if (transmittals.Count == 0)
        {
            TempData["Msg"] = "No PDF files found for this project.";
            return RedirectToAction(nameof(TransmittalLog), new { projectId });
        }

        var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "transmittal-files");

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var tr in transmittals)
            {
                if (string.IsNullOrWhiteSpace(tr.TR_BlobName)) continue;
                var filePath = Path.Combine(folder, Path.GetFileName(tr.TR_BlobName));
                if (!System.IO.File.Exists(filePath)) continue;

                var entryName = tr.TR_FileName ?? $"TR_{tr.ID}.pdf";
                var baseName = Path.GetFileNameWithoutExtension(entryName);
                var ext = Path.GetExtension(entryName);
                var candidate = entryName;
                int seq = 1;
                while (!usedNames.Add(candidate))
                {
                    candidate = $"{baseName}_{seq++}{ext}";
                }

                var entry = archive.CreateEntry(candidate, CompressionLevel.Fastest);
                using var entryStream = entry.Open();
                using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                await fileStream.CopyToAsync(entryStream);
            }
        }

        ms.Seek(0, SeekOrigin.Begin);
        return File(ms.ToArray(), "application/zip", $"Transmittal_PDFs_{projectId}_{AppClock.Now:yyyyMMdd}.zip");
    }

    /// <summary>
    /// Derives a default Discipline value from the user's Access field.
    /// </summary>
    private static string ResolveDisciplineFromAccess(string? access)
    {
        if (string.IsNullOrWhiteSpace(access)) return string.Empty;
        var a = access.Trim();
        if (a.Contains("QC", StringComparison.OrdinalIgnoreCase)) return "QC";
        if (a.Contains("Project", StringComparison.OrdinalIgnoreCase)) return "Project";
        if (a.Contains("Engineering", StringComparison.OrdinalIgnoreCase)) return "Engineering";
        if (a.Contains("Construction", StringComparison.OrdinalIgnoreCase)) return "Construction";
        return string.Empty;
    }

    private static string? TrNullIfEmpty(string? s, int? maxLength = null)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        return maxLength.HasValue && s.Length > maxLength.Value ? s[..maxLength.Value] : s;
    }
}
