using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
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
    // GET: /Home/WelderList
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> WelderList([FromQuery] int? projectId)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");
        ViewBag.FullName = fullName;

        var pid = await GetDefaultProjectIdAsync(projectId);
        ViewBag.ProjectId = pid;

        // Build the project dropdown: Welders_Project_ID - Project_Name (grouped/distinct, matches Control Lookups)
        var projList = await _context.Projects_tbl.AsNoTracking()
            .Where(p => p.Welders_Project_ID != null)
            .GroupBy(p => p.Welders_Project_ID)
            .Select(g => new
            {
                Welders_Project_ID = g.Key,
                Project_Name = g.OrderBy(p => p.Project_ID).First().Project_Name,
                Default_P = g.Any(p => p.Default_P)
            })
            .OrderBy(p => p.Welders_Project_ID)
            .ToListAsync();

        if (!pid.HasValue && projList.Count > 0)
        {
            var fallback = projList.FirstOrDefault(p => p.Default_P) ?? projList.Last();
            pid = fallback?.Welders_Project_ID;
            ViewBag.ProjectId = pid;
        }
        else if (pid.HasValue)
        {
            // Map a Project_ID to its Welders_Project_ID when coming from session/query
            var mapped = await _context.Projects_tbl.AsNoTracking()
                .Where(p => p.Project_ID == pid.Value)
                .Select(p => p.Welders_Project_ID)
                .FirstOrDefaultAsync();
            if (mapped.HasValue) { pid = mapped.Value; ViewBag.ProjectId = pid; }
        }

        ViewBag.ProjectsSelect = projList.Select(p => new
        {
            Id = p.Welders_Project_ID!.Value,
            Label = string.Concat(p.Welders_Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty)
        }).ToList();
        ViewBag.DefaultProjectId = pid;

        const string welderListSql = @"
SELECT
    w.Welder_ID,
    ISNULL(w.Welder_Symbol, '') AS Welder_Symbol,
    ISNULL(w.Name, '') AS Name,
    ISNULL(w.Welder_Location, '') AS Welder_Location,
    w.Mobilization_Date,
    w.Demobilization_Date,
    ISNULL(w.Status, '') AS Status,
    w.Project_Welder,
    STRING_AGG(NULLIF(LTRIM(RTRIM(q.Welding_Process)), ''), ', ') AS Welding_Process,
    STRING_AGG(CASE WHEN q.Batch_No IS NULL THEN NULL ELSE LTRIM(RTRIM(CONVERT(varchar(16), q.Batch_No))) END, ', ') AS Batch_Nos,
    MIN(q.Batch_No) AS FirstBatch,
    MIN(DATEADD(month, 6, COALESCE(q.DATE_OF_LAST_CONTINUITY, q.Test_Date))) AS NextContinuity
FROM dbo.Welders_tbl w
LEFT JOIN dbo.Welder_List_tbl q ON q.Welder_ID_WL = w.Welder_ID
WHERE w.Welder_Symbol IS NOT NULL AND w.Name IS NOT NULL
  AND (@ProjectId IS NULL OR LTRIM(RTRIM(CAST(w.Project_Welder AS varchar(32)))) = LTRIM(RTRIM(CAST(@ProjectId AS varchar(32)))))
GROUP BY w.Welder_ID, w.Welder_Symbol, w.Name, w.Welder_Location, w.Mobilization_Date, w.Demobilization_Date, w.Status, w.Project_Welder
ORDER BY w.Welder_Symbol;";

        var rows = new List<WelderListRow>();
        await using var conn = _context.Database.GetDbConnection();
        if (conn.State != ConnectionState.Open)
            await conn.OpenAsync();

        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = welderListSql;
            var pidParam = cmd.CreateParameter();
            pidParam.ParameterName = "@ProjectId";
            pidParam.Value = pid.HasValue ? pid.Value : DBNull.Value;
            pidParam.DbType = DbType.Int32;
            cmd.Parameters.Add(pidParam);

            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                rows.Add(new WelderListRow
                {
                    Welder_ID = reader.GetInt32(0),
                    Welder_Symbol = reader.GetString(1),
                    Name = reader.GetString(2),
                    Welder_Location = reader.GetString(3),
                    Mobilization_Date = reader.IsDBNull(4) ? null : reader.GetDateTime(4),
                    Demobilization_Date = reader.IsDBNull(5) ? null : reader.GetDateTime(5),
                    Status = reader.GetString(6),
                    Project_Welder = reader.IsDBNull(7) ? null : Convert.ToInt32(reader.GetValue(7)).ToString(),
                    Welding_Process = reader.IsDBNull(8) ? string.Empty : reader.GetString(8),
                    Batch_Nos = reader.IsDBNull(9) ? string.Empty : reader.GetString(9),
                    Batch_No = reader.IsDBNull(10) ? (int?)null : reader.GetInt32(10),
                    Next_Continuity = reader.IsDBNull(11) ? (DateTime?)null : reader.GetDateTime(11)
                });
            }
        }

        return View(rows);
    }

    // Optional: used by WelderList delete button
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteWelder([FromForm] int id)
    {
        try
        {
            var welder = await _context.Welders_tbl.FirstOrDefaultAsync(w => w.Welder_ID == id);
            if (welder == null) return NotFound();

            var qualsToDelete = await _context.Welder_List_tbl
                .Where(q => q.Welder_ID_WL == id)
                .ToListAsync();
            if (qualsToDelete.Count > 0)
            {
                _context.Welder_List_tbl.RemoveRange(qualsToDelete);
            }

            string symbol = welder.Welder_Symbol ?? string.Empty;
            _context.Welders_tbl.Remove(welder);
            await _context.SaveChangesAsync();

            TempData["Msg"] = string.IsNullOrWhiteSpace(symbol)
                ? $"Deleted welder ID {id}."
                : $"Deleted welder Symbol {symbol}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting welder {Id}", id);
            TempData["Msg"] = "Failed to delete welder.";
        }
        return RedirectToAction(nameof(WelderList));
    }

    // GET: /Home/ExportWelders  (Excel export from stored procedure "PMS_Welders_List_Q")
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportWelders([FromQuery] int? projectId, [FromQuery] string? q, [FromQuery] string? batch, [FromQuery] bool due = false)
    {
        var dt = new DataTable();
        try
        {
            var conn = _context.Database.GetDbConnection();
            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "[PMS_Welders_List_Q]";
            cmd.CommandType = CommandType.StoredProcedure;
            using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportWelders: failed executing stored procedure");
            return Content("Failed to run stored procedure for welder export.");
        }

        if (projectId.HasValue && dt.Rows.Count > 0)
        {
            var pidText = projectId.Value.ToString();
            var projectCol = dt.Columns.Cast<DataColumn>()
                .FirstOrDefault(c => c.ColumnName.Contains("project", StringComparison.OrdinalIgnoreCase));
            if (projectCol != null)
            {
                var filtered = dt.Clone();
                foreach (DataRow row in dt.Rows)
                {
                    var raw = row[projectCol]?.ToString()?.Trim() ?? string.Empty;
                    var matches = int.TryParse(raw, out var rowPid)
                        ? rowPid == projectId.Value
                        : raw.StartsWith(pidText, StringComparison.OrdinalIgnoreCase);
                    if (matches)
                        filtered.ImportRow(row);
                }
                dt = filtered;
            }
        }

        var hasSearch = !string.IsNullOrWhiteSpace(q);
        string[] searchTerms = hasSearch
            ? q!.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            : [];

        var batchTrim = batch?.Trim() ?? string.Empty;

        if (dt.Rows.Count > 0 && (searchTerms.Length > 0 || batchTrim.Length > 0))
        {
            var filtered = dt.Clone();
            foreach (DataRow row in dt.Rows)
            {
                bool ok = true;
                if (ok && searchTerms.Length > 0)
                {
                    foreach (var term in searchTerms)
                    {
                        bool foundTerm = false;
                        foreach (DataColumn col in dt.Columns)
                        {
                            var v = row[col];
                            if (v != null && v != DBNull.Value)
                            {
                                var text = v.ToString();
                                if (!string.IsNullOrEmpty(text) && text.Contains(term, StringComparison.OrdinalIgnoreCase))
                                {
                                    foundTerm = true;
                                    break;
                                }
                            }
                        }
                        if (!foundTerm) { ok = false; break; }
                    }
                }
                if (ok && batchTrim.Length > 0 && dt.Columns.Contains("Batch No"))
                {
                    var raw = row["Batch No"]?.ToString() ?? string.Empty;
                    var parts = raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                    ok = parts.Any(p => string.Equals(p, batchTrim, StringComparison.OrdinalIgnoreCase));
                }
                if (ok) filtered.ImportRow(row);
            }
            dt = filtered;
        }

        if (due && dt.Rows.Count > 0 && dt.Columns.Contains("Welder Symbol"))
        {
            var symbols = dt.AsEnumerable()
                .Select(r => r["Welder Symbol"]?.ToString()?.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var qualDates = await _context.Welder_List_tbl
                .AsNoTracking()
                .Where(qr => qr.Welder != null && symbols.Contains(qr.Welder!.Welder_Symbol))
                .Select(qr => new { qr.Welder!.Welder_Symbol, qr.DATE_OF_LAST_CONTINUITY, qr.Test_Date })
                .ToListAsync();

            var now = AppClock.Now.Date;
            var dueSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var grouped = qualDates.GroupBy(x => x.Welder_Symbol);
            foreach (var g in grouped)
            {
                DateTime? nextContinuity = null;
                foreach (var item in g)
                {
                    var baseDate = item.DATE_OF_LAST_CONTINUITY ?? item.Test_Date;
                    if (baseDate.HasValue)
                    {
                        var cand = baseDate.Value.AddMonths(6).Date;
                        if (nextContinuity == null || cand < nextContinuity) nextContinuity = cand;
                    }
                }
                if (nextContinuity.HasValue)
                {
                    var diff = (nextContinuity.Value - now).TotalDays;
                    if (diff <= 10)
                        dueSet.Add(g.Key);
                }
            }

            if (dueSet.Count > 0)
            {
                var filtered = dt.Clone();
                foreach (DataRow row in dt.Rows)
                {
                    var sym = row["Welder Symbol"]?.ToString()?.Trim();
                    if (sym != null && dueSet.Contains(sym))
                        filtered.ImportRow(row);
                }
                dt = filtered;
            }
            else
            {
                dt = dt.Clone();
            }
        }

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Welder Register");

        if (dt.Columns.Count == 0)
        {
            ws.Cell(1, 1).SetValue("No data");
        }
        else
        {
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                ws.Cell(1, c + 1).SetValue(dt.Columns[c].ColumnName);
            }
            int row = 2;
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dr[c];
                    var cell = ws.Cell(row, c + 1);
                    if (val == DBNull.Value || val == null)
                    {
                        cell.SetValue(string.Empty);
                    }
                    else if (val is DateTime dtVal)
                    {
                        cell.SetValue(dtVal);
                        cell.Style.DateFormat.Format = "yyyy-MM-dd";
                    }
                    else if (val is int i)
                    {
                        cell.SetValue(i);
                    }
                    else if (val is long l)
                    {
                        cell.SetValue(l);
                    }
                    else if (val is short s)
                    {
                        cell.SetValue((int)s);
                    }
                    else if (val is byte b8)
                    {
                        cell.SetValue((int)b8);
                    }
                    else if (val is double d)
                    {
                        cell.SetValue(d);
                    }
                    else if (val is float f)
                    {
                        cell.SetValue((double)f);
                    }
                    else if (val is decimal dec)
                    {
                        cell.SetValue((double)dec);
                    }
                    else if (val is bool bl)
                    {
                        cell.SetValue(bl);
                    }
                    else
                    {
                        cell.SetValue(val.ToString() ?? string.Empty);
                    }
                }
                ws.Row(row).Height = 17;
                row++;
            }
            int lastRow = dt.Rows.Count + 1;
            var fullRange = ws.Range(1, 1, lastRow, dt.Columns.Count);
            var table = fullRange.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowTotalsRow = false;
            ws.Row(1).Height = 30;

            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var name = dt.Columns[c].ColumnName.ToLowerInvariant();
                var col = ws.Column(c + 1);
                var type = dt.Columns[c].DataType;
                if (name.Contains("id") || name.Contains("no") || name.Contains("batch") || name.Contains("project"))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                else if (type == typeof(int) || type == typeof(long))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                else if (type == typeof(double) || type == typeof(decimal) || type == typeof(float))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms2 = new MemoryStream();
        wb.SaveAs(ms2);
        var bytes2 = ms2.ToArray();
        var fileName2 = $"WelderRegister_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName2);
    }

    // INSERTED: Download all qualification files for a welder symbol as a zip
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DownloadWelderQualifications([FromQuery] string symbol)
    {
        try
        {
            var symbolClean = StripWhitespace((symbol ?? string.Empty).Trim());
            if (string.IsNullOrWhiteSpace(symbolClean)) return BadRequest("Symbol required");

            var welder = await _context.Welders_tbl
                .AsNoTracking()
                .FirstOrDefaultAsync(w => w.Welder_Symbol != null && EF.Functions.Collate(w.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == symbolClean);
            if (welder == null) return NotFound("Welder not found");

            var quals = await _context.Welder_List_tbl
                .AsNoTracking()
                .Where(q => q.Welder_ID_WL == welder.Welder_ID && !string.IsNullOrWhiteSpace(q.JCC_BlobName))
                .Select(q => new { q.JCC_No, q.JCC_BlobName, q.JCC_FileName })
                .ToListAsync();
            if (quals.Count == 0) return NotFound("No qualification files");

            var root = Path.Combine(_env.ContentRootPath, "App_Data", "WelderQualifications");
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);

            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int added = 0;
                foreach (var q in quals)
                {
                    try
                    {
                        var blob = q.JCC_BlobName!.Trim();
                        var fullPath = Path.Combine(root, blob);
                        if (!System.IO.File.Exists(fullPath))
                        {
                            _logger.LogWarning("Qualification blob missing on disk: {Path}", fullPath);
                            continue;
                        }

                        var ext = Path.GetExtension(fullPath);
                        var desiredNameRaw = (q.JCC_FileName ?? string.Empty).Trim();
                        var desiredName = string.IsNullOrWhiteSpace(desiredNameRaw) ? ($"{q.JCC_No}{ext}") : desiredNameRaw;

                        desiredName = desiredName.Replace('\\', '_').Replace('/', '_');
                        if (desiredName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                        {
                            var clean = new string(desiredName.Where(c => !Path.GetInvalidFileNameChars().Contains(c)).ToArray());
                            desiredName = string.IsNullOrWhiteSpace(clean) ? Path.GetFileName(blob) : clean;
                        }

                        var entryName = desiredName;
                        int dup = 1;
                        while (usedNames.Contains(entryName))
                        {
                            entryName = Path.GetFileNameWithoutExtension(desiredName) + $"_{dup}" + Path.GetExtension(desiredName);
                            dup++;
                        }
                        usedNames.Add(entryName);

                        var entry = zip.CreateEntry(entryName, CompressionLevel.Fastest);
                        using var entryStream = entry.Open();
                        await using var fileStream = System.IO.File.OpenRead(fullPath);
                        await fileStream.CopyToAsync(entryStream);
                        added++;
                    }
                    catch (Exception inner)
                    {
                        _logger.LogWarning(inner, "Skipping qualification file for JCC {Jcc} due to error", q.JCC_No);
                    }
                }

                if (added == 0)
                {
                    return NotFound("Qualification files not found on disk");
                }
            }

            ms.Seek(0, SeekOrigin.Begin);
            var fileName = $"Welder_{symbolClean}_Qualifications_{AppClock.Now:yyyyMMdd_HHmmss}.zip";
            return File(ms.ToArray(), "application/zip", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DownloadWelderQualifications failed for symbol {Symbol}", symbol);
            return StatusCode(500, "Error building archive");
        }
    }
}
