using System;
using System.Collections.Generic;
using System.IO;
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
    // GET: /Home/TransmittalStatus
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> TransmittalStatus([FromQuery] int? projectId)
    {
        var allProjects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .Select(p => new { p.Project_ID, p.Project_Name })
            .ToListAsync();

        var defaultProjectId = await GetDefaultProjectIdAsync(projectId);

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

        ViewBag.DefaultProjectId = defaultProjectId;

        var statusGroups = new List<TransmittalStatusGroup>();

        if (defaultProjectId.HasValue)
        {
            // Replicate SP logic: latest revision per Doc_No for the selected project
            var allRows = await _context.Transmittal_Log_tbl
                .AsNoTracking()
                .Where(t => t.TR_Project_No == defaultProjectId.Value)
                .ToListAsync();

            var latestPerDoc = allRows
                .Where(t => !string.IsNullOrWhiteSpace(t.Doc_No))
                .GroupBy(t => t.Doc_No!.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(t => t.TR_REV).First())
                .ToList();

            // Also include rows with no Doc_No (same as SP includes them via ROW_NUMBER)
            var noDocRows = allRows
                .Where(t => string.IsNullOrWhiteSpace(t.Doc_No))
                .ToList();

            var rows = latestPerDoc.Concat(noDocRows).ToList();

            // Group by Categorize (mirrors VBA Column O split)
            var groups = rows
                .GroupBy(t => (t.Categorize ?? "Uncategorized").Trim(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Key)
                .ToList();

            foreach (var grp in groups)
            {
                var grpRows = grp.ToList();
                statusGroups.Add(BuildTransmittalStatusGroup(grp.Key, grpRows));
            }
        }

        ViewBag.StatusData = statusGroups;
        ViewBag.StatusDate = AppClock.Now.ToString("dd-MM-yy");

        var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
        var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
        var composed = (fn + " " + ln).Trim();
        if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
        ViewBag.FullName = composed;

        return View();
    }

    private static TransmittalStatusGroup BuildTransmittalStatusGroup(string category, List<TransmittalLog> rows)
    {
        // Total Submitted = count of rows with a non-blank Transmittal (mirrors COUNTA(A))
        var totalSubmitted = rows.Count(r => !string.IsNullOrWhiteSpace(r.Transmittal));

        // Received = count of rows with a non-blank Transmittal_No (mirrors COUNTA(H))
        var received = rows.Count(r => !string.IsNullOrWhiteSpace(r.Transmittal_No));

        // Pending = Total Submitted - Received
        var pending = totalSubmitted - received;
        if (pending < 0) pending = 0;

        // Status breakdown based on TR_Comment (mirrors COUNTIF on column I)
        var approved = rows.Count(r =>
            TrCommentEquals(r.TR_Comment, "REVIEWED WITHOUT COMMENTS")
            || TrCommentEquals(r.TR_Comment, "FOR INFORMATION"));

        var approvedWithComments = rows.Count(r =>
            TrCommentEquals(r.TR_Comment, "REVIEWED WITH COMMENTS")
            || TrCommentEquals(r.TR_Comment, "REVIEWED WITH MINOR COMMENTS")
            || TrCommentEquals(r.TR_Comment, "REVIEWED WITH MAJOR COMMENTS")
            || TrCommentEquals(r.TR_Comment, "COMMENTS AS NOTED"));

        var rejected = rows.Count(r =>
            TrCommentEquals(r.TR_Comment, "Rejected"));

        return new TransmittalStatusGroup
        {
            Category = category,
            Rows =
            [
                new() { Label = "Total Submitted", Value = totalSubmitted },
                new() { Label = "Pending", Value = pending },
                new() { Label = "Received", Value = received },
                new() { Label = "Approved", Value = approved, IsIndented = true },
                new() { Label = "Approved with Comments", Value = approvedWithComments, IsIndented = true },
                new() { Label = "Rejected", Value = rejected, IsIndented = true },
            ]
        };
    }

    private static bool TrCommentEquals(string? comment, string value)
        => string.Equals((comment ?? "").Trim(), value, StringComparison.OrdinalIgnoreCase);

    // GET: /Home/ExportTransmittalStatus?projectId=
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportTransmittalStatus([FromQuery] int? projectId)
    {
        if (!projectId.HasValue)
            return BadRequest("projectId is required.");

        // Replicate the same query logic used by the TransmittalStatus page
        var allRows = await _context.Transmittal_Log_tbl
            .AsNoTracking()
            .Where(t => t.TR_Project_No == projectId.Value)
            .ToListAsync();

        var latestPerDoc = allRows
            .Where(t => !string.IsNullOrWhiteSpace(t.Doc_No))
            .GroupBy(t => t.Doc_No!.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(t => t.TR_REV).First())
            .ToList();

        var noDocRows = allRows
            .Where(t => string.IsNullOrWhiteSpace(t.Doc_No))
            .ToList();

        var rows = latestPerDoc.Concat(noDocRows).ToList();

        // Group by Categorize (mirrors VBA Column O split)
        var groups = rows
            .GroupBy(t => (t.Categorize ?? "Uncategorized").Trim(), StringComparer.OrdinalIgnoreCase)
            .OrderBy(g => g.Key)
            .ToList();

        // Headers for each category sheet (Categorize column excluded)
        string[] headers = {
            "Transmittal", "Date", "Doc No", "eGesDoc No",
            "Documents Title", "REV", "Remarks", "Transmittal No",
            "TR Comment", "Date TR", "TR REV",
            "SAPMT Comment", "Date SAPMT", "SAPMT REV",
            "Discipline"
        };

        using var wb = new XLWorkbook();

        foreach (var grp in groups)
        {
            // Sheet name limited to 31 chars (Excel limit)
            var sheetName = grp.Key.Length > 31 ? grp.Key[..31] : grp.Key;
            var ws = wb.Worksheets.Add(sheetName);

            // Write headers
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            // Write data rows
            var grpRows = grp.ToList();
            int row = 2;
            foreach (var t in grpRows)
            {
                ws.Cell(row, 1).Value = t.Transmittal ?? "";
                if (t.Date.HasValue) { ws.Cell(row, 2).Value = t.Date.Value; ws.Cell(row, 2).Style.NumberFormat.Format = "yyyy-MM-dd"; }
                ws.Cell(row, 3).Value = t.Doc_No ?? "";
                ws.Cell(row, 4).Value = t.eGesDoc_No ?? "";
                ws.Cell(row, 5).Value = t.Documents_Title ?? "";
                ws.Cell(row, 6).Value = t.REV ?? "";
                ws.Cell(row, 7).Value = t.Remarks ?? "";
                ws.Cell(row, 8).Value = t.Transmittal_No ?? "";
                ws.Cell(row, 9).Value = t.TR_Comment ?? "";
                if (t.Date_TR.HasValue) { ws.Cell(row, 10).Value = t.Date_TR.Value; ws.Cell(row, 10).Style.NumberFormat.Format = "yyyy-MM-dd"; }
                ws.Cell(row, 11).Value = t.TR_REV ?? "";
                ws.Cell(row, 12).Value = t.SAPMT_Comment ?? "";
                if (t.Date_SAPMT.HasValue) { ws.Cell(row, 13).Value = t.Date_SAPMT.Value; ws.Cell(row, 13).Style.NumberFormat.Format = "yyyy-MM-dd"; }
                ws.Cell(row, 14).Value = t.SAPMT_REV ?? "";
                ws.Cell(row, 15).Value = t.Discipline ?? "";
                ws.Row(row).Height = 17;
                row++;
            }

            int lastDataRow = row - 1;

            // Create table with the same theme as Transmittal Log export
            if (lastDataRow >= 2)
            {
                var dataRange = ws.Range(1, 1, lastDataRow, headers.Length);
                var table = dataRange.CreateTable($"TransmittalTable_{sheetName.Replace(" ", "_")}");
                table.Theme = XLTableTheme.TableStyleMedium2;
                table.ShowTotalsRow = false;
            }

            // --- Status summary table (mirrors VBA AddStatusTable) ---
            var statusGroup = BuildTransmittalStatusGroup(grp.Key, grpRows);
            int statusStartRow = lastDataRow + 3;

            // Status header row
            string[] statusHeaders = { "Submitted / Responded", "Status", "Remarks" };
            for (int i = 0; i < statusHeaders.Length; i++)
                ws.Cell(statusStartRow, i + 1).Value = statusHeaders[i];

            // Status data rows
            int statusRow = statusStartRow + 1;
            foreach (var sr in statusGroup.Rows)
            {
                var label = sr.IsIndented ? $"  – {sr.Label}" : sr.Label;
                ws.Cell(statusRow, 1).Value = label;
                ws.Cell(statusRow, 2).Value = sr.Value;
                ws.Cell(statusRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(statusRow, 3).Value = sr.Remarks ?? "";
                ws.Row(statusRow).Height = 17;
                statusRow++;
            }

            int statusEndRow = statusRow - 1;

            // Create status table with the same theme
            if (statusEndRow >= statusStartRow + 1)
            {
                var statusRange = ws.Range(statusStartRow, 1, statusEndRow, 3);
                var statusTable = statusRange.CreateTable($"StatusTable_{sheetName.Replace(" ", "_")}");
                statusTable.Theme = XLTableTheme.TableStyleMedium2;
                statusTable.ShowTotalsRow = false;
            }

            ws.Row(1).Height = 30;
            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);
        }

        // If no groups, create one empty sheet
        if (groups.Count == 0)
        {
            var ws = wb.Worksheets.Add("No Data");
            ws.Cell(1, 1).Value = "No transmittal data found for this project.";
        }

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"Transmittal_Status_{projectId}_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}
