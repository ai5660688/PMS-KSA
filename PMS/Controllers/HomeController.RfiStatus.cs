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
    // GET: /Home/RfiStatus
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> RfiStatus([FromQuery] int? projectId)
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

        var list = await query.ToListAsync();

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

        var now = AppClock.Now;

        var statusData = new List<RfiStatusGroup>();

        if (list.Count > 0)
        {
            var shopAll = list.Where(r =>
                (r.RFI_LOCATION ?? "").IndexOf("Shop", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            var siteAll = list.Where(r =>
                (r.RFI_LOCATION ?? "").IndexOf("Shop", StringComparison.OrdinalIgnoreCase) < 0).ToList();

            if (shopAll.Count > 0)
                statusData.Add(BuildRfiStatusGroup("Sub Contractor", "Shop", shopAll, now));
            if (siteAll.Count > 0)
                statusData.Add(BuildRfiStatusGroup("Sub Contractor", "Field", siteAll, now));
        }

        ViewBag.StatusData = statusData;
        ViewBag.StatusDate = now.ToString("dd-MM-yy");

        var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
        var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
        var composed = (fn + " " + ln).Trim();
        if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
        ViewBag.FullName = composed;

        return View();
    }

    private static RfiStatusGroup BuildRfiStatusGroup(string contractor, string locationType, List<Rfi> rows, DateTime now)
    {
        var disciplines = rows
            .Where(r => !string.IsNullOrWhiteSpace(r.SubDiscipline))
            .Select(r => r.SubDiscipline!.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(d => d)
            .ToList();

        var group = new RfiStatusGroup
        {
            Contractor = contractor,
            LocationType = locationType,
            Rows = []
        };

        foreach (var disc in disciplines)
        {
            var discRows = rows.Where(r =>
                string.Equals(r.SubDiscipline?.Trim(), disc, StringComparison.OrdinalIgnoreCase)).ToList();
            group.Rows.Add(BuildRfiStatusRow(disc, discRows, now));
        }

        group.Total = new RfiStatusRow
        {
            Discipline = "TOTAL",
            Raised = group.Rows.Sum(r => r.Raised),
            Closed = group.Rows.Sum(r => r.Closed),
            Cancelled = group.Rows.Sum(r => r.Cancelled),
            UnderReviewInspDate = group.Rows.Sum(r => r.UnderReviewInspDate),
            UnderReviewLess72 = group.Rows.Sum(r => r.UnderReviewLess72),
            UnderReviewMore72 = group.Rows.Sum(r => r.UnderReviewMore72),
            UnderQcPidSign = group.Rows.Sum(r => r.UnderQcPidSign),
            PmtSignBalance = group.Rows.Sum(r => r.PmtSignBalance),
            UnderScanningWithPmtSign = group.Rows.Sum(r => r.UnderScanningWithPmtSign),
            TotalOpen = group.Rows.Sum(r => r.TotalOpen)
        };

        return group;
    }

    private static RfiStatusRow BuildRfiStatusRow(string discipline, List<Rfi> rows, DateTime now)
    {
        var raised = rows.Count;

        // "Closed" — also accept legacy "Closed"
        var closed = rows.Count(r =>
            RfiStatusEquals(r.INSPECTION_STATUS, "Closed")
            || RfiStatusEquals(r.INSPECTION_STATUS, "Closed"));

        var cancelled = rows.Count(r =>
            RfiStatusEquals(r.INSPECTION_STATUS, "Cancelled"));

        // Open RFIs = all RFIs that are not Closed and not Cancelled
        var openRfis = rows.Where(r =>
            !RfiStatusEquals(r.INSPECTION_STATUS, "Closed")
            && !RfiStatusEquals(r.INSPECTION_STATUS, "Cancelled")).ToList();

        int inspDate = 0, less72 = 0, more72 = 0;
        foreach (var r in openRfis)
        {
            if (r.Date == null)
            {
                inspDate++;
            }
            else if (r.Date.Value.Date == now.Date)
            {
                inspDate++;
            }
            else
            {
                var hours = (now - r.Date.Value).TotalHours;
                if (hours <= 0) inspDate++; // future date
                else if (hours <= 72) less72++;
                else more72++;
            }
        }

        // "Under TR QC / PID SIGN" — contains "/ PID SIGN" catches both legacy & new
        var underQc = rows.Count(r =>
            RfiStatusContains(r.INSPECTION_STATUS, "/ PID SIGN"));

        var pmtSign = rows.Count(r =>
            RfiStatusEquals(r.INSPECTION_STATUS, "PMT SIGN BALANCE"));

        // "Under Scanning" — also accept legacy catch-all "Under *"
        // (starts with "Under" but not already counted by the contractor or QC buckets)
        var underScanning = rows.Count(r =>
        {
            if (RfiStatusEquals(r.INSPECTION_STATUS, "Under Scanning"))
                return true;
            // Legacy: any "Under ..." not matched by the other buckets
            var s = (r.INSPECTION_STATUS ?? "").Trim();
            return s.StartsWith("Under", StringComparison.OrdinalIgnoreCase)
                && !RfiStatusEquals(r.INSPECTION_STATUS, "Under Contractor")
                && !RfiStatusEquals(r.INSPECTION_STATUS, "UNDER REVIEW")
                && !RfiStatusContains(r.INSPECTION_STATUS, "/ PID SIGN");
        });

        var totalOpen = raised - closed - cancelled;
        if (totalOpen < 0) totalOpen = 0;

        return new RfiStatusRow
        {
            Discipline = discipline,
            Raised = raised,
            Closed = closed,
            Cancelled = cancelled,
            UnderReviewInspDate = inspDate,
            UnderReviewLess72 = less72,
            UnderReviewMore72 = more72,
            UnderQcPidSign = underQc,
            PmtSignBalance = pmtSign,
            UnderScanningWithPmtSign = underScanning,
            TotalOpen = totalOpen
        };
    }

    private static bool RfiStatusEquals(string? status, string value)
        => string.Equals((status ?? "").Trim(), value, StringComparison.OrdinalIgnoreCase);

    private static bool RfiStatusContains(string? status, string value)
        => (status ?? "").IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;

    // GET: /Home/ExportRfiStatus?projectId=
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportRfiStatus([FromQuery] int? projectId)
    {
        var query = _context.RFI_tbl.AsNoTracking();
        if (projectId.HasValue)
            query = query.Where(r => r.RFI_Project_No == projectId.Value);

        var list = await query.OrderBy(r => r.RFI_ID).ToListAsync();
        var now = AppClock.Now;

        var shopAll = list.Where(r =>
            (r.RFI_LOCATION ?? "").IndexOf("Shop", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
        var fieldAll = list.Where(r =>
            (r.RFI_LOCATION ?? "").IndexOf("Shop", StringComparison.OrdinalIgnoreCase) < 0).ToList();

        using var wb = new XLWorkbook();

        // --- Shop status sheet ---
        if (shopAll.Count > 0)
        {
            var shopGroup = BuildRfiStatusGroup("Sub Contractor", "Shop", shopAll, now);
            WriteRfiStatusSheet(wb, "Shop", shopGroup);
        }

        // --- Field status sheet ---
        if (fieldAll.Count > 0)
        {
            var fieldGroup = BuildRfiStatusGroup("Sub Contractor", "Field", fieldAll, now);
            WriteRfiStatusSheet(wb, "Field", fieldGroup);
        }

        // --- RFI Log sheets distributed by Shop/Field + Sub-Discipline ---
        var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Reserve the status sheet names already added
        foreach (var ws in wb.Worksheets)
            usedSheetNames.Add(ws.Name);

        foreach (var (locLabel, locRows) in new[] { ("Shop", shopAll), ("Field", fieldAll) })
        {
            var byDisc = locRows
                .GroupBy(r => string.IsNullOrWhiteSpace(r.SubDiscipline) ? "(Other)" : r.SubDiscipline.Trim(),
                         StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Key)
                .ToList();

            foreach (var discGroup in byDisc)
            {
                var sheetName = SanitizeSheetName($"{locLabel} - {discGroup.Key}", usedSheetNames);
                usedSheetNames.Add(sheetName);
                WriteRfiLogSheet(wb, sheetName, discGroup.ToList());
            }
        }

        if (wb.Worksheets.Count == 0)
            wb.Worksheets.Add("Empty").Cell(1, 1).Value = "No RFI data found.";

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"RFI_Status_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    private static void WriteRfiStatusSheet(XLWorkbook wb, string sheetName, RfiStatusGroup group)
    {
        var ws = wb.Worksheets.Add(sheetName);

        // Title row
        ws.Cell(1, 1).Value = $"RFI STATUS — {sheetName.ToUpperInvariant()}";
        ws.Range(1, 1, 1, 12).Merge();
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;
        ws.Cell(1, 1).Style.Font.FontColor = XLColor.White;
        ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#00bfff");
        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(1).Height = 30;

        // Header row 1
        string[] headers = {
            "Sr No.", "Sub-Discipline", "Raised RFIs", "Closed", "Cancelled",
            "Under Review — Inspection Date", "Under Review — Less than 72HRS", "Under Review — More than 72HRS",
            "Under Contractor / PID SIGN", "PMT SIGN BALANCE", "Under Scanning", "TOTAL OPEN RFIs"
        };

        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(2, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#e6f6fa");
            cell.Style.Font.FontColor = XLColor.FromHtml("#176d8a");
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#cce7ef");
        }
        // Highlight "Under Review" header columns
        for (int c = 6; c <= 8; c++)
        {
            ws.Cell(2, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#ffc107");
            ws.Cell(2, c).Style.Font.FontColor = XLColor.FromHtml("#333333");
        }
        ws.Row(2).Height = 36;

        // Data rows
        int row = 3;
        int sr = 1;
        foreach (var r in group.Rows)
        {
            ws.Cell(row, 1).Value = sr++;
            ws.Cell(row, 2).Value = r.Discipline;
            ws.Cell(row, 3).Value = r.Raised;
            ws.Cell(row, 4).Value = r.Closed;
            ws.Cell(row, 5).Value = r.Cancelled;
            ws.Cell(row, 6).Value = r.UnderReviewInspDate;
            ws.Cell(row, 7).Value = r.UnderReviewLess72;
            ws.Cell(row, 8).Value = r.UnderReviewMore72;
            ws.Cell(row, 9).Value = r.UnderQcPidSign;
            ws.Cell(row, 10).Value = r.PmtSignBalance;
            ws.Cell(row, 11).Value = r.UnderScanningWithPmtSign;
            ws.Cell(row, 12).Value = r.TotalOpen;

            for (int c = 1; c <= 12; c++)
            {
                ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(row, c).Style.Border.OutsideBorderColor = XLColor.FromHtml("#cce7ef");
                ws.Cell(row, c).Style.Alignment.Horizontal = c == 2
                    ? XLAlignmentHorizontalValues.Left
                    : XLAlignmentHorizontalValues.Center;
            }
            // Highlight TOTAL OPEN column
            ws.Cell(row, 12).Style.Fill.BackgroundColor = XLColor.FromHtml("#ffe600");
            ws.Cell(row, 12).Style.Font.Bold = true;
            ws.Row(row).Height = 20;
            row++;
        }

        // Total row
        ws.Cell(row, 1).Value = "";
        ws.Cell(row, 2).Value = "TOTAL";
        ws.Cell(row, 3).Value = group.Total.Raised;
        ws.Cell(row, 4).Value = group.Total.Closed;
        ws.Cell(row, 5).Value = group.Total.Cancelled;
        ws.Cell(row, 6).Value = group.Total.UnderReviewInspDate;
        ws.Cell(row, 7).Value = group.Total.UnderReviewLess72;
        ws.Cell(row, 8).Value = group.Total.UnderReviewMore72;
        ws.Cell(row, 9).Value = group.Total.UnderQcPidSign;
        ws.Cell(row, 10).Value = group.Total.PmtSignBalance;
        ws.Cell(row, 11).Value = group.Total.UnderScanningWithPmtSign;
        ws.Cell(row, 12).Value = group.Total.TotalOpen;

        for (int c = 1; c <= 12; c++)
        {
            ws.Cell(row, c).Style.Font.Bold = true;
            ws.Cell(row, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#e6f6fa");
            ws.Cell(row, c).Style.Border.TopBorder = XLBorderStyleValues.Medium;
            ws.Cell(row, c).Style.Border.TopBorderColor = XLColor.FromHtml("#176d8a");
            ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(row, c).Style.Border.OutsideBorderColor = XLColor.FromHtml("#cce7ef");
            ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        ws.Cell(row, 12).Style.Fill.BackgroundColor = XLColor.FromHtml("#ff4444");
        ws.Cell(row, 12).Style.Font.FontColor = XLColor.White;
        ws.Cell(row, 12).Style.Font.Bold = true;
        ws.Row(row).Height = 24;

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(2);
    }

    private static void WriteRfiLogSheet(XLWorkbook wb, string sheetName, List<Rfi> rows)
    {
        var ws = wb.Worksheets.Add(sheetName);

        string[] headers = {
            "RFI ID", "Discipline", "Sub Contractor", "Sub Discipline",
            "RFI Contractor", "SubCon RFI No", "EPM No", "SATIP", "SAIC",
            "Activity", "Location", "Unit", "Element", "Description",
            "Date", "Time", "Company Insp Level", "Contractor Insp Level", "SubCon Insp Level",
            "TR QC", "Sub Con QC", "PID", "PMT", "Inspection Status",
            "Ref Drawing No", "Scan Copy", "Remarks"
        };

        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#e6f6fa");
            cell.Style.Font.FontColor = XLColor.FromHtml("#176d8a");
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#cce7ef");
        }
        ws.Row(1).Height = 30;

        int row = 2;
        foreach (var r in rows)
        {
            ws.Cell(row, 1).Value = r.RFI_ID;
            ws.Cell(row, 2).Value = r.DISCIPLINE ?? "";
            ws.Cell(row, 3).Value = r.Sub_Contractor ?? "";
            ws.Cell(row, 4).Value = r.SubDiscipline ?? "";
            ws.Cell(row, 5).Value = r.RFI_CONTRACTOR ?? "";
            ws.Cell(row, 6).Value = r.SubCon_RFI_No ?? "";
            ws.Cell(row, 7).Value = r.EPMNo ?? 0;
            ws.Cell(row, 8).Value = r.SATIP ?? "";
            ws.Cell(row, 9).Value = r.SAIC ?? "";
            ws.Cell(row, 10).Value = r.ACTIVITY ?? 0;
            ws.Cell(row, 11).Value = r.RFI_LOCATION ?? "";
            ws.Cell(row, 12).Value = r.UNIT ?? "";
            ws.Cell(row, 13).Value = r.ELEMENT ?? "";
            ws.Cell(row, 14).Value = r.RFI_DESCRIPTION ?? "";
            if (r.Date.HasValue)
            {
                ws.Cell(row, 15).Value = r.Date.Value;
                ws.Cell(row, 15).Style.NumberFormat.Format = "yyyy-MM-dd";
            }
            if (r.Time.HasValue)
            {
                ws.Cell(row, 16).Value = r.Time.Value.ToString("HH:mm");
            }
            ws.Cell(row, 17).Value = r.COMPANY_INSPECTION_LEVEL ?? "";
            ws.Cell(row, 18).Value = r.CONTRACTOR_INSPECTION_LEVEL ?? "";
            ws.Cell(row, 19).Value = r.SUBCON_INSPECTION_LEVEL ?? "";
            ws.Cell(row, 20).Value = r.TR_QC ?? "";
            ws.Cell(row, 21).Value = r.SUB_CON_QC ?? "";
            ws.Cell(row, 22).Value = r.PID ?? "";
            ws.Cell(row, 23).Value = r.PMT ?? "";
            ws.Cell(row, 24).Value = r.INSPECTION_STATUS ?? "";
            ws.Cell(row, 25).Value = r.REFRENCE_DRAWING_No ?? "";
            ws.Cell(row, 26).Value = r.SCAN_COPY ?? "";
            ws.Cell(row, 27).Value = r.REMARKS ?? "";

            for (int c = 1; c <= headers.Length; c++)
            {
                ws.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(row, c).Style.Border.OutsideBorderColor = XLColor.FromHtml("#cce7ef");
            }
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

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);
    }

    private static string SanitizeSheetName(string name, HashSet<string> usedNames)
    {
        // Excel sheet names: max 31 chars, no \ / ? * [ ]
        var sanitized = name
            .Replace("\\", "").Replace("/", "").Replace("?", "")
            .Replace("*", "").Replace("[", "").Replace("]", "");
        if (sanitized.Length > 31) sanitized = sanitized[..31];
        sanitized = sanitized.TrimEnd('\'');
        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "Sheet";

        var candidate = sanitized;
        int seq = 1;
        while (usedNames.Contains(candidate))
        {
            var suffix = $"_{seq++}";
            var maxBase = 31 - suffix.Length;
            candidate = (sanitized.Length > maxBase ? sanitized[..maxBase] : sanitized) + suffix;
        }
        return candidate;
    }
}
