using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    // ── Hydrotest Flag ──────────────────────────────────────────────

    private static readonly string[] HtExcelColumns =
    {
        "Line_ID", "Joint_ID", "Project No", "LAYOUT_NO", "Sheet No", "Weld No",
        "System", "Sub_System", "Test_Package_No", "Test_Type",
        "Spool_No", "P_ID", "Test_Package_No_WS"
    };

    private static readonly string[] HtExcelHeaders =
    {
        "Line ID", "Joint ID", "Project No", "Layout No", "Sheet No", "Weld No",
        "System", "Sub System", "Test Package No", "Test Type",
        "Spool No", "P&ID", "Test Package No WS"
    };

    // GET: /Home/HydrotestFlag
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> HydrotestFlag(
        [FromQuery] List<int>? projectNo,
        [FromQuery] List<string>? layoutNo,
        [FromQuery] List<string>? sheetNo,
        [FromQuery] List<string>? weldNumber)
    {
        var selectedProjectIds = (projectNo ?? []).Where(id => id > 0).ToList();
        var selLayoutNos = (layoutNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selSheetNos = (sheetNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selWeldNumbers = (weldNumber ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();

        bool hasFilter = selectedProjectIds.Count > 0
            || selLayoutNos.Count > 0
            || selSheetNos.Count > 0
            || selWeldNumbers.Count > 0;

        // Project list (always full)
        var projectList = await (from p in _context.Projects_tbl
            join ls in _context.Line_Sheet_tbl on p.Project_ID equals ls.Project_No
            join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
            select new { p.Project_ID, p.Project_Name })
            .Distinct()
            .OrderBy(x => x.Project_ID)
            .ToListAsync();
        ViewBag.ProjectNos = projectList
            .Select(x => new ProjectOption { Id = x.Project_ID, Name = x.Project_Name ?? string.Empty })
            .ToList();

        // Layout Nos – filtered by selected projects
        {
            var layoutQry = from ls in _context.Line_Sheet_tbl
                            join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                            where ll.LAYOUT_NO != null && ll.LAYOUT_NO != ""
                            select new { ls, ll };
            if (selectedProjectIds.Count > 0)
                layoutQry = layoutQry.Where(x => selectedProjectIds.Contains(x.ls.Project_No!.Value));
            ViewBag.LayoutNos = await layoutQry.Select(x => x.ll.LAYOUT_NO!)
                .Distinct().OrderBy(v => v).ToListAsync();
        }

        // Sheet Nos – filtered by selected projects + selected layouts
        {
            var sheetQry = from ls in _context.Line_Sheet_tbl
                           join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                           where ls.LS_SHEET != null && ls.LS_SHEET != ""
                           select new { ls, ll };
            if (selectedProjectIds.Count > 0)
                sheetQry = sheetQry.Where(x => selectedProjectIds.Contains(x.ls.Project_No!.Value));
            if (selLayoutNos.Count > 0)
                sheetQry = sheetQry.Where(x => selLayoutNos.Contains(x.ll.LAYOUT_NO!));
            ViewBag.SheetNos = await sheetQry.Select(x => x.ls.LS_SHEET!)
                .Distinct().OrderBy(v => v).ToListAsync();
        }

        // Weld Numbers – filtered by selected projects + layouts + sheets
        {
            var weldQry = from ls in _context.Line_Sheet_tbl
                          join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                          join dfr in _context.DFR_tbl on ls.Line_Sheet_ID equals dfr.Line_Sheet_ID_DFR
                          where dfr.WELD_NUMBER != null && dfr.WELD_NUMBER != ""
                          select new { ls, ll, dfr };
            if (selectedProjectIds.Count > 0)
                weldQry = weldQry.Where(x => selectedProjectIds.Contains(x.ls.Project_No!.Value));
            if (selLayoutNos.Count > 0)
                weldQry = weldQry.Where(x => selLayoutNos.Contains(x.ll.LAYOUT_NO!));
            if (selSheetNos.Count > 0)
                weldQry = weldQry.Where(x => selSheetNos.Contains(x.ls.LS_SHEET!));
            ViewBag.WeldNumbers = await weldQry.Select(x => x.dfr.WELD_NUMBER!)
                .Distinct().OrderBy(v => v).ToListAsync();
        }

        // Current filter values
        ViewBag.SelectedProjectIds = selectedProjectIds;
        ViewBag.SelectedLayoutNos = selLayoutNos;
        ViewBag.SelectedSheetNos = selSheetNos;
        ViewBag.SelectedWeldNumbers = selWeldNumbers;

        var rows = new List<HydrotestFlagRow>();
        if (hasFilter)
        {
            bool needDfr = selSheetNos.Count > 0 || selWeldNumbers.Count > 0;

            if (needDfr)
            {
                // LEFT JOIN DFR when Sheet or Weld filter is active
                var qry = from p in _context.Projects_tbl
                          join ls in _context.Line_Sheet_tbl on p.Project_ID equals ls.Project_No
                          join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                          join dfr in _context.DFR_tbl on ls.Line_Sheet_ID equals dfr.Line_Sheet_ID_DFR into dfrGroup
                          from d in dfrGroup.DefaultIfEmpty()
                          select new { p, ls, ll, d };

                if (selectedProjectIds.Count > 0)
                    qry = qry.Where(x => selectedProjectIds.Contains(x.p.Project_ID));
                if (selLayoutNos.Count > 0)
                    qry = qry.Where(x => selLayoutNos.Contains(x.ll.LAYOUT_NO!));
                if (selSheetNos.Count > 0)
                    qry = qry.Where(x => selSheetNos.Contains(x.ls.LS_SHEET!));
                if (selWeldNumbers.Count > 0)
                    qry = qry.Where(x => x.d != null && selWeldNumbers.Contains(x.d.WELD_NUMBER!));

                rows = await qry
                    .OrderBy(x => x.ll.Line_ID)
                    .Select(x => new HydrotestFlagRow
                    {
                        LineId = x.ll.Line_ID,
                        JointId = x.d != null ? (int?)x.d.Joint_ID : null,
                        ProjectNo = x.p.Project_ID + " - " + (x.p.Project_Name ?? ""),
                        LayoutNo = x.ll.LAYOUT_NO,
                        SheetNo = x.ls.LS_SHEET,
                        WeldNo = x.d == null ? null
                            : (x.d.J_Add != "New"
                                ? (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "") + (x.d.J_Add ?? "")
                                : (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "")),
                        System = x.ll.System,
                        SubSystem = x.ll.Sub_System,
                        TestPackageNo = x.ll.Test_Package_No,
                        TestType = x.ll.Test_Type,
                        SpoolNo = x.ll.Spool_No,
                        PId = x.ll.P_ID,
                        TestPackageNoWs = x.ll.Test_Package_No_WS
                    })
                    .ToListAsync();
            }
            else
            {
                // No DFR join needed — only Project/Layout filters
                var qry = from p in _context.Projects_tbl
                          join ls in _context.Line_Sheet_tbl on p.Project_ID equals ls.Project_No
                          join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                          select new { p, ls, ll };

                if (selectedProjectIds.Count > 0)
                    qry = qry.Where(x => selectedProjectIds.Contains(x.p.Project_ID));
                if (selLayoutNos.Count > 0)
                    qry = qry.Where(x => selLayoutNos.Contains(x.ll.LAYOUT_NO!));

                rows = await qry
                    .OrderBy(x => x.ll.Line_ID)
                    .Select(x => new HydrotestFlagRow
                    {
                        LineId = x.ll.Line_ID,
                        ProjectNo = x.p.Project_ID + " - " + (x.p.Project_Name ?? ""),
                        LayoutNo = x.ll.LAYOUT_NO,
                        SheetNo = x.ls.LS_SHEET,
                        System = x.ll.System,
                        SubSystem = x.ll.Sub_System,
                        TestPackageNo = x.ll.Test_Package_No,
                        TestType = x.ll.Test_Type,
                        SpoolNo = x.ll.Spool_No,
                        PId = x.ll.P_ID,
                        TestPackageNoWs = x.ll.Test_Package_No_WS
                    })
                    .ToListAsync();
            }
        }

        return View("HydrotestFlag", rows);
    }

    // ── Cascading JSON endpoints ────────────────────────────────────

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetHtLayoutNos([FromQuery] List<int>? projectNo)
    {
        var pids = (projectNo ?? []).Where(id => id > 0).ToList();
        var qry = from ls in _context.Line_Sheet_tbl
                  join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                  where ll.LAYOUT_NO != null && ll.LAYOUT_NO != ""
                  select new { ls, ll };
        if (pids.Count > 0)
            qry = qry.Where(x => pids.Contains(x.ls.Project_No!.Value));
        var result = await qry.Select(x => x.ll.LAYOUT_NO!).Distinct().OrderBy(v => v).ToListAsync();
        return Json(result);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetHtSheetNos(
        [FromQuery] List<int>? projectNo,
        [FromQuery] List<string>? layoutNo)
    {
        var pids = (projectNo ?? []).Where(id => id > 0).ToList();
        var lnos = (layoutNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();

        var qry = from ls in _context.Line_Sheet_tbl
                  join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                  where ls.LS_SHEET != null && ls.LS_SHEET != ""
                  select new { ls, ll };
        if (pids.Count > 0)
            qry = qry.Where(x => pids.Contains(x.ls.Project_No!.Value));
        if (lnos.Count > 0)
            qry = qry.Where(x => lnos.Contains(x.ll.LAYOUT_NO!));

        var result = await qry.Select(x => x.ls.LS_SHEET!).Distinct().OrderBy(v => v).ToListAsync();
        return Json(result);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetHtWeldNumbers(
        [FromQuery] List<int>? projectNo,
        [FromQuery] List<string>? layoutNo,
        [FromQuery] List<string>? sheetNo)
    {
        var pids = (projectNo ?? []).Where(id => id > 0).ToList();
        var lnos = (layoutNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var snos = (sheetNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();

        var qry = from ls in _context.Line_Sheet_tbl
                  join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                  join dfr in _context.DFR_tbl on ls.Line_Sheet_ID equals dfr.Line_Sheet_ID_DFR
                  where dfr.WELD_NUMBER != null && dfr.WELD_NUMBER != ""
                  select new { ls, ll, dfr };
        if (pids.Count > 0)
            qry = qry.Where(x => pids.Contains(x.ls.Project_No!.Value));
        if (lnos.Count > 0)
            qry = qry.Where(x => lnos.Contains(x.ll.LAYOUT_NO!));
        if (snos.Count > 0)
            qry = qry.Where(x => snos.Contains(x.ls.LS_SHEET!));

        var result = await qry.Select(x => x.dfr.WELD_NUMBER!)
            .Distinct().OrderBy(v => v).ToListAsync();
        return Json(result);
    }

    // ── Save (inline) ───────────────────────────────────────────────

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveHydrotestFlag(
        List<int> ids,
        List<int> jointIds,
        List<string?> systemVals,
        List<string?> subSystemVals,
        List<string?> testPackageNoVals,
        List<string?> testTypeVals,
        List<string?> spoolNoVals,
        List<string?> pIdVals,
        List<string?> testPackageNoWsVals)
    {
        if (ids == null || ids.Count == 0)
        {
            TempData["Msg"] = "No rows to save.";
            return RedirectToAction(nameof(HydrotestFlag));
        }

        int updated = 0;
        var userId = HttpContext.Session.GetInt32("UserID");

        // Deduplicate LINE_LIST updates – keep the last occurrence index
        var lineIdxMap = new Dictionary<int, int>();
        for (int i = 0; i < ids.Count; i++)
            lineIdxMap[ids[i]] = i;

        foreach (var kv in lineIdxMap)
        {
            var row = await _context.LINE_LIST_tbl.FirstOrDefaultAsync(r => r.Line_ID == kv.Key);
            if (row == null) continue;

            int i = kv.Value;
            row.System = SafeAt(systemVals, i);
            row.Sub_System = SafeAt(subSystemVals, i);
            row.Test_Package_No = SafeAt(testPackageNoVals, i);
            row.Test_Type = SafeAt(testTypeVals, i);
            row.Spool_No = SafeAt(spoolNoVals, i);
            row.P_ID = SafeAt(pIdVals, i);
            row.Test_Package_No_WS = SafeAt(testPackageNoWsVals, i);
            row.Line_List_Updated_Date = AppClock.UtcNow;
            row.Line_List_Updated_By = userId;
            updated++;
        }

        // Also update DFR_tbl for Test Package No / Test Package No WS
        for (int i = 0; i < ids.Count; i++)
        {
            int jid = (jointIds != null && i < jointIds.Count) ? jointIds[i] : 0;
            if (jid <= 0) continue;

            var dfr = await _context.DFR_tbl.FirstOrDefaultAsync(d => d.Joint_ID == jid);
            if (dfr == null) continue;

            dfr.DFR_Test_Package_No = SafeAt(testPackageNoVals, i);
            dfr.DFR_Test_Package_No_WS = SafeAt(testPackageNoWsVals, i);
            dfr.DFR_Updated_Date = AppClock.UtcNow;
            dfr.DFR_Updated_By = userId;
        }

        await _context.SaveChangesAsync();
        TempData["Msg"] = $"{updated} row(s) updated.";
        return RedirectToAction(nameof(HydrotestFlag));
    }

    // ── Bulk Fill ───────────────────────────────────────────────────

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkFillHydrotestFlag(
        string? bulkIds,
        string? bulkJointIds,
        string? bfSystem,
        string? bfSubSystem,
        string? bfTestPackageNo,
        string? bfTestType,
        string? bfSpoolNo,
        string? bfPId,
        string? bfTestPackageNoWs)
    {
        if (string.IsNullOrWhiteSpace(bulkIds))
        {
            TempData["Msg"] = "No rows matched the filter for bulk fill.";
            return RedirectToAction(nameof(HydrotestFlag));
        }

        var idList = bulkIds.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => int.TryParse(s, out var v) ? v : 0)
            .Where(v => v > 0)
            .Distinct()
            .ToList();

        if (idList.Count == 0)
        {
            TempData["Msg"] = "No valid IDs for bulk fill.";
            return RedirectToAction(nameof(HydrotestFlag));
        }

        var rows = await _context.LINE_LIST_tbl.Where(r => idList.Contains(r.Line_ID)).ToListAsync();
        int updated = 0;
        var userId = HttpContext.Session.GetInt32("UserID");

        foreach (var row in rows)
        {
            if (!string.IsNullOrEmpty(bfSystem)) row.System = bfSystem.Trim();
            if (!string.IsNullOrEmpty(bfSubSystem)) row.Sub_System = bfSubSystem.Trim();
            if (!string.IsNullOrEmpty(bfTestPackageNo)) row.Test_Package_No = bfTestPackageNo.Trim();
            if (!string.IsNullOrEmpty(bfTestType)) row.Test_Type = bfTestType.Trim();
            if (!string.IsNullOrEmpty(bfSpoolNo)) row.Spool_No = bfSpoolNo.Trim();
            if (!string.IsNullOrEmpty(bfPId)) row.P_ID = bfPId.Trim();
            if (!string.IsNullOrEmpty(bfTestPackageNoWs)) row.Test_Package_No_WS = bfTestPackageNoWs.Trim();

            row.Line_List_Updated_Date = AppClock.UtcNow;
            row.Line_List_Updated_By = userId;
            updated++;
        }

        // Also update DFR rows for Test Package No / Test Package No WS
        if (!string.IsNullOrEmpty(bfTestPackageNo) || !string.IsNullOrEmpty(bfTestPackageNoWs))
        {
            var jointIdList = (bulkJointIds ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(s => int.TryParse(s, out var v) ? v : 0)
                .Where(v => v > 0)
                .Distinct()
                .ToList();

            if (jointIdList.Count > 0)
            {
                var dfrRows = await _context.DFR_tbl.Where(d => jointIdList.Contains(d.Joint_ID)).ToListAsync();
                foreach (var dfr in dfrRows)
                {
                    if (!string.IsNullOrEmpty(bfTestPackageNo)) dfr.DFR_Test_Package_No = bfTestPackageNo.Trim();
                    if (!string.IsNullOrEmpty(bfTestPackageNoWs)) dfr.DFR_Test_Package_No_WS = bfTestPackageNoWs.Trim();
                    dfr.DFR_Updated_Date = AppClock.UtcNow;
                    dfr.DFR_Updated_By = userId;
                }
            }
        }

        await _context.SaveChangesAsync();
        TempData["Msg"] = $"Bulk fill complete. {updated} row(s) updated.";
        return RedirectToAction(nameof(HydrotestFlag));
    }

    // ── Import Excel ────────────────────────────────────────────────

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImportHydrotestFlag(IFormFile? excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            TempData["Msg"] = "Please select an Excel file to import.";
            return RedirectToAction(nameof(HydrotestFlag));
        }

        try
        {
            using var stream = excelFile.OpenReadStream();
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();

            var headerRow = ws.Row(1);
            var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            for (int c = 1; c <= lastCol; c++)
            {
                var hdr = (headerRow.Cell(c).GetString() ?? "").Trim();
                if (!string.IsNullOrWhiteSpace(hdr)) colMap[hdr] = c;
            }

            int colLineId = FindCol(colMap, "Line_ID", "Line ID", "LINE_ID");
            if (colLineId < 1)
            {
                TempData["Msg"] = "Excel file must contain a 'Line ID' or 'Line_ID' column.";
                return RedirectToAction(nameof(HydrotestFlag));
            }

            int colJointId = FindCol(colMap, "Joint_ID", "Joint ID", "JOINT_ID");
            int colSystem = FindCol(colMap, "System");
            int colSubSystem = FindCol(colMap, "Sub_System", "Sub System");
            int colTestPkg = FindCol(colMap, "Test_Package_No", "Test Package No");
            int colTestType = FindCol(colMap, "Test_Type", "Test Type");
            int colSpool = FindCol(colMap, "Spool_No", "Spool No");
            int colPId = FindCol(colMap, "P_ID", "P&ID");
            int colTestPkgWs = FindCol(colMap, "Test_Package_No_WS", "Test Package No WS");

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            int updated = 0;
            var userId = HttpContext.Session.GetInt32("UserID");
            var updatedLineIds = new HashSet<int>();

            for (int r = 2; r <= lastRow; r++)
            {
                var idRaw = ws.Cell(r, colLineId).GetString()?.Trim();
                if (!int.TryParse(idRaw, out var lineId) || lineId <= 0) continue;

                // Only update LINE_LIST once per Line_ID
                if (updatedLineIds.Add(lineId))
                {
                    var row = await _context.LINE_LIST_tbl.FirstOrDefaultAsync(x => x.Line_ID == lineId);
                    if (row != null)
                    {
                        if (colSystem >= 1) { var v = ws.Cell(r, colSystem).GetString()?.Trim(); if (v != null) row.System = v; }
                        if (colSubSystem >= 1) { var v = ws.Cell(r, colSubSystem).GetString()?.Trim(); if (v != null) row.Sub_System = v; }
                        if (colTestPkg >= 1) { var v = ws.Cell(r, colTestPkg).GetString()?.Trim(); if (v != null) row.Test_Package_No = v; }
                        if (colTestType >= 1) { var v = ws.Cell(r, colTestType).GetString()?.Trim(); if (v != null) row.Test_Type = v; }
                        if (colSpool >= 1) { var v = ws.Cell(r, colSpool).GetString()?.Trim(); if (v != null) row.Spool_No = v; }
                        if (colPId >= 1) { var v = ws.Cell(r, colPId).GetString()?.Trim(); if (v != null) row.P_ID = v; }
                        if (colTestPkgWs >= 1) { var v = ws.Cell(r, colTestPkgWs).GetString()?.Trim(); if (v != null) row.Test_Package_No_WS = v; }
                        row.Line_List_Updated_Date = AppClock.UtcNow;
                        row.Line_List_Updated_By = userId;
                        updated++;
                    }
                }

                // Update DFR row if Joint_ID column present
                if (colJointId >= 1)
                {
                    var jidRaw = ws.Cell(r, colJointId).GetString()?.Trim();
                    if (int.TryParse(jidRaw, out var jid) && jid > 0)
                    {
                        var dfr = await _context.DFR_tbl.FirstOrDefaultAsync(d => d.Joint_ID == jid);
                        if (dfr != null)
                        {
                            if (colTestPkg >= 1) { var v = ws.Cell(r, colTestPkg).GetString()?.Trim(); if (v != null) dfr.DFR_Test_Package_No = v; }
                            if (colTestPkgWs >= 1) { var v = ws.Cell(r, colTestPkgWs).GetString()?.Trim(); if (v != null) dfr.DFR_Test_Package_No_WS = v; }
                            dfr.DFR_Updated_Date = AppClock.UtcNow;
                            dfr.DFR_Updated_By = userId;
                        }
                    }
                }
            }

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Import complete. {updated} row(s) updated.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ImportHydrotestFlag failed");
            TempData["Msg"] = $"Import failed. {ex.GetBaseException()?.Message ?? ex.Message}";
        }

        return RedirectToAction(nameof(HydrotestFlag));
    }

    // ── Export Excel ────────────────────────────────────────────────

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportHydrotestFlag(
        [FromQuery] List<int>? projectNo,
        [FromQuery] List<string>? layoutNo,
        [FromQuery] List<string>? sheetNo,
        [FromQuery] List<string>? weldNumber)
    {
        var selectedProjectIds = (projectNo ?? []).Where(id => id > 0).ToList();
        var selLayoutNos = (layoutNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selSheetNos = (sheetNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selWeldNumbers = (weldNumber ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();

        var qry = from p in _context.Projects_tbl
                  join ls in _context.Line_Sheet_tbl on p.Project_ID equals ls.Project_No
                  join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                  join dfr in _context.DFR_tbl on ls.Line_Sheet_ID equals dfr.Line_Sheet_ID_DFR into dfrGroup
                  from d in dfrGroup.DefaultIfEmpty()
                  select new { p, ls, ll, d };

        if (selectedProjectIds.Count > 0)
            qry = qry.Where(x => selectedProjectIds.Contains(x.p.Project_ID));
        if (selLayoutNos.Count > 0)
            qry = qry.Where(x => selLayoutNos.Contains(x.ll.LAYOUT_NO!));
        if (selSheetNos.Count > 0)
            qry = qry.Where(x => selSheetNos.Contains(x.ls.LS_SHEET!));
        if (selWeldNumbers.Count > 0)
            qry = qry.Where(x => x.d != null && selWeldNumbers.Contains(x.d.WELD_NUMBER!));

        var rows = await qry
            .OrderBy(x => x.ll.Line_ID)
            .Select(x => new HydrotestFlagRow
            {
                LineId = x.ll.Line_ID,
                JointId = x.d != null ? (int?)x.d.Joint_ID : null,
                ProjectNo = x.p.Project_ID + " - " + (x.p.Project_Name ?? ""),
                LayoutNo = x.ll.LAYOUT_NO,
                SheetNo = x.ls.LS_SHEET,
                WeldNo = x.d == null ? null
                    : (x.d.J_Add != "New"
                        ? (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "") + (x.d.J_Add ?? "")
                        : (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "")),
                System = x.ll.System,
                SubSystem = x.ll.Sub_System,
                TestPackageNo = x.ll.Test_Package_No,
                TestType = x.ll.Test_Type,
                SpoolNo = x.ll.Spool_No,
                PId = x.ll.P_ID,
                TestPackageNoWs = x.ll.Test_Package_No_WS
            })
            .ToListAsync();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Hydrotest Flag");

        for (int i = 0; i < HtExcelHeaders.Length; i++)
            ws.Cell(1, i + 1).Value = HtExcelHeaders[i];

        int rowIdx = 2;
        foreach (var r in rows)
        {
            ws.Cell(rowIdx, 1).Value = r.LineId;
            ws.Cell(rowIdx, 2).Value = r.JointId.HasValue ? r.JointId.Value : "";
            ws.Cell(rowIdx, 3).Value = r.ProjectNo;
            ws.Cell(rowIdx, 4).Value = r.LayoutNo ?? "";
            ws.Cell(rowIdx, 5).Value = r.SheetNo ?? "";
            ws.Cell(rowIdx, 6).Value = r.WeldNo ?? "";
            ws.Cell(rowIdx, 7).Value = r.System ?? "";
            ws.Cell(rowIdx, 8).Value = r.SubSystem ?? "";
            ws.Cell(rowIdx, 9).Value = r.TestPackageNo ?? "";
            ws.Cell(rowIdx, 10).Value = r.TestType ?? "";
            ws.Cell(rowIdx, 11).Value = r.SpoolNo ?? "";
            ws.Cell(rowIdx, 12).Value = r.PId ?? "";
            ws.Cell(rowIdx, 13).Value = r.TestPackageNoWs ?? "";
            rowIdx++;
        }

        int lastRow = rowIdx - 1;
        if (lastRow >= 2)
        {
            var fullRange = ws.Range(1, 1, lastRow, HtExcelHeaders.Length);
            var table = fullRange.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowTotalsRow = false;
        }

        ws.Row(1).Height = 30;
        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return File(ms.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"HydrotestFlag_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx");
    }
}
