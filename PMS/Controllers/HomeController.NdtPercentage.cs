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
    // ── NDT Percentage ──────────────────────────────────────────────

    private static readonly string[] NdtExcelColumns =
    {
        "Line_ID", "Joint_ID", "LAYOUT_NO", "Line_Class", "Sheet_No", "Weld_No",
        "Material", "Material_Des",
        "Fluid", "Design_Temperature_F", "Category",
        "RT_Shop", "RT_Field", "RT_Field_Shop_SW",
        "MT_Shop", "MT_Field", "PT_Shop", "PT_Field",
        "PWHT_Y_N", "PWHT_20mm", "HT", "HT_After_PWHT",
        "PMI", "DWG_Remarks", "Special_RT"
    };

    private static readonly string[] NdtExcelHeaders =
    {
        "Line ID", "Joint ID", "Layout No", "Line Class", "Sheet No", "Weld No",
        "Material", "Material Des",
        "Fluid", "Design Temp (°F)", "Category",
        "RT Shop", "RT Field", "RT Field/Shop SW",
        "MT Shop", "MT Field", "PT Shop", "PT Field",
        "PWHT Y/N", "PWHT >20mm", "HT", "HT After PWHT",
        "PMI", "DWG Remarks", "Special Joint"
    };

    // GET: /Home/NdtPercentage
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> NdtPercentage(
        [FromQuery] List<int>? projectNo,
        [FromQuery] List<string>? layoutNo, [FromQuery] List<string>? lineClass,
        [FromQuery] List<string>? material, [FromQuery] List<string>? materialDes,
        [FromQuery] List<string>? fluid, [FromQuery] List<string>? designTemp,
        [FromQuery] List<string>? category,
        [FromQuery] List<string>? sheetNo,
        [FromQuery] List<string>? weldNumber)
    {
        var selectedProjectIds = (projectNo ?? []).Where(id => id > 0).ToList();
        var selLayoutNos = (layoutNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selLineClasses = (lineClass ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selMaterials = (material ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selMaterialDescs = (materialDes ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selFluids = (fluid ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selDesignTemps = (designTemp ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selCategories = (category ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selSheetNos = (sheetNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selWeldNumbers = (weldNumber ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();

        bool hasFilter = selectedProjectIds.Count > 0
            || selLayoutNos.Count > 0
            || selLineClasses.Count > 0
            || selMaterials.Count > 0
            || selMaterialDescs.Count > 0
            || selFluids.Count > 0
            || selDesignTemps.Count > 0
            || selCategories.Count > 0
            || selSheetNos.Count > 0
            || selWeldNumbers.Count > 0;

        // Distinct values for filter dropdowns
        var allRows = _context.LINE_LIST_tbl.AsNoTracking();

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

        ViewBag.LayoutNos = await allRows.Where(r => r.LAYOUT_NO != null && r.LAYOUT_NO != "")
            .Select(r => r.LAYOUT_NO!).Distinct().OrderBy(v => v).ToListAsync();
        ViewBag.LineClasses = await allRows.Where(r => r.Line_Class != null && r.Line_Class != "")
            .Select(r => r.Line_Class!).Distinct().OrderBy(v => v).ToListAsync();
        ViewBag.Materials = await allRows.Where(r => r.Material != null && r.Material != "")
            .Select(r => r.Material!).Distinct().OrderBy(v => v).ToListAsync();
        ViewBag.MaterialDescs = await allRows.Where(r => r.Material_Des != null && r.Material_Des != "")
            .Select(r => r.Material_Des!).Distinct().OrderBy(v => v).ToListAsync();
        ViewBag.Fluids = await allRows.Where(r => r.Fluid != null && r.Fluid != "")
            .Select(r => r.Fluid!).Distinct().OrderBy(v => v).ToListAsync();
        ViewBag.DesignTemps = await allRows.Where(r => r.Design_Temperature_F != null && r.Design_Temperature_F != "")
            .Select(r => r.Design_Temperature_F!).Distinct().OrderBy(v => v).ToListAsync();
        ViewBag.Categories = await allRows.Where(r => r.Category != null && r.Category != "")
            .Select(r => r.Category!).Distinct().OrderBy(v => v).ToListAsync();

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
        ViewBag.SelectedLineClasses = selLineClasses;
        ViewBag.SelectedMaterials = selMaterials;
        ViewBag.SelectedMaterialDescs = selMaterialDescs;
        ViewBag.SelectedFluids = selFluids;
        ViewBag.SelectedDesignTemps = selDesignTemps;
        ViewBag.SelectedCategories = selCategories;
        ViewBag.SelectedSheetNos = selSheetNos;
        ViewBag.SelectedWeldNumbers = selWeldNumbers;

        var rows = new List<NdtPercentageRow>();
        if (hasFilter)
        {
            bool needDfr = selSheetNos.Count > 0 || selWeldNumbers.Count > 0;

            if (needDfr)
            {
                // Join through Line_Sheet → LINE_LIST → DFR when Sheet or Weld filter is active
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
                if (selLineClasses.Count > 0)
                    qry = qry.Where(x => selLineClasses.Contains(x.ll.Line_Class!));
                if (selMaterials.Count > 0)
                    qry = qry.Where(x => selMaterials.Contains(x.ll.Material!));
                if (selMaterialDescs.Count > 0)
                    qry = qry.Where(x => selMaterialDescs.Contains(x.ll.Material_Des!));
                if (selFluids.Count > 0)
                    qry = qry.Where(x => selFluids.Contains(x.ll.Fluid!));
                if (selDesignTemps.Count > 0)
                    qry = qry.Where(x => selDesignTemps.Contains(x.ll.Design_Temperature_F!));
                if (selCategories.Count > 0)
                    qry = qry.Where(x => selCategories.Contains(x.ll.Category!));
                if (selSheetNos.Count > 0)
                    qry = qry.Where(x => selSheetNos.Contains(x.ls.LS_SHEET!));
                if (selWeldNumbers.Count > 0)
                    qry = qry.Where(x => x.d != null && selWeldNumbers.Contains(x.d.WELD_NUMBER!));

                rows = await qry
                    .OrderBy(x => x.ll.Line_ID)
                    .Select(x => new NdtPercentageRow
                    {
                        LineId = x.ll.Line_ID,
                        JointId = x.d != null ? (int?)x.d.Joint_ID : null,
                        LayoutNo = x.ll.LAYOUT_NO,
                        LineClass = x.ll.Line_Class,
                        SheetNo = x.ls.LS_SHEET,
                        WeldNo = x.d == null ? null
                            : (x.d.J_Add != "New"
                                ? (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "") + (x.d.J_Add ?? "")
                                : (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "")),
                        MaterialDes = x.ll.Material_Des,
                        Material = x.ll.Material,
                        Fluid = x.ll.Fluid,
                        DesignTemp = x.ll.Design_Temperature_F,
                        Category = x.ll.Category,
                        RtShop = x.ll.RT_Shop,
                        RtField = x.ll.RT_Field,
                        RtFieldShopSw = x.ll.RT_Field_Shop_SW,
                        MtShop = x.ll.MT_Shop,
                        MtField = x.ll.MT_Field,
                        PtShop = x.ll.PT_Shop,
                        PtField = x.ll.PT_Field,
                        PwhtYn = x.ll.PWHT_Y_N,
                        Pwht20 = x.ll.PWHT_20mm,
                        Ht = x.ll.HT,
                        HtAfterPwht = x.ll.HT_After_PWHT,
                        Pmi = x.ll.PMI,
                        DwgRemarks = x.ll.DWG_Remarks,
                        SpecialRt = x.d != null ? x.d.Special_RT : null
                    })
                    .ToListAsync();
            }
            else
            {
                // No DFR join needed
                IQueryable<LineList> qry = _context.LINE_LIST_tbl.AsQueryable();
                if (selectedProjectIds.Count > 0)
                {
                    var lineIdsForProject = _context.Line_Sheet_tbl
                        .Where(ls => selectedProjectIds.Contains(ls.Project_No!.Value) && ls.Line_ID_LS.HasValue)
                        .Select(ls => ls.Line_ID_LS!.Value);
                    qry = qry.Where(r => lineIdsForProject.Contains(r.Line_ID));
                }
                if (selLayoutNos.Count > 0)
                    qry = qry.Where(r => selLayoutNos.Contains(r.LAYOUT_NO!));
                if (selLineClasses.Count > 0)
                    qry = qry.Where(r => selLineClasses.Contains(r.Line_Class!));
                if (selMaterials.Count > 0)
                    qry = qry.Where(r => selMaterials.Contains(r.Material!));
                if (selMaterialDescs.Count > 0)
                    qry = qry.Where(r => selMaterialDescs.Contains(r.Material_Des!));
                if (selFluids.Count > 0)
                    qry = qry.Where(r => selFluids.Contains(r.Fluid!));
                if (selDesignTemps.Count > 0)
                    qry = qry.Where(r => selDesignTemps.Contains(r.Design_Temperature_F!));
                if (selCategories.Count > 0)
                    qry = qry.Where(r => selCategories.Contains(r.Category!));

                rows = await qry.OrderBy(r => r.Line_ID)
                    .Select(r => new NdtPercentageRow
                    {
                        LineId = r.Line_ID,
                        LayoutNo = r.LAYOUT_NO,
                        LineClass = r.Line_Class,
                        MaterialDes = r.Material_Des,
                        Material = r.Material,
                        Fluid = r.Fluid,
                        DesignTemp = r.Design_Temperature_F,
                        Category = r.Category,
                        RtShop = r.RT_Shop,
                        RtField = r.RT_Field,
                        RtFieldShopSw = r.RT_Field_Shop_SW,
                        MtShop = r.MT_Shop,
                        MtField = r.MT_Field,
                        PtShop = r.PT_Shop,
                        PtField = r.PT_Field,
                        PwhtYn = r.PWHT_Y_N,
                        Pwht20 = r.PWHT_20mm,
                        Ht = r.HT,
                        HtAfterPwht = r.HT_After_PWHT,
                        Pmi = r.PMI,
                        DwgRemarks = r.DWG_Remarks
                    })
                    .ToListAsync();
            }
        }

        return View("NdtPercentage", rows);
    }

    // ── Cascading JSON endpoints for NDT Percentage ─────────────────

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetNdtSheetNos(
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
    public async Task<IActionResult> GetNdtWeldNumbers(
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

    // POST: /Home/SaveNdtPercentage  (inline table save)
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveNdtPercentage(List<int> ids,
        List<int> jointIds,
        List<string?> materialDesVals, List<string?> materialVals,
        List<string?> designTempVals, List<string?> categoryVals,
        List<double?> rtShopVals, List<double?> rtFieldVals, List<double?> rtFieldShopSwVals,
        List<double?> mtShopVals, List<double?> mtFieldVals,
        List<double?> ptShopVals, List<double?> ptFieldVals,
        List<bool?> pwhtYnVals, List<bool?> pwht20Vals,
        List<double?> htVals, List<double?> htAfterVals,
        List<bool?> pmiVals, List<string?> dwgRemarksVals,
        List<double?> specialRtVals)
    {
        if (ids == null || ids.Count == 0)
        {
            TempData["Msg"] = "No rows to save.";
            return RedirectToAction(nameof(NdtPercentage));
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
            row.Material_Des = SafeAt(materialDesVals, i);
            row.Material = SafeAt(materialVals, i);
            row.Design_Temperature_F = SafeAt(designTempVals, i);
            row.Category = SafeAt(categoryVals, i);
            row.RT_Shop = SafeDoubleAt(rtShopVals, i);
            row.RT_Field = SafeDoubleAt(rtFieldVals, i);
            row.RT_Field_Shop_SW = SafeDoubleAt(rtFieldShopSwVals, i);
            row.MT_Shop = SafeDoubleAt(mtShopVals, i);
            row.MT_Field = SafeDoubleAt(mtFieldVals, i);
            row.PT_Shop = SafeDoubleAt(ptShopVals, i);
            row.PT_Field = SafeDoubleAt(ptFieldVals, i);
            row.PWHT_Y_N = SafeBoolAt(pwhtYnVals, i);
            row.PWHT_20mm = SafeBoolAt(pwht20Vals, i);
            row.HT = SafeDoubleAt(htVals, i);
            row.HT_After_PWHT = SafeDoubleAt(htAfterVals, i);
            row.PMI = SafeBoolAt(pmiVals, i);
            row.DWG_Remarks = SafeAt(dwgRemarksVals, i);
            row.Line_List_Updated_Date = AppClock.UtcNow;
            row.Line_List_Updated_By = userId;
            updated++;
        }

        // Update DFR_tbl.Special_RT for each joint
        for (int i = 0; i < ids.Count; i++)
        {
            int jid = (jointIds != null && i < jointIds.Count) ? jointIds[i] : 0;
            if (jid <= 0) continue;

            var dfr = await _context.DFR_tbl.FirstOrDefaultAsync(d => d.Joint_ID == jid);
            if (dfr == null) continue;

            dfr.Special_RT = SafeDoubleAt(specialRtVals, i);
            dfr.DFR_Updated_Date = AppClock.UtcNow;
            dfr.DFR_Updated_By = userId;
        }

        await _context.SaveChangesAsync();
        TempData["Msg"] = $"{updated} row(s) updated.";
        return RedirectToAction(nameof(NdtPercentage));
    }

    // POST: /Home/BulkFillNdtPercentage  (fill all filtered rows with same values)
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkFillNdtPercentage(
        string? bulkIds,
        string? bulkJointIds,
        string? bfMaterialDes, string? bfMaterial, string? bfDesignTemp, string? bfCategory,
        double? bfRtShop, double? bfRtField, double? bfRtFieldShopSw,
        double? bfMtShop, double? bfMtField,
        double? bfPtShop, double? bfPtField,
        string? bfPwhtYn, string? bfPwht20,
        double? bfHt, double? bfHtAfter,
        string? bfPmi, string? bfDwgRemarks,
        double? bfSpecialRt)
    {
        if (string.IsNullOrWhiteSpace(bulkIds))
        {
            TempData["Msg"] = "No rows matched the filter for bulk fill.";
            return RedirectToAction(nameof(NdtPercentage));
        }

        var idList = bulkIds.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => int.TryParse(s, out var v) ? v : 0)
            .Where(v => v > 0)
            .ToList();

        if (idList.Count == 0)
        {
            TempData["Msg"] = "No valid IDs for bulk fill.";
            return RedirectToAction(nameof(NdtPercentage));
        }

        var rows = await _context.LINE_LIST_tbl.Where(r => idList.Contains(r.Line_ID)).ToListAsync();
        int updated = 0;
        var userId = HttpContext.Session.GetInt32("UserID");

        foreach (var row in rows)
        {
            if (!string.IsNullOrEmpty(bfMaterialDes)) row.Material_Des = bfMaterialDes.Trim();
            if (!string.IsNullOrEmpty(bfMaterial)) row.Material = bfMaterial.Trim();
            if (!string.IsNullOrEmpty(bfDesignTemp)) row.Design_Temperature_F = bfDesignTemp.Trim();
            if (!string.IsNullOrEmpty(bfCategory)) row.Category = bfCategory.Trim();
            if (bfRtShop.HasValue) row.RT_Shop = bfRtShop;
            if (bfRtField.HasValue) row.RT_Field = bfRtField;
            if (bfRtFieldShopSw.HasValue) row.RT_Field_Shop_SW = bfRtFieldShopSw;
            if (bfMtShop.HasValue) row.MT_Shop = bfMtShop;
            if (bfMtField.HasValue) row.MT_Field = bfMtField;
            if (bfPtShop.HasValue) row.PT_Shop = bfPtShop;
            if (bfPtField.HasValue) row.PT_Field = bfPtField;
            if (!string.IsNullOrEmpty(bfPwhtYn))
                row.PWHT_Y_N = bfPwhtYn.Trim().Equals("true", StringComparison.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(bfPwht20))
                row.PWHT_20mm = bfPwht20.Trim().Equals("true", StringComparison.OrdinalIgnoreCase);
            if (bfHt.HasValue) row.HT = bfHt;
            if (bfHtAfter.HasValue) row.HT_After_PWHT = bfHtAfter;
            if (!string.IsNullOrEmpty(bfPmi))
                row.PMI = bfPmi.Trim().Equals("true", StringComparison.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(bfDwgRemarks)) row.DWG_Remarks = bfDwgRemarks.Trim();

            row.Line_List_Updated_Date = AppClock.UtcNow;
            row.Line_List_Updated_By = userId;
            updated++;
        }

        // Bulk update DFR_tbl.Special_RT
        if (bfSpecialRt.HasValue)
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
                    dfr.Special_RT = bfSpecialRt.Value;
                    dfr.DFR_Updated_Date = AppClock.UtcNow;
                    dfr.DFR_Updated_By = userId;
                }
            }
        }

        await _context.SaveChangesAsync();
        TempData["Msg"] = $"Bulk fill complete. {updated} row(s) updated.";
        return RedirectToAction(nameof(NdtPercentage));
    }

    // POST: /Home/ImportNdtPercentage  (upload Excel)
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImportNdtPercentage(IFormFile? excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            TempData["Msg"] = "Please select an Excel file to import.";
            return RedirectToAction(nameof(NdtPercentage));
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

            // Find the Line_ID column
            int colId = FindCol(colMap, "Line_ID", "Line ID", "LINE_ID");
            if (colId < 1)
            {
                TempData["Msg"] = "Excel file must contain a 'Line ID' or 'Line_ID' column.";
                return RedirectToAction(nameof(NdtPercentage));
            }

            int colJointId = FindCol(colMap, "Joint_ID", "Joint ID", "JOINT_ID");
            int colSpecialRt = FindCol(colMap, "Special_RT", "Special RT", "Special Joint", "Special_Joint");

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            int updated = 0;
            var userId = HttpContext.Session.GetInt32("UserID");
            var updatedLineIds = new HashSet<int>();

            for (int r = 2; r <= lastRow; r++)
            {
                var idRaw = ws.Cell(r, colId).GetString()?.Trim();
                if (!int.TryParse(idRaw, out var lineId) || lineId <= 0) continue;

                if (updatedLineIds.Add(lineId))
                {
                    var row = await _context.LINE_LIST_tbl.FirstOrDefaultAsync(x => x.Line_ID == lineId);
                    if (row != null)
                    {
                        SetIfPresent(ws, r, colMap, row, "Material_Des", "Material Des");
                        SetIfPresent(ws, r, colMap, row, "Material", null);
                        SetIfPresent(ws, r, colMap, row, "Design_Temperature_F", "Design Temp (°F)");
                        SetIfPresent(ws, r, colMap, row, "Category", null);
                        SetDoubleIfPresent(ws, r, colMap, row, "RT_Shop", "RT Shop");
                        SetDoubleIfPresent(ws, r, colMap, row, "RT_Field", "RT Field");
                        SetDoubleIfPresent(ws, r, colMap, row, "RT_Field_Shop_SW", "RT Field/Shop SW");
                        SetDoubleIfPresent(ws, r, colMap, row, "MT_Shop", "MT Shop");
                        SetDoubleIfPresent(ws, r, colMap, row, "MT_Field", "MT Field");
                        SetDoubleIfPresent(ws, r, colMap, row, "PT_Shop", "PT Shop");
                        SetDoubleIfPresent(ws, r, colMap, row, "PT_Field", "PT Field");
                        SetBoolIfPresent(ws, r, colMap, row, "PWHT_Y_N", "PWHT Y/N");
                        SetBoolIfPresent(ws, r, colMap, row, "PWHT_20mm", "PWHT >20mm");
                        SetDoubleIfPresent(ws, r, colMap, row, "HT", null);
                        SetDoubleIfPresent(ws, r, colMap, row, "HT_After_PWHT", "HT After PWHT");
                        SetBoolIfPresent(ws, r, colMap, row, "PMI", null);
                        SetIfPresent(ws, r, colMap, row, "DWG_Remarks", "DWG Remarks");

                        row.Line_List_Updated_Date = AppClock.UtcNow;
                        row.Line_List_Updated_By = userId;
                        updated++;
                    }
                }

                // Update DFR_tbl.Special_RT if Joint_ID column present
                if (colJointId >= 1 && colSpecialRt >= 1)
                {
                    var jidRaw = ws.Cell(r, colJointId).GetString()?.Trim();
                    if (int.TryParse(jidRaw, out var jid) && jid > 0)
                    {
                        var dfr = await _context.DFR_tbl.FirstOrDefaultAsync(d => d.Joint_ID == jid);
                        if (dfr != null)
                        {
                            var srtRaw = (ws.Cell(r, colSpecialRt).GetString() ?? "").Trim();
                            if (!string.IsNullOrWhiteSpace(srtRaw))
                            {
                                if (double.TryParse(srtRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var srtVal))
                                    dfr.Special_RT = srtVal;
                                else
                                    dfr.Special_RT = srtRaw.ToUpperInvariant() is "YES" or "TRUE" or "Y" ? 1 : 0;
                                dfr.DFR_Updated_Date = AppClock.UtcNow;
                                dfr.DFR_Updated_By = userId;
                            }
                        }
                    }
                }
            }

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Import complete. {updated} row(s) updated.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ImportNdtPercentage failed");
            TempData["Msg"] = $"Import failed. {ex.GetBaseException()?.Message ?? ex.Message}";
        }

        return RedirectToAction(nameof(NdtPercentage));
    }

    // GET: /Home/ExportNdtPercentage
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportNdtPercentage(
        [FromQuery] List<int>? projectNo,
        [FromQuery] List<string>? layoutNo, [FromQuery] List<string>? lineClass,
        [FromQuery] List<string>? material, [FromQuery] List<string>? materialDes,
        [FromQuery] List<string>? fluid, [FromQuery] List<string>? designTemp,
        [FromQuery] List<string>? category,
        [FromQuery] List<string>? sheetNo,
        [FromQuery] List<string>? weldNumber)
    {
        var selectedProjectIds = (projectNo ?? []).Where(id => id > 0).ToList();
        var selLayoutNos = (layoutNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selLineClasses = (lineClass ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selMaterials = (material ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selMaterialDescs = (materialDes ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selFluids = (fluid ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selDesignTemps = (designTemp ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selCategories = (category ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selSheetNos = (sheetNo ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();
        var selWeldNumbers = (weldNumber ?? []).Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()).ToList();

        // Always use the joined query for export so Sheet/Weld/Special_RT are included
        var qry = from p in _context.Projects_tbl
                  join ls in _context.Line_Sheet_tbl on p.Project_ID equals ls.Project_No
                  join ll in _context.LINE_LIST_tbl on ls.Line_ID_LS equals ll.Line_ID
                  join dfr in _context.DFR_tbl on ls.Line_Sheet_ID equals dfr.Line_Sheet_ID_DFR into dfrGroup
                  from d in dfrGroup.DefaultIfEmpty()
                  select new { p, ls, ll, d };

        if (selectedProjectIds.Count > 0)
            qry = qry.Where(x => selectedProjectIds.Contains(x.p.Project_ID));
        if (selLayoutNos.Count > 0) qry = qry.Where(x => selLayoutNos.Contains(x.ll.LAYOUT_NO!));
        if (selLineClasses.Count > 0) qry = qry.Where(x => selLineClasses.Contains(x.ll.Line_Class!));
        if (selMaterials.Count > 0) qry = qry.Where(x => selMaterials.Contains(x.ll.Material!));
        if (selMaterialDescs.Count > 0) qry = qry.Where(x => selMaterialDescs.Contains(x.ll.Material_Des!));
        if (selFluids.Count > 0) qry = qry.Where(x => selFluids.Contains(x.ll.Fluid!));
        if (selDesignTemps.Count > 0) qry = qry.Where(x => selDesignTemps.Contains(x.ll.Design_Temperature_F!));
        if (selCategories.Count > 0) qry = qry.Where(x => selCategories.Contains(x.ll.Category!));
        if (selSheetNos.Count > 0) qry = qry.Where(x => selSheetNos.Contains(x.ls.LS_SHEET!));
        if (selWeldNumbers.Count > 0) qry = qry.Where(x => x.d != null && selWeldNumbers.Contains(x.d.WELD_NUMBER!));

        var rows = await qry
            .OrderBy(x => x.ll.Line_ID)
            .Select(x => new NdtPercentageRow
            {
                LineId = x.ll.Line_ID,
                JointId = x.d != null ? (int?)x.d.Joint_ID : null,
                LayoutNo = x.ll.LAYOUT_NO,
                LineClass = x.ll.Line_Class,
                SheetNo = x.ls.LS_SHEET,
                WeldNo = x.d == null ? null
                    : (x.d.J_Add != "New"
                        ? (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "") + (x.d.J_Add ?? "")
                        : (x.d.LOCATION ?? "") + "-" + (x.d.WELD_NUMBER ?? "")),
                MaterialDes = x.ll.Material_Des,
                Material = x.ll.Material,
                Fluid = x.ll.Fluid,
                DesignTemp = x.ll.Design_Temperature_F,
                Category = x.ll.Category,
                RtShop = x.ll.RT_Shop,
                RtField = x.ll.RT_Field,
                RtFieldShopSw = x.ll.RT_Field_Shop_SW,
                MtShop = x.ll.MT_Shop,
                MtField = x.ll.MT_Field,
                PtShop = x.ll.PT_Shop,
                PtField = x.ll.PT_Field,
                PwhtYn = x.ll.PWHT_Y_N,
                Pwht20 = x.ll.PWHT_20mm,
                Ht = x.ll.HT,
                HtAfterPwht = x.ll.HT_After_PWHT,
                Pmi = x.ll.PMI,
                DwgRemarks = x.ll.DWG_Remarks,
                SpecialRt = x.d != null ? x.d.Special_RT : null
            })
            .ToListAsync();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("NDT Percentage");

        for (int i = 0; i < NdtExcelHeaders.Length; i++)
            ws.Cell(1, i + 1).Value = NdtExcelHeaders[i];

        int rowIdx = 2;
        foreach (var r in rows)
        {
            ws.Cell(rowIdx, 1).Value = r.LineId;
            ws.Cell(rowIdx, 2).Value = r.JointId.HasValue ? r.JointId.Value : "";
            ws.Cell(rowIdx, 3).Value = r.LayoutNo ?? "";
            ws.Cell(rowIdx, 4).Value = r.LineClass ?? "";
            ws.Cell(rowIdx, 5).Value = r.SheetNo ?? "";
            ws.Cell(rowIdx, 6).Value = r.WeldNo ?? "";
            ws.Cell(rowIdx, 7).Value = r.Material ?? "";
            ws.Cell(rowIdx, 8).Value = r.MaterialDes ?? "";
            ws.Cell(rowIdx, 9).Value = r.Fluid ?? "";
            ws.Cell(rowIdx, 10).Value = r.DesignTemp ?? "";
            ws.Cell(rowIdx, 11).Value = r.Category ?? "";
            if (r.RtShop.HasValue) ws.Cell(rowIdx, 12).Value = r.RtShop.Value;
            if (r.RtField.HasValue) ws.Cell(rowIdx, 13).Value = r.RtField.Value;
            if (r.RtFieldShopSw.HasValue) ws.Cell(rowIdx, 14).Value = r.RtFieldShopSw.Value;
            if (r.MtShop.HasValue) ws.Cell(rowIdx, 15).Value = r.MtShop.Value;
            if (r.MtField.HasValue) ws.Cell(rowIdx, 16).Value = r.MtField.Value;
            if (r.PtShop.HasValue) ws.Cell(rowIdx, 17).Value = r.PtShop.Value;
            if (r.PtField.HasValue) ws.Cell(rowIdx, 18).Value = r.PtField.Value;
            ws.Cell(rowIdx, 19).Value = r.PwhtYn == true ? "Yes" : r.PwhtYn == false ? "No" : "";
            ws.Cell(rowIdx, 20).Value = r.Pwht20 == true ? "Yes" : r.Pwht20 == false ? "No" : "";
            if (r.Ht.HasValue) ws.Cell(rowIdx, 21).Value = r.Ht.Value;
            if (r.HtAfterPwht.HasValue) ws.Cell(rowIdx, 22).Value = r.HtAfterPwht.Value;
            ws.Cell(rowIdx, 23).Value = r.Pmi == true ? "Yes" : r.Pmi == false ? "No" : "";
            ws.Cell(rowIdx, 24).Value = r.DwgRemarks ?? "";
            if (r.SpecialRt.HasValue) ws.Cell(rowIdx, 25).Value = r.SpecialRt.Value;
            rowIdx++;
        }

        int lastRow = rowIdx - 1;
        if (lastRow >= 2)
        {
            var fullRange = ws.Range(1, 1, lastRow, NdtExcelHeaders.Length);
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
            $"NdtPercentage_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx");
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private static string? SafeAt(List<string?>? list, int idx) =>
        list != null && idx < list.Count ? list[idx] : null;

    private static double? SafeDoubleAt(List<double?>? list, int idx) =>
        list != null && idx < list.Count ? list[idx] : null;

    private static bool? SafeBoolAt(List<bool?>? list, int idx) =>
        list != null && idx < list.Count ? list[idx] : null;

    private static int FindCol(Dictionary<string, int> map, params string[] names)
    {
        foreach (var n in names)
            if (map.TryGetValue(n, out var c)) return c;
        return -1;
    }

    private static void SetIfPresent(IXLWorksheet ws, int r,
        Dictionary<string, int> map, LineList row, string dbCol, string? altCol)
    {
        int c = -1;
        if (map.TryGetValue(dbCol, out c) || (altCol != null && map.TryGetValue(altCol, out c))) { }
        if (c < 1) return;
        var val = ws.Cell(r, c).GetString()?.Trim();
        if (val == null) return;
        switch (dbCol)
        {
            case "Material_Des": row.Material_Des = val; break;
            case "Material": row.Material = val; break;
            case "Design_Temperature_F": row.Design_Temperature_F = val; break;
            case "Category": row.Category = val; break;
            case "DWG_Remarks": row.DWG_Remarks = val; break;
        }
    }

    private static void SetDoubleIfPresent(IXLWorksheet ws, int r,
        Dictionary<string, int> map, LineList row, string dbCol, string? altCol)
    {
        int c = -1;
        if (map.TryGetValue(dbCol, out c) || (altCol != null && map.TryGetValue(altCol, out c))) { }
        if (c < 1) return;
        var raw = ws.Cell(r, c).GetString()?.Trim();
        if (string.IsNullOrWhiteSpace(raw)) return;
        if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var val)) return;
        switch (dbCol)
        {
            case "RT_Shop": row.RT_Shop = val; break;
            case "RT_Field": row.RT_Field = val; break;
            case "RT_Field_Shop_SW": row.RT_Field_Shop_SW = val; break;
            case "MT_Shop": row.MT_Shop = val; break;
            case "MT_Field": row.MT_Field = val; break;
            case "PT_Shop": row.PT_Shop = val; break;
            case "PT_Field": row.PT_Field = val; break;
            case "HT": row.HT = val; break;
            case "HT_After_PWHT": row.HT_After_PWHT = val; break;
        }
    }

    private static void SetBoolIfPresent(IXLWorksheet ws, int r,
        Dictionary<string, int> map, LineList row, string dbCol, string? altCol)
    {
        int c = -1;
        if (map.TryGetValue(dbCol, out c) || (altCol != null && map.TryGetValue(altCol, out c))) { }
        if (c < 1) return;
        var raw = (ws.Cell(r, c).GetString() ?? "").Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(raw)) return;
        bool? boolVal = raw is "YES" or "TRUE" or "1" or "Y" ? true
                      : raw is "NO" or "FALSE" or "0" or "N" ? false : null;
        if (boolVal == null) return;
        switch (dbCol)
        {
            case "PWHT_Y_N": row.PWHT_Y_N = boolVal; break;
            case "PWHT_20mm": row.PWHT_20mm = boolVal; break;
            case "PMI": row.PMI = boolVal; break;
        }
    }
}
