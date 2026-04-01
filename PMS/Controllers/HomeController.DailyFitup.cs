using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DfrForm()
    {
        try
        {
            var fullName = HttpContext.Session.GetString("FullName");
            if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");
            ViewBag.FullName = fullName;

            var projects = await _context.Projects_tbl
                .AsNoTracking()
                .OrderBy(p => p.Project_Name)
                .Select(p => new ProjectOption { Id = p.Project_ID, Name = p.Project_Name ?? string.Empty })
                .ToListAsync();

            var weldTypes = await _context.PMS_Weld_Type_tbl
                .AsNoTracking()
                .Where(wt => wt.W_Weld_Type != null && wt.W_Weld_Type != "")
                .GroupBy(wt => wt.W_Weld_Type!)
                .Select(g => new { Name = g.Key, MinId = g.Min(x => x.W_Type_ID) })
                .OrderBy(x => x.MinId)
                .Select(x => x.Name)
                .ToListAsync();

            var materialDescriptions = await _context.Material_Des_tbl
                .AsNoTracking()
                .Where(m => m.MATD_Description != null && m.MATD_Description != "")
                .GroupBy(m => m.MATD_Description!)
                .Select(g => new { Name = g.Key, MinSn = g.Min(x => x.MATD_SN) })
                .OrderBy(x => x.MinSn)
                .Select(x => x.Name)
                .ToListAsync();

            var materialGrades = await _context.Material_tbl
                .AsNoTracking()
                .Where(m => m.MAT_GRADE != null && m.MAT_GRADE != "")
                .GroupBy(m => m.MAT_GRADE!)
                .Select(g => new { Name = g.Key, MinSn = g.Min(x => x.MAT_SN) })
                .OrderBy(x => x.MinSn)
                .Select(x => x.Name)
                .ToListAsync();

            var vm = new DfrDailyFitupViewModel
            {
                Projects = projects,
                Locations = ["All", "Shop", "Field"],
                SelectedLocation = "All",
                HeaderView = "DWG",
                LayoutOptions = [],
                SheetOptions = [],
                LocationOptions = ["WS", "FW"],
                JAddOptions = ["NEW", "R1", "R2"],
                WeldTypeOptions = weldTypes,
                MaterialDescriptions = materialDescriptions,
                MaterialGrades = materialGrades
            };

            return View(vm);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load DfrForm");
            return View(new DfrDailyFitupViewModel());
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DailyFitup(
        int? projectId,
        string? location,
        string? headerView,
        string? layout,
        string? sheet,
        string? fitupDateFilter,
        string? fitupReportFilter,
        string? doLoad,
        string? date,
        string? FitupReportHeader,
        string? FitupReportCombined,
        int? SelectedRfiId,
        string? SelectedTacker,
        string? actualDateFilter = null)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");
        ViewBag.FullName = fullName;

        var hv = NormalizeHeaderView(headerView);
        var headerLoc = NormalizeHeaderLocation(location);
        layout = TrimToNull(layout, 50);
        sheet = TrimToNull(sheet, 50);
        fitupDateFilter = TrimToNull(fitupDateFilter, 50);
        fitupReportFilter = TrimToNull(fitupReportFilter, 50);
        date = TrimToNull(date, 50);
        FitupReportHeader = TrimToNull(FitupReportHeader, 50);
        FitupReportCombined = TrimToNull(FitupReportCombined, 100);

        var projects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_Name)
            .Select(p => new ProjectOption { Id = p.Project_ID, Name = p.Project_Name ?? string.Empty })
            .ToListAsync();

        var selectedProjectId = await GetDefaultProjectIdAsync(projectId) ?? (projects.Count > 0 ? projects.Max(p => p.Id) : 0);

        var weldersProjectId = await ResolveWeldersProjectIdAsync(selectedProjectId);

        var weldTypeQuery = _context.PMS_Weld_Type_tbl.AsNoTracking()
            .Where(wt => wt.W_Weld_Type != null && wt.W_Weld_Type != "");
        if (weldersProjectId > 0)
        {
            weldTypeQuery = weldTypeQuery.Where(wt => wt.W_Project_No == weldersProjectId);
        }

        var weldTypeOptions = await weldTypeQuery
            .GroupBy(wt => wt.W_Weld_Type!)
            .Select(g => new { Name = g.Key, MinId = g.Min(x => x.W_Type_ID) })
            .OrderBy(x => x.MinId)
            .Select(x => x.Name)
            .ToListAsync();

        if (weldTypeOptions.Count == 0)
        {
            weldTypeOptions = await _context.PMS_Weld_Type_tbl.AsNoTracking()
                .Where(wt => wt.W_Weld_Type != null && wt.W_Weld_Type != "")
                .GroupBy(wt => wt.W_Weld_Type!)
                .Select(g => new { Name = g.Key, MinId = g.Min(x => x.W_Type_ID) })
                .OrderBy(x => x.MinId)
                .Select(x => x.Name)
                .ToListAsync();
        }

        var materialDescriptions = await _context.Material_Des_tbl.AsNoTracking()
            .Where(m => m.MATD_Description != null && m.MATD_Description != "")
            .GroupBy(m => m.MATD_Description!)
            .Select(g => new { Name = g.Key, MinSn = g.Min(x => x.MATD_SN) })
            .OrderBy(x => x.MinSn)
            .Select(x => x.Name)
            .ToListAsync();

        var materialGrades = await _context.Material_tbl
            .AsNoTracking()
            .Where(m => m.MAT_GRADE != null && m.MAT_GRADE != "")
            .GroupBy(m => m.MAT_GRADE!)
            .Select(g => new { Name = g.Key, MinSn = g.Min(x => x.MAT_SN) })
            .OrderBy(x => x.MinSn)
            .Select(x => x.Name)
            .ToListAsync();

        var tackerQuery = _context.Welders_tbl.AsNoTracking()
            .Where(w => w.Welder_Symbol != null && w.Welder_Symbol != ""
                && (!w.Demobilization_Date.HasValue || (w.Status != null && w.Status.Contains("TACKER"))));

        if (selectedProjectId > 0)
        {
            tackerQuery = tackerQuery.Where(w => w.Project_Welder == weldersProjectId);
        }

        var tackerOptions = await tackerQuery
            .Select(w => new WelderSymbolOption { Symbol = w.Welder_Symbol! })
            .Distinct()
            .OrderBy(x => x.Symbol)
            .ToListAsync();

        var locationOptions = await _context.PMS_Location_tbl.AsNoTracking()
            .Where(x => x.LO_Project_No == weldersProjectId)
            .Where(x => x.LO_Location != null && x.LO_Location != "")
            .OrderBy(x => x.LO_ID)
            .Select(x => x.LO_Location!)
            .ToListAsync();
        if (locationOptions.Count == 0)
        {
            locationOptions = ["WS", "FW"];
        }

        var jAddOptions = await _context.PMS_J_Add_tbl.AsNoTracking()
            .Where(x => x.Add_Project_No == weldersProjectId)
            .Where(x => x.Add_J_Add != null && x.Add_J_Add != "")
            .GroupBy(x => x.Add_J_Add!)
            .Select(g => new { Name = g.Key, MinId = g.Min(x => x.Add_ID) })
            .OrderBy(x => x.MinId)
            .Select(x => x.Name)
            .ToListAsync();
        if (jAddOptions.Count == 0)
        {
            jAddOptions = ["NEW", "R1", "R2"];
        }

        static string? CombineReportNumbers(string? shop, string? field, string? threaded)
        {
            if (string.IsNullOrWhiteSpace(shop) && string.IsNullOrWhiteSpace(field) && string.IsNullOrWhiteSpace(threaded))
            {
                return null;
            }

            return string.Join("/", new[]
            {
                shop ?? string.Empty,
                field ?? string.Empty,
                threaded ?? string.Empty
            });
        }
        var vm = new DfrDailyFitupViewModel
        {
            Projects = projects,
            Locations = ["All", "Shop", "Field"],
            SelectedProjectId = selectedProjectId,
            SelectedLocation = headerLoc,
            HeaderView = hv,
            SearchLayout = layout,
            SearchSheet = sheet,
            SearchFitupDate = fitupDateFilter,
            SearchFitupReport = fitupReportFilter,
            LocationOptions = locationOptions,
            JAddOptions = jAddOptions,
            WeldTypeOptions = weldTypeOptions,
            TackerOptions = tackerOptions,
            MaterialDescriptions = materialDescriptions,
            MaterialGrades = materialGrades,
            SelectedRfiId = SelectedRfiId,
            SelectedTacker = SelectedTacker
        };

        var lastReportProjectId = HttpContext.Session.GetInt32("DailyFitupReportProjectId");
        bool preserveSubmittedReportNumbers = lastReportProjectId.HasValue && lastReportProjectId.Value == vm.SelectedProjectId;

        if (!string.IsNullOrWhiteSpace(date)) vm.FitupDateHeader = date;

        if (string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
        {
            if (preserveSubmittedReportNumbers && !string.IsNullOrWhiteSpace(FitupReportCombined))
            {
                var parts = FitupReportCombined.Split('/', 3, StringSplitOptions.TrimEntries);
                vm.FitupReportShop = parts.Length >= 1 ? parts[0] : null;
                vm.FitupReportField = parts.Length >= 2 ? parts[1] : null;
                vm.FitupReportThreaded = parts.Length >= 3 ? parts[2] : null;
                vm.FitupReportHeader = FitupReportCombined;
            }
        }
        else if (preserveSubmittedReportNumbers && !string.IsNullOrWhiteSpace(FitupReportHeader))
        {
            vm.FitupReportHeader = FitupReportHeader;
        }

        var defWelds = await _context.PMS_Weld_Type_tbl.AsNoTracking()
            .Where(w => w.Default_Value && (vm.SelectedProjectId <= 0 || w.W_Project_No == weldersProjectId))
            .Select(w => w.W_Weld_Type!)
            .Where(s => s != null && s != "")
            .Distinct()
            .OrderBy(s => s)
            .ToListAsync();
        if (defWelds.Count > 0)
        {
            vm.DefaultWeldTypes = defWelds;
            vm.DefaultWeldType = defWelds.FirstOrDefault();
        }

        DateTime refDate;
        if (!string.IsNullOrWhiteSpace(vm.FitupDateHeader) && DateTime.TryParse(vm.FitupDateHeader, out var tmp))
            refDate = tmp;
        else
            refDate = DateTime.Now;

        // Use a fixed reference for building the RFI dropdown so filtering/Load doesn't remove newer/older RFIs.
        // The RFI list should be based on 'now' (complete) rather than the selected FitupDateHeader which is a user filter.
        var rfiRefDate = DateTime.Now;

        // Include all RFIs with a non-null date so the dropdown is complete (don't exclude future-dated RFIs).
        var rfiBaseAll = _context.RFI_tbl.AsNoTracking()
            .Where(r => r.Date != null);
        if (weldersProjectId > 0)
        {
            rfiBaseAll = rfiBaseAll.Where(r => r.RFI_Project_No == weldersProjectId);
        }

        var rfiBaseWelding = rfiBaseAll
            .Where(r => r.SubDiscipline != null && EF.Functions.Like(r.SubDiscipline!, "%Welding%"))
            .Where(r => r.ACTIVITY.HasValue && Math.Abs(r.ACTIVITY.Value - 3.6) < 0.000001);

        var rfiBasePiping = rfiBaseAll
            .Where(r => r.SubDiscipline != null && EF.Functions.Like(r.SubDiscipline!, "%Piping%"))
            .Where(r => r.ACTIVITY.HasValue && Math.Abs(r.ACTIVITY.Value - 3.11) < 0.000001);

        var rfiListW = await rfiBaseWelding
            .OrderByDescending(r => r.Date)
            .ThenByDescending(r => r.Time)
            .Take(1000)
            .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.Date, r.Time, r.RFI_LOCATION, r.RFI_DESCRIPTION })
            .ToListAsync();

        var rfiListP = await rfiBasePiping
            .OrderByDescending(r => r.Date)
            .ThenByDescending(r => r.Time)
            .Take(1000)
            .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.Date, r.Time, r.RFI_LOCATION, r.RFI_DESCRIPTION })
            .ToListAsync();

        var weldingSeq = rfiListW.AsEnumerable();
        if (string.Equals(location, "Shop", StringComparison.OrdinalIgnoreCase))
        {
            weldingSeq = weldingSeq.Where(r => (r.RFI_LOCATION ?? string.Empty).Contains("Shop", StringComparison.OrdinalIgnoreCase));
        }
        else if (string.Equals(location, "Field", StringComparison.OrdinalIgnoreCase))
        {
            weldingSeq = weldingSeq.Where(r => !(r.RFI_LOCATION ?? string.Empty).Contains("Shop", StringComparison.OrdinalIgnoreCase));
        }

        var options = new List<RfiOption>();
        static DateTime CombineDt(DateTime? d, DateTime? t)
        {
            if (d.HasValue)
            {
                var baseD = d.Value.Date;
                if (t.HasValue) return baseD + t.Value.TimeOfDay;
                return d.Value;
            }
            if (t.HasValue) return t.Value;
            return DateTime.MinValue;
        }

        if (string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
        {
            var unified = rfiListW.Select(r => new { r, piping = false, when = CombineDt(r.Date, r.Time) })
                .Concat(rfiListP.Select(r => new { r, piping = true, when = CombineDt(r.Date, r.Time) }))
                .OrderByDescending(x => x.when);
            foreach (var x in unified)
            {
                var r = x.r;
                var prefix = x.piping
                    ? "TH | "
                    : ((r.RFI_LOCATION ?? string.Empty).Contains("Shop", StringComparison.OrdinalIgnoreCase) ? "WS | " : "FW | ");
                options.Add(new RfiOption
                {
                    Id = r.RFI_ID,
                    Display = prefix
                        + (r.SubCon_RFI_No ?? string.Empty)
                        + (r.Date.HasValue ? (" | " + r.Date.Value.ToString("dd-MMM-yyyy")) : string.Empty)
                        + (r.Time.HasValue ? (" " + r.Time.Value.ToString("HH:mm")) : string.Empty)
                        + (string.IsNullOrWhiteSpace(r.RFI_DESCRIPTION) ? string.Empty : (" | " + r.RFI_DESCRIPTION)),
                    Value = r.SubCon_RFI_No ?? string.Empty
                });
            }
        }
        else if (string.Equals(headerLoc, "Threaded", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var r in rfiListP.OrderByDescending(r => CombineDt(r.Date, r.Time)).ThenByDescending(r => r.RFI_ID))
            {
                options.Add(new RfiOption
                {
                    Id = r.RFI_ID,
                    Display = (r.SubCon_RFI_No ?? string.Empty)
                        + (r.Date.HasValue ? (" | " + r.Date.Value.ToString("dd-Mmm-yyyy")) : string.Empty)
                        + (r.Time.HasValue ? (" " + r.Time.Value.ToString("HH:mm")) : string.Empty)
                        + (string.IsNullOrWhiteSpace(r.RFI_DESCRIPTION) ? string.Empty : (" | " + r.RFI_DESCRIPTION)),
                    Value = r.SubCon_RFI_No ?? string.Empty
                });
            }
        }
        else
        {
            foreach (var r in weldingSeq.OrderByDescending(r => CombineDt(r.Date, r.Time)))
            {
                options.Add(new RfiOption
                {
                    Id = r.RFI_ID,
                    Display = (r.SubCon_RFI_No ?? string.Empty)
                        + (r.Date.HasValue ? (" | " + r.Date.Value.ToString("dd-Mmm-yyyy")) : string.Empty)
                        + (r.Time.HasValue ? (" " + r.Time.Value.ToString("HH:mm")) : string.Empty)
                        + (string.IsNullOrWhiteSpace(r.RFI_DESCRIPTION) ? string.Empty : (" | " + r.RFI_DESCRIPTION)),
                    Value = r.SubCon_RFI_No ?? string.Empty
                });
            }
        }

        vm.RfiOptions = [.. options];

        if (string.Equals(hv, "DWG", StringComparison.OrdinalIgnoreCase))
        {
            if (!vm.SelectedRfiId.HasValue && vm.RfiOptions.Count > 0)
            {
                if (string.Equals(headerLoc, "Threaded", StringComparison.OrdinalIgnoreCase))
                    vm.SelectedRfiId = rfiListP.OrderByDescending(r => CombineDt(r.Date, r.Time)).FirstOrDefault()?.RFI_ID;
                else if (string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
                    vm.SelectedRfiId = options.FirstOrDefault()?.Id;
                else
                    vm.SelectedRfiId = weldingSeq.OrderByDescending(r => CombineDt(r.Date, r.Time)).FirstOrDefault()?.RFI_ID;
            }

            if (vm.SelectedRfiId.HasValue)
            {
                var sel = rfiListW.FirstOrDefault(r => r.RFI_ID == vm.SelectedRfiId.Value)
                          ?? rfiListP.FirstOrDefault(r => r.RFI_ID == vm.SelectedRfiId.Value);
                if (sel != null && (sel.Date.HasValue || sel.Time.HasValue))
                {
                    var d2 = sel.Date ?? DateTime.Now;
                    if (sel.Time.HasValue) d2 = d2.Date + sel.Time.Value.TimeOfDay;
                    vm.FitupDateHeader = d2.ToString("dd-MMM-yyyy hh:mm tt");
                }
            }
        }

        if (string.IsNullOrWhiteSpace(vm.FitupDateHeader))
        {
            vm.FitupDateHeader = DateTime.Now.ToString("dd-Mmm-yyyy hh:mm tt");
        }

        if (!string.IsNullOrWhiteSpace(layout) && vm.SelectedProjectId > 0)
        {
            var ls = await _context.Line_Sheet_tbl.AsNoTracking()
                .Where(x => x.LS_LAYOUT_NO == layout && (string.IsNullOrWhiteSpace(sheet) || x.LS_SHEET == sheet))
                .OrderBy(x => x.LS_SHEET)
                .Select(x => new { x.Line_Sheet_ID, x.Line_ID_LS, x.LS_REV, x.LS_SHEET })
                .FirstOrDefaultAsync();
            if (ls != null)
            {
                var ll = await _context.LINE_LIST_tbl.AsNoTracking()
                    .Where(l => l.Line_ID == ls.Line_ID_LS)
                    .Select(l => new
                    {
                        l.Line_Class,
                        l.Material,
                        l.Category,
                        l.RT_Shop,
                        l.RT_Field,
                        l.PWHT_Y_N,
                        l.PWHT_20mm
                    })
                    .FirstOrDefaultAsync();
                if (ll != null)
                {
                    vm.LineClass = ll.Line_Class;
                    vm.LineMaterial = ll.Material;
                    vm.LineCategory = ll.Category;
                    vm.RtShop = ll.RT_Shop;
                    vm.RtField = ll.RT_Field;
                    vm.PwhtYN = ll.PWHT_Y_N;
                    vm.Pwht20mm = ll.PWHT_20mm;
                }
                vm.LsRev = ls.LS_REV;
            }
        }

        if (vm.SelectedProjectId > 0)
        {
            vm.LayoutOptions = await _context.DFR_tbl.AsNoTracking()
                .Where(d => d.Project_No == vm.SelectedProjectId && d.LAYOUT_NUMBER != null && d.LAYOUT_NUMBER != "")
                .Select(d => d.LAYOUT_NUMBER!)
                .Distinct()
                .OrderBy(x => x)
                .ToListAsync();
            if (string.IsNullOrWhiteSpace(layout) && vm.LayoutOptions.Count > 0)
            {
                layout = vm.LayoutOptions.First();
                vm.SearchLayout = layout;
            }

            if (!string.IsNullOrWhiteSpace(layout))
            {
                var layoutKeyUpper = layout.Trim().ToUpperInvariant();
                #pragma warning disable CA1862 // ToUpper is required for EF Core SQL translation
                                vm.SheetOptions = await _context.DFR_tbl.AsNoTracking()
                                    .Where(d => d.Project_No == vm.SelectedProjectId && d.LAYOUT_NUMBER != null && d.LAYOUT_NUMBER.Trim().ToUpper() == layoutKeyUpper && d.SHEET != null && d.SHEET != "")
                                    .Select(d => d.SHEET!)
                                    .Distinct()
                                    .OrderBy(x => x)
                                    .ToListAsync();
                #pragma warning restore CA1862

                if (vm.SheetOptions.Count > 0 && string.IsNullOrWhiteSpace(vm.SearchSheet))
                {
                    vm.SearchSheet = vm.SheetOptions.First();
                }
            }
        }

        var baseReports = _context.DFR_tbl.AsNoTracking().Where(d => d.Project_No == vm.SelectedProjectId);
        bool headerIsThreaded = headerLoc.Equals("Threaded", StringComparison.OrdinalIgnoreCase);

        if (string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(vm.FitupReportShop))
                vm.FitupReportShop = await GetNextReportNumberAsync(baseReports.Where(d => d.LOCATION == "WS"));
            if (string.IsNullOrWhiteSpace(vm.FitupReportField))
                vm.FitupReportField = await GetNextReportNumberAsync(baseReports.Where(d => d.LOCATION != null && d.LOCATION != "WS"));
            if (string.IsNullOrWhiteSpace(vm.FitupReportThreaded))
                vm.FitupReportThreaded = await GetNextReportNumberAsync(baseReports.Where(d => d.LOCATION != null && d.LOCATION != "WS" && d.WELD_TYPE != null && EF.Functions.Like(d.WELD_TYPE, "%TH%")));
        }
        else if (string.Equals(headerLoc, "Shop", StringComparison.OrdinalIgnoreCase) || string.Equals(headerLoc, "WS", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(vm.FitupReportHeader))
                vm.FitupReportHeader = await GetNextReportNumberAsync(baseReports.Where(d => d.LOCATION == "WS"));
        }
        else if (headerIsThreaded)
        {
            if (string.IsNullOrWhiteSpace(vm.FitupReportHeader))
                vm.FitupReportHeader = await GetNextReportNumberAsync(baseReports.Where(d => d.LOCATION != null && d.LOCATION != "WS" && d.WELD_TYPE != null && EF.Functions.Like(d.WELD_TYPE, "%TH%")));
        }
        else if (string.IsNullOrWhiteSpace(vm.FitupReportHeader))
        {
            vm.FitupReportHeader = await GetNextReportNumberAsync(baseReports.Where(d => d.LOCATION != null && d.LOCATION != "WS"));
        }

        if (string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
        {
            vm.FitupReportHeader = CombineReportNumbers(vm.FitupReportShop, vm.FitupReportField, vm.FitupReportThreaded);
        }

        if (vm.SelectedProjectId > 0)
        {
            HttpContext.Session.SetInt32("DailyFitupReportProjectId", vm.SelectedProjectId);
        }

        bool shouldLoad = string.Equals(doLoad, "1", StringComparison.OrdinalIgnoreCase);
        if (!shouldLoad)
        {
            return View(vm);
        }

        var rowsQ = _context.DFR_tbl.AsNoTracking().Where(d => d.Project_No == vm.SelectedProjectId);

        bool isThreaded = headerLoc.Equals("Threaded", StringComparison.OrdinalIgnoreCase);
        if (isThreaded)
        {
            rowsQ = rowsQ.Where(d => d.WELD_TYPE != null && EF.Functions.Like(d.WELD_TYPE, "%TH%"));
        }
        else
        {
            var locCode = MapHeaderLocation(headerLoc);
            if (!string.IsNullOrWhiteSpace(locCode) && !string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(locCode, "FW", StringComparison.OrdinalIgnoreCase))
                {
                    rowsQ = rowsQ.Where(d => d.LOCATION == null || d.LOCATION == "FW");
                }
                else
                {
                    rowsQ = rowsQ.Where(d => d.LOCATION != null && d.LOCATION == locCode);
                }
            }
        }

        if (string.Equals(hv, "DWG", StringComparison.OrdinalIgnoreCase))
        {
            if (!string.IsNullOrWhiteSpace(layout)) rowsQ = rowsQ.Where(d => d.LAYOUT_NUMBER == layout);
            if (!string.IsNullOrWhiteSpace(sheet)) rowsQ = rowsQ.Where(d => d.SHEET == sheet);
        }
        else if (string.Equals(hv, "Date", StringComparison.OrdinalIgnoreCase))
        {
            var appliedDateFilter = false;

            if (!string.IsNullOrWhiteSpace(actualDateFilter) && DateTime.TryParse(actualDateFilter, out var actual))
            {
                var start = actual.Date;
                var end = start.AddDays(1);
                rowsQ = rowsQ.Where(x => _context.DWR_tbl.Any(w => w.Joint_ID_DWR == x.Joint_ID && w.ACTUAL_DATE_WELDED >= start && w.ACTUAL_DATE_WELDED < end));
                appliedDateFilter = true;
            }

            if (!appliedDateFilter)
            {
                if (!string.IsNullOrWhiteSpace(fitupDateFilter) && DateTime.TryParse(fitupDateFilter, out var d))
                {
                    var start = d.Date;
                    var end = start.AddDays(1);
                    rowsQ = rowsQ.Where(x => x.FITUP_DATE >= start && x.FITUP_DATE < end);
                }
                else
                {
                    rowsQ = rowsQ.Where(x => false);
                }
            }
        }
        else if (string.Equals(hv, "Report", StringComparison.OrdinalIgnoreCase))
        {
            if (!string.IsNullOrWhiteSpace(fitupReportFilter))
            {
                var key = fitupReportFilter.Trim();
                rowsQ = rowsQ.Where(x => x.FITUP_INSPECTION_QR_NUMBER != null && x.FITUP_INSPECTION_QR_NUMBER.Contains(key));
            }
            else
            {
                rowsQ = rowsQ.Where(x => false);
            }
        }

        const int rowLimit = 1000;
        vm.RowLimit = rowLimit;

        var rowsRaw = await rowsQ
            .OrderBy(d => d.LAYOUT_NUMBER)
            .ThenBy(d => d.SHEET)
            .ThenBy(d => d.WELD_NUMBER)
            .Select(d => new
            {
                d.Joint_ID,
                d.WELD_NUMBER,
                d.LOCATION,
                d.LAYOUT_NUMBER,
                d.J_Add,
                d.WELD_TYPE,
                d.SHEET,
                d.DFR_REV,
                d.SPOOL_NUMBER,
                d.Spool_ID_DFR,
                d.Line_Sheet_ID_DFR,
                d.DIAMETER,
                d.SCHEDULE,
                d.MATERIAL_A,
                d.MATERIAL_B,
                d.GRADE_A,
                d.GRADE_B,
                d.HEAT_NUMBER_A,
                d.HEAT_NUMBER_B,
                d.WPS_ID_DFR,
                d.FITUP_DATE,
                d.FITUP_INSPECTION_QR_NUMBER,
                d.RFI_ID_DFR,
                d.TACK_WELDER,
                d.OL_DIAMETER,
                d.OL_SCHEDULE,
                d.OL_Thick,
                d.Deleted,
                d.Cancelled,
                d.Fitup_Confirmed,
                d.HOLD_DFR,
                d.HOLD_DFR_Date_D,
                d.DFR_REMARKS,
                d.DFR_Updated_By,
                d.DFR_Updated_Date,
                d.DFR_Hold_Date,
                d.DFR_Hold_Release_Date
            })
            .Take(rowLimit + 1)
            .ToListAsync();

        vm.IsTruncated = rowsRaw.Count > rowLimit;
        vm.RowsReturned = Math.Min(rowsRaw.Count, rowLimit);
        if (vm.IsTruncated)
        {
            rowsRaw = rowsRaw.Take(rowLimit).ToList();
        }

        var jointIds = rowsRaw.Select(x => x.Joint_ID).Distinct().ToList();
        var wpsIds = rowsRaw.Where(x => x.WPS_ID_DFR.HasValue).Select(x => x.WPS_ID_DFR!.Value).Distinct().ToList();
        var rfiIds = rowsRaw.Where(x => x.RFI_ID_DFR.HasValue).Select(x => x.RFI_ID_DFR!.Value).Distinct().ToList();
        var spoolIds = rowsRaw.Where(x => x.Spool_ID_DFR.HasValue).Select(x => x.Spool_ID_DFR!.Value).Distinct().ToList();
        var lineSheetIds = rowsRaw.Where(x => x.Line_Sheet_ID_DFR.HasValue).Select(x => x.Line_Sheet_ID_DFR!.Value).Distinct().ToList();
        var layoutKeyList = rowsRaw
            .Select(x => (x.LAYOUT_NUMBER ?? string.Empty).Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var wpsLookup = wpsIds.Count == 0
            ? new Dictionary<int, string>()
            : await _context.WPS_tbl.AsNoTracking()
                .Where(w => wpsIds.Contains(w.WPS_ID))
                .Select(w => new { w.WPS_ID, w.WPS })
                .ToDictionaryAsync(x => x.WPS_ID, x => x.WPS ?? string.Empty);

        var rfiLookup = rfiIds.Count == 0
            ? new Dictionary<int, string>()
            : await _context.RFI_tbl.AsNoTracking()
                .Where(r => rfiIds.Contains(r.RFI_ID))
                .Select(r => new { r.RFI_ID, r.SubCon_RFI_No })
                .ToDictionaryAsync(x => x.RFI_ID, x => x.SubCon_RFI_No ?? string.Empty);

        var dwrRows = jointIds.Count == 0
            ? new List<Dwr>()
            : await _context.DWR_tbl.AsNoTracking()
                .Where(w => jointIds.Contains(w.Joint_ID_DWR))
                .ToListAsync();

        var dwrLookup = dwrRows
            .GroupBy(x => x.Joint_ID_DWR)
            .ToDictionary(g => g.Key, g => g.OrderByDescending(x => x.DWR_Updated_Date ?? DateTime.MinValue).First());

        var userIds = new HashSet<int>();
        foreach (var row in rowsRaw)
        {
            if (row.DFR_Updated_By.HasValue)
            {
                userIds.Add(row.DFR_Updated_By.Value);
            }
        }
        foreach (var dwr in dwrRows)
        {
            if (dwr.DWR_Updated_By.HasValue)
            {
                userIds.Add(dwr.DWR_Updated_By.Value);
            }
        }

        var userLookup = userIds.Count == 0
            ? new Dictionary<int, string>()
            : await _context.PMS_Login_tbl.AsNoTracking()
                .Where(u => userIds.Contains(u.UserID))
                .Select(u => new
                {
                    u.UserID,
                    Full = (((u.FirstName ?? string.Empty).Trim() + " " + (u.LastName ?? string.Empty).Trim()).Trim())
                })
                .ToDictionaryAsync(x => x.UserID, x => x.Full);

        var spoolHoldLookup = spoolIds.Count == 0
            ? new Dictionary<int, bool>()
            : (await _context.SP_Release_tbl.AsNoTracking()
                .Where(s => spoolIds.Contains(s.Spool_ID))
                .Select(s => new { s.Spool_ID, s.SP_Hold_Date, s.SP_Hold_Release_Date })
                .ToListAsync())
                .GroupBy(x => x.Spool_ID)
                .ToDictionary(g => g.Key, g => g.Any(x => x.SP_Hold_Date != null && x.SP_Hold_Release_Date == null));

        var sheetHoldLookup = lineSheetIds.Count == 0
            ? new Dictionary<int, bool>()
            : (await _context.Line_Sheet_tbl.AsNoTracking()
                .Where(ls => lineSheetIds.Contains(ls.Line_Sheet_ID))
                .Select(ls => new { ls.Line_Sheet_ID, ls.LS_Hold_Date, ls.LS_Hold_Release_Date })
                .ToListAsync())
                .ToDictionary(x => x.Line_Sheet_ID, x => x.LS_Hold_Date != null && x.LS_Hold_Release_Date == null);

        var layoutClassLookup = layoutKeyList.Count == 0
            ? new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
            : (await _context.LINE_LIST_tbl.AsNoTracking()
                .Where(l => layoutKeyList.Contains((l.LAYOUT_NO ?? string.Empty).Trim()))
                .Select(l => new { Layout = (l.LAYOUT_NO ?? string.Empty).Trim(), l.Line_Class })
                .ToListAsync())
                .GroupBy(x => x.Layout, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Select(x => x.Line_Class).FirstOrDefault(), StringComparer.OrdinalIgnoreCase);

        var rows = rowsRaw.Select(x =>
        {
            var layoutKey = (x.LAYOUT_NUMBER ?? string.Empty).Trim();
            layoutClassLookup.TryGetValue(layoutKey, out var lineClass);
            var hasSheetHold = x.Line_Sheet_ID_DFR.HasValue && sheetHoldLookup.TryGetValue(x.Line_Sheet_ID_DFR.Value, out var sHold) && sHold;
            var hasSpoolHold = x.Spool_ID_DFR.HasValue && spoolHoldLookup.TryGetValue(x.Spool_ID_DFR.Value, out var spHold) && spHold;
            var updatedBy = x.DFR_Updated_By.HasValue && userLookup.TryGetValue(x.DFR_Updated_By.Value, out var updater) ? updater : null;
            var wps = x.WPS_ID_DFR.HasValue && wpsLookup.TryGetValue(x.WPS_ID_DFR.Value, out var wpsVal) ? wpsVal : null;
            var rfiNo = x.RFI_ID_DFR.HasValue && rfiLookup.TryGetValue(x.RFI_ID_DFR.Value, out var rfiVal) ? rfiVal : null;
            var dwr = dwrLookup.TryGetValue(x.Joint_ID, out var dwrVal) ? dwrVal : null;

            return new DfrRowVm
            {
                JointId = x.Joint_ID,
                JointNo = (x.LOCATION ?? string.Empty) + (string.IsNullOrEmpty(x.WELD_NUMBER) ? string.Empty : ("-" + x.WELD_NUMBER)) + (string.IsNullOrWhiteSpace(x.J_Add) || x.J_Add == "NEW" ? string.Empty : x.J_Add),
                WeldNumber = x.WELD_NUMBER,
                Location = x.LOCATION,
                LayoutNumber = x.LAYOUT_NUMBER,
                JAdd = x.J_Add,
                WeldType = x.WELD_TYPE,
                Sheet = x.SHEET,
                Rev = x.DFR_REV,
                SpoolNo = x.SPOOL_NUMBER,
                DiaIn = x.DIAMETER,
                Sch = x.SCHEDULE,
                MaterialA = x.MATERIAL_A,
                MaterialB = x.MATERIAL_B,
                GradeA = x.GRADE_A,
                GradeB = x.GRADE_B,
                HeatNumberA = x.HEAT_NUMBER_A,
                HeatNumberB = x.HEAT_NUMBER_B,
                WpsId = x.WPS_ID_DFR,
                Wps = wps,
                FitupDate = x.FITUP_DATE.HasValue ? x.FITUP_DATE.Value.ToString("yyyy-MM-dd'T'HH:mm") : null,
                ActualDate = dwr?.ACTUAL_DATE_WELDED.HasValue == true ? dwr.ACTUAL_DATE_WELDED.Value.ToString("yyyy-MM-dd'T'HH:mm") : null,
                FitupReport = x.FITUP_INSPECTION_QR_NUMBER,
                RfiId = x.RFI_ID_DFR,
                RfiNo = rfiNo,
                TackWelder = x.TACK_WELDER,
                OlDia = x.OL_DIAMETER,
                OlSch = x.OL_SCHEDULE,
                OlThick = x.OL_Thick,
                Deleted = x.Deleted,
                Cancelled = x.Cancelled,
                FitupConfirmed = x.Fitup_Confirmed,
                HoldDfr = x.HOLD_DFR,
                HoldDfrDate = x.HOLD_DFR_Date_D.HasValue ? x.HOLD_DFR_Date_D.Value.ToString("yyyy-MM-dd") : null,
                Remarks = x.DFR_REMARKS,
                PREHEAT_TEMP_C = dwr?.PREHEAT_TEMP_C,
                UpdatedBy = string.IsNullOrWhiteSpace(updatedBy) ? null : updatedBy,
                UpdatedDate = x.DFR_Updated_Date.HasValue ? x.DFR_Updated_Date.Value.ToString("dd-Mmm-yyyy hh:mm tt", CultureInfo.InvariantCulture) : null,
                OpenClosed = dwr?.Open_Closed,
                IPOrT = dwr?.IP_or_T,
                DfrOnHold = x.DFR_Hold_Date != null && x.DFR_Hold_Release_Date == null,
                SpOnHold = hasSpoolHold,
                SheetOnHold = hasSheetHold,
                LineClass = lineClass,
                ROOT_A = dwr?.ROOT_A,
                ROOT_B = dwr?.ROOT_B,
                FILL_A = dwr?.FILL_A,
                FILL_B = dwr?.FILL_B,
                CAP_A = dwr?.CAP_A,
                CAP_B = dwr?.CAP_B,
                Weld_Confirmed = dwr?.Weld_Confirmed ?? false,
                POST_VISUAL_INSPECTION_QR_NO = dwr?.POST_VISUAL_INSPECTION_QR_NO
            };
        }).ToList();

        vm.Rows = rows;

        return View(vm);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetLayoutsForProject(int projectId)
    {
        if (projectId <= 0) return Json(Array.Empty<string>());
        var list = await _context.DFR_tbl.AsNoTracking()
            .Where(d => d.Project_No == projectId && d.LAYOUT_NUMBER != null && d.LAYOUT_NUMBER != "")
            .Select(d => d.LAYOUT_NUMBER!)
            .Distinct()
            .OrderBy(x => x)
            .ToListAsync();
        return Json(list);
    }
    private static string NormalizeHeaderView(string? headerView)
    {
        var hv = (headerView ?? string.Empty).Trim();
        if (hv.Equals("date", StringComparison.OrdinalIgnoreCase)) return "Date";
        if (hv.Equals("report", StringComparison.OrdinalIgnoreCase)) return "Report";
        return "DWG";
    }

    private static string NormalizeHeaderLocation(string? location)
    {
        var loc = (location ?? string.Empty).Trim();
        if (loc.Equals("shop", StringComparison.OrdinalIgnoreCase) || loc.Equals("ws", StringComparison.OrdinalIgnoreCase)) return "Shop";
        if (loc.Equals("field", StringComparison.OrdinalIgnoreCase) || loc.Equals("fw", StringComparison.OrdinalIgnoreCase)) return "Field";
        if (loc.Equals("threaded", StringComparison.OrdinalIgnoreCase) || loc.Equals("th", StringComparison.OrdinalIgnoreCase)) return "Threaded";
        return "All";
    }

    private static string? TrimToNull(string? value, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var trimmed = value.Trim();
        if (trimmed.Length > maxLength) trimmed = trimmed[..maxLength];
        return trimmed.Length == 0 ? null : trimmed;
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetSheetsForLayout(int projectId, string layout)
    {
        var layoutKey = TrimToNull(layout, 50);
        if (projectId <= 0 || string.IsNullOrWhiteSpace(layoutKey)) return Json(Array.Empty<string>());

        var layoutKeyUpper = layoutKey.ToUpperInvariant();

        #pragma warning disable CA1862 // ToUpper is required for EF Core SQL translation
                var list = await _context.DFR_tbl.AsNoTracking()
                    .Where(d => d.Project_No == projectId)
                    .Where(d => d.LAYOUT_NUMBER != null && d.LAYOUT_NUMBER.Trim().ToUpper() == layoutKeyUpper)
                    .Where(d => d.SHEET != null && d.SHEET != "")
                    .Select(d => d.SHEET!)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToListAsync();
        #pragma warning restore CA1862
        return Json(list);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AddDfrRow(int projectId, string location, string layout, string sheet)
    {
        string? weldNumber = null;
        string locCode = string.Empty;
        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            if (!userId.HasValue)
                return Json(new { success = false, message = "Session expired" });
            if (projectId <= 0)
                return Json(new { success = false, message = "Project required" });
            if (string.IsNullOrWhiteSpace(layout) || string.IsNullOrWhiteSpace(sheet))
                return Json(new { success = false, message = "Layout/Sheet required" });
            var locCodeRaw = (location ?? string.Empty).Trim();
            if (string.Equals(locCodeRaw, "All", StringComparison.OrdinalIgnoreCase))
                return Json(new { success = false, message = "Select specific location (Shop/Field)" });

            var normalizedLoc = locCodeRaw.ToUpperInvariant();
            if (normalizedLoc == "SHOP" || normalizedLoc.StartsWith("WS")) locCode = "WS";
            else if (normalizedLoc == "FIELD" || normalizedLoc.StartsWith("FW")) locCode = "FW";
            else if (normalizedLoc.StartsWith("TH") || normalizedLoc.Contains("THREAD")) locCode = "TH";
            else locCode = "FW";

            string weldType = locCode switch
            {
                "WS" => "BW",
                "TH" => "TH",
                _ => "FW"
            };
            const int maxAttempts = 5;
            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var scope = _context.DFR_tbl.AsNoTracking()
                    .Where(d => d.Project_No == projectId && d.LAYOUT_NUMBER == layout && d.SHEET == sheet && d.J_Add == "NEW" && d.WELD_NUMBER != null && d.WELD_NUMBER != "");

                var maxExisting = await scope
                    .Select(d => d.WELD_NUMBER!)
                    .OrderByDescending(s => s.Length)
                    .ThenByDescending(s => s)
                    .FirstOrDefaultAsync();

                int numericWidth = Math.Clamp(string.IsNullOrEmpty(maxExisting) ? 4 : maxExisting!.Length, 4, 6);
                var candidateValue = ExtractNumericValue(maxExisting) + 1;
                weldNumber = candidateValue.ToString("D" + numericWidth);

                var duplicateExists = await scope.AnyAsync(d => d.WELD_NUMBER == weldNumber);
                if (duplicateExists)
                {
                    _logger.LogWarning("AddDfrRow duplicate weld number candidate {WeldNumber} detected (attempt {Attempt}). Retrying.", weldNumber, attempt);
                    await Task.Delay(25);
                    continue;
                }

                var entity = new Dfr
                {
                    Project_No = projectId,
                    LOCATION = locCode,
                    LAYOUT_NUMBER = Clean(layout, 10),
                    SHEET = Clean(sheet, 5),
                    WELD_NUMBER = Clean(weldNumber, 6),
                    WELD_TYPE = NormalizeWeldType(weldType),
                    J_Add = "NEW",
                    Deleted = false,
                    Cancelled = false,
                    Fitup_Confirmed = false,
                    DFR_Updated_By = userId.Value,
                    DFR_Updated_Date = AppClock.Now
                };

                await using var tx = await _context.Database.BeginTransactionAsync();
                _context.DFR_tbl.Add(entity);
                try
                {
                    await _context.SaveChangesAsync();
                    try
                    {
                        await UpdateLineSheetAndSpoolRefsAsync(entity);
                        await tx.CommitAsync();
                    }
                    catch (Exception linkEx)
                    {
                        _logger.LogWarning(linkEx, "AddDfrRow link refs failed Joint_ID={JointId}; rolling back", entity.Joint_ID);
                        await tx.RollbackAsync();
                        return Json(new { success = false, message = "Error linking joint references: " + (linkEx.Message.Length < 200 ? linkEx.Message : "Linking error.") });
                    }
                }
                catch (DbUpdateException dbEx)
                {
                    string inner = dbEx.InnerException?.Message ?? dbEx.Message;
                    _logger.LogError(dbEx, "AddDfrRow DbUpdateException Project={Project} Layout={Layout} Sheet={Sheet} Weld={Weld} Loc={Loc}", projectId, layout, sheet, weldNumber, locCode);
                    await tx.RollbackAsync();
                    _context.Entry(entity).State = EntityState.Detached;
                    if (inner.Contains("UNIQUE", StringComparison.OrdinalIgnoreCase) || inner.Contains("duplicate", StringComparison.OrdinalIgnoreCase))
                    {
                        if (attempt == maxAttempts)
                        {
                            return Json(new { success = false, message = "Duplicate weld number. Reload and try again." });
                        }
                        await Task.Delay(25);
                        continue;
                    }
                    return Json(new { success = false, message = "DB save failed: " + inner.Split('\n')[0] });
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "AddDfrRow save failed Project={Project} Layout={Layout} Sheet={Sheet} Weld={Weld} Loc={Loc}", projectId, layout, sheet, weldNumber, locCode);
                    await tx.RollbackAsync();
                    _context.Entry(entity).State = EntityState.Detached;
                    return Json(new { success = false, message = "Error saving joint: " + (ex.Message.Length < 200 ? ex.Message : "Save error.") });
                }

                return Json(new { success = true, id = entity.Joint_ID, weldNumber = entity.WELD_NUMBER, sheet = entity.SHEET, location = entity.LOCATION });
            }

            return Json(new { success = false, message = "Failed to generate weld number" });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "AddDfrRow failed for Project={Project} Layout={Layout} Sheet={Sheet} Weld={WeldNumber} Location={Location}. Exception: {ExceptionMessage} StackTrace: {StackTrace}",
                projectId, layout, sheet, weldNumber, locCode, ex.Message, ex.StackTrace);

            var userMessage = "An unexpected error occurred while adding the joint.";
            if (!string.IsNullOrWhiteSpace(ex.Message))
            {
                if (ex.Message.Contains("UNIQUE", StringComparison.OrdinalIgnoreCase) || ex.Message.Contains("duplicate", StringComparison.OrdinalIgnoreCase))
                    userMessage = "Duplicate weld number. Please reload and try again.";
                else if (ex.Message.Length < 200)
                    userMessage = ex.Message;
            }
            return Json(new { success = false, message = userMessage });
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteDfrRow(int id)
    {
        try
        {
            var d = await _context.DFR_tbl.FirstOrDefaultAsync(x => x.Joint_ID == id);
            if (d == null) return NotFound();

            int projectId = d.Project_No;
            var layout = (d.LAYOUT_NUMBER ?? string.Empty).Trim();
            var sheet = (d.SHEET ?? string.Empty).Trim();

            _context.DFR_tbl.Remove(d);
            await _context.SaveChangesAsync();

            if (!string.IsNullOrWhiteSpace(layout))
            {
                try { await PruneUnusedSpReleaseRowsForScopeAsync(projectId, layout, sheet); }
                catch (Exception exPrune) { _logger.LogWarning(exPrune, "Scoped prune after delete failed for Layout={Layout} Sheet={Sheet}", layout, sheet); }
            }

            return Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteDfrRow failed for {Id}", id);
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetDailyFitupExists(int projectId, string location, string fitupDateIso)
    {
        try
        {
            if (projectId <= 0 || string.IsNullOrWhiteSpace(fitupDateIso)) return BadRequest("Invalid");
            if (!DateTime.TryParse(fitupDateIso, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt)) return BadRequest("Invalid date");
            var day = dt.Date;
            var locCode = MapHeaderLocation(location) ?? "FW";
            var row = await _context.PMS_Updated_Confirmed_tbl
                .FirstOrDefaultAsync(x => x.U_C_Project_No == projectId && x.U_C_Location == locCode && x.Updated_Confirmed_Date.Date == day);
            bool updatedExists = row != null && row.Fitup_Updated_Date.HasValue;
            bool confirmedExists = row != null && row.Fitup_Confirmed_Date.HasValue;
            return Json(new { updatedExists, confirmedExists });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetDailyFitupExists failed");
            return Json(new { updatedExists = false, confirmedExists = false });
        }
    }

    private async Task<decimal?> ComputeFitupDiaAsync(int projectId, DateTime day, string? locCode)
    {
        var start = day.Date;
        var end = start.AddDays(1);
        IQueryable<Dfr> q = _context.DFR_tbl.AsNoTracking()
            .Where(d => d.Project_No == projectId && d.FITUP_DATE >= start && d.FITUP_DATE < end && d.DIAMETER.HasValue);

        if (!string.IsNullOrWhiteSpace(locCode))
        {
            if (string.Equals(locCode, "WS", StringComparison.OrdinalIgnoreCase))
            {
                q = q.Where(d => d.LOCATION == "WS");
                q = q.Where(d => d.WELD_TYPE == null || (
                    !EF.Functions.Like(d.WELD_TYPE!, "SP%") &&
                    !EF.Functions.Like(d.WELD_TYPE!, "FJ%") &&
                    !EF.Functions.Like(d.WELD_TYPE!, "%TH%")));
            }
            else if (string.Equals(locCode, "FW", StringComparison.OrdinalIgnoreCase))
            {
                q = q.Where(d => d.LOCATION == null || d.LOCATION != "WS");
                q = q.Where(d => d.WELD_TYPE == null || (
                    !EF.Functions.Like(d.WELD_TYPE!, "SP%") &&
                    !EF.Functions.Like(d.WELD_TYPE!, "FJ%") &&
                    !EF.Functions.Like(d.WELD_TYPE!, "%TH%")));
            }
            else if (string.Equals(locCode, "TH", StringComparison.OrdinalIgnoreCase))
            {
                q = q.Where(d => d.WELD_TYPE != null && EF.Functions.Like(d.WELD_TYPE!, "%TH%"));
            }
        }

        return await q.SumAsync(d => (decimal?)d.DIAMETER);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public Task<IActionResult> CompleteDailyFitup(int projectId, string location, string fitupDateIso)
    {
        return HandleDailyFitupStatusAsync(projectId, location, fitupDateIso, isConfirm: false);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public Task<IActionResult> ConfirmDailyFitup(int projectId, string location, string fitupDateIso)
    {
        return HandleDailyFitupStatusAsync(projectId, location, fitupDateIso, isConfirm: true);
    }

    private async Task<IActionResult> HandleDailyFitupStatusAsync(int projectId, string location, string fitupDateIso, bool isConfirm)
    {
        try
        {
            if (projectId <= 0 || string.IsNullOrWhiteSpace(fitupDateIso)) return BadRequest("Invalid");
            if (!DateTime.TryParse(fitupDateIso, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt)) return BadRequest("Invalid date");
            var day = dt.Date;
            var headerRaw = (location ?? string.Empty).Trim();
            var userId = HttpContext.Session.GetInt32("UserID");
            var now = AppClock.Now;
            var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);

            if (string.Equals(headerRaw, "All", StringComparison.OrdinalIgnoreCase))
            {
                var start = day.Date;
                var end = start.AddDays(1);
                var hasWs = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Project_No == projectId && d.FITUP_DATE >= start && d.FITUP_DATE < end && d.LOCATION == "WS");
                var hasFw = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Project_No == projectId && d.FITUP_DATE >= start && d.FITUP_DATE < end && (d.LOCATION == null || d.LOCATION != "WS"));
                var hasThType = await _context.PMS_Weld_Type_tbl.AsNoTracking().AnyAsync(w => w.W_Project_No == weldersProjectId && w.W_Weld_Type != null && EF.Functions.Like(w.W_Weld_Type, "%TH%"));
                var hasThDay = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Project_No == projectId && d.FITUP_DATE >= start && d.FITUP_DATE < end && d.WELD_TYPE != null && EF.Functions.Like(d.WELD_TYPE, "%TH%"));
                var targets = new List<string>();
                if (hasWs) targets.Add("WS");
                if (hasFw) targets.Add("FW");
                if (hasThType && hasThDay) targets.Add("TH");
                foreach (var lc in targets)
                {
                    var row = await _context.PMS_Updated_Confirmed_tbl.FirstOrDefaultAsync(x => x.U_C_Project_No == projectId && x.U_C_Location == lc && x.Updated_Confirmed_Date.Date == day);
                    if (row == null)
                    {
                        row = new UpdatedConfirmed { U_C_Project_No = projectId, U_C_Location = lc, Updated_Confirmed_Date = day };
                        _context.PMS_Updated_Confirmed_tbl.Add(row);
                    }

                    if (isConfirm)
                    {
                        row.Fitup_Confirmed_Date = now;
                        row.Fitup_Confirmed_By = userId;
                        row.Fitup_Confirmed_Dia = await ComputeFitupDiaAsync(projectId, row.Updated_Confirmed_Date.Date, lc);
                    }
                    else
                    {
                        row.Fitup_Updated_Date = now;
                        row.Fitup_Updated_By = userId;
                        row.Fitup_Dia = await ComputeFitupDiaAsync(projectId, row.Updated_Confirmed_Date.Date, lc);
                    }
                }
                await _context.SaveChangesAsync();
                return Ok();
            }

            string singleLoc = (string.Equals(headerRaw, "Threaded", StringComparison.OrdinalIgnoreCase) || headerRaw.StartsWith("TH", StringComparison.OrdinalIgnoreCase))
                ? "TH"
                : (MapHeaderLocation(location) ?? "FW");

            var rowSingle = await _context.PMS_Updated_Confirmed_tbl.FirstOrDefaultAsync(x => x.U_C_Project_No == projectId && x.U_C_Location == singleLoc && x.Updated_Confirmed_Date.Date == day);
            var sStart = day.Date; var sEnd = sStart.AddDays(1);
            bool canInsert = true;
            if (singleLoc == "WS")
                canInsert = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Project_No == projectId && d.FITUP_DATE >= sStart && d.FITUP_DATE < sEnd && d.LOCATION == "WS");
            else if (singleLoc == "FW")
                canInsert = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Project_No == projectId && d.FITUP_DATE >= sStart && d.FITUP_DATE < sEnd && (d.LOCATION == null || d.LOCATION != "WS"));
            else if (singleLoc == "TH")
            {
                var hasThTypeS = await _context.PMS_Weld_Type_tbl.AsNoTracking().AnyAsync(w => w.W_Project_No == weldersProjectId && w.W_Weld_Type != null && EF.Functions.Like(w.W_Weld_Type, "%TH%"));
                var hasThDayS = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Project_No == projectId && d.FITUP_DATE >= sStart && d.FITUP_DATE < sEnd && d.WELD_TYPE != null && EF.Functions.Like(d.WELD_TYPE, "%TH%"));
                canInsert = hasThTypeS && hasThDayS;
            }
            if (rowSingle == null)
            {
                if (!canInsert)
                {
                    return Ok();
                }
                rowSingle = new UpdatedConfirmed { U_C_Project_No = projectId, U_C_Location = singleLoc, Updated_Confirmed_Date = day };
                _context.PMS_Updated_Confirmed_tbl.Add(rowSingle);
            }

            if (isConfirm)
            {
                rowSingle.Fitup_Confirmed_Date = now;
                rowSingle.Fitup_Confirmed_By = userId;
                rowSingle.Fitup_Confirmed_Dia = await ComputeFitupDiaAsync(projectId, rowSingle.Updated_Confirmed_Date.Date, singleLoc);
            }
            else
            {
                rowSingle.Fitup_Updated_Date = now;
                rowSingle.Fitup_Updated_By = userId;
                rowSingle.Fitup_Dia = await ComputeFitupDiaAsync(projectId, rowSingle.Updated_Confirmed_Date.Date, singleLoc);
            }
            await _context.SaveChangesAsync();
            return Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CompleteDailyFitup {Action} failed", isConfirm ? "confirm" : "complete");
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetRfiDateTime(int id)
    {
        try
        {
            var rfi = await _context.Set<PMS.Models.Rfi>()
                .AsNoTracking()
                .Where(r => r.RFI_ID == id)
                .Select(r => new { r.Date, r.Time })
                .FirstOrDefaultAsync();
            if (rfi == null || rfi.Date == null)
                return NotFound();
            var baseDate = rfi.Date.Value;
            var timePart = rfi.Time?.TimeOfDay ?? TimeSpan.Zero;
            var combined = baseDate.Date + timePart;
            var iso = combined.ToString("yyyy-MM-dd'T'HH:mm");
            return Ok(new { iso });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in GetRfiDateTime for id {Id}", id);
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetRfiOptions(int projectId, string location, string? fitupDateIso)
    {
        try
        {
            if (projectId <= 0) return BadRequest("Invalid project");

            var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);
            if (weldersProjectId <= 0) return BadRequest("Invalid project");

            string loc = (location ?? string.Empty).Trim().ToUpperInvariant();
            if (loc == "SHOP") loc = "WS";
            if (loc == "FIELD") loc = "FW";
            if (loc.StartsWith("WS")) loc = "WS";
            else if (loc.StartsWith("FW")) loc = "FW";
            else if (loc.StartsWith("TH") || loc.Contains("THREAD")) loc = "TH";

            DateTime? fitDay = null;
            if (!string.IsNullOrWhiteSpace(fitupDateIso) && DateTime.TryParse(fitupDateIso, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
            {
                fitDay = dt.Date;
            }

            IQueryable<Rfi> baseQuery = _context.RFI_tbl.AsNoTracking()
                .Where(r => r.RFI_Project_No == weldersProjectId);

            if (loc == "TH")
            {
                baseQuery = baseQuery.Where(r => r.SubDiscipline != null && EF.Functions.Like(r.SubDiscipline, "%Piping%"))
                    .Where(r => r.ACTIVITY == 3.11 || r.ACTIVITY == 3.110 || r.ACTIVITY == 3.1100);
            }
            else
            {
                baseQuery = baseQuery.Where(r => r.SubDiscipline != null && EF.Functions.Like(r.SubDiscipline, "%Welding%"))
                    .Where(r => r.ACTIVITY == 3.6 || r.ACTIVITY == 3.60 || r.ACTIVITY == 3.600);
            }

            if (loc == "WS")
            {
                baseQuery = baseQuery.Where(r => r.RFI_LOCATION != null && (
                    EF.Functions.Like(r.RFI_LOCATION!.ToUpper(), "WS%") ||
                    EF.Functions.Like(r.RFI_LOCATION!.ToUpper(), "%SHOP%") ||
                    EF.Functions.Like(r.RFI_LOCATION!.ToUpper(), "%WORK%")));
            }
            else if (loc == "FW")
            {
                baseQuery = baseQuery.Where(r =>
                    r.RFI_LOCATION == null ||
                    (
                        !EF.Functions.Like(r.RFI_LOCATION!.ToUpper(), "WS%") &&
                        !EF.Functions.Like(r.RFI_LOCATION!.ToUpper(), "%SHOP%") &&
                        !EF.Functions.Like(r.RFI_LOCATION!.ToUpper(), "%WORK%")
                    ));
            }

            IQueryable<Rfi> working = baseQuery;
            // Note: previous behaviour narrowed results to the selected fitup date when any matches existed.
            // That produced small per-row lists. For fuller row dropdowns (matching header behaviour) do not
            // restrict to the fitup date here – keep the full recent list filtered only by project/location/activity.
            // (fitDay is still accepted but not used to narrow results)

            working = working.OrderByDescending(r => r.Date)
                .ThenByDescending(r => r.Time)
                .ThenByDescending(r => r.RFI_ID);

            var listRaw = await working.Take(2000)
                .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.Date, r.Time, r.RFI_LOCATION, r.RFI_DESCRIPTION })
                .ToListAsync();

            string PrefixFor(string? rfiLoc)
            {
                if (loc == "TH") return "TH | ";
                if (loc == "WS") return "WS | ";
                if (loc == "FW") return "FW | ";
                var raw = (rfiLoc ?? string.Empty).ToUpperInvariant();
                if (raw.Contains("SHOP") || raw.Contains("WORK") || raw.StartsWith("WS")) return "WS | ";
                if (raw.Contains("FIELD") || raw.StartsWith("FW")) return "FW | ";
                return "WS | ";
            }

            var list = listRaw.Select(r => new RfiOption
            {
                Id = r.RFI_ID,
                Value = r.SubCon_RFI_No ?? string.Empty,
                Display = PrefixFor(r.RFI_LOCATION) +
                          (r.SubCon_RFI_No ?? r.RFI_ID.ToString()) +
                          (r.Date.HasValue ? (" | " + r.Date.Value.ToString("dd-MMM-yyyy")) : string.Empty) +
                          (r.Time.HasValue ? (" " + r.Time.Value.ToString("HH:mm")) : string.Empty) +
                          (string.IsNullOrWhiteSpace(r.RFI_DESCRIPTION) ? string.Empty : (" | " + r.RFI_DESCRIPTION)),
            }).ToList();

            return Ok(list);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in GetRfiOptions loc={Loc} fitup={FitupDateIso}", location, fitupDateIso);
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetSchedulesForDiameter(double dia)
    {
        try
        {
            if (dia <= 0) return Json(Array.Empty<string>());
            var list = await _context.Schedule_tbl.AsNoTracking()
                .Where(s => s.NPS == dia && s.SCH != null && s.SCH != "")
                .Select(s => s.SCH!)
                .Distinct()
                .OrderBy(x => x)
                .ToListAsync();
            return Json(list);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetSchedulesForDiameter failed");
            return Json(Array.Empty<string>());
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetScheduleThickness(string? sch, double? dia, string? material)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sch) || !dia.HasValue || dia.Value <= 0)
            {
                return Json(new { thickness = (double?)null });
            }

            var (thickness, _) = await _wpsSelectorService.ResolveOletThicknessAsync(material, sch, dia.Value);
            return Json(new { thickness });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetScheduleThickness failed for Sch={Sch} Dia={Dia}", sch, dia);
            return Json(new { thickness = (double?)null });
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkUpdateDfrRows([FromBody] List<DfrRowUpdateDto> items)
    {
        if (items == null || items.Count == 0) return BadRequest("No rows");

        List<object> errors = [];
        List<int> updatedIds = [];
        List<int> pendingSupersedeConfirm = [];
        var pendingDetails = new List<object>();
        var touchedScopes = new HashSet<(int ProjectId, string Layout, string? Sheet)>();
        int processed = 0;
        foreach (var dto in items)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(dto.FitupReport) && !dto.FitupDate.HasValue)
                {
                    errors.Add(new { id = dto.JointId, message = "Skipped: Fit-up Date is required when Fit-up Report is provided." });
                    continue;
                }
                if (dto.FitupDate.HasValue && string.IsNullOrWhiteSpace(dto.FitupReport))
                {
                    errors.Add(new { id = dto.JointId, message = "Skipped: Fit-up Report is required when Fit-up Date is provided." });
                    continue;
                }
                var entity = await _context.DFR_tbl.FirstOrDefaultAsync(d => d.Joint_ID == dto.JointId);
                if (entity == null)
                {
                    errors.Add(new { id = dto.JointId, message = $"Joint {dto.JointId} not found" });
                    continue;
                }
                if (entity.Deleted || entity.Cancelled)
                {
                    bool isAttemptingToUncheck =
                        (entity.Deleted && dto.Deleted.HasValue && dto.Deleted.Value == false) ||
                        (entity.Cancelled && dto.Cancelled.HasValue && dto.Cancelled.Value == false);
                    if (!isAttemptingToUncheck)
                    {
                        errors.Add(new { id = dto.JointId, message = $"Joint {await GetJointDisplayAsync(dto.JointId)}: Row locked" });
                        continue;
                    }
                }
                if (entity.Fitup_Confirmed && !(dto.FitupConfirmed.HasValue && dto.FitupConfirmed.Value == false))
                {
                    var lockedJoint = await GetJointDisplayAsync(dto.JointId);
                    errors.Add(new { id = dto.JointId, message = $"Joint {lockedJoint}: Row locked" });
                    continue;
                }

                await using var tx = await _context.Database.BeginTransactionAsync();

                if (entity.Fitup_Confirmed && dto.FitupConfirmed == false)
                {
                    entity.Fitup_Confirmed = false;
                }
                else
                {
                    entity.WELD_TYPE = NormalizeWeldType(dto.WeldType ?? entity.WELD_TYPE);
                    entity.DFR_REV = dto.Rev;
                    entity.SPOOL_NUMBER = dto.SpoolNo;
                    entity.DIAMETER = dto.DiaIn;
                    entity.SCHEDULE = dto.Sch;
                    entity.MATERIAL_A = dto.MaterialA;
                    entity.MATERIAL_B = dto.MaterialB;
                    entity.GRADE_A = dto.GradeA;
                    entity.GRADE_B = dto.GradeB;
                    entity.HEAT_NUMBER_A = dto.HeatNumberA;
                    entity.HEAT_NUMBER_B = dto.HeatNumberB;
                    if (dto.WpsId.HasValue && dto.WpsId.Value > 0)
                    {
                        entity.WPS_ID_DFR = dto.WpsId.Value;
                    }
                    else if (!string.IsNullOrWhiteSpace(dto.Wps))
                    {
                        var wpsId = await _context.WPS_tbl.AsNoTracking()
                            .Where(w => w.WPS == dto.Wps)
                            .Select(w => (int?)w.WPS_ID)
                            .FirstOrDefaultAsync();
                        entity.WPS_ID_DFR = wpsId;
                    }
                    else
                    {
                        entity.WPS_ID_DFR = null;
                    }
                    if (dto.FitupDate.HasValue)
                    {
                        var dtLocal = AppClock.ToProjectLocal(dto.FitupDate.Value);
                        var (ok, message) = await ValidateFitupDateAgainstConstraintsAsync(dto.JointId, dtLocal);
                        if (!ok)
                        {
                            var jointDisplay = await GetJointDisplayAsync(dto.JointId);
                            return BadRequest($"Joint {jointDisplay}: {message}");
                        }
                        entity.FITUP_DATE = dtLocal;
                    }
                    else
                    {
                        // When certain NDE/RT dates exist on related tables, FitupDate must be present per business rule.
                        var weldDates = await _context.DWR_tbl.AsNoTracking()
                            .Where(w => w.Joint_ID_DWR == dto.JointId)
                            .Select(w => new { w.DATE_WELDED, w.ACTUAL_DATE_WELDED })
                            .FirstOrDefaultAsync();
                        var hasDwrWelded = weldDates?.DATE_WELDED != null;
                        var hasDwrActual = weldDates?.ACTUAL_DATE_WELDED != null;
                        var otherNde = await _context.Other_NDE_tbl.AsNoTracking()
                            .Where(n => n.Joint_ID_NDE == dto.JointId)
                            .Select(n => new { n.Root_PT_DATE, n.OTHER_NDE_DATE, n.PMI_DATE, n.Bevel_PT_DATE })
                            .FirstOrDefaultAsync();
                        var rtEntry = await _context.RT_tbl.AsNoTracking()
                            .Where(r => r.Joint_ID_RT == dto.JointId)
                            .Select(r => new { r.NDE_REQUEST, r.BSR_NDE_REQUEST })
                            .FirstOrDefaultAsync();

                        bool anyRelatedDatesPresent = hasDwrWelded
                            || hasDwrActual
                            || (otherNde != null && (otherNde.Root_PT_DATE != null || otherNde.OTHER_NDE_DATE != null || otherNde.PMI_DATE != null || otherNde.Bevel_PT_DATE != null))
                            || (rtEntry != null && (rtEntry.NDE_REQUEST != null || rtEntry.BSR_NDE_REQUEST != null));

                        if (anyRelatedDatesPresent)
                        {
                            errors.Add(new { id = dto.JointId, message = "Skipped: Fit-up Date is required when weld/actual or related NDE/RT dates exist." });
                            await tx.RollbackAsync();
                            continue;
                        }
                        entity.FITUP_DATE = null;
                    }
                    entity.FITUP_INSPECTION_QR_NUMBER = dto.FitupReport;
                    entity.TACK_WELDER = dto.TackWelder;
                    entity.OL_DIAMETER = dto.OlDia;
                    entity.OL_SCHEDULE = dto.OlSch;
                    entity.OL_Thick = dto.OlThick;
                    entity.DFR_REMARKS = dto.Remarks;
                    entity.LOCATION = Clean(dto.Location, 4) ?? entity.LOCATION;
                    entity.LAYOUT_NUMBER = Clean(dto.LayoutNumber, 10) ?? entity.LAYOUT_NUMBER;
                    if (!string.IsNullOrWhiteSpace(dto.JAdd)) entity.J_Add = CleanSelect(dto.JAdd, 8)?.ToUpperInvariant();
                    if (!string.IsNullOrWhiteSpace(dto.Sheet)) entity.SHEET = Clean(dto.Sheet, 5) ?? entity.SHEET;
                    if (!string.IsNullOrWhiteSpace(dto.WeldNumber)) entity.WELD_NUMBER = Clean(dto.WeldNumber, 6);
                    if (dto.Deleted.HasValue)
                    {
                        if (dto.Deleted.Value && entity.FITUP_DATE == null)
                        {
                            var jd = await GetJointDisplayAsync(dto.JointId);
                            errors.Add(new { id = dto.JointId, message = $"Joint {jd}: Cannot mark as Deleted — Fit-up Date is required (joint must be fitted-up first)." });
                            await tx.RollbackAsync();
                            continue;
                        }
                        entity.Deleted = dto.Deleted.Value;
                    }
                    if (dto.Cancelled.HasValue)
                    {
                        if (dto.Cancelled.Value && entity.FITUP_DATE != null)
                        {
                            var jd = await GetJointDisplayAsync(dto.JointId);
                            errors.Add(new { id = dto.JointId, message = $"Joint {jd}: Cannot mark as Cancelled — Fit-up Date must be empty (joint already fitted-up, use Deleted instead)." });
                            await tx.RollbackAsync();
                            continue;
                        }
                        entity.Cancelled = dto.Cancelled.Value;
                    }
                    if (dto.FitupConfirmed.HasValue) entity.Fitup_Confirmed = dto.FitupConfirmed.Value;
                    entity.RFI_ID_DFR = (dto.RfiId.HasValue && dto.RfiId.Value > 0) ? dto.RfiId.Value : (int?)null;
                }
                entity.DFR_Updated_By = HttpContext.Session.GetInt32("UserID");
                entity.DFR_Updated_Date = AppClock.Now;

                await _context.SaveChangesAsync();

                await UpdateLineSheetAndSpoolRefsAsync(entity);

                if (!string.IsNullOrWhiteSpace(entity.LAYOUT_NUMBER) && !string.IsNullOrWhiteSpace(entity.SHEET))
                {
                    _ = await PruneUnusedSpReleaseRowsForScopeAsync(entity.Project_No, entity.LAYOUT_NUMBER.Trim(), entity.SHEET.Trim());
                }

                bool need = await ShouldSupersedeAsync(entity);
                if (need && dto.ConfirmSupersede != true)
                {
                    await tx.RollbackAsync();
                    pendingSupersedeConfirm.Add(entity.Joint_ID);
                    var j = await GetJointDisplayAsync(entity.Joint_ID);
                    pendingDetails.Add(new { id = entity.Joint_ID, joint = j });
                    continue;
                }

                if (need)
                {
                    await SupersedeSpoolRelatedDataIfNeededAsync(entity);
                }

                await tx.CommitAsync();
                processed++;
                updatedIds.Add(entity.Joint_ID);

                if (!string.IsNullOrWhiteSpace(entity.LAYOUT_NUMBER))
                {
                    var layoutKey = entity.LAYOUT_NUMBER.Trim();
                    var sheetKey = string.IsNullOrWhiteSpace(entity.SHEET) ? null : entity.SHEET.Trim();
                    touchedScopes.Add((entity.Project_No, layoutKey, sheetKey));
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "BulkUpdateDfrRows row failed for {JointId}", dto.JointId);
                errors.Add(new { id = dto.JointId, message = $"Joint {await GetJointDisplayAsync(dto.JointId)}: Error" });
            }
        }

        if (pendingSupersedeConfirm.Count > 0)
        {
            // Return 200 with a flag instead of HTTP 409 to avoid client fetch errors
            return Json(new
            {
                code = "requireSupersedeConfirm",
                ids = pendingSupersedeConfirm,
                details = pendingDetails,
                message = "You are making modification on released spool(s), are you sure you want to supersede the previous reports for the selected joint(s)?",
                updatedIds,
                success = false
            });
        }

        foreach (var scope in touchedScopes)
        {
            try { await PruneUnusedSpReleaseRowsForScopeAsync(scope.ProjectId, scope.Layout, scope.Sheet); }
            catch (Exception exPrune) { _logger.LogWarning(exPrune, "Scoped prune failed for Project={ProjectId} Layout={Layout} Sheet={Sheet}", scope.ProjectId, scope.Layout, scope.Sheet); }
        }

        var skipped = errors.Count;
        var updated = processed;
        return Json(new { success = skipped == 0, updated, skipped, errors, updatedIds });
    }

    private static int ExtractNumericValue(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return 0;
        var match = NumberInStringRegex().Match(input);
        if (match.Success && int.TryParse(match.Value, out var numeric))
        {
            return numeric;
        }
        return int.TryParse(input, out var fallback) ? fallback : 0;
    }

    private static string? MapHeaderLocation(string? headerLoc)
    {
        if (string.IsNullOrWhiteSpace(headerLoc)) return null;
        var upper = headerLoc.Trim().ToUpperInvariant();
        if (upper == "ALL") return null;
        if (upper.StartsWith("WS") || upper.StartsWith("SHOP") || upper.StartsWith("WORK")) return "WS";
        if (upper.StartsWith("FW") || upper.StartsWith("FIELD")) return "FW";
        if (upper.StartsWith("TH") || upper.Contains("THREAD")) return "TH";
        return upper;
    }

    private async Task<int> ResolveWeldersProjectIdAsync(int projectId)
    {
        if (projectId <= 0) return 0;

        var weldersProjectId = await _context.Projects_tbl.AsNoTracking()
            .Where(p => p.Project_ID == projectId)
            .Select(p => p.Welders_Project_ID)
            .FirstOrDefaultAsync();

        if (weldersProjectId.HasValue && weldersProjectId.Value > 0)
        {
            return weldersProjectId.Value;
        }

        return projectId;
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetWpsCandidates(
        int projectId,
        string? lineClass,
        bool? _pwht,
        double? thickness,
        string? sch,
        double? dia,
        double? olThick)
    {
        try
        {
            var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);

            double? effectiveThickness = thickness;
            if (!effectiveThickness.HasValue)
            {
                var resolved = await _wpsSelectorService.ResolveThicknessAsync(
                    projectId: weldersProjectId,
                    explicitLineClass: lineClass,
                    explicitThickness: thickness,
                    sch: sch,
                    dia: dia,
                    olSch: null,
                    olDia: null,
                    olThick: olThick);
                effectiveThickness = resolved.thickness;
            }

            var candidates = await _wpsSelectorService.GetCandidatesAsync(weldersProjectId, lineClass, effectiveThickness);

            if (_pwht.HasValue)
            {
                candidates = candidates.Where(c => c.Pwht == _pwht.Value).ToList();
            }

            var items = candidates.Select(c => new
            {
                id = c.Id,
                wps = c.Wps,
                thicknessRange = c.ThicknessRange,
                pwht = c.Pwht
            }).ToList();

            return Json(items);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetWpsCandidates failed");
            return Json(Array.Empty<object>());
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> PruneSpoolRefs()
    {
        try
        {
            int projectId = 0;
            if (int.TryParse(HttpContext.Request.Form["projectId"], out var p) && p > 0) projectId = p;
            if (projectId <= 0)
            {
                projectId = await _context.Projects_tbl.AsNoTracking().MaxAsync(x => x.Project_ID);
            }
            var removed = await PruneUnusedSpReleaseRowsAsync(projectId);
            return Ok(new { removed });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PruneSpoolRefs failed");
            return StatusCode(500, "Error");
        }
    }

    private static async Task<string> GetNextReportNumberAsync(IQueryable<Dfr> query)
    {
        var maxExisting = await query
            .Where(d => d.FITUP_INSPECTION_QR_NUMBER != null && d.FITUP_INSPECTION_QR_NUMBER != "")
            .Select(d => d.FITUP_INSPECTION_QR_NUMBER!)
            .OrderByDescending(s => s.Length)
            .ThenByDescending(s => s)
            .FirstOrDefaultAsync();

        var next = ExtractNumericValue(maxExisting) + 1;
        return next.ToString("D5");
    }
    private async Task<int> PruneUnusedSpReleaseRowsAsync(int projectId)
    {
        int removed = 0;
        var candidates = await _context.SP_Release_tbl
            .Where(s => s.SP_Project_No == projectId
                     && s.SP_Date == null
                     && s.SP_Date_SUPERSEDED == null
                     && s.SP_TYPE == "W")
            .Select(s => new
            {
                s.Spool_ID,
                Layout = (s.SP_LAYOUT_NUMBER ?? "").Trim(),
                Sheet = (s.SP_SHEET ?? "").Trim(),
                SpoolNo = (s.SP_SPOOL_NUMBER ?? "").Trim().ToUpper(),
            })
            .ToListAsync();
        if (candidates.Count == 0) return 0;
        var dfrKeys = await _context.DFR_tbl.AsNoTracking()
            .Where(d => d.Project_No == projectId)
            .Select(d => new
            {
                Layout = (d.LAYOUT_NUMBER ?? "").Trim(),
                Sheet = (d.SHEET ?? "").Trim(),
                SpoolNo = (d.SPOOL_NUMBER ?? "").Trim().ToUpper()
            })
            .Distinct()
            .ToListAsync();
        var dfrSet = new HashSet<string>(dfrKeys.Select(k => $"{k.Layout}\u001F{k.Sheet}\u001F{k.SpoolNo}"));
        var referencedIds = new HashSet<int>(await _context.DFR_tbl.AsNoTracking()
            .Where(d => d.Project_No == projectId && d.Spool_ID_DFR.HasValue)
            .Select(d => d.Spool_ID_DFR!.Value)
            .Distinct()
            .ToListAsync());
        var toRemove = candidates
            .Where(c => !referencedIds.Contains(c.Spool_ID) && !dfrSet.Contains($"{c.Layout}\u001F{c.Sheet}\u001F{c.SpoolNo}"))
            .Select(c => c.Spool_ID)
            .ToList();
        if (toRemove.Count == 0) return 0;
        foreach (var id in toRemove)
        {
            var entity = new SpRelease { Spool_ID = id };
            _context.SP_Release_tbl.Attach(entity);
            _context.SP_Release_tbl.Remove(entity);
        }
        removed = await _context.SaveChangesAsync();
        _logger.LogInformation("PruneUnusedSpRelease: removed {Count} stale rows for Project {projectId}", removed, projectId);
        return removed;
    }

    private async Task<int> PruneUnusedSpReleaseRowsForScopeAsync(int projectId, string layout, string? sheet)
    {
        if (projectId <= 0 || string.IsNullOrWhiteSpace(layout)) return 0;
        var layoutKey = layout.Trim();
        var sheetKey = (sheet ?? string.Empty).Trim();
        var candidates = await _context.SP_Release_tbl
            .Where(s => s.SP_Project_No == projectId
                     && s.SP_TYPE == "W"
                     && s.SP_Date == null
                     && s.SP_Date_SUPERSEDED == null
                     && (s.SP_LAYOUT_NUMBER ?? "").Trim() == layoutKey
                     && ((s.SP_SHEET ?? "").Trim() == sheetKey))
            .Select(s => new { s.Spool_ID, SpoolNo = (s.SP_SPOOL_NUMBER ?? "").Trim().ToUpper() })
            .ToListAsync();
        if (candidates.Count == 0) return 0;
        var referencedIds = new HashSet<int>(await _context.DFR_tbl.AsNoTracking()
            .Where(d => d.Project_No == projectId && (d.LAYOUT_NUMBER ?? "").Trim() == layoutKey && (d.SHEET ?? "").Trim() == sheetKey && d.Spool_ID_DFR.HasValue)
            .Select(d => d.Spool_ID_DFR!.Value)
            .Distinct()
            .ToListAsync());
        var dfrSpools = new HashSet<string>(await _context.DFR_tbl.AsNoTracking()
            .Where(d => d.Project_No == projectId && (d.LAYOUT_NUMBER ?? "").Trim() == layoutKey && (d.SHEET ?? "").Trim() == sheetKey)
            .Select(d => (d.SPOOL_NUMBER ?? "").Trim().ToUpper())
            .Distinct()
            .ToListAsync());
        var toRemove = candidates
            .Where(c => !referencedIds.Contains(c.Spool_ID) && !dfrSpools.Contains(c.SpoolNo))
            .Select(c => c.Spool_ID)
            .ToList();
        if (toRemove.Count == 0) return 0;
        foreach (var id in toRemove)
        {
            var stub = new SpRelease { Spool_ID = id };
            _context.SP_Release_tbl.Attach(stub);
            _context.SP_Release_tbl.Remove(stub);
        }
        var removed = await _context.SaveChangesAsync();
        _logger.LogInformation("Pruned {Count} orphan SP_Release rows for Project={ProjectId} Layout={Layout} Sheet={Sheet}", removed, projectId, layoutKey, sheetKey);
        return removed;
    }
}