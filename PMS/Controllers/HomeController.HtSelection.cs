using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> HtSelection(int? projectId = null)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");
        ViewBag.FullName = fullName;

        var projects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_Name)
            .Select(p => new ProjectOption { Id = p.Project_ID, Name = p.Project_Name ?? string.Empty })
            .ToListAsync();

        var selectedProject = await GetDefaultProjectIdAsync(projectId) ?? (projects.Count > 0 ? projects[0].Id : 0);
        var (weldTypes, defaultWeldTypes) = selectedProject > 0 ? await GetJointProgressWeldTypesAsync(selectedProject) : (new List<string>(), new List<string>());
        weldTypes = weldTypes.Where(w => w.Equals("BW", StringComparison.OrdinalIgnoreCase)).ToList();
        if (weldTypes.Count == 0 && defaultWeldTypes.Count > 0)
        {
            weldTypes = defaultWeldTypes.Where(w => w.Equals("BW", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        var defaultWt = weldTypes.Where(w => w.Equals("BW", StringComparison.OrdinalIgnoreCase)).Take(1).ToList();
        if (defaultWt.Count == 0) defaultWt = weldTypes.Count > 0 ? weldTypes.Take(1).ToList() : defaultWeldTypes;

        var lotProjectId = selectedProject > 0
            ? await _context.Projects_tbl.AsNoTracking()
                .Where(p => p.Project_ID == selectedProject)
                .Select(p => p.Welders_Project_ID ?? p.Project_ID)
                .FirstOrDefaultAsync()
            : 0;

        var today = DateTime.Today;
        var lotOptions = lotProjectId > 0
            ? await _context.Lot_No_tbl.AsNoTracking()
                .Where(l => l.Lot_Project_No == lotProjectId && l.From_Date <= today)
                .OrderBy(l => l.From_Date)
                .ThenBy(l => l.Lot_No)
                .Select(l => new LotOption { Id = l.Lot_ID, LotNo = l.Lot_No, From = l.From_Date, To = l.To_Date })
                .ToListAsync()
            : new List<LotOption>();

        var now = AppClock.Now;
        var currentLot = lotOptions.FirstOrDefault(l => IsDateWithinRange(now, l.From, l.To));

        var vm = new HtSelectionViewModel
        {
            Projects = projects,
            SelectedProjectId = selectedProject,
            SelectedProjectIds = selectedProject > 0 ? new List<int> { selectedProject } : new List<int>(),
            SelectedLotId = currentLot?.Id,
            Location = "All",
            LotCategory = "Welding",
            HtMode = "WithoutPWHT",
            WeldTypeOptions = weldTypes,
            SelectedWeldTypes = defaultWt,
            LotOptions = lotOptions
        };

        return View(vm);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetHtSelectionData([FromQuery] HtSelectionQuery query)
    {
        var resolvedIds = (query.ProjectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (resolvedIds.Count == 0 && query.ProjectId > 0) resolvedIds.Add(query.ProjectId);
        query.ProjectIds = resolvedIds;
        if (resolvedIds.Count > 0) query.ProjectId = resolvedIds[0];

        if (query.ProjectId <= 0)
        {
            return Json(new { rows = new List<HtSelectionRowDto>(), stats = new List<HtSelectionStatDto>(), lot = (object?)null });
        }

        var payload = await LoadHtSelectionRowsAsync(query);
        var rows = payload.Rows;

        var statsMap = new Dictionary<(string LineClass, string Root, string FillCap), (int Count, int Selected)>();

        foreach (var row in rows)
        {
            var key = (row.LineClass ?? string.Empty, row.Root ?? string.Empty, row.FillCap ?? string.Empty);
            var selectedIncrement = row.HtSelection ? 1 : 0;
            if (statsMap.TryGetValue(key, out var aggregate))
            {
                statsMap[key] = (aggregate.Count + 1, aggregate.Selected + selectedIncrement);
            }
            else
            {
                statsMap[key] = (1, selectedIncrement);
            }
        }

        var stats = statsMap
            .Select(kvp => new HtSelectionStatDto
            {
                LineClass = kvp.Key.LineClass ?? string.Empty,
                Root = kvp.Key.Root ?? string.Empty,
                FillCap = kvp.Key.FillCap ?? string.Empty,
                GroupCount = kvp.Value.Count,
                SelectedCount = kvp.Value.Selected
            })
            .OrderBy(s => s.LineClass)
            .ThenBy(s => s.Root)
            .ThenBy(s => s.FillCap)
            .ToList();

        var lotInfo = payload.LotFilter != null ? new { payload.LotFilter.Lot_ID, payload.LotFilter.Lot_No, payload.LotFilter.From_Date, payload.LotFilter.To_Date } : null;
        var today = DateTime.Today;
        var lotsPayload = payload.Lots
            .Where(l => l.From_Date <= today)
            .OrderBy(l => l.From_Date)
            .ThenBy(l => l.Lot_No)
            .Select(l => new { l.Lot_ID, l.Lot_No, l.From_Date, l.To_Date })
            .ToList();

        return Json(new { rows, stats, lot = lotInfo, lots = lotsPayload, defaultLotId = payload.DefaultLotId, truncated = payload.Truncated });
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> AutoSelectHtSelection([FromBody] HtSelectionQuery? query)
    {
        if (query == null) return BadRequest("Invalid request");

        var resolvedIds = (query.ProjectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (resolvedIds.Count == 0 && query.ProjectId > 0) resolvedIds.Add(query.ProjectId);
        query.ProjectIds = resolvedIds;
        if (resolvedIds.Count > 0) query.ProjectId = resolvedIds[0];

        if (query.ProjectId <= 0)
        {
            return BadRequest("Invalid request");
        }

        var payload = await LoadHtSelectionRowsAsync(query);
        var rows = payload.Rows;

        if (rows.Count == 0)
        {
            return Json(new { ok = true, selectedIds = new List<int>() });
        }

        var lineTargets = rows
            .GroupBy(r => NormalizeLineClass(r.LineClass))
            .ToDictionary(
                g => g.Key,
                g =>
                {
                    var required = g.Sum(row => Math.Max(0d, row.HtFraction));
                    if (required <= 0d) return 0;
                    return Math.Max(1, (int)Math.Ceiling(required));
                },
                StringComparer.OrdinalIgnoreCase);

        var rowWelderSets = rows
            .Select(r => new
            {
                r.JointId,
                LineClass = NormalizeLineClass(r.LineClass),
                Diameter = r.Diameter ?? 0d,
                Welders = BuildWelderTokenSet(new RtSelectionRowDto { Root = r.Root, FillCap = r.FillCap })
            })
            .ToList();

        var welderCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rowWelderSets)
        {
            foreach (var welder in row.Welders)
            {
                welderCounts[welder] = welderCounts.TryGetValue(welder, out var count) ? count + 1 : 1;
            }
        }

        var welderTargets = welderCounts.ToDictionary(
            kvp => kvp.Key,
            kvp => kvp.Value <= 0 ? 0 : Math.Max(1, (int)Math.Ceiling(kvp.Value * 0.10)),
            StringComparer.OrdinalIgnoreCase);

        var selectedIds = new List<int>();
        var selectedLineCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var selectedWelderCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        bool AllTargetsMet()
        {
            var linesOk = lineTargets.All(kvp => GetSelected(selectedLineCounts, kvp.Key) >= kvp.Value);
            var weldersOk = welderTargets.All(kvp => GetSelected(selectedWelderCounts, kvp.Key) >= kvp.Value);
            return linesOk && weldersOk;
        }

        foreach (var candidate in rowWelderSets.OrderByDescending(r => r.Diameter).ThenBy(r => r.JointId))
        {
            var lineKey = candidate.LineClass;
            var needsLine = lineTargets.TryGetValue(lineKey, out var lineTarget) && GetSelected(selectedLineCounts, lineKey) < lineTarget;
            var needsWelder = candidate.Welders.Any(w => welderTargets.TryGetValue(w, out var welderTarget) && GetSelected(selectedWelderCounts, w) < welderTarget);

            if (!needsLine && !needsWelder)
            {
                continue;
            }

            selectedIds.Add(candidate.JointId);

            if (needsLine)
            {
                selectedLineCounts[lineKey] = GetSelected(selectedLineCounts, lineKey) + 1;
            }

            foreach (var welder in candidate.Welders)
            {
                if (welderTargets.TryGetValue(welder, out var welderTarget) && GetSelected(selectedWelderCounts, welder) < welderTarget)
                {
                    selectedWelderCounts[welder] = GetSelected(selectedWelderCounts, welder) + 1;
                }
            }

            if (AllTargetsMet())
            {
                break;
            }
        }

        return Json(new { ok = true, selectedIds });
    }

    private async Task<HtSelectionDataPayload> LoadHtSelectionRowsAsync(HtSelectionQuery query)
    {
        var empty = new HtSelectionDataPayload();
        var projectIds = (query.ProjectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (projectIds.Count == 0 && query.ProjectId > 0) projectIds.Add(query.ProjectId);
        if (projectIds.Count == 0)
        {
            return empty;
        }

        var lotCategory = (query.LotCategory ?? "Welding").Trim();
        bool useFitupLot = lotCategory.Equals("Fit-up", StringComparison.OrdinalIgnoreCase);
        bool useWeldingLot = lotCategory.Equals("Welding", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(lotCategory);
        if (!useFitupLot && !useWeldingLot)
        {
            useWeldingLot = true;
        }

        var loc = (query.Location ?? "All").Trim();
        var includeShop = loc.Equals("Shop", StringComparison.OrdinalIgnoreCase) || loc.Equals("All", StringComparison.OrdinalIgnoreCase);
        var includeField = loc.Equals("Field", StringComparison.OrdinalIgnoreCase) || loc.Equals("All", StringComparison.OrdinalIgnoreCase);
        if (!includeShop && !includeField)
        {
            includeShop = includeField = true;
        }

        var lotProjectIds = await _context.Projects_tbl.AsNoTracking()
            .Where(p => projectIds.Contains(p.Project_ID))
            .Select(p => p.Welders_Project_ID ?? p.Project_ID)
            .Distinct()
            .ToListAsync();
        var lotProjectId = lotProjectIds.Count > 0 ? lotProjectIds[0] : 0;

        var lotList = await _context.Lot_No_tbl.AsNoTracking()
            .Where(l => l.Lot_Project_No.HasValue && lotProjectIds.Contains(l.Lot_Project_No.Value))
            .OrderBy(l => l.From_Date)
            .ThenBy(l => l.To_Date)
            .ThenBy(l => l.Lot_No)
            .ToListAsync();

        var now = AppClock.Now;
        var defaultLot = lotList.FirstOrDefault(l => IsDateWithinRange(now, l.From_Date, l.To_Date));
        var defaultLotId = defaultLot?.Lot_ID;

        LotNo? lotFilter = null;
        if (query.LotId.HasValue)
        {
            lotFilter = lotList.FirstOrDefault(l => l.Lot_ID == query.LotId.Value);
        }

        var weldTypeFilter = (query.WeldTypes ?? new List<string>())
            .Select(NormalizeWeldTypeToken)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
        bool filterWeldType = weldTypeFilter.Count > 0;

        string pipeToken = (query.PipeClass ?? string.Empty).Trim();
        string welderToken = (query.Welder ?? string.Empty).Trim();
        string layoutToken = (query.Layout ?? string.Empty).Trim();
        string sheetToken = (query.Sheet ?? string.Empty).Trim();
        string jointToken = (query.JointNo ?? string.Empty).Trim();

        const double htFullFraction = 1d;
        const double htTolerance = 0.00001d;
        var htMode = (query.HtMode ?? "WithoutPWHT").Trim();
        var isAfterPwht = htMode.Equals("AfterPWHT", StringComparison.OrdinalIgnoreCase);

        var dataQuery = from p in _context.Projects_tbl.AsNoTracking()
                        join wt in _context.PMS_Weld_Type_tbl.AsNoTracking() on (p.Welders_Project_ID ?? p.Project_ID) equals wt.W_Project_No
                        join d in _context.DFR_tbl.AsNoTracking() on new { Project_ID = p.Project_ID, WELD_TYPE = wt.W_Weld_Type } equals new { Project_ID = d.Project_No, WELD_TYPE = d.WELD_TYPE }
                        join ls in _context.Line_Sheet_tbl.AsNoTracking() on d.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
                        join ll in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
                        join dw in _context.DWR_tbl.AsNoTracking() on d.Joint_ID equals dw.Joint_ID_DWR into dwJoin
                        from dw in dwJoin.DefaultIfEmpty()
                        join wps in _context.WPS_tbl.AsNoTracking() on dw.WPS_ID_DWR equals wps.WPS_ID into wpsJoin
                        from wps in wpsJoin.DefaultIfEmpty()
                        join pw in _context.PWHT_HT_tbl.AsNoTracking() on d.Joint_ID equals pw.Joint_ID_PWHT_HT into pwJoin
                        from pw in pwJoin.DefaultIfEmpty()
                        let rawLocation = d.LOCATION
                        let normalizedLocation = rawLocation == null ? null : rawLocation.Trim().ToUpper()
                        let normalizedWeldType = (wt.W_Weld_Type ?? string.Empty).Trim().ToUpper()
                        let lotDate = useFitupLot ? d.FITUP_DATE : dw.ACTUAL_DATE_WELDED
                        where projectIds.Contains(p.Project_ID)
                              && normalizedWeldType == "BW"
                              && (
                                    (!isAfterPwht
                                        && ll.HT.HasValue
                                        && Math.Abs(ll.HT.GetValueOrDefault()) > htTolerance
                                        && Math.Abs(ll.HT.GetValueOrDefault() - htFullFraction) > htTolerance)
                                    ||
                                    (isAfterPwht
                                        && ll.HT_After_PWHT.HasValue
                                        && Math.Abs(ll.HT_After_PWHT.GetValueOrDefault()) > htTolerance
                                        && Math.Abs(ll.HT_After_PWHT.GetValueOrDefault() - htFullFraction) > htTolerance
                                        && pw != null
                                        && pw.PWHT_DATE != null
                                        && (
                                            ll.PWHT_Y_N == true
                                            || (ll.PWHT_20mm == true && d.OL_Thick.GetValueOrDefault() > 20)
                                        )
                                    )
                                 )
                        select new { d, ll, dw, wps, lotDate, pw };

        if (useFitupLot)
        {
            dataQuery = dataQuery.Where(x => x.d.FITUP_DATE != null);
        }
        else if (useWeldingLot)
        {
            dataQuery = dataQuery.Where(x => x.dw != null && x.dw.ACTUAL_DATE_WELDED != null);
        }

        if (filterWeldType)
        {
            dataQuery = dataQuery.Where(x => weldTypeFilter.Contains((x.d.WELD_TYPE ?? string.Empty).Trim().ToUpper()));
        }

        if (!includeShop || !includeField)
        {
            dataQuery = dataQuery.Where(x =>
                (includeShop && ShopLocationTokens.Contains(((x.d.LOCATION ?? string.Empty).Trim().ToUpper()))) ||
                (includeField && !ShopLocationTokens.Contains(((x.d.LOCATION ?? string.Empty).Trim().ToUpper()))));
        }

        if (!string.IsNullOrWhiteSpace(pipeToken))
        {
            var pattern = BuildContainsLikePattern(pipeToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.ll.Line_Class ?? string.Empty, pattern));
        }

        if (!string.IsNullOrWhiteSpace(layoutToken))
        {
            var pattern = BuildContainsLikePattern(layoutToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.d.LAYOUT_NUMBER ?? string.Empty, pattern));
        }

        if (!string.IsNullOrWhiteSpace(sheetToken))
        {
            var pattern = BuildContainsLikePattern(sheetToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.d.SHEET ?? string.Empty, pattern));
        }

        if (!string.IsNullOrWhiteSpace(welderToken))
        {
            var pattern = BuildContainsLikePattern(welderToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(
                ((x.dw == null ? string.Empty : x.dw.ROOT_A ?? string.Empty) + "," +
                 (x.dw == null ? string.Empty : x.dw.ROOT_B ?? string.Empty) + "," +
                 (x.dw == null ? string.Empty : x.dw.FILL_A ?? string.Empty) + "," +
                 (x.dw == null ? string.Empty : x.dw.FILL_B ?? string.Empty) + "," +
                 (x.dw == null ? string.Empty : x.dw.CAP_A ?? string.Empty) + "," +
                 (x.dw == null ? string.Empty : x.dw.CAP_B ?? string.Empty)),
                pattern));
        }

        if (!string.IsNullOrWhiteSpace(jointToken))
        {
            var pattern = BuildContainsLikePattern(jointToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(
                (x.d.LOCATION ?? string.Empty) + "-" + (x.d.WELD_NUMBER ?? string.Empty) + (x.d.J_Add ?? string.Empty),
                pattern));
        }

        if (lotFilter != null)
        {
            var lotFrom = lotFilter.From_Date;
            var lotTo = lotFilter.To_Date;

            if (useFitupLot)
            {
                if (lotFrom.HasValue)
                {
                    var from = lotFrom.Value;
                    dataQuery = dataQuery.Where(x => x.d.FITUP_DATE >= from);
                }
                if (lotTo.HasValue)
                {
                    var to = lotTo.Value;
                    dataQuery = dataQuery.Where(x => x.d.FITUP_DATE <= to);
                }
            }
            else
            {
                if (lotFrom.HasValue)
                {
                    var from = lotFrom.Value;
                    dataQuery = dataQuery.Where(x => x.dw != null && x.dw.ACTUAL_DATE_WELDED >= from);
                }
                if (lotTo.HasValue)
                {
                    var to = lotTo.Value;
                    dataQuery = dataQuery.Where(x => x.dw != null && x.dw.ACTUAL_DATE_WELDED <= to);
                }
            }
        }

        bool isUploadMode = query.IsUpload;
        var uploadLayoutTokens = (query.UploadedLayouts ?? new List<string>())
            .Select(l => (l ?? string.Empty).Trim().ToUpper())
            .Where(l => l.Length > 0)
            .Distinct()
            .ToList();

        if (isUploadMode && uploadLayoutTokens.Count > 0)
        {
            dataQuery = dataQuery.Where(x =>
                uploadLayoutTokens.Contains((x.d.LAYOUT_NUMBER ?? string.Empty).Trim().ToUpper()));
        }

        int rowLimit = isUploadMode ? 50000 : 4000;

        var rawRowsLimited = await dataQuery
            .Select(x => new
            {
                x.d.Joint_ID,
                x.d.LAYOUT_NUMBER,
                x.d.SHEET,
                x.d.SPOOL_NUMBER,
                x.d.LOCATION,
                x.d.WELD_NUMBER,
                x.d.J_Add,
                x.d.DIAMETER,
                x.d.OL_Thick,
                x.d.WELD_TYPE,
                x.d.FITUP_DATE,
                LineClass = x.ll.Line_Class,
                Ht = x.ll.HT,
                HtAfterPwht = x.ll.HT_After_PWHT,
                PwhtYn = x.ll.PWHT_Y_N,
                Pwht20mm = x.ll.PWHT_20mm,
                RootA = x.dw == null ? null : x.dw.ROOT_A,
                RootB = x.dw == null ? null : x.dw.ROOT_B,
                FillA = x.dw == null ? null : x.dw.FILL_A,
                FillB = x.dw == null ? null : x.dw.FILL_B,
                CapA = x.dw == null ? null : x.dw.CAP_A,
                CapB = x.dw == null ? null : x.dw.CAP_B,
                HtSelection = x.dw != null && x.dw.HT_SELECTION,
                HtSelectionDate = x.dw == null ? null : x.dw.HT_SELECTION_DATE,
                ActualDate = x.dw == null ? null : x.dw.ACTUAL_DATE_WELDED,
                PwhtDate = x.pw == null ? null : x.pw.PWHT_DATE,
                WpsProcess = x.wps == null ? null : x.wps.Weld_Process
            })
            .Take(rowLimit)
            .ToListAsync();

        var rawRows = rawRowsLimited
            .Select(r => new HtSelectionRawRow
            {
                JointId = r.Joint_ID,
                LayoutNumber = r.LAYOUT_NUMBER,
                Sheet = r.SHEET,
                SpoolNumber = r.SPOOL_NUMBER,
                Location = r.LOCATION,
                WeldNumber = r.WELD_NUMBER,
                JAdd = r.J_Add,
                EffectiveJAdd = r.J_Add,
                Diameter = r.DIAMETER,
                ActualThick = r.OL_Thick,
                WeldType = r.WELD_TYPE,
                FitupDate = r.FITUP_DATE,
                LineClass = r.LineClass,
                Ht = r.Ht,
                HtAfterPwht = r.HtAfterPwht,
                PwhtYn = r.PwhtYn,
                Pwht20mm = r.Pwht20mm,
                RootA = r.RootA,
                RootB = r.RootB,
                FillA = r.FillA,
                FillB = r.FillB,
                CapA = r.CapA,
                CapB = r.CapB,
                HtSelection = r.HtSelection,
                ActualDate = r.ActualDate,
                HtSelectionDate = r.HtSelectionDate,
                PwhtDate = r.PwhtDate,
                WpsProcess = r.WpsProcess
            })
            .ToList();

        var latestRawRows = rawRows
            .GroupBy(r => new
            {
                Layout = NormalizeGroupKey(r.LayoutNumber),
                Sheet = NormalizeGroupKey(r.Sheet),
                Weld = NormalizeGroupKey(r.WeldNumber)
            })
            .Select(g =>
            {
                var rRows = g.Where(row => IsRejectRevision(row.JAdd)).ToList();
                if (rRows.Count > 0)
                {
                    return rRows.OrderByDescending(row => ExtractRevisionOrder(row.JAdd)).ThenByDescending(row => row.JointId).First();
                }

                var cRows = g.Where(row => IsCRevision(row.JAdd)).ToList();
                if (cRows.Count > 0)
                {
                    return cRows.OrderByDescending(row => ExtractRevisionOrder(row.JAdd)).ThenByDescending(row => row.JointId).First();
                }

                var cqRows = g.Where(row => IsCqRevision(row.JAdd)).ToList();
                if (cqRows.Count > 0)
                {
                    return cqRows.OrderByDescending(row => ExtractRevisionOrder(row.JAdd)).ThenByDescending(row => row.JointId).First();
                }

                var fallback = g
                    .OrderByDescending(row => ExtractRevisionOrder(row.EffectiveJAdd ?? row.JAdd))
                    .ThenByDescending(row => row.JointId)
                    .First();
                fallback.EffectiveJAdd ??= fallback.JAdd;
                return fallback;
            })
            .Where(row => row != null)
            .Select(row => row!)
            .ToList();

        var rows = new List<HtSelectionRowDto>();
        var sortedLots = lotList;

        foreach (var r in latestRawRows)
        {
            var lotDate = useFitupLot ? r.FitupDate : r.ActualDate;
            var lot = FindLotMatchBinary(sortedLots, lotDate);
            if (lotFilter != null)
            {
                if (lot == null || lot.Lot_ID != lotFilter.Lot_ID) continue;
            }

            var htFraction = ComputeHtFraction(r, isAfterPwht);

            var row = new HtSelectionRowDto
            {
                JointId = r.JointId,
                LineClass = r.LineClass ?? string.Empty,
                LayoutNumber = r.LayoutNumber ?? string.Empty,
                Sheet = r.Sheet ?? string.Empty,
                SpoolNumber = r.SpoolNumber ?? string.Empty,
                JointNo = BuildJointNo(r.Location, r.WeldNumber, r.EffectiveJAdd ?? r.JAdd),
                Diameter = r.Diameter,
                ActualThick = r.ActualThick,
                Root = BuildRoot(r.RootA, r.RootB),
                FillCap = BuildFillCap(r.FillA, r.FillB, r.CapA, r.CapB, r.WpsProcess),
                HtSelection = r.HtSelection,
                HtSelectionDate = r.HtSelectionDate,
                LotNo = lot?.Lot_No ?? string.Empty,
                LotFrom = lot?.From_Date,
                LotTo = lot?.To_Date,
                Location = r.Location ?? string.Empty,
                WeldType = r.WeldType ?? string.Empty,
                HtFraction = htFraction
            };
            rows.Add(row);
        }

        var orderedRows = rows
            .OrderBy(x => x.LineClass)
            .ThenBy(x => x.LayoutNumber)
            .ThenBy(x => x.SpoolNumber)
            .ThenBy(x => x.JointNo)
            .ToList();

        return new HtSelectionDataPayload
        {
            Rows = orderedRows,
            Lots = lotList,
            LotFilter = lotFilter,
            DefaultLotId = defaultLotId,
            Truncated = rawRowsLimited.Count >= rowLimit
        };
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> UpdateHtSelection([FromBody] HtSelectionUpdateRequest? req)
    {
        if (req == null || req.JointId <= 0) return BadRequest("Invalid request");

        var entity = await _context.DWR_tbl.FirstOrDefaultAsync(d => d.Joint_ID_DWR == req.JointId);
        if (entity == null)
        {
            var exists = await _context.DFR_tbl.AsNoTracking().AnyAsync(d => d.Joint_ID == req.JointId);
            if (!exists) return NotFound("Joint not found");

            entity = new Dwr
            {
                Joint_ID_DWR = req.JointId
            };
            await _context.DWR_tbl.AddAsync(entity);
        }

        entity.HT_SELECTION = req.HtSelection;
        entity.HT_SELECTION_DATE = req.HtSelection ? (req.HtSelectionDate ?? AppClock.Now) : null;
        entity.DWR_Updated_Date = AppClock.Now;
        entity.DWR_Updated_By = HttpContext.Session.GetInt32("UserID");

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateException ex)
        {
            _logger.LogError(ex, "UpdateHtSelection failed for Joint_ID {JointId}", req.JointId);
            return StatusCode(StatusCodes.Status500InternalServerError, new
            {
                ok = false,
                message = "Failed to update HT Selection.",
                detail = GetInnermostMessage(ex)
            });
        }

        return Json(new { ok = true });
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> UpdateHtSelectionBulk([FromBody] HtSelectionBulkUpdateRequest? req)
    {
        if (req == null || req.Items == null || req.Items.Count == 0)
        {
            return BadRequest("Invalid request");
        }

        var items = req.Items
            .Where(i => i != null && i.JointId > 0)
            .ToList();
        if (items.Count == 0) return BadRequest("No valid joints provided");

        var updateMap = new Dictionary<int, HtSelectionUpdateRequest>();
        foreach (var item in items)
        {
            updateMap[item.JointId] = item;
        }
        if (updateMap.Count == 0) return BadRequest("No valid joints provided");

        var jointIds = updateMap.Keys.ToHashSet();

        var existingDwrs = await _context.DWR_tbl
            .Where(d => jointIds.Contains(d.Joint_ID_DWR))
            .ToListAsync();

        var existingDwrIds = existingDwrs.Select(d => d.Joint_ID_DWR).ToHashSet();
        var missingIds = jointIds.Except(existingDwrIds).ToList();

        if (missingIds.Count > 0)
        {
            var validMissing = await _context.DFR_tbl.AsNoTracking()
                .Where(d => missingIds.Contains(d.Joint_ID))
                .Select(d => d.Joint_ID)
                .ToListAsync();

            foreach (var mid in validMissing)
            {
                var newEntity = new Dwr { Joint_ID_DWR = mid };
                await _context.DWR_tbl.AddAsync(newEntity);
                existingDwrs.Add(newEntity);
                existingDwrIds.Add(mid);
            }

            var notFound = missingIds.Except(validMissing).ToList();
            if (notFound.Count > 0)
            {
                return NotFound(new { ok = false, missing = notFound });
            }
        }

        var now = AppClock.Now;
        var userId = HttpContext.Session.GetInt32("UserID");

        foreach (var dwr in existingDwrs)
        {
            if (!updateMap.TryGetValue(dwr.Joint_ID_DWR, out var update)) continue;
            dwr.HT_SELECTION = update.HtSelection;
            dwr.HT_SELECTION_DATE = update.HtSelection ? (update.HtSelectionDate ?? now) : null;
            dwr.DWR_Updated_Date = now;
            dwr.DWR_Updated_By = userId;
        }

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateException ex)
        {
            _logger.LogError(ex, "UpdateHtSelectionBulk failed for {Count} joints", updateMap.Count);
            return StatusCode(StatusCodes.Status500InternalServerError, new
            {
                ok = false,
                message = "Failed to save HT selections.",
                detail = GetInnermostMessage(ex)
            });
        }

        return Json(new { ok = true, updated = updateMap.Count });
    }

    private static bool IsCqRevision(string? jAdd)
    {
        var token = (jAdd ?? string.Empty).Trim();
        if (token.Length == 0) return false;
        return token.StartsWith("CQ", StringComparison.OrdinalIgnoreCase);
    }

    private static int ExtractRevisionOrder(string? jAdd)
    {
        var token = (jAdd ?? string.Empty).Trim();
        if (token.Length == 0 || token.Equals("NEW", StringComparison.OrdinalIgnoreCase)) return 0;

        var idx = token.Length - 1;
        while (idx >= 0 && char.IsDigit(token[idx]))
        {
            idx--;
        }

        var start = idx + 1;
        if (start >= token.Length) return 0;

        var numericPart = token[start..];
        return int.TryParse(numericPart, out var parsed) ? parsed : 0;
    }

    private static double ComputeHtFraction(HtSelectionRawRow row, bool isAfterPwht)
    {
        var value = isAfterPwht ? row.HtAfterPwht : row.Ht;
        if (!value.HasValue || value.Value <= 0d) return 0d;
        return Math.Clamp(value.Value, 0d, 1d);
    }

    private sealed class HtSelectionRawRow
    {
        public int JointId { get; init; }
        public string? LayoutNumber { get; init; }
        public string? Sheet { get; init; }
        public string? SpoolNumber { get; init; }
        public string? Location { get; init; }
        public string? WeldNumber { get; init; }
        public string? JAdd { get; init; }
        public string? EffectiveJAdd { get; set; }
        public double? Diameter { get; init; }
        public double? ActualThick { get; init; }
        public string? WeldType { get; init; }
        public DateTime? FitupDate { get; init; }
        public string? LineClass { get; init; }
        public double? Ht { get; init; }
        public double? HtAfterPwht { get; init; }
        public bool? PwhtYn { get; init; }
        public bool? Pwht20mm { get; init; }
        public string? RootA { get; init; }
        public string? RootB { get; init; }
        public string? FillA { get; init; }
        public string? FillB { get; init; }
        public string? CapA { get; init; }
        public string? CapB { get; init; }
        public bool HtSelection { get; init; }
        public DateTime? ActualDate { get; init; }
        public DateTime? HtSelectionDate { get; init; }
        public DateTime? PwhtDate { get; init; }
        public string? WpsProcess { get; init; }
    }

    public sealed class HtSelectionQuery
    {
        public int ProjectId { get; set; }
        public List<int>? ProjectIds { get; set; }
        public string? Location { get; set; }
        public string? LotCategory { get; set; }
        public List<string>? WeldTypes { get; set; }
        public int? LotId { get; set; }
        public string? PipeClass { get; set; }
        public string? Welder { get; set; }
        public string? Layout { get; set; }
        public string? Sheet { get; set; }
        public string? JointNo { get; set; }
        public string? HtMode { get; set; }
        public bool IsUpload { get; set; }
        public List<string>? UploadedLayouts { get; set; }
    }

    public sealed class HtSelectionUpdateRequest
    {
        public int JointId { get; set; }
        public bool HtSelection { get; set; }
        public DateTime? HtSelectionDate { get; set; }
    }

    public sealed class HtSelectionBulkUpdateRequest
    {
        public List<HtSelectionUpdateRequest> Items { get; set; } = new();
    }

    private sealed class HtSelectionDataPayload
    {
        public List<HtSelectionRowDto> Rows { get; init; } = new();
        public List<LotNo> Lots { get; init; } = new();
        public LotNo? LotFilter { get; init; }
        public int? DefaultLotId { get; init; }
        public bool Truncated { get; init; }
    }
}
