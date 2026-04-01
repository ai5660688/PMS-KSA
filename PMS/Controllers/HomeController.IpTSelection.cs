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
    public async Task<IActionResult> IpTSelection(int? projectId = null)
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
        var defaultWt = weldTypes.Where(w => w.Equals("BW", StringComparison.OrdinalIgnoreCase)).Take(1).ToList();
        if (defaultWt.Count == 0) defaultWt = defaultWeldTypes.Count > 0 ? defaultWeldTypes : weldTypes;

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

        var vm = new IpTSelectionViewModel
        {
            Projects = projects,
            SelectedProjectId = selectedProject,
            SelectedProjectIds = selectedProject > 0 ? new List<int> { selectedProject } : new List<int>(),
            SelectedLotId = currentLot?.Id,
            LotOptions = lotOptions,
            WeldTypeOptions = weldTypes,
            SelectedWeldTypes = defaultWt
        };

        return View(vm);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetIpTSelectionData([FromQuery] IpTSelectionQuery query)
    {
        var projectIds = (query.ProjectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (projectIds.Count == 0 && query.ProjectId > 0) projectIds.Add(query.ProjectId);
        query.ProjectIds = projectIds;
        if (projectIds.Count > 0) query.ProjectId = projectIds[0];

        if (query.ProjectId <= 0)
        {
            return Json(new { rows = new List<IpTSelectionRowDto>(), lot = (object?)null });
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

        var loc = (query.Location ?? "All").Trim();
        var includeShop = loc.Equals("Shop", StringComparison.OrdinalIgnoreCase) || loc.Equals("All", StringComparison.OrdinalIgnoreCase);
        var includeField = loc.Equals("Field", StringComparison.OrdinalIgnoreCase) || loc.Equals("All", StringComparison.OrdinalIgnoreCase);
        if (!includeShop && !includeField)
        {
            includeShop = includeField = true;
        }

        // Main query
        //   J_Add NOT LIKE 'R%'
        //   RT eligibility:
        //     (RT_Shop <> 1 AND LOCATION = 'WS')
        //     OR (RT_Field <> 1 AND LOCATION <> 'WS')
        //     OR (RT_Field_Shop_SW <> 1 AND WELD_TYPE IN ('SW','SOF','TH'))
        //   RT result = 'R'
        var dataQuery = from dwr in _context.DWR_tbl.AsNoTracking()
                        join dfr in _context.DFR_tbl.AsNoTracking() on dwr.Joint_ID_DWR equals dfr.Joint_ID
                        join ls in _context.Line_Sheet_tbl.AsNoTracking() on dfr.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
                        join ll in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
                        join rt in _context.RT_tbl.AsNoTracking() on dfr.Joint_ID equals rt.Joint_ID_RT into rtJoin
                        from rt in rtJoin.DefaultIfEmpty()
                        where projectIds.Contains(dfr.Project_No)
                              && !EF.Functions.Like(dfr.J_Add ?? string.Empty, "R%")
                              && (
                                  (ll.RT_Shop.HasValue && Math.Abs(ll.RT_Shop.Value - 1.0) >= 0.001
                                      && (dfr.LOCATION ?? string.Empty) == "WS")
                                  || (ll.RT_Field.HasValue && Math.Abs(ll.RT_Field.Value - 1.0) >= 0.001
                                      && (dfr.LOCATION ?? string.Empty) != "WS")
                                  || (ll.RT_Field_Shop_SW.HasValue && Math.Abs(ll.RT_Field_Shop_SW.Value - 1.0) >= 0.001
                                      && ((dfr.WELD_TYPE ?? string.Empty) == "SW"
                                          || (dfr.WELD_TYPE ?? string.Empty) == "SOF"
                                          || (dfr.WELD_TYPE ?? string.Empty) == "TH"))
                              )
                              && (rt.Final_RT_RESULT == "R" || EF.Functions.Like(rt.Final_RT_RESULT ?? string.Empty, "R/%"))
                              && rt.Repair_Welder != null
                              && (rt.Repair_Welder == dwr.ROOT_A
                                  || rt.Repair_Welder == dwr.ROOT_B
                                  || rt.Repair_Welder == dwr.FILL_A
                                  || rt.Repair_Welder == dwr.FILL_B
                                  || rt.Repair_Welder == dwr.CAP_A
                                  || rt.Repair_Welder == dwr.CAP_B)
                              && (
                                  dwr.IP_or_T == null
                                  || !EF.Functions.Like(dwr.IP_or_T, "IP%")
                                  || (EF.Functions.Like(dwr.IP_or_T, "IP%")
                                      && dwr.DWR_REMARKS != null
                                      && dwr.DWR_REMARKS != (dwr.ROOT_A ?? string.Empty)
                                      && dwr.DWR_REMARKS != (dwr.ROOT_B ?? string.Empty)
                                      && dwr.DWR_REMARKS != (dwr.FILL_A ?? string.Empty)
                                      && dwr.DWR_REMARKS != (dwr.FILL_B ?? string.Empty)
                                      && dwr.DWR_REMARKS != (dwr.CAP_A ?? string.Empty)
                                      && dwr.DWR_REMARKS != (dwr.CAP_B ?? string.Empty))
                              )
                        select new
                        {
                            dfr.Joint_ID,
                            dfr.J_Add,
                            dfr.DIAMETER,
                            dfr.LOCATION,
                            dfr.WELD_NUMBER,
                            dfr.LAYOUT_NUMBER,
                            dfr.WELD_TYPE,
                            dwr.ACTUAL_DATE_WELDED,
                            dwr.IP_or_T,
                            dwr.DWR_REMARKS,
                            dwr.ROOT_A,
                            dwr.ROOT_B,
                            dwr.FILL_A,
                            dwr.FILL_B,
                            dwr.CAP_A,
                            dwr.CAP_B,
                            FinalRtResult = rt == null ? null : rt.Final_RT_RESULT,
                            RepairWelder = rt == null ? null : rt.Repair_Welder,
                            LineClass = ll.Line_Class,
                            dfr.SHEET
                        };

        // Apply lot date filter
        if (lotFilter != null)
        {
            if (lotFilter.From_Date.HasValue)
            {
                var from = lotFilter.From_Date.Value;
                dataQuery = dataQuery.Where(x => x.ACTUAL_DATE_WELDED >= from);
            }
            if (lotFilter.To_Date.HasValue)
            {
                var to = lotFilter.To_Date.Value;
                dataQuery = dataQuery.Where(x => x.ACTUAL_DATE_WELDED <= to);
            }
        }

        // Apply location filter
        if (!includeShop || !includeField)
        {
            dataQuery = dataQuery.Where(x =>
                (includeShop && ShopLocationTokens.Contains(((x.LOCATION ?? string.Empty).Trim().ToUpper()))) ||
                (includeField && !ShopLocationTokens.Contains(((x.LOCATION ?? string.Empty).Trim().ToUpper()))));
        }

        // Apply weld type filter
        var weldTypeFilter = (query.WeldTypes ?? new List<string>())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s.Trim().ToUpperInvariant())
            .ToList();
        if (weldTypeFilter.Count > 0)
        {
            dataQuery = dataQuery.Where(x => weldTypeFilter.Contains((x.WELD_TYPE ?? string.Empty).Trim().ToUpper()));
        }

        // Build a base-population query from eligible NON-REJECTED joints
        // to compute Joint_Availability, Total_T_Joints, and R_Joint_Count for status.
        // Rejected joints (RT result = 'R' or 'R/%') are excluded because they cannot
        // serve as tracer candidates — matching BuildMatchedRowsPreviewAsync logic.
        // We use LEFT JOIN to RT_tbl so that joints with an accepted re-test remain
        // in the population (only the rejected RT rows are filtered out).
        var peerStatusQuery =
            from dwr in _context.DWR_tbl.AsNoTracking()
            join dfr in _context.DFR_tbl.AsNoTracking() on dwr.Joint_ID_DWR equals dfr.Joint_ID
            join ls  in _context.Line_Sheet_tbl.AsNoTracking() on dfr.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
            join ll  in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
            join rt  in _context.RT_tbl.AsNoTracking() on dfr.Joint_ID equals rt.Joint_ID_RT into rtJoin
            from rt in rtJoin.DefaultIfEmpty()
            where projectIds.Contains(dfr.Project_No)
                  && !EF.Functions.Like(dfr.J_Add ?? string.Empty, "R%")
                  && (
                      (ll.RT_Shop.HasValue && Math.Abs(ll.RT_Shop.Value - 1.0) >= 0.001
                          && (dfr.LOCATION ?? string.Empty) == "WS")
                      || (ll.RT_Field.HasValue && Math.Abs(ll.RT_Field.Value - 1.0) >= 0.001
                          && (dfr.LOCATION ?? string.Empty) != "WS")
                      || (ll.RT_Field_Shop_SW.HasValue && Math.Abs(ll.RT_Field_Shop_SW.Value - 1.0) >= 0.001
                          && ((dfr.WELD_TYPE ?? string.Empty) == "SW"
                              || (dfr.WELD_TYPE ?? string.Empty) == "SOF"
                              || (dfr.WELD_TYPE ?? string.Empty) == "TH"))
                  )
                  && (
                      dwr.IP_or_T == null
                      || !EF.Functions.Like(dwr.IP_or_T, "IP%")
                      || (EF.Functions.Like(dwr.IP_or_T, "IP%")
                          && dwr.DWR_REMARKS != null
                          && dwr.DWR_REMARKS != (dwr.ROOT_A ?? string.Empty)
                          && dwr.DWR_REMARKS != (dwr.ROOT_B ?? string.Empty)
                          && dwr.DWR_REMARKS != (dwr.FILL_A ?? string.Empty)
                          && dwr.DWR_REMARKS != (dwr.FILL_B ?? string.Empty)
                          && dwr.DWR_REMARKS != (dwr.CAP_A ?? string.Empty)
                          && dwr.DWR_REMARKS != (dwr.CAP_B ?? string.Empty))
                  )
                  // Exclude rejected RT rows — only accepted/null rows remain
                  && (rt == null || (rt.Final_RT_RESULT != "R" && !EF.Functions.Like(rt.Final_RT_RESULT ?? string.Empty, "R/%")))
            select new
            {
                JointId = dfr.Joint_ID,
                dfr.DIAMETER,
                dfr.LOCATION,
                dfr.WELD_TYPE,
                dwr.IP_or_T,
                dwr.ACTUAL_DATE_WELDED,
                dwr.ROOT_A,
                dwr.ROOT_B,
                dwr.FILL_A,
                dwr.FILL_B,
                dwr.CAP_A,
                dwr.CAP_B,
                LineClass = ll.Line_Class
            };

        // Apply location filter to peer population (must match dataQuery scope)
        if (!includeShop || !includeField)
        {
            peerStatusQuery = peerStatusQuery.Where(x =>
                (includeShop && ShopLocationTokens.Contains(((x.LOCATION ?? string.Empty).Trim().ToUpper()))) ||
                (includeField && !ShopLocationTokens.Contains(((x.LOCATION ?? string.Empty).Trim().ToUpper()))));
        }

        if (weldTypeFilter.Count > 0)
        {
            peerStatusQuery = peerStatusQuery.Where(x =>
                weldTypeFilter.Contains((x.WELD_TYPE ?? string.Empty).Trim().ToUpper()));
        }
        if (lotFilter != null)
        {
            if (lotFilter.From_Date.HasValue)
            {
                var psFrom = lotFilter.From_Date.Value;
                peerStatusQuery = peerStatusQuery.Where(x => x.ACTUAL_DATE_WELDED >= psFrom);
            }
            if (lotFilter.To_Date.HasValue)
            {
                var psTo = lotFilter.To_Date.Value;
                peerStatusQuery = peerStatusQuery.Where(x => x.ACTUAL_DATE_WELDED <= psTo);
            }
        }

        // Apply search filters
        string lineClassToken = (query.LineClass ?? string.Empty).Trim();
        string layoutToken = (query.Layout ?? string.Empty).Trim();
        string sheetToken = (query.Sheet ?? string.Empty).Trim();
        string jointToken = (query.JointNo ?? string.Empty).Trim();
        string repairWelderToken = (query.RepairWelder ?? string.Empty).Trim();

        if (!string.IsNullOrWhiteSpace(lineClassToken))
        {
            var pattern = BuildContainsLikePattern(lineClassToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.LineClass ?? string.Empty, pattern));
        }

        if (!string.IsNullOrWhiteSpace(layoutToken))
        {
            var pattern = BuildContainsLikePattern(layoutToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.LAYOUT_NUMBER ?? string.Empty, pattern));
        }

        if (!string.IsNullOrWhiteSpace(sheetToken))
        {
            var pattern = BuildContainsLikePattern(sheetToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.SHEET ?? string.Empty, pattern));
        }

        if (!string.IsNullOrWhiteSpace(jointToken))
        {
            var pattern = BuildContainsLikePattern(jointToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(
                (x.J_Add ?? string.Empty) != "New" && (x.J_Add ?? string.Empty) != string.Empty
                    ? (x.LOCATION ?? string.Empty) + "-" + (x.WELD_NUMBER ?? string.Empty) + (x.J_Add ?? string.Empty)
                    : (x.LOCATION ?? string.Empty) + "-" + (x.WELD_NUMBER ?? string.Empty),
                pattern));
        }

        if (!string.IsNullOrWhiteSpace(repairWelderToken))
        {
            var pattern = BuildContainsLikePattern(repairWelderToken);
            dataQuery = dataQuery.Where(x => EF.Functions.Like(x.RepairWelder ?? string.Empty, pattern));
        }

        var rawData = await dataQuery.Take(5000).ToListAsync();
        // No Take() limit — the LEFT JOIN to RT_tbl can multiply rows, and a low
        // limit was silently truncating the peer population, producing wrong statuses.
        // The query is already scoped by project + RT-eligibility + IP/T-eligibility +
        // lot/location/weld-type filters, so the result set is bounded.
        var rawPeerStatusData = await peerStatusQuery.ToListAsync();

        // Build one row per joint, using RT.Repair_Welder
        var expanded = new List<IpTSelectionRowDto>();
        foreach (var r in rawData)
        {
            // Find lot
            var lot = FindLotMatchBinary(lotList, r.ACTUAL_DATE_WELDED);
            if (lotFilter != null && (lot == null || lot.Lot_ID != lotFilter.Lot_ID))
                continue;

            expanded.Add(new IpTSelectionRowDto
            {
                JointId = r.Joint_ID,
                JAdd = r.J_Add ?? string.Empty,
                ActualDateWelded = r.ACTUAL_DATE_WELDED,
                Diameter = r.DIAMETER,
                LineClass = r.LineClass ?? string.Empty,
                WelderSymbol = (r.RepairWelder ?? string.Empty).Trim(),
                FinalRtResult = r.FinalRtResult,
                IpOrT = r.IP_or_T,
                RepairWelder = r.RepairWelder,
                RootA = r.ROOT_A,
                RootB = r.ROOT_B,
                FillA = r.FILL_A,
                FillB = r.FILL_B,
                CapA = r.CAP_A,
                CapB = r.CAP_B,
                LotNo = lot?.Lot_No ?? string.Empty,
                Location = r.LOCATION ?? string.Empty,
                WeldNumber = r.WELD_NUMBER ?? string.Empty,
                LayoutNumber = r.LAYOUT_NUMBER ?? string.Empty,
                Sheet = r.SHEET,
                HasRepairCq = false,
                DwrRemarks = r.DWR_REMARKS,
                TracerStatus = string.Empty
            });
        }

        // Pre-group ALL eligible joints by the enabled match criteria dimensions.
        // Status rules (see ComputeTracerStatus):
        //   Not Available    — Joint_Availability == 0 (repair welder has zero joints in group)
        //   Done             — Req_Joint_Count (R_Joint_Count × 2) ≤ Total_T_Joints
        //   2nd Not Available— Joint_Availability < 2 AND no unassigned joints (IP_or_T IS NULL) exist
        //   Pending          — otherwise
        var peerByKey =
            new Dictionary<(string lot, DateTime? date, double? dia, string lc),
                           List<(int jointId, string? ipOrT, string? ra, string? rb, string? fa, string? fb, string? ca, string? cb)>>();

        // Broad lookup (without date dimension) to catch null-IP_or_T joints that the
        // matched-rows popup can see but the strict (lot, date, dia, lineClass) group
        // might miss due to date mismatch or grouping differences.
        var broadNullIpOrT =
            new Dictionary<(string lot, double? dia, string lc, string welder),
                           HashSet<int>>();

        foreach (var p in rawPeerStatusData)
        {
            var pLot = FindLotMatchBinary(lotList, p.ACTUAL_DATE_WELDED);
            if (lotFilter != null && (pLot == null || pLot.Lot_ID != lotFilter.Lot_ID))
                continue;
            var pDia = p.DIAMETER.HasValue ? Math.Round(p.DIAMETER.Value, 3) : (double?)null;
            var pLc  = (p.LineClass ?? "").Trim().ToUpperInvariant();
            var pLotNo = (pLot?.Lot_No ?? string.Empty).Trim().ToUpperInvariant();

            // Status grouping matches SQL TotalT / WelderTotalJoints CTEs:
            // group by (lot, date, diameter, lineClass) to match the UI match criteria.
            // Collapse dimensions whose match criterion is unchecked so all values
            // map to the same bucket, mirroring BuildMatchedRowsPreviewAsync scope.
            var pKey = (
                query.MatchLot ? pLotNo : "",
                query.MatchDate ? p.ACTUAL_DATE_WELDED?.Date : (DateTime?)null,
                query.MatchDiameter ? pDia : (double?)null,
                query.MatchLineClass ? pLc : ""
            );
            if (!peerByKey.TryGetValue(pKey, out var list))
                peerByKey[pKey] = list = new();
            list.Add((p.JointId, p.IP_or_T, p.ROOT_A, p.ROOT_B, p.FILL_A, p.FILL_B, p.CAP_A, p.CAP_B));

            // Build broad null-IP_or_T lookup (keyed without date, per welder)
            if (string.IsNullOrWhiteSpace(p.IP_or_T))
            {
                var welders = new[] { p.ROOT_A, p.ROOT_B, p.FILL_A, p.FILL_B, p.CAP_A, p.CAP_B };
                foreach (var w in welders)
                {
                    if (string.IsNullOrWhiteSpace(w)) continue;
                    var wKey = (
                        query.MatchLot ? pLotNo : "",
                        query.MatchDiameter ? pDia : (double?)null,
                        query.MatchLineClass ? pLc : "",
                        query.MatchWelder ? w.Trim().ToUpperInvariant() : ""
                    );
                    if (!broadNullIpOrT.TryGetValue(wKey, out var ids))
                        broadNullIpOrT[wKey] = ids = new HashSet<int>();
                    ids.Add(p.JointId);
                }
            }
        }

        static bool IsWelderInAnyPass(string? welder, string? rootA, string? rootB, string? fillA, string? fillB, string? capA, string? capB)
        {
            if (string.IsNullOrWhiteSpace(welder))
                return false;

            var ws = welder.Trim();
            return (!string.IsNullOrWhiteSpace(rootA) && rootA.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(rootB) && rootB.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(fillA) && fillA.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(fillB) && fillB.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(capA) && capA.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(capB) && capB.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase));
        }

        // Pre-compute R_Joint_Count per (match-criteria group, repair welder),
        // considering only clear repair records where RepairWelder is among original welders.
        var rJointCountByKey = expanded
            .Where(e => IsWelderInAnyPass(e.WelderSymbol, e.RootA, e.RootB, e.FillA, e.FillB, e.CapA, e.CapB))
            .GroupBy(e => (
                lot: query.MatchLot ? (e.LotNo ?? "").Trim().ToUpperInvariant() : "",
                date: query.MatchDate ? e.ActualDateWelded?.Date : (DateTime?)null,
                dia: query.MatchDiameter ? (e.Diameter.HasValue ? Math.Round(e.Diameter.Value, 3) : (double?)null) : (double?)null,
                lc: query.MatchLineClass ? (e.LineClass ?? "").Trim().ToUpperInvariant() : "",
                rw: (e.WelderSymbol ?? "").Trim().ToUpperInvariant()
            ))
            .ToDictionary(g => g.Key, g => g.Select(e => e.JointId).Distinct().Count());

        foreach (var row in expanded)
        {
            var rDia = row.Diameter.HasValue ? Math.Round(row.Diameter.Value, 3) : (double?)null;
            var rLc  = (row.LineClass ?? "").Trim().ToUpperInvariant();
            var rKey = (
                query.MatchLot ? (row.LotNo ?? "").Trim().ToUpperInvariant() : "",
                query.MatchDate ? row.ActualDateWelded?.Date : (DateTime?)null,
                query.MatchDiameter ? rDia : (double?)null,
                query.MatchLineClass ? rLc : ""
            );

            // R_Joint_Count for this repair welder in this group
            var rwUpper = (row.WelderSymbol ?? "").Trim().ToUpperInvariant();
            var rJointCount = rJointCountByKey.GetValueOrDefault((rKey.Item1, rKey.Item2, rKey.Item3, rKey.Item4, rwUpper), 0);
            var reqJointCount = rJointCount * 2;

            if (!peerByKey.TryGetValue(rKey, out var group) || group.Count == 0)
            {
                var broadNullNoPeer = 0;
                if (!query.MatchDate)
                {
                    var broadKeyNoPeer = (
                        rKey.Item1,
                        query.MatchDiameter ? rDia : (double?)null,
                        query.MatchLineClass ? rLc : "",
                        query.MatchWelder ? rwUpper : ""
                    );
                    broadNullNoPeer = broadNullIpOrT.TryGetValue(broadKeyNoPeer, out var bIds) ? bIds.Count : 0;
                }
                row.TracerStatus = ComputeTracerStatus(0, reqJointCount, 0, broadNullNoPeer);
                continue;
            }

            // Joint_Availability (mirrors SQL WelderTotalJoints): count distinct joints this repair welder
            // worked on in ANY pass position within the peer group — no process restriction.
            var ws = row.WelderSymbol;

            // Total_T_Joints: distinct joints in group where IP_or_T starts with 'T'.
            // This mirrors the SQL TotalT CTE which counts ALL T-joints in the
            // (lot, date, diameter, lineClass) group — NOT scoped to the repair welder.
            var totalTJoints = group
                .Where(p => !string.IsNullOrWhiteSpace(p.ipOrT) && p.ipOrT.StartsWith("T", StringComparison.OrdinalIgnoreCase))
                .Select(p => p.jointId)
                .Distinct()
                .Count();
            int jointAvailability;
            if (!query.MatchWelder || string.IsNullOrWhiteSpace(ws))
            {
                // MatchWelder OFF or no welder → count all joints in the group
                jointAvailability = group.Select(p => p.jointId).Distinct().Count();
            }
            else
            {
                jointAvailability = group.Where(p =>
                        (!string.IsNullOrWhiteSpace(p.ra) && p.ra.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrWhiteSpace(p.rb) && p.rb.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrWhiteSpace(p.fa) && p.fa.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrWhiteSpace(p.fb) && p.fb.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrWhiteSpace(p.ca) && p.ca.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrWhiteSpace(p.cb) && p.cb.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)))
                    .Select(p => p.jointId)
                    .Distinct()
                    .Count();
            }

            // Check for unassigned (null IP_or_T) joints using two approaches:
            // 1. Narrow: within the exact (lot, date, diameter, lineClass) group, scoped to repair welder
            // 2. Broad:  within the (lot, diameter, lineClass) scope (ignoring date)
            //    — catches matched-rows-popup joints that may differ by date.
            var narrowNullCount = (!query.MatchWelder || string.IsNullOrWhiteSpace(ws))
                ? group
                    .Where(p => string.IsNullOrWhiteSpace(p.ipOrT))
                    .Select(p => p.jointId)
                    .Distinct()
                    .Count()
                : group
                    .Where(p => string.IsNullOrWhiteSpace(p.ipOrT) &&
                        ((!string.IsNullOrWhiteSpace(p.ra) && p.ra.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                         (!string.IsNullOrWhiteSpace(p.rb) && p.rb.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                         (!string.IsNullOrWhiteSpace(p.fa) && p.fa.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                         (!string.IsNullOrWhiteSpace(p.fb) && p.fb.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                         (!string.IsNullOrWhiteSpace(p.ca) && p.ca.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase)) ||
                         (!string.IsNullOrWhiteSpace(p.cb) && p.cb.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase))))
                    .Select(p => p.jointId)
                    .Distinct()
                    .Count();
            // Only use the broad (date-ignoring) lookup when MatchDate is OFF.
            // When MatchDate is ON the matched-rows popup also filters by date, so
            // joints from other dates would never appear as candidates.
            var broadNullCount = 0;
            if (!query.MatchDate)
            {
                var broadKey = (
                    rKey.Item1,
                    query.MatchDiameter ? rDia : (double?)null,
                    query.MatchLineClass ? rLc : "",
                    query.MatchWelder ? rwUpper : ""
                );
                broadNullCount = broadNullIpOrT.TryGetValue(broadKey, out var broadIds) ? broadIds.Count : 0;
            }
            var nullIpOrTAvailability = Math.Max(narrowNullCount, broadNullCount);

            row.TracerStatus = ComputeTracerStatus(jointAvailability, reqJointCount, totalTJoints, nullIpOrTAvailability);
        }

        var orderedRows = expanded
            .OrderBy(x => x.RepairWelder)
            .ThenBy(x => x.LineClass)
            .ThenBy(x => x.LayoutNumber)
            .ThenBy(x => x.WeldNumber)
            .ThenBy(x => x.ActualDateWelded)
            .ToList();

        var today = DateTime.Today;
        var lotsPayload = lotList
            .Where(l => l.From_Date <= today)
            .OrderBy(l => l.From_Date)
            .ThenBy(l => l.Lot_No)
            .Select(l => new { l.Lot_ID, l.Lot_No, l.From_Date, l.To_Date })
            .ToList();

        var lotInfo = lotFilter != null ? new { lotFilter.Lot_ID, lotFilter.Lot_No, lotFilter.From_Date, lotFilter.To_Date } : null;

        return Json(new { rows = orderedRows, lot = lotInfo, lots = lotsPayload, defaultLotId });
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> GetIpTDropdownOptions([FromBody] IpTDropdownRequest? req)
    {
        if (req == null || req.JointId <= 0 || req.ProjectId <= 0)
            return BadRequest("Invalid request");

        var weldersProjectId = await ResolveWeldersProjectIdAsync(req.ProjectId);

        // Load PMS_IP_T_tbl for this project
        var ipTAll = await _context.PMS_IP_T_tbl.AsNoTracking()
            .Where(x => x.IP_T_Project_No == weldersProjectId || x.IP_T_Project_No == 0)
            .ToListAsync();

        // Get the joint's DFR and DWR data
        var dfr = await (from d in _context.DFR_tbl.AsNoTracking()
                         join ls in _context.Line_Sheet_tbl.AsNoTracking() on d.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
                         join ll in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
                         where d.Joint_ID == req.JointId
                         select new
                         {
                             d.LOCATION,
                             d.WELD_NUMBER,
                             d.J_Add,
                             d.Project_No,
                             d.DIAMETER,
                             ll.Line_Class
                         }).FirstOrDefaultAsync();

        if (dfr == null) return NotFound("Joint not found");

        var dwrSource = await _context.DWR_tbl.AsNoTracking()
            .Where(d => d.Joint_ID_DWR == req.JointId)
            .Select(d => new { d.IP_or_T, d.ACTUAL_DATE_WELDED, d.ROOT_A, d.ROOT_B, d.FILL_A, d.FILL_B, d.CAP_A, d.CAP_B })
            .FirstOrDefaultAsync();

        var currentIpOrT = dwrSource?.IP_or_T;

        var sourceRepairWelder = await _context.RT_tbl.AsNoTracking()
            .Where(r => r.Joint_ID_RT == req.JointId)
            .Select(r => r.Repair_Welder)
            .FirstOrDefaultAsync();

        // Check RT result to determine if this is a repair joint
        var rtResult = await _context.RT_tbl.AsNoTracking()
            .Where(r => r.Joint_ID_RT == req.JointId)
            .Select(r => r.Final_RT_RESULT)
            .FirstOrDefaultAsync();

        var trimmedRtResult = (rtResult ?? "").Trim();
        bool isRepair = trimmedRtResult.Equals("R", StringComparison.OrdinalIgnoreCase)
                     || trimmedRtResult.StartsWith("R/", StringComparison.OrdinalIgnoreCase);

        bool flNumbering = false;
        string? flPrefix = null;

        // Precompute tracer value lists from PMS_IP_T_tbl
        var firstTracerValues = ipTAll
            .Where(x => (x.IP_T_Notes ?? "").Equals("First Tracer", StringComparison.OrdinalIgnoreCase))
            .Select(x => (x.P_IP_T_List ?? "").Trim())
            .Where(x => x.Length > 0)
            .ToList();

        var allSecondTracerValues = ipTAll
            .Where(x => (x.IP_T_Notes ?? "").Equals("Second Tracer", StringComparison.OrdinalIgnoreCase)
                        && (x.P_IP_T_List ?? "").StartsWith("IP", StringComparison.OrdinalIgnoreCase))
            .Select(x => (x.P_IP_T_List ?? "").Trim())
            .Where(x => x.Length > 0)
            .ToList();

        // Determine FL numbering based on repair joint's own Current T
        var isSecondTracerSource = allSecondTracerValues.Any(v => v.Equals(currentIpOrT, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(currentIpOrT) && isSecondTracerSource)
        {
            flPrefix = ipTAll
                .Where(x => (x.IP_T_Notes ?? "").Equals("Full Lot", StringComparison.OrdinalIgnoreCase))
                .Select(x => (x.P_IP_T_List ?? "").Trim())
                .FirstOrDefault(x => x.Length > 0) ?? "FL";

            flNumbering = true;
        }

        // Rule 4: When current T is Second Tracer (Full Lot stage),
        // ignore "Actual Date Welded" match criterion even if checked.
        if (flNumbering)
        {
            req.MatchDate = false;
        }

        // --- Build matched rows preview based on Match Criteria ---
        var matchedRows = await BuildMatchedRowsPreviewAsync(
            req, dfr.Project_No, dfr.DIAMETER, dfr.Line_Class,
            dwrSource?.ACTUAL_DATE_WELDED, req.WelderSymbol,
            sourceRepairWelder,
            dwrSource?.ROOT_A, dwrSource?.ROOT_B, dwrSource?.FILL_A, dwrSource?.FILL_B, dwrSource?.CAP_A, dwrSource?.CAP_B,
            isRepair, dfr.LOCATION, dfr.WELD_NUMBER);

        // Compute per-row options based on each MATCHED ROW's CurrentIpOrT
        var optionsPerRow = new Dictionary<int, List<string>>();
        var allOptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in matchedRows)
        {
            var rowOptions = GetTracerOptionsForCurrentT(row.CurrentIpOrT, ipTAll, firstTracerValues, allSecondTracerValues);
            optionsPerRow[row.JointId] = rowOptions;
            foreach (var o in rowOptions) allOptions.Add(o);
        }

        // Fallback global options (union of all per-row options)
        var options = allOptions.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();

        return Json(new { options, optionsPerRow, flNumbering, flPrefix, matchedRows, hasRepairCq = isRepair, ignoreDateMatch = flNumbering });
    }

    /// <summary>
    /// Determines tracer dropdown options based on a row's current IP_or_T value.
    /// Rule 2: NULL → First Tracer options
    /// Rule 3: First Tracer → Second Tracer options (prefixed)
    /// Rule 4: Second Tracer → Full Lot (empty, handled by FL numbering)
    /// Fallback: all distinct options
    /// </summary>
    private static List<string> GetTracerOptionsForCurrentT(
        string? currentIpOrT,
        List<PmsIpT> ipTAll,
        List<string> firstTracerValues,
        List<string> allSecondTracerValues)
    {
        if (string.IsNullOrWhiteSpace(currentIpOrT))
        {
            // Rule 2: CURRENT T IS NULL → First Tracer options
            return firstTracerValues;
        }

        if (firstTracerValues.Any(v => v.Equals(currentIpOrT, StringComparison.OrdinalIgnoreCase)))
        {
            // Rule 3: CURRENT T = First Tracer → Second Tracer options prefixed with currentIpOrT
            return ipTAll
                .Where(x => (x.IP_T_Notes ?? "").Equals("Second Tracer", StringComparison.OrdinalIgnoreCase)
                            && (x.P_IP_T_List ?? "").StartsWith(currentIpOrT!, StringComparison.OrdinalIgnoreCase))
                .Select(x => (x.P_IP_T_List ?? "").Trim())
                .Where(x => x.Length > 0)
                .ToList();
        }

        if (allSecondTracerValues.Any(v => v.Equals(currentIpOrT, StringComparison.OrdinalIgnoreCase)))
        {
            // Rule 4: CURRENT T = Second Tracer → Full Lot (options empty; FL numbering handles this)
            return new List<string>();
        }

        // Fallback: all distinct options
        return ipTAll
            .Select(x => (x.P_IP_T_List ?? "").Trim())
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private async Task<List<IpTMatchedRowDto>> BuildMatchedRowsPreviewAsync(
        IpTDropdownRequest req, int projectNo, double? diameter, string? lineClass,
        DateTime? actualDate, string welderSymbol,
        string? sourceRepairWelder,
        string? srcRootA, string? srcRootB, string? srcFillA, string? srcFillB, string? srcCapA, string? srcCapB,
        bool isRepair, string? sourceLocation, string? sourceWeldNumber)
    {
        // Load lot info for mapping dates → lot numbers
        var lotProjectId = await _context.Projects_tbl.AsNoTracking()
            .Where(p => p.Project_ID == projectNo)
            .Select(p => p.Welders_Project_ID ?? p.Project_ID)
            .FirstOrDefaultAsync();

        var lotList = await _context.Lot_No_tbl.AsNoTracking()
            .Where(l => l.Lot_Project_No == lotProjectId)
            .OrderBy(l => l.From_Date)
            .ThenBy(l => l.To_Date)
            .ThenBy(l => l.Lot_No)
            .ToListAsync();

        var sourceLot = FindLotMatchBinary(lotList, actualDate);

        // Weld type tokens from selected checkboxes
        var swSofTh = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "SW", "SOF", "TH" };
        var weldTypeFilter = (req.WeldTypes ?? new List<string>())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s.Trim().ToUpperInvariant())
            .ToList();

        // Base query:
        //   J_Add NOT LIKE 'R%'
        //   AND RT eligibility per location/weld-type
        //   AND IP_or_T eligibility
        var matchQuery = from d in _context.DFR_tbl.AsNoTracking()
                         join dw in _context.DWR_tbl.AsNoTracking() on d.Joint_ID equals dw.Joint_ID_DWR
                         join ls in _context.Line_Sheet_tbl.AsNoTracking() on d.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
                         join ll in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
                         join rt in _context.RT_tbl.AsNoTracking() on d.Joint_ID equals rt.Joint_ID_RT into rtJoin
                         from rt in rtJoin.DefaultIfEmpty()
                         where d.Project_No == projectNo
                               // J_Add NOT LIKE 'R%'
                               && !EF.Functions.Like(d.J_Add ?? string.Empty, "R%")
                               // RT eligibility per location / weld type:
                                //   (RT_Shop <> 1 AND LOCATION = 'WS')
                                //   OR (RT_Field <> 1 AND LOCATION <> 'WS')
                                //   OR (RT_Field_Shop_SW <> 1 AND WELD_TYPE IN ('SW','SOF','TH'))
                                && (
                                    (ll.RT_Shop.HasValue && Math.Abs(ll.RT_Shop.Value - 1.0) >= 0.001
                                        && (d.LOCATION ?? string.Empty) == "WS")
                                    || (ll.RT_Field.HasValue && Math.Abs(ll.RT_Field.Value - 1.0) >= 0.001
                                        && (d.LOCATION ?? string.Empty) != "WS")
                                    || (ll.RT_Field_Shop_SW.HasValue && Math.Abs(ll.RT_Field_Shop_SW.Value - 1.0) >= 0.001
                                        && ((d.WELD_TYPE ?? string.Empty) == "SW"
                                            || (d.WELD_TYPE ?? string.Empty) == "SOF"
                                            || (d.WELD_TYPE ?? string.Empty) == "TH"))
                                )
                                // IP_or_T eligibility:
                                 //   IP_or_T NOT LIKE 'IP%'
                                 //   OR (DWR_REMARKS NOT IN (ROOT_A,ROOT_B,FILL_A,FILL_B,CAP_A,CAP_B) AND IP_or_T LIKE 'IP%')
                                 //   OR IP_or_T IS NULL
                                 && (
                                     dw.IP_or_T == null
                                     || !EF.Functions.Like(dw.IP_or_T, "IP%")
                                     || (EF.Functions.Like(dw.IP_or_T, "IP%")
                                         && dw.DWR_REMARKS != null
                                         && dw.DWR_REMARKS != (dw.ROOT_A ?? string.Empty)
                                         && dw.DWR_REMARKS != (dw.ROOT_B ?? string.Empty)
                                         && dw.DWR_REMARKS != (dw.FILL_A ?? string.Empty)
                                         && dw.DWR_REMARKS != (dw.FILL_B ?? string.Empty)
                                         && dw.DWR_REMARKS != (dw.CAP_A ?? string.Empty)
                                         && dw.DWR_REMARKS != (dw.CAP_B ?? string.Empty))
                                 )
                                 // Exclude rejected joints from matched rows
                                 && (rt == null || (rt.Final_RT_RESULT != "R" && !EF.Functions.Like(rt.Final_RT_RESULT ?? string.Empty, "R/%")))
                          select new
                         {
                             d.Joint_ID,
                             d.J_Add,
                             d.LOCATION,
                             d.WELD_NUMBER,
                             d.LAYOUT_NUMBER,
                             d.SHEET,
                             d.WELD_TYPE,
                             d.DIAMETER,
                             dw.ACTUAL_DATE_WELDED,
                             dw.IP_or_T,
                             dw.DWR_REMARKS,
                             dw.ROOT_A,
                             dw.ROOT_B,
                             dw.FILL_A,
                             dw.FILL_B,
                             dw.CAP_A,
                             dw.CAP_B,
                             LineClass = ll.Line_Class,
                             RepairWelder = rt == null ? null : rt.Repair_Welder,
                             RequestedDate = rt == null ? (DateTime?)null : rt.DATE_NDE_WAS_REQUESTED
                         };

        // Filter by selected Weld Types
        if (weldTypeFilter.Count > 0)
        {
            matchQuery = matchQuery.Where(x => weldTypeFilter.Contains((x.WELD_TYPE ?? string.Empty).Trim().ToUpper()));
        }

        // Match Diameter
        if (req.MatchDiameter && diameter.HasValue)
        {
            var dia = diameter.Value;
            matchQuery = matchQuery.Where(x => x.DIAMETER.HasValue && Math.Abs(x.DIAMETER.Value - dia) < 0.001);
        }

        // Match Line Class
        if (req.MatchLineClass && !string.IsNullOrWhiteSpace(lineClass))
        {
            var lc = lineClass;
            matchQuery = matchQuery.Where(x => x.LineClass == lc);
        }

        // Match Date (same actual date welded)
        if (req.MatchDate && actualDate.HasValue)
        {
            var dt = actualDate.Value.Date;
            matchQuery = matchQuery.Where(x => x.ACTUAL_DATE_WELDED.HasValue && x.ACTUAL_DATE_WELDED.Value.Date == dt);
        }

        var candidates = await matchQuery.Take(500).ToListAsync();

        // Match Welder (in-memory – checks all welder columns)
        if (req.MatchWelder && !string.IsNullOrWhiteSpace(welderSymbol))
        {
            var ws = welderSymbol.Trim();
            candidates = candidates
                .Where(c =>
                {
                    var welders = new[] { c.ROOT_A, c.ROOT_B, c.FILL_A, c.FILL_B, c.CAP_A, c.CAP_B };
                    return welders.Any(w => !string.IsNullOrWhiteSpace(w) && w.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase));
                })
                .ToList();
        }

        // Match Process (in-memory – checks root vs fill/cap)
        if (req.MatchProcess && !string.IsNullOrWhiteSpace(sourceRepairWelder))
        {
            var rw = sourceRepairWelder.Trim();
            bool isRoot = (!string.IsNullOrWhiteSpace(srcRootA) && srcRootA.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                       || (!string.IsNullOrWhiteSpace(srcRootB) && srcRootB.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase));
            bool isFillCap = (!string.IsNullOrWhiteSpace(srcFillA) && srcFillA.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                          || (!string.IsNullOrWhiteSpace(srcFillB) && srcFillB.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                          || (!string.IsNullOrWhiteSpace(srcCapA) && srcCapA.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                          || (!string.IsNullOrWhiteSpace(srcCapB) && srcCapB.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase));

            if (isRoot)
            {
                candidates = candidates
                    .Where(c => (!string.IsNullOrWhiteSpace(c.ROOT_A) && c.ROOT_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.ROOT_B) && c.ROOT_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
            }
            else if (isFillCap)
            {
                candidates = candidates
                    .Where(c => (!string.IsNullOrWhiteSpace(c.FILL_A) && c.FILL_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.FILL_B) && c.FILL_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.CAP_A) && c.CAP_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.CAP_B) && c.CAP_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
            }
        }

        // Match Lot (in-memory – compare lot numbers)
        if (req.MatchLot && sourceLot != null)
        {
            var srcLotNo = sourceLot.Lot_No;
            candidates = candidates
                .Where(c =>
                {
                    var lot = FindLotMatchBinary(lotList, c.ACTUAL_DATE_WELDED);
                    return lot != null && lot.Lot_No == srcLotNo;
                })
                .ToList();
        }

        // Build result DTOs
        var result = candidates
            .Select(c =>
            {
                var loc = (c.LOCATION ?? "").Trim();
                var weld = (c.WELD_NUMBER ?? "").Trim();
                var jAdd = (c.J_Add ?? "").Trim();
                var baseJointNo = loc.Length > 0 && weld.Length > 0 ? loc + "-" + weld : (loc + weld);
                var jointNo = jAdd.Length > 0 && !jAdd.Equals("New", StringComparison.OrdinalIgnoreCase)
                    ? baseJointNo + jAdd
                    : baseJointNo;
                var lot = FindLotMatchBinary(lotList, c.ACTUAL_DATE_WELDED);
                return new IpTMatchedRowDto
                {
                    JointId = c.Joint_ID,
                    JointNo = jointNo,
                    JAdd = c.J_Add ?? string.Empty,
                    LayoutNumber = c.LAYOUT_NUMBER ?? string.Empty,
                    Sheet = c.SHEET,
                    ActualDateWelded = c.ACTUAL_DATE_WELDED,
                    Diameter = c.DIAMETER,
                    LineClass = c.LineClass ?? string.Empty,
                    Root = string.Join(", ", new[] { c.ROOT_A, c.ROOT_B }
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Select(v => v!.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)),
                    FillCap = string.Join(", ", new[] { c.FILL_A, c.FILL_B, c.CAP_A, c.CAP_B }
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Select(v => v!.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)),
                    CurrentIpOrT = c.IP_or_T,
                    LotNo = lot?.Lot_No ?? string.Empty,
                    RequestedDate = c.RequestedDate
                };
            })
            .GroupBy(r => r.JointId)
            .Select(g => g.First())
            .OrderBy(r => r.LineClass)
            .ThenBy(r => r.JointNo)
            .ToList();

        return result;
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> UpdateIpOrT([FromBody] IpTUpdateRequest? req)
    {
        if (req == null || req.JointId <= 0) return BadRequest("Invalid request");

        var userId = HttpContext.Session.GetInt32("UserID");
        var now = AppClock.Now;

        static string? Clip(string? value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        // Get the joint info for matching
        var dfrInfo = await (from dfr in _context.DFR_tbl.AsNoTracking()
                            join ls in _context.Line_Sheet_tbl.AsNoTracking() on dfr.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
                            join ll in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
                            where dfr.Joint_ID == req.JointId
                            select new
                            {
                                dfr.Joint_ID,
                                dfr.LOCATION,
                                dfr.WELD_NUMBER,
                                dfr.DIAMETER,
                                dfr.Project_No,
                                ll.Line_Class
                            }).FirstOrDefaultAsync();

        if (dfrInfo == null) return NotFound("Joint not found");

        var dwrSource = await _context.DWR_tbl.AsNoTracking()
            .Where(d => d.Joint_ID_DWR == req.JointId)
            .Select(d => new { d.ACTUAL_DATE_WELDED, d.ROOT_A, d.ROOT_B, d.FILL_A, d.FILL_B, d.CAP_A, d.CAP_B })
            .FirstOrDefaultAsync();

        var sourceRepairWelder = await _context.RT_tbl.AsNoTracking()
            .Where(r => r.Joint_ID_RT == req.JointId)
            .Select(r => r.Repair_Welder)
            .FirstOrDefaultAsync();

        // Check if FL numbering is requested
        var weldersProjectId = await ResolveWeldersProjectIdAsync(req.ProjectId);
        var ipTAll = await _context.PMS_IP_T_tbl.AsNoTracking()
            .Where(x => x.IP_T_Project_No == weldersProjectId || x.IP_T_Project_No == 0)
            .ToListAsync();

        var secondTracerValues = ipTAll
            .Where(x => (x.IP_T_Notes ?? "").Equals("Second Tracer", StringComparison.OrdinalIgnoreCase)
                        && (x.P_IP_T_List ?? "").StartsWith("IP", StringComparison.OrdinalIgnoreCase))
            .Select(x => (x.P_IP_T_List ?? "").Trim())
            .Where(x => x.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        bool isFlNumbering = !string.IsNullOrWhiteSpace(req.IpOrTValue)
                             && req.IpOrTValue.Equals("FL_AUTO", StringComparison.OrdinalIgnoreCase);

        if (isFlNumbering)
        {
            // Rule 5: Assign FL-01, FL-02, etc. to all matching rows by (Lot, Diameter, LineClass, Welder)
            bool? sourceIsRootProcess = null;
            if (dwrSource != null && !string.IsNullOrWhiteSpace(sourceRepairWelder))
            {
                var rw = sourceRepairWelder.Trim();
                bool isRoot = (!string.IsNullOrWhiteSpace(dwrSource.ROOT_A) && dwrSource.ROOT_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                           || (!string.IsNullOrWhiteSpace(dwrSource.ROOT_B) && dwrSource.ROOT_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase));
                bool isFillCap = (!string.IsNullOrWhiteSpace(dwrSource.FILL_A) && dwrSource.FILL_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                              || (!string.IsNullOrWhiteSpace(dwrSource.FILL_B) && dwrSource.FILL_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                              || (!string.IsNullOrWhiteSpace(dwrSource.CAP_A) && dwrSource.CAP_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                              || (!string.IsNullOrWhiteSpace(dwrSource.CAP_B) && dwrSource.CAP_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase));
                if (isRoot) sourceIsRootProcess = true;
                else if (isFillCap) sourceIsRootProcess = false;
            }

            return await ApplyFlNumberingAsync(req, dfrInfo.Project_No, dfrInfo.DIAMETER,
                dfrInfo.Line_Class, dwrSource?.ACTUAL_DATE_WELDED, req.WelderSymbol,
                sourceRepairWelder, sourceIsRootProcess, userId, now);
        }

        // Build the same matched rows as the preview to determine which joints to update
        var dropdownReq = new IpTDropdownRequest
        {
            JointId = req.JointId,
            ProjectId = req.ProjectId,
            WelderSymbol = req.WelderSymbol,
            WeldTypes = req.WeldTypes,
            MatchLot = req.MatchLot,
            MatchDate = req.MatchDate,
            MatchDiameter = req.MatchDiameter,
            MatchLineClass = req.MatchLineClass,
            MatchWelder = req.MatchWelder,
            MatchProcess = req.MatchProcess
        };

        var rtResult = await _context.RT_tbl.AsNoTracking()
            .Where(r => r.Joint_ID_RT == req.JointId)
            .Select(r => r.Final_RT_RESULT)
            .FirstOrDefaultAsync();
        var trimmedRt = (rtResult ?? "").Trim();
        bool isRepair = trimmedRt.Equals("R", StringComparison.OrdinalIgnoreCase)
                     || trimmedRt.StartsWith("R/", StringComparison.OrdinalIgnoreCase);

        var matchedRows = await BuildMatchedRowsPreviewAsync(
            dropdownReq, dfrInfo.Project_No, dfrInfo.DIAMETER, dfrInfo.Line_Class,
            dwrSource?.ACTUAL_DATE_WELDED, req.WelderSymbol,
            sourceRepairWelder,
            dwrSource?.ROOT_A, dwrSource?.ROOT_B, dwrSource?.FILL_A, dwrSource?.FILL_B, dwrSource?.CAP_A, dwrSource?.CAP_B,
            isRepair, dfrInfo.LOCATION, dfrInfo.WELD_NUMBER);

        var jointIdsToUpdate = matchedRows.Select(r => r.JointId).Distinct().ToList();

        // Update all matched DWR rows
        var dwrEntities = await _context.DWR_tbl
            .Where(d => jointIdsToUpdate.Contains(d.Joint_ID_DWR))
            .ToListAsync();

        var existingIds = dwrEntities.Select(d => d.Joint_ID_DWR).ToHashSet();
        var missingIds = jointIdsToUpdate.Where(id => !existingIds.Contains(id)).ToList();

        if (missingIds.Count > 0)
        {
            var validMissing = await _context.DFR_tbl.AsNoTracking()
                .Where(d => missingIds.Contains(d.Joint_ID))
                .Select(d => d.Joint_ID)
                .ToListAsync();

            foreach (var mid in validMissing)
            {
                var newEntity = new Dwr { Joint_ID_DWR = mid };
                _context.DWR_tbl.Add(newEntity);
                dwrEntities.Add(newEntity);
            }
        }

        var clippedValue = Clip(req.IpOrTValue, 8);
        foreach (var dwr in dwrEntities)
        {
            dwr.IP_or_T = clippedValue;
            dwr.DWR_Updated_By = userId;
            dwr.DWR_Updated_Date = now;
        }

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateException ex)
        {
            _logger.LogError(ex, "UpdateIpOrT failed for Joint_ID {JointId}", req.JointId);
            return StatusCode(StatusCodes.Status500InternalServerError, new
            {
                ok = false,
                message = "Failed to update IP/T.",
                detail = GetInnermostMessage(ex)
            });
        }

        return Json(new { ok = true, updatedCount = dwrEntities.Count, updatedIds = dwrEntities.Select(d => d.Joint_ID_DWR).ToList() });
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public async Task<IActionResult> BatchUpdateIpOrT([FromBody] IpTBatchUpdateRequest? req)
    {
        if (req == null || req.Assignments == null || req.Assignments.Count == 0)
            return BadRequest("Invalid request");

        var userId = HttpContext.Session.GetInt32("UserID");
        var now = AppClock.Now;

        static string? Clip(string? value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        // Only process rows that have a non-empty value assigned
        var validAssignments = req.Assignments
            .Where(a => a.JointId > 0 && !string.IsNullOrWhiteSpace(a.IpOrTValue))
            .ToList();

        if (validAssignments.Count == 0)
            return Json(new { ok = true, updatedCount = 0, updatedIds = new List<int>() });

        var jointIds = validAssignments.Select(a => a.JointId).Distinct().ToList();

        var dwrEntities = await _context.DWR_tbl
            .Where(d => jointIds.Contains(d.Joint_ID_DWR))
            .ToListAsync();

        var existingIds = dwrEntities.Select(d => d.Joint_ID_DWR).ToHashSet();
        var missingIds = jointIds.Where(id => !existingIds.Contains(id)).ToList();

        if (missingIds.Count > 0)
        {
            var validMissing = await _context.DFR_tbl.AsNoTracking()
                .Where(d => missingIds.Contains(d.Joint_ID))
                .Select(d => d.Joint_ID)
                .ToListAsync();

            foreach (var mid in validMissing)
            {
                var newEntity = new Dwr { Joint_ID_DWR = mid };
                _context.DWR_tbl.Add(newEntity);
                dwrEntities.Add(newEntity);
            }
        }

        // Build a lookup of jointId → value
        var assignmentMap = validAssignments.ToDictionary(a => a.JointId, a => Clip(a.IpOrTValue, 8));

        foreach (var dwr in dwrEntities)
        {
            if (assignmentMap.TryGetValue(dwr.Joint_ID_DWR, out var value))
            {
                dwr.IP_or_T = value;
                dwr.DWR_Updated_By = userId;
                dwr.DWR_Updated_Date = now;
            }
        }

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateException ex)
        {
            _logger.LogError(ex, "BatchUpdateIpOrT failed");
            return StatusCode(StatusCodes.Status500InternalServerError, new
            {
                ok = false,
                message = "Failed to update IP/T.",
                detail = GetInnermostMessage(ex)
            });
        }

        return Json(new { ok = true, updatedCount = validAssignments.Count, updatedIds = jointIds });
    }

    private async Task<IActionResult> ApplyFlNumberingAsync(
        IpTUpdateRequest req, int projectNo, double? diameter, string? lineClass,
        DateTime? actualDate, string welderSymbol,
        string? sourceRepairWelder, bool? sourceIsRootProcess, int? userId, DateTime now)
    {
        // Find all matching rows by (Lot / Diameter / LineClass / Welder) - Rule 5
        var matchingQuery = from dfr in _context.DFR_tbl.AsNoTracking()
                            join dw in _context.DWR_tbl.AsNoTracking() on dfr.Joint_ID equals dw.Joint_ID_DWR
                            join ls in _context.Line_Sheet_tbl.AsNoTracking() on dfr.Line_Sheet_ID_DFR equals ls.Line_Sheet_ID
                            join ll in _context.LINE_LIST_tbl.AsNoTracking() on ls.Line_ID_LS equals ll.Line_ID
                            where dfr.Project_No == projectNo
                                  && dfr.WELD_TYPE == "BW"
                            select new
                            {
                                dfr.Joint_ID,
                                dfr.DIAMETER,
                                dw.ACTUAL_DATE_WELDED,
                                ll.Line_Class,
                                dw.ROOT_A,
                                dw.ROOT_B,
                                dw.FILL_A,
                                dw.FILL_B,
                                dw.CAP_A,
                                dw.CAP_B
                            };

        if (req.MatchDiameter && diameter.HasValue)
        {
            var dia = diameter.Value;
            matchingQuery = matchingQuery.Where(x => x.DIAMETER.HasValue && Math.Abs(x.DIAMETER.Value - dia) < 0.001);
        }

        if (req.MatchLineClass && !string.IsNullOrWhiteSpace(lineClass))
        {
            var lc = lineClass;
            matchingQuery = matchingQuery.Where(x => x.Line_Class == lc);
        }

        var candidates = await matchingQuery.ToListAsync();

        if (req.MatchWelder && !string.IsNullOrWhiteSpace(welderSymbol))
        {
            var ws = welderSymbol.Trim();
            candidates = candidates
                .Where(c =>
                {
                    var welders = new[] { c.ROOT_A, c.ROOT_B, c.FILL_A, c.FILL_B, c.CAP_A, c.CAP_B };
                    return welders.Any(w => !string.IsNullOrWhiteSpace(w) && w.Trim().Equals(ws, StringComparison.OrdinalIgnoreCase));
                })
                .ToList();
        }

        if (req.MatchProcess && sourceIsRootProcess.HasValue && !string.IsNullOrWhiteSpace(sourceRepairWelder))
        {
            var rw = sourceRepairWelder.Trim();
            if (sourceIsRootProcess.Value)
            {
                candidates = candidates
                    .Where(c => (!string.IsNullOrWhiteSpace(c.ROOT_A) && c.ROOT_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.ROOT_B) && c.ROOT_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
            }
            else
            {
                candidates = candidates
                    .Where(c => (!string.IsNullOrWhiteSpace(c.FILL_A) && c.FILL_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.FILL_B) && c.FILL_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.CAP_A) && c.CAP_A.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase))
                             || (!string.IsNullOrWhiteSpace(c.CAP_B) && c.CAP_B.Trim().Equals(rw, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
            }
        }

        // Match Lot (in-memory – compare lot numbers)
        if (req.MatchLot && actualDate.HasValue)
        {
            var lotProjectIdFl = await _context.Projects_tbl.AsNoTracking()
                .Where(p => p.Project_ID == projectNo)
                .Select(p => p.Welders_Project_ID ?? p.Project_ID)
                .FirstOrDefaultAsync();

            var lotListFl = await _context.Lot_No_tbl.AsNoTracking()
                .Where(l => l.Lot_Project_No == lotProjectIdFl)
                .OrderBy(l => l.From_Date)
                .ThenBy(l => l.To_Date)
                .ThenBy(l => l.Lot_No)
                .ToListAsync();

            var sourceLotFl = FindLotMatchBinary(lotListFl, actualDate);
            if (sourceLotFl != null)
            {
                var srcLotNo = sourceLotFl.Lot_No;
                candidates = candidates
                    .Where(c =>
                    {
                        var lot = FindLotMatchBinary(lotListFl, c.ACTUAL_DATE_WELDED);
                        return lot != null && lot.Lot_No == srcLotNo;
                    })
                    .ToList();
            }
        }

        var jointIds = candidates.Select(c => c.Joint_ID).Distinct().ToList();

        var dwrEntities = await _context.DWR_tbl
            .Where(d => jointIds.Contains(d.Joint_ID_DWR))
            .ToListAsync();

        var existingIds = dwrEntities.Select(d => d.Joint_ID_DWR).ToHashSet();
        var missingIds = jointIds.Where(id => !existingIds.Contains(id)).ToList();

        if (missingIds.Count > 0)
        {
            var validMissing = await _context.DFR_tbl.AsNoTracking()
                .Where(d => missingIds.Contains(d.Joint_ID))
                .Select(d => d.Joint_ID)
                .ToListAsync();

            foreach (var mid in validMissing)
            {
                var newEntity = new Dwr { Joint_ID_DWR = mid };
                _context.DWR_tbl.Add(newEntity);
                dwrEntities.Add(newEntity);
            }
        }

        // Resolve Full Lot prefix from PMS_IP_T_tbl
        var weldersProjectIdFl = await ResolveWeldersProjectIdAsync(req.ProjectId);
        var ipTAllFl = await _context.PMS_IP_T_tbl.AsNoTracking()
            .Where(x => x.IP_T_Project_No == weldersProjectIdFl || x.IP_T_Project_No == 0)
            .ToListAsync();

        var flPrefix = ipTAllFl
            .Where(x => (x.IP_T_Notes ?? "").Equals("Full Lot", StringComparison.OrdinalIgnoreCase))
            .Select(x => (x.P_IP_T_List ?? "").Trim())
            .FirstOrDefault(x => x.Length > 0) ?? "FL";

        // Assign {prefix}-01, {prefix}-02, etc.
        var orderedEntities = dwrEntities.OrderBy(d => d.Joint_ID_DWR).ToList();
        for (int i = 0; i < orderedEntities.Count; i++)
        {
            var flValue = $"{flPrefix}-{(i + 1):D2}";
            if (flValue.Length > 8) flValue = flValue[..8];
            orderedEntities[i].IP_or_T = flValue;
            orderedEntities[i].DWR_Updated_By = userId;
            orderedEntities[i].DWR_Updated_Date = now;
        }

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateException ex)
        {
            _logger.LogError(ex, "ApplyFlNumbering failed for Joint_ID {JointId}", req.JointId);
            return StatusCode(StatusCodes.Status500InternalServerError, new
            {
                ok = false,
                message = "Failed to apply FL numbering.",
                detail = GetInnermostMessage(ex)
            });
        }

        return Json(new { ok = true, updatedCount = orderedEntities.Count, updatedIds = orderedEntities.Select(d => d.Joint_ID_DWR).ToList() });
    }

    /// <summary>
    /// Returns the status for a row based on the peer group formed by matching criteria.
    /// <para><b>Not Available</b> — Joint_Availability == 0 (repair welder has zero joints in the group).</para>
    /// <para><b>Done</b> — Req_Joint_Count (R_Joint_Count × 2) ≤ Total_T_Joints, OR all candidate joints are already assigned (no null IP_or_T).</para>
    /// <para><b>2nd Not Available</b> — Joint_Availability &lt; 2 AND no unassigned joints (IP_or_T IS NULL) exist for this welder.</para>
    /// <para><b>Pending</b> — otherwise.</para>
    /// </summary>
    private static string ComputeTracerStatus(int jointAvailability, int reqJointCount, int totalTJoints, int nullIpOrTAvailability)
    {
        if (jointAvailability == 0)
            return "Not Available";

        if (reqJointCount <= totalTJoints)
            return "Done";

        if (jointAvailability < 2 && nullIpOrTAvailability == 0)
            return "2nd Not Available";

        // All candidate joints already have an IP/T value (e.g. LOT, T1, T2) —
        // nothing left to select, so the tracer selection is complete.
        if (nullIpOrTAvailability == 0)
            return "Done";

        return "Pending";
    }
}