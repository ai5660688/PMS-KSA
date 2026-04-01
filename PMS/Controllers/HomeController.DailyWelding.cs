using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;
using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PMS.Controllers;

public partial class HomeController
{
    private static readonly string[] DailyWeldingDateFormats =
    {
        "yyyy-MM-dd",
        "yyyy-MM-ddTHH:mm",
        "yyyy-MM-ddTHH:mm:ss"
    };

    // Use source-generated regex to avoid runtime Regex compilation and satisfy SYSLIB1045 analyzer
    [System.Text.RegularExpressions.GeneratedRegex("\\d+")]
    private static partial System.Text.RegularExpressions.Regex _rx_digits();

    [SessionAuthorization]
    [HttpGet]
    [Route("/Input/Welding/DailyWelding")]
    public async Task<IActionResult> DailyWelding(
        int? projectId,
        string? location,
        string? headerView,
        string? layout,
        string? sheet,
        string? fitupDateFilter,
        string? actualDateFilter,
        string? fitupReportFilter,
        string? doLoad,
        string? date,
        string? FitupReportHeader,
        string? FitupReportCombined,
        int? SelectedRfiId,
        string? SelectedTacker)
    {
        var dateFilterValue = !string.IsNullOrWhiteSpace(actualDateFilter) ? actualDateFilter : fitupDateFilter;
        var result = await DailyFitup(projectId, location, headerView, layout, sheet, dateFilterValue, fitupReportFilter, doLoad, date, FitupReportHeader, FitupReportCombined, SelectedRfiId, SelectedTacker, actualDateFilter);
        if (result is not ViewResult { Model: DfrDailyFitupViewModel vm })
        {
            return result;
        }

        var weldersProjectId = await ResolveWeldersProjectIdAsync(vm.SelectedProjectId);

        if (vm.SelectedProjectId > 0)
        {
            var hv = string.IsNullOrWhiteSpace(vm.HeaderView) ? "DWG" : vm.HeaderView!.Trim();
            var headerLoc = string.IsNullOrWhiteSpace(vm.SelectedLocation) ? (location ?? "All") : vm.SelectedLocation!;
            var locCode = MapHeaderLocation(headerLoc);

            var q = from w in _context.DWR_tbl.AsNoTracking()
                    join d in _context.DFR_tbl.AsNoTracking() on w.Joint_ID_DWR equals d.Joint_ID
                    where d.Project_No == vm.SelectedProjectId
                       && w.POST_VISUAL_INSPECTION_QR_NO != null && w.POST_VISUAL_INSPECTION_QR_NO != ""
                    select new
                    {
                        d.LOCATION,
                        d.WELD_TYPE,
                        d.LAYOUT_NUMBER,
                        d.SHEET,
                        d.FITUP_DATE,
                        w.POST_VISUAL_INSPECTION_QR_NO,
                        w.DATE_WELDED,
                        w.ACTUAL_DATE_WELDED,
                        d.Joint_ID
                    };

            if (string.Equals(hv, "DWG", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(layout)) q = q.Where(x => x.LAYOUT_NUMBER == layout);
                if (!string.IsNullOrWhiteSpace(sheet)) q = q.Where(x => x.SHEET == sheet);
            }
            else if (string.Equals(hv, "Date", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(dateFilterValue) && DateTime.TryParse(dateFilterValue, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var day))
                {
                    var start = day.Date;
                    var end = start.AddDays(1);
                    q = q.Where(x => x.ACTUAL_DATE_WELDED >= start && x.ACTUAL_DATE_WELDED < end);
                }
            }
            else if (string.Equals(hv, "Report", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(fitupReportFilter))
                {
                    var key = fitupReportFilter.Trim();
                    q = q.Where(x => x.POST_VISUAL_INSPECTION_QR_NO!.Contains(key));
                }
            }

            static string ComputeNextQr(System.Collections.Generic.List<string> items)
            {
                if (items == null || items.Count == 0)
                {
                    return "00001";
                }
                long max = -1;
                var rx = _rx_digits();
                foreach (var s in items)
                {
                    if (string.IsNullOrWhiteSpace(s)) continue;
                    var matches = rx.Matches(s);
                    foreach (Match m in matches)
                    {
                        if (long.TryParse(m.Value, out var v))
                        {
                            if (v > max) max = v;
                        }
                    }
                }
                if (max >= 0)
                {
                    var next = max + 1;
                    if (next < 1) next = 1;
                    return next.ToString("D5");
                }
                return "00001";
            }

            static bool ContainsToken(string? source, string token) =>
                !string.IsNullOrWhiteSpace(source) && source.Contains(token, StringComparison.OrdinalIgnoreCase);

            static bool IsShop(string? location)
            {
                if (string.IsNullOrWhiteSpace(location)) return false;
                var value = location.Trim();
                return ContainsToken(value, "shop") || ContainsToken(value, "ws") || ContainsToken(value, "work");
            }

            static bool IsField(string? location) => !string.IsNullOrWhiteSpace(location) && !IsShop(location);

            static bool IsThreaded(string? weldType, string? location) => ContainsToken(weldType, "TH") || ContainsToken(location, "thread");

            static System.Collections.Generic.List<string> BuildQrList(System.Collections.Generic.IEnumerable<string?> source)
            {
                var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var list = new System.Collections.Generic.List<string>();
                if (source == null) return list;
                foreach (var item in source)
                {
                    if (string.IsNullOrWhiteSpace(item)) continue;
                    var trimmed = item.Trim();
                    if (trimmed.Length == 0) continue;
                    if (seen.Add(trimmed)) list.Add(trimmed);
                }
                return list;
            }

            // Load QR candidates once to avoid repeated LIKE scans that caused SQL timeouts.
            var qrCandidates = await q
                .Where(x => !string.IsNullOrWhiteSpace(x.POST_VISUAL_INSPECTION_QR_NO))
                .Select(x => new
                {
                    Location = x.LOCATION,
                    WeldType = x.WELD_TYPE,
                    Qr = x.POST_VISUAL_INSPECTION_QR_NO!
                })
                .ToListAsync();

            if (string.Equals(headerLoc, "All", StringComparison.OrdinalIgnoreCase))
            {
                var shopList = BuildQrList(qrCandidates.Where(x => IsShop(x.Location)).Select(x => x.Qr));
                var fieldList = BuildQrList(qrCandidates.Where(x => IsField(x.Location)).Select(x => x.Qr));
                var threadedList = BuildQrList(qrCandidates.Where(x => IsThreaded(x.WeldType, x.Location)).Select(x => x.Qr));

                var nextShop = ComputeNextQr(shopList);
                var nextField = ComputeNextQr(fieldList);
                var nextThreaded = ComputeNextQr(threadedList);

                if (!string.IsNullOrWhiteSpace(nextShop)) vm.FitupReportShop = nextShop;
                if (!string.IsNullOrWhiteSpace(nextField)) vm.FitupReportField = nextField;
                if (!string.IsNullOrWhiteSpace(nextThreaded)) vm.FitupReportThreaded = nextThreaded;
            }
            else if (string.Equals(headerLoc, "Threaded", StringComparison.OrdinalIgnoreCase))
            {
                var threadedList = BuildQrList(qrCandidates.Where(x => IsThreaded(x.WeldType, x.Location)).Select(x => x.Qr));
                var nextThreaded = ComputeNextQr(threadedList);
                if (!string.IsNullOrWhiteSpace(nextThreaded)) vm.FitupReportHeader = nextThreaded;
            }
            else if (string.Equals(locCode, "WS", StringComparison.OrdinalIgnoreCase) || string.Equals(headerLoc, "Shop", StringComparison.OrdinalIgnoreCase))
            {
                var shopList = BuildQrList(qrCandidates.Where(x => IsShop(x.Location)).Select(x => x.Qr));
                var nextShop = ComputeNextQr(shopList);
                if (!string.IsNullOrWhiteSpace(nextShop)) vm.FitupReportHeader = nextShop;
            }
            else
            {
                var fieldList = BuildQrList(qrCandidates.Where(x => IsField(x.Location)).Select(x => x.Qr));
                var nextField = ComputeNextQr(fieldList);
                if (!string.IsNullOrWhiteSpace(nextField)) vm.FitupReportHeader = nextField;
            }
        }

        {
            var rfiQuery = _context.RFI_tbl.AsNoTracking()
                .Where(r => r.SubDiscipline != null && EF.Functions.Like(r.SubDiscipline!, "%Welding%"))
                .Where(r => r.ACTIVITY == 3.9 || r.ACTIVITY == 3.90 || r.ACTIVITY == 3.900);

            if (weldersProjectId > 0)
            {
                rfiQuery = rfiQuery.Where(r => r.RFI_Project_No == weldersProjectId);
            }

            var rfiRaw = await rfiQuery
                .OrderByDescending(r => r.Date)
                .ThenByDescending(r => r.Time)
                .ThenByDescending(r => r.RFI_ID)
                .Take(2000)
                .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.Date, r.Time, r.RFI_LOCATION, r.RFI_DESCRIPTION })
                .ToListAsync();

            static string PrefixFor(string? rfiLoc)
            {
                var raw = (rfiLoc ?? string.Empty).ToUpperInvariant();
                if (raw.Contains("SHOP") || raw.Contains("WORK") || raw.StartsWith("WS")) return "WS | ";
                if (raw.Contains("FIELD") || raw.StartsWith("FW")) return "FW | ";
                return string.Empty;
            }

            var opts = rfiRaw.Select(r => new RfiOption
            {
                Id = r.RFI_ID,
                Value = r.SubCon_RFI_No ?? string.Empty,
                Display = PrefixFor(r.RFI_LOCATION)
                          + (r.SubCon_RFI_No ?? r.RFI_ID.ToString())
                          + (r.Date.HasValue ? (" | " + r.Date.Value.ToString("dd-MMM-yyyy")) : string.Empty)
                          + (r.Time.HasValue ? (" " + r.Time.Value.ToString("HH:mm")) : string.Empty)
                          + (string.IsNullOrWhiteSpace(r.RFI_DESCRIPTION) ? string.Empty : (" | " + r.RFI_DESCRIPTION))
            }).ToList();

            if (opts.Count > 0)
            {
                vm.RfiOptions = opts;
                if (!vm.SelectedRfiId.HasValue)
                {
                    vm.SelectedRfiId = opts.First().Id;
                }
            }
        }

        try
        {
            var ipQuery = _context.PMS_IP_T_tbl.AsNoTracking();
            if (weldersProjectId > 0)
            {
                ipQuery = ipQuery.Where(x => x.IP_T_Project_No == weldersProjectId || x.IP_T_Project_No == 0);
            }

            var ipOptions = await ipQuery
                .Select(x => x.IP_T_List)
                .Where(x => x != null && x != "")
                .ToListAsync();

            vm.IpOrTOptions = ipOptions
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x!.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "DailyWelding: failed to load IP/T list");
        }

        if (vm.WpsOptions == null || vm.WpsOptions.Count == 0)
        {
            try
            {
                var wpsQuery = _context.WPS_tbl.AsNoTracking();
                if (weldersProjectId > 0)
                {
                    wpsQuery = wpsQuery.Where(w => w.Project_WPS == weldersProjectId);
                }
                var wpsList = await wpsQuery
                    .OrderBy(w => w.WPS)
                    .Select(w => new WpsOption
                    {
                        Id = w.WPS_ID,
                        Wps = w.WPS,
                        ThicknessRange = w.Thickness_Range,
                        Pwht = w.PWHT
                    })
                    .ToListAsync();
                vm.WpsOptions = wpsList;
            }
            catch
            {
            }
        }

        try
        {
            if (vm.Rows != null && vm.Rows.Count > 0)
            {
                var jointIds = vm.Rows.Select(r => r.JointId).Where(id => id > 0).Distinct().ToList();
                if (jointIds.Count > 0)
                {
                    const int chunkSize = 900;
                    var dwrRows = new System.Collections.Generic.List<(int JointId, string? Qr, int? RfiId, DateTime? DateWelded, DateTime? ActualDate, int? WpsId, int? UpdatedBy, DateTime? UpdatedDate, string? RootA, string? RootB, string? FillA, string? FillB, string? CapA, string? CapB, double? Preheat, string? IPorT, string? OpenClosed, bool? WeldConfirmed, string? Remarks)>();
                    foreach (var batch in jointIds.Chunk(chunkSize))
                    {
                        var chunk = await _context.DWR_tbl.AsNoTracking()
                            .Where(w => batch.Contains(w.Joint_ID_DWR))
                            .Select(w => new
                            {
                                w.Joint_ID_DWR,
                                w.POST_VISUAL_INSPECTION_QR_NO,
                                w.RFI_ID_DWR,
                                w.DATE_WELDED,
                                w.ACTUAL_DATE_WELDED,
                                w.WPS_ID_DWR,
                                w.DWR_Updated_By,
                                w.DWR_Updated_Date,
                                w.ROOT_A,
                                w.ROOT_B,
                                w.FILL_A,
                                w.FILL_B,
                                w.CAP_A,
                                w.CAP_B,
                                w.PREHEAT_TEMP_C,
                                w.IP_or_T,
                                w.Open_Closed,
                                w.Weld_Confirmed,
                                w.DWR_REMARKS
                            })
                            .ToListAsync();

                        dwrRows.AddRange(chunk.Select(x => (
                            JointId: x.Joint_ID_DWR,
                            Qr: x.POST_VISUAL_INSPECTION_QR_NO,
                            RfiId: x.RFI_ID_DWR,
                            DateWelded: x.DATE_WELDED,
                            ActualDate: x.ACTUAL_DATE_WELDED,
                            WpsId: x.WPS_ID_DWR,
                            UpdatedBy: x.DWR_Updated_By,
                            UpdatedDate: x.DWR_Updated_Date,
                            RootA: x.ROOT_A,
                            RootB: x.ROOT_B,
                            FillA: x.FILL_A,
                            FillB: x.FILL_B,
                            CapA: x.CAP_A,
                            CapB: x.CAP_B,
                            Preheat: x.PREHEAT_TEMP_C,
                            IPorT: x.IP_or_T,
                            OpenClosed: x.Open_Closed,
                            WeldConfirmed: (bool?)x.Weld_Confirmed,
                            Remarks: x.DWR_REMARKS)));
                    }

                    var dwrMap = dwrRows
                        .GroupBy(x => x.JointId)
                        .ToDictionary(
                            g => g.Key,
                            g => g.OrderByDescending(row => row.UpdatedDate ?? DateTime.MinValue).First());

                    var dwrUpdaterIds = dwrRows
                        .Select(x => x.UpdatedBy)
                        .Where(id => id.HasValue && id.Value > 0)
                        .Select(id => id!.Value)
                        .Distinct()
                        .ToList();

                    var dwrUpdaterLookup = dwrUpdaterIds.Count == 0
                        ? new System.Collections.Generic.Dictionary<int, string>()
                        : await _context.PMS_Login_tbl.AsNoTracking()
                            .Where(u => dwrUpdaterIds.Contains(u.UserID))
                            .Select(u => new
                            {
                                u.UserID,
                                Full = (((u.FirstName ?? string.Empty).Trim() + " " + (u.LastName ?? string.Empty).Trim()).Trim())
                            })
                            .ToDictionaryAsync(x => x.UserID, x => x.Full);

                    vm.RfiOptions ??= new System.Collections.Generic.List<RfiOption>();
                    var existingRfiIds = new System.Collections.Generic.HashSet<int>(vm.RfiOptions.Select(o => o.Id));
                    var missingRfiIds = dwrMap.Values
                        .Select(x => x.RfiId)
                        .Where(id => id.HasValue && id.Value > 0 && !existingRfiIds.Contains(id.Value))
                        .Select(id => id!.Value)
                        .Distinct()
                        .ToList();

                    if (missingRfiIds.Count > 0)
                    {
                        static string BuildRfiDisplay(string? rfiLocation, string? subConNo, DateTime? rfiDate, DateTime? rfiTime, string? description)
                        {
                            var locUpper = (rfiLocation ?? string.Empty).ToUpperInvariant();
                            string prefix = string.Empty;
                            if (locUpper.Contains("SHOP") || locUpper.Contains("WORK") || locUpper.StartsWith("WS")) prefix = "WS | ";
                            else if (locUpper.Contains("FIELD") || locUpper.StartsWith("FW")) prefix = "FW | ";

                            var parts = new System.Collections.Generic.List<string>();
                            if (!string.IsNullOrWhiteSpace(subConNo)) parts.Add(subConNo!);
                            if (rfiDate.HasValue) parts.Add(rfiDate.Value.ToString("dd-MMM-yyyy"));
                            if (rfiTime.HasValue) parts.Add(rfiTime.Value.ToString("HH:mm"));
                            if (!string.IsNullOrWhiteSpace(description)) parts.Add(description!);

                            var body = string.Join(" | ", parts);
                            return prefix + body;
                        }

                        var extraRfis = await _context.RFI_tbl.AsNoTracking()
                            .Where(r => missingRfiIds.Contains(r.RFI_ID))
                            .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.Date, r.Time, r.RFI_LOCATION, r.RFI_DESCRIPTION })
                            .ToListAsync();

                        foreach (var rfi in extraRfis)
                        {
                            vm.RfiOptions.Add(new RfiOption
                            {
                                Id = rfi.RFI_ID,
                                Value = rfi.SubCon_RFI_No ?? rfi.RFI_ID.ToString(),
                                Display = BuildRfiDisplay(rfi.RFI_LOCATION, rfi.SubCon_RFI_No, rfi.Date, rfi.Time, rfi.RFI_DESCRIPTION)
                            });
                            existingRfiIds.Add(rfi.RFI_ID);
                        }

                        foreach (var orphanId in missingRfiIds.Except(extraRfis.Select(r => r.RFI_ID)))
                        {
                            vm.RfiOptions.Add(new RfiOption
                            {
                                Id = orphanId,
                                Value = orphanId.ToString(),
                                Display = orphanId.ToString()
                            });
                            existingRfiIds.Add(orphanId);
                        }
                    }

                    var rfiLookup = vm.RfiOptions
                        .GroupBy(o => o.Id)
                        .ToDictionary(g => g.Key, g => g.First());
                    // Build a simple id -> short value lookup so per-row RFI column shows SubCon_RFI_No (or id) only
                    var rfiValueLookup = vm.RfiOptions
                        .GroupBy(o => o.Id)
                        .ToDictionary(g => g.Key, g => (g.First().Value ?? string.Empty));

                    vm.WpsOptions ??= new System.Collections.Generic.List<WpsOption>();
                    var existingWpsIds = new System.Collections.Generic.HashSet<int>(vm.WpsOptions.Select(x => x.Id));
                    var missingWpsIds = dwrMap.Values
                        .Select(x => x.WpsId)
                        .Where(id => id.HasValue && id.Value > 0 && !existingWpsIds.Contains(id.Value))
                        .Select(id => id!.Value)
                        .Distinct()
                        .ToList();
                    if (missingWpsIds.Count > 0)
                    {
                        var extra = await _context.WPS_tbl.AsNoTracking()
                            .Where(w => missingWpsIds.Contains(w.WPS_ID))
                            .Select(w => new WpsOption
                            {
                                Id = w.WPS_ID,
                                Wps = w.WPS,
                                ThicknessRange = w.Thickness_Range,
                                Pwht = w.PWHT
                            })
                            .ToListAsync();
                        foreach (var opt in extra)
                        {
                            if (existingWpsIds.Add(opt.Id))
                            {
                                vm.WpsOptions.Add(opt);
                            }
                        }
                    }

                    foreach (var row in vm.Rows)
                    {
                        if (row == null) continue;
                        row.UpdatedBy = null;
                        row.UpdatedDate = null;
                        if (!dwrMap.TryGetValue(row.JointId, out var dwr))
                        {
                            try { _logger.LogDebug("DailyWelding: No DWR row for JointId {JointId}", row.JointId); } catch { }
                            continue;
                        }

                        row.POST_VISUAL_INSPECTION_QR_NO = dwr.Qr;
                        row.FitupReport = dwr.Qr;
                        row.RFI_ID_DWR = dwr.RfiId;
                        row.RfiId = dwr.RfiId;
                        if (dwr.RfiId.HasValue)
                        {
                            if (rfiValueLookup.TryGetValue(dwr.RfiId.Value, out var shortVal) && !string.IsNullOrWhiteSpace(shortVal))
                            {
                                // Prefer the short SubCon_RFI_No value (matching DailyFitup behavior)
                                row.RfiNo = shortVal;
                            }
                            else
                            {
                                // Fallback to numeric RFI id when no short value available
                                row.RfiNo = dwr.RfiId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            }
                        }

                        row.DATE_WELDED = dwr.DateWelded;
                        // Keep DFR_tbl.FITUP_DATE in FitupDate so rows stay editable when present; DATE_WELDED still available via DATE_WELDED property
                        row.ACTUAL_DATE_WELDED = dwr.ActualDate;
                        row.ActualDate = dwr.ActualDate?.ToString("yyyy-MM-dd'T'HH:mm");

                        row.WPS_ID_DWR = dwr.WpsId;
                        if (dwr.WpsId.HasValue && dwr.WpsId.Value > 0)
                        {
                            row.WpsId = dwr.WpsId.Value;
                            var wpsOpt = vm.WpsOptions.FirstOrDefault(x => x.Id == dwr.WpsId.Value);
                            if (wpsOpt != null && !string.IsNullOrWhiteSpace(wpsOpt.Wps))
                            {
                                row.Wps = wpsOpt.Wps;
                            }
                            row.WpsCandidates ??= new System.Collections.Generic.List<WpsOption>();
                            if (!row.WpsCandidates.Any(x => x.Id == dwr.WpsId.Value || (!string.IsNullOrWhiteSpace(row.Wps) && string.Equals(x.Wps, row.Wps, StringComparison.OrdinalIgnoreCase))))
                            {
                                row.WpsCandidates.Insert(0, new WpsOption { Id = dwr.WpsId.Value, Wps = row.Wps ?? string.Empty });
                            }
                        }

                        row.ROOT_A = dwr.RootA;
                        row.ROOT_B = dwr.RootB;
                        row.FILL_A = dwr.FillA;
                        row.FILL_B = dwr.FillB;
                        row.CAP_A = dwr.CapA;
                        row.CAP_B = dwr.CapB;
                        row.PREHEAT_TEMP_C = dwr.Preheat;
                        row.IPOrT = dwr.IPorT;
                        row.OpenClosed = dwr.OpenClosed;
                        row.Weld_Confirmed = dwr.WeldConfirmed.GetValueOrDefault();
                        row.Remarks = dwr.Remarks;

                        string? updatedByDisplay = null;
                        if (dwr.UpdatedBy.HasValue && dwr.UpdatedBy.Value > 0 && dwrUpdaterLookup.TryGetValue(dwr.UpdatedBy.Value, out var name) && !string.IsNullOrWhiteSpace(name))
                        {
                            updatedByDisplay = name;
                        }

                        row.UpdatedBy = updatedByDisplay;
                        row.UpdatedDate = dwr.UpdatedDate?.ToString("dd-Mmm-yyyy hh:mm tt", CultureInfo.InvariantCulture);
                    }

                    // NEW: populate per-row WPS candidates using _wpsSelectorService similar to DailyFitup.
                    // Limit processing to avoid heavy calls when many rows present.
                    try
                    {
                        if (vm.Rows != null && vm.Rows.Count > 0)
                        {
                            const int RowCandidateLimit = 500; // safeguard
                            var toProcess = vm.Rows.Take(RowCandidateLimit).ToList();
                            foreach (var row in toProcess)
                            {
                                try
                                {
                                    var existingCandidates = row.WpsCandidates?.ToList() ?? new System.Collections.Generic.List<WpsOption>();
                                    // If we already have a richer list (>1), keep it; otherwise augment single/saved entries so saved rows show the same list as new rows
                                    var hasRichCandidates = existingCandidates.Count > 1;
                                    if (hasRichCandidates) continue;

                                    // Resolve effective thickness using existing helper on the service
                                    var resolved = await _wpsSelectorService.ResolveThicknessAsync(
                                        projectId: weldersProjectId,
                                        explicitLineClass: row.LineClass,
                                        explicitThickness: row.PREHEAT_TEMP_C,
                                        sch: row.Sch,
                                        dia: row.DiaIn,
                                        olSch: row.OlSch,
                                        olDia: row.OlDia,
                                        olThick: row.OlThick);

                                    var candidates = await _wpsSelectorService.GetCandidatesAsync(weldersProjectId, row.LineClass, resolved.thickness);

                                    if (candidates != null && candidates.Count > 0)
                                    {
                                        var merged = candidates.Select(c => new WpsOption { Id = c.Id, Wps = c.Wps, ThicknessRange = c.ThicknessRange, Pwht = c.Pwht }).ToList();

                                        // Merge any existing single saved candidate (from DWR) to keep it visible
                                        foreach (var existing in existingCandidates)
                                        {
                                            if (existing == null) continue;
                                            if (merged.Any(x => (x.Id > 0 && existing.Id > 0 && x.Id == existing.Id) || (!string.IsNullOrWhiteSpace(x.Wps) && string.Equals(x.Wps, existing.Wps, StringComparison.OrdinalIgnoreCase))))
                                            {
                                                continue;
                                            }
                                            merged.Insert(0, new WpsOption { Id = existing.Id, Wps = existing.Wps, ThicknessRange = existing.ThicknessRange, Pwht = existing.Pwht });
                                        }

                                        row.WpsCandidates = merged;

                                        // Ensure any saved WPS is represented and placed first
                                        if (!string.IsNullOrWhiteSpace(row.Wps))
                                        {
                                            var existingIndex = row.WpsCandidates.FindIndex(x => string.Equals(x.Wps, row.Wps, StringComparison.OrdinalIgnoreCase));
                                            if (existingIndex > 0)
                                            {
                                                var existing = row.WpsCandidates[existingIndex];
                                                row.WpsCandidates.RemoveAt(existingIndex);
                                                row.WpsCandidates.Insert(0, existing);
                                            }
                                            else if (existingIndex == -1)
                                            {
                                                // Insert a candidate for the saved text value
                                                row.WpsCandidates.Insert(0, new WpsOption { Id = row.WpsId ?? 0, Wps = row.Wps });
                                            }
                                        }
                                    }
                                }
                                catch (Exception exRow)
                                {
                                    try { _logger.LogDebug(exRow, "Get WPS candidates failed for Joint {JointId}", row.JointId); } catch { }
                                }
                            }

                            // NOTE: Do NOT merge per-row candidate lists into the header vm.WpsOptions here.
                            // DailyFitup keeps header WpsOptions derived from project WPS list plus any explicit WPS IDs
                            // found on rows. Keeping the header list unchanged avoids adding many similar or irrelevant
                            // candidates and preserves the DailyFitup behaviour.
                        }
                    }
                    catch (Exception exCandidates)
                    {
                        try { _logger.LogWarning(exCandidates, "Populate WPS candidates failed"); } catch { }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "DailyWelding DWR overlay failed");
        }

        // SANITY CHECK: ensure header WpsOptions do not contain per-row free-text candidates
        try
        {
            vm.WpsOptions = SanitizeWpsOptions(vm.WpsOptions);
        }
        catch (Exception exSan)
        {
            try { _logger.LogWarning(exSan, "DailyWelding: WpsOptions sanitization failed"); } catch { }
        }

        var isLoadRequested = string.Equals(doLoad, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(doLoad, "true", StringComparison.OrdinalIgnoreCase);
        if (isLoadRequested)
        {
            return View("DwrForm", vm);
        }

        return View("DailyWelding", vm);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetWeldingRfiOptions(string location, string? fitupDateIso, int? projectId)
    {
        try
        {
            string loc = (location ?? string.Empty).Trim().ToUpperInvariant();
            if (loc == "SHOP") loc = "WS";
            if (loc == "FIELD") loc = "FW";
            if (loc.StartsWith("WS")) loc = "WS";
            else if (loc.StartsWith("FW")) loc = "FW";
            else if (loc.StartsWith("TH") || loc.Contains("THREAD")) loc = "FW";

            var weldersProjectId = projectId.HasValue ? await ResolveWeldersProjectIdAsync(projectId.Value) : 0;

            // Note: we intentionally ignore any date filter here and always return the full list
            // (subject to project and location filtering). This keeps the RFI No. dropdown unfiltered by date.

            IQueryable<Rfi> baseQuery = _context.RFI_tbl.AsNoTracking()
                .Where(r => r.SubDiscipline != null && EF.Functions.Like(r.SubDiscipline!, "%Welding%"))
                .Where(r => r.ACTIVITY == 3.9 || r.ACTIVITY == 3.90 || r.ACTIVITY == 3.900);

            if (weldersProjectId > 0)
            {
                baseQuery = baseQuery.Where(r => r.RFI_Project_No == weldersProjectId);
            }

            if (loc == "WS")
            {
                baseQuery = baseQuery.Where(r => r.RFI_LOCATION != null && (
                    r.RFI_LOCATION.StartsWith("WS") ||
                    r.RFI_LOCATION.Contains("SHOP") ||
                    r.RFI_LOCATION.Contains("WORK")));
            }
            else if (loc == "FW")
            {
                baseQuery = baseQuery.Where(r => r.RFI_LOCATION != null && !(
                    r.RFI_LOCATION.StartsWith("WS") ||
                    r.RFI_LOCATION.Contains("SHOP") ||
                    r.RFI_LOCATION.Contains("WORK")));
            }

            // Always return the complete (location / project filtered) list sorted by date/time/id
            var working = baseQuery.OrderByDescending(r => r.Date)
                                   .ThenByDescending(r => r.Time)
                                   .ThenByDescending(r => r.RFI_ID);

            var listRaw = await working.Take(2000)
                .Select(r => new { r.RFI_ID, r.SubCon_RFI_No, r.Date, r.Time, r.RFI_LOCATION, r.RFI_DESCRIPTION })
                .ToListAsync();

            string PrefixFor(string? rfiLoc)
            {
                if (loc == "WS") return "WS | ";
                if (loc == "FW") return "FW | ";
                var raw = (rfiLoc ?? string.Empty).ToUpperInvariant();
                if (raw.Contains("SHOP") || raw.Contains("WORK") || raw.StartsWith("WS")) return "WS | ";
                if (raw.Contains("FIELD") || raw.StartsWith("FW")) return "FW | ";
                return string.Empty;
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

            // Ensure callers always receive a selectable blank option so the table dropdown can be cleared without relying on client-side injection
            list.Insert(0, new RfiOption { Id = 0, Value = string.Empty, Display = string.Empty });

            return Ok(list);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in GetWeldingRfiOptions loc={Loc} fitup={FitupDateIso} project={ProjectId}", location, fitupDateIso, projectId);
            return StatusCode(500, "Error");
        }
    }

    // Helper to sanitize header WpsOptions: keep only DB-backed Id>0 and dedupe by Wps text
    public static System.Collections.Generic.List<WpsOption> SanitizeWpsOptions(System.Collections.Generic.List<WpsOption>? options)
    {
        var keep = new System.Collections.Generic.List<WpsOption>();
        if (options == null || options.Count == 0) return keep;
        var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var opt in options)
        {
            if (opt == null) continue;
            if (opt.Id <= 0) continue; // drop free-text entries
            var key = (opt.Wps ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(key)) continue;
            if (seen.Add(key)) keep.Add(new WpsOption { Id = opt.Id, Wps = key, ThicknessRange = opt.ThicknessRange, Pwht = opt.Pwht });
        }
        return keep;
    }

    private async Task<decimal?> ComputeWeldingDiaAsync(int projectId, DateTime day, string? locCode)
    {
        var start = day.Date;
        var end = start.AddDays(1);

        var query = from dwr in _context.DWR_tbl.AsNoTracking()
                    join dfr in _context.DFR_tbl.AsNoTracking() on dwr.Joint_ID_DWR equals dfr.Joint_ID
                    where dfr.Project_No == projectId
                          && dwr.ACTUAL_DATE_WELDED.HasValue
                          && dwr.ACTUAL_DATE_WELDED.Value >= start
                          && dwr.ACTUAL_DATE_WELDED.Value < end
                          && dfr.DIAMETER.HasValue
                    select new { dwr, dfr };

        if (!string.IsNullOrWhiteSpace(locCode))
        {
            if (string.Equals(locCode, "WS", StringComparison.OrdinalIgnoreCase))
            {
                query = query.Where(x => x.dfr.LOCATION == "WS");
                query = query.Where(x => x.dfr.WELD_TYPE == null || (
                    !EF.Functions.Like(x.dfr.WELD_TYPE!, "SP%") &&
                    !EF.Functions.Like(x.dfr.WELD_TYPE!, "FJ%") &&
                    !EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%")));
            }
            else if (string.Equals(locCode, "FW", StringComparison.OrdinalIgnoreCase))
            {
                query = query.Where(x => x.dfr.LOCATION == null || x.dfr.LOCATION != "WS");
                query = query.Where(x => x.dfr.WELD_TYPE == null || (
                    !EF.Functions.Like(x.dfr.WELD_TYPE!, "SP%") &&
                    !EF.Functions.Like(x.dfr.WELD_TYPE!, "FJ%") &&
                    !EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%")));
            }
            else if (string.Equals(locCode, "TH", StringComparison.OrdinalIgnoreCase))
            {
                query = query.Where(x => x.dfr.WELD_TYPE != null && EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%"));
            }
        }

        return await query.SumAsync(x => (decimal?)x.dfr.DIAMETER);
    }

    private async Task<bool> HasWeldingForLocationAsync(int projectId, string locCode, DateTime start, DateTime end)
    {
        var query = from dwr in _context.DWR_tbl.AsNoTracking()
                    join dfr in _context.DFR_tbl.AsNoTracking() on dwr.Joint_ID_DWR equals dfr.Joint_ID
                    where dfr.Project_No == projectId
                          && dwr.ACTUAL_DATE_WELDED.HasValue
                          && dwr.ACTUAL_DATE_WELDED.Value >= start
                          && dwr.ACTUAL_DATE_WELDED.Value < end
                    select new { dwr, dfr };

        if (string.Equals(locCode, "WS", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(x => x.dfr.LOCATION == "WS");
            query = query.Where(x => x.dfr.WELD_TYPE == null || (
                !EF.Functions.Like(x.dfr.WELD_TYPE!, "SP%") &&
                !EF.Functions.Like(x.dfr.WELD_TYPE!, "FJ%") &&
                !EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%")));
        }
        else if (string.Equals(locCode, "FW", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(x => x.dfr.LOCATION == null || x.dfr.LOCATION != "WS");
            query = query.Where(x => x.dfr.WELD_TYPE == null || (
                !EF.Functions.Like(x.dfr.WELD_TYPE!, "SP%") &&
                !EF.Functions.Like(x.dfr.WELD_TYPE!, "FJ%") &&
                !EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%")));
        }
        else if (string.Equals(locCode, "TH", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(x => x.dfr.WELD_TYPE != null && EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%"));
        }

        return await query.Take(1).AnyAsync();
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetDailyWeldingExists(int projectId, string location, string fitupDateIso, string? actualDateIso = null)
    {
        try
        {
            var dateText = string.IsNullOrWhiteSpace(actualDateIso) ? fitupDateIso : actualDateIso;
            if (projectId <= 0 || string.IsNullOrWhiteSpace(dateText)) return BadRequest("Invalid");
            if (!DateTime.TryParse(dateText, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt)) return BadRequest("Invalid date");
            var day = dt.Date;
            var locCode = MapHeaderLocation(location) ?? "FW";
            var row = await _context.PMS_Updated_Confirmed_tbl
                .FirstOrDefaultAsync(x => x.U_C_Project_No == projectId && x.U_C_Location == locCode && x.Updated_Confirmed_Date.Date == day);
            bool updatedExists = row != null && row.Welding_Updated_Date.HasValue;
            bool confirmedExists = row != null && row.Welding_Confirmed_Date.HasValue;
            return Json(new { updatedExists, confirmedExists });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetDailyWeldingExists failed");
            return Json(new { updatedExists = false, confirmedExists = false });
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CompleteDailyWelding(int projectId, string location, string fitupDateIso, string? actionType, string? actualDateIso = null)
    {
        try
        {
            var dateText = string.IsNullOrWhiteSpace(actualDateIso) ? fitupDateIso : actualDateIso;
            if (projectId <= 0 || string.IsNullOrWhiteSpace(dateText)) return BadRequest("Invalid");
            if (!DateTime.TryParse(dateText, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt)) return BadRequest("Invalid date");
            var day = dt.Date;
            var headerRaw = (location ?? string.Empty).Trim();
            var userId = HttpContext.Session.GetInt32("UserID");
            var isConfirm = string.Equals(actionType, "confirmed", StringComparison.OrdinalIgnoreCase);
            var now = AppClock.Now;
            var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);

            if (string.Equals(headerRaw, "All", StringComparison.OrdinalIgnoreCase))
            {
                var start = day.Date;
                var end = start.AddDays(1);

                var baseQuery = from dwr in _context.DWR_tbl.AsNoTracking()
                                join dfr in _context.DFR_tbl.AsNoTracking() on dwr.Joint_ID_DWR equals dfr.Joint_ID
                                where dfr.Project_No == projectId
                                      && dwr.ACTUAL_DATE_WELDED.HasValue
                                      && dwr.ACTUAL_DATE_WELDED.Value >= start
                                      && dwr.ACTUAL_DATE_WELDED.Value < end
                                select new { dwr, dfr };

                var hasWs = await baseQuery
                    .Where(x => x.dfr.LOCATION == "WS")
                    .Where(x => x.dfr.WELD_TYPE == null || (
                        !EF.Functions.Like(x.dfr.WELD_TYPE!, "SP%") &&
                        !EF.Functions.Like(x.dfr.WELD_TYPE!, "FJ%") &&
                        !EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%")))
                    .AnyAsync();
                var hasFw = await baseQuery
                    .Where(x => x.dfr.LOCATION == null || x.dfr.LOCATION != "WS")
                    .Where(x => x.dfr.WELD_TYPE == null || (
                        !EF.Functions.Like(x.dfr.WELD_TYPE!, "SP%") &&
                        !EF.Functions.Like(x.dfr.WELD_TYPE!, "FJ%") &&
                        !EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%")))
                    .AnyAsync();
                var hasThType = await _context.PMS_Weld_Type_tbl.AsNoTracking()
                    .AnyAsync(w => w.W_Project_No == weldersProjectId && w.W_Weld_Type != null && EF.Functions.Like(w.W_Weld_Type!, "%TH%"));
                var hasThDay = await baseQuery
                    .Where(x => x.dfr.WELD_TYPE != null && EF.Functions.Like(x.dfr.WELD_TYPE!, "%TH%"))
                    .AnyAsync();

                var targets = new System.Collections.Generic.List<string>();
                if (hasWs) targets.Add("WS");
                if (hasFw) targets.Add("FW");
                if (hasThType && hasThDay) targets.Add("TH");

                foreach (var lc in targets)
                {
                    var row = await _context.PMS_Updated_Confirmed_tbl
                        .FirstOrDefaultAsync(x => x.U_C_Project_No == projectId && x.U_C_Location == lc && x.Updated_Confirmed_Date.Date == day);
                    if (row == null)
                    {
                        row = new UpdatedConfirmed { U_C_Project_No = projectId, U_C_Location = lc, Updated_Confirmed_Date = day };
                        _context.PMS_Updated_Confirmed_tbl.Add(row);
                    }

                    if (isConfirm)
                    {
                        row.Welding_Confirmed_Date = now;
                        row.Welding_Confirmed_By = userId;
                        row.Welding_Confirmed_Dia = await ComputeWeldingDiaAsync(projectId, row.Updated_Confirmed_Date.Date, lc);
                    }
                    else
                    {
                        row.Welding_Updated_Date = now;
                        row.Welding_Updated_By = userId;
                        row.Welding_Total_Dia = await ComputeWeldingDiaAsync(projectId, row.Updated_Confirmed_Date.Date, lc);
                    }
                }

                if (targets.Count > 0)
                {
                    await _context.SaveChangesAsync();
                }

                return Ok();
            }

            string singleLoc = (string.Equals(headerRaw, "Threaded", StringComparison.OrdinalIgnoreCase) || headerRaw.StartsWith("TH", StringComparison.OrdinalIgnoreCase))
                ? "TH"
                : (MapHeaderLocation(location) ?? "FW");

            var rowSingle = await _context.PMS_Updated_Confirmed_tbl
                .FirstOrDefaultAsync(x => x.U_C_Project_No == projectId && x.U_C_Location == singleLoc && x.Updated_Confirmed_Date.Date == day);

            var sStart = day.Date;
            var sEnd = sStart.AddDays(1);
            var hasData = await HasWeldingForLocationAsync(projectId, singleLoc, sStart, sEnd);
            if (!hasData)
            {
                return Ok();
            }

            if (rowSingle == null)
            {
                rowSingle = new UpdatedConfirmed { U_C_Project_No = projectId, U_C_Location = singleLoc, Updated_Confirmed_Date = day };
                _context.PMS_Updated_Confirmed_tbl.Add(rowSingle);
            }

            if (isConfirm)
            {
                rowSingle.Welding_Confirmed_Date = now;
                rowSingle.Welding_Confirmed_By = userId;
                rowSingle.Welding_Confirmed_Dia = await ComputeWeldingDiaAsync(projectId, rowSingle.Updated_Confirmed_Date.Date, singleLoc);
            }
            else
            {
                rowSingle.Welding_Updated_Date = now;
                rowSingle.Welding_Updated_By = userId;
                rowSingle.Welding_Total_Dia = await ComputeWeldingDiaAsync(projectId, rowSingle.Updated_Confirmed_Date.Date, singleLoc);
            }

            await _context.SaveChangesAsync();
            return Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CompleteDailyWelding failed");
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SendDailyWeldingEmail([FromForm] int projectId, [FromForm] string location, [FromForm] string fitupDateIso, [FromForm] string? actionLabel, [FromForm] string kind, [FromForm] string? actualDateIso = null)
    {
        try
        {
            var dateText = string.IsNullOrWhiteSpace(actualDateIso) ? fitupDateIso : actualDateIso;
            if (projectId <= 0 || string.IsNullOrWhiteSpace(dateText))
                return Json(new { success = false, message = "Missing inputs" });

            var isCompleted = string.Equals(kind, "completed", StringComparison.OrdinalIgnoreCase);
            var razorKey = isCompleted ? "Daily_Welding_Completed" : "Daily_Welding_Confirmed";

            if (!DateTime.TryParse(dateText, out var fitupDate))
            {
                if (!DateTime.TryParseExact(dateText, DailyWeldingDateFormats, null, DateTimeStyles.AssumeLocal, out fitupDate))
                {
                    fitupDate = DateTime.Now;
                }
            }
            var day = fitupDate.Date;

            if (string.Equals(actionLabel?.Trim(), "Update Daily Welding", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(actionLabel?.Trim(), "Updated Daily Welding", StringComparison.OrdinalIgnoreCase))
            {
                actionLabel = "Daily Welding has been Updated";
            }
            else if (string.Equals(actionLabel?.Trim(), "Update Confirmed Welding", StringComparison.OrdinalIgnoreCase))
            {
                actionLabel = "Confirmed Welding has been Updated";
            }

            static string MapHeaderLocation(string raw)
            {
                var upper = (raw ?? string.Empty).Trim().ToUpperInvariant();
                if (upper.StartsWith("WS") || upper.Contains("SHOP") || upper.Contains("WORK")) return "WS";
                if (upper.StartsWith("FW") || upper.Contains("FIELD")) return "FW";
                if (upper.StartsWith("TH") || upper.Contains("THREAD")) return "TH";
                if (upper == "ALL") return "All";
                return upper;
            }

            var locCodeFromHeader = MapHeaderLocation(location);
            bool headerAll = string.Equals(locCodeFromHeader, "All", StringComparison.OrdinalIgnoreCase);

            var ucRows = await _context.PMS_Updated_Confirmed_tbl.AsNoTracking()
                .Where(x => x.U_C_Project_No == projectId && x.Updated_Confirmed_Date.Date == day)
                .ToListAsync();

            UpdatedConfirmed? uc = null;
            if (!headerAll)
            {
                var targetLoc = (locCodeFromHeader == "WS" || locCodeFromHeader == "FW" || locCodeFromHeader == "TH") ? locCodeFromHeader : null;
                var qry = ucRows.AsQueryable();
                if (targetLoc != null) qry = qry.Where(r => r.U_C_Location == targetLoc).AsQueryable();
                uc = qry.OrderByDescending(r => r.Updated_Confirmed_Date).FirstOrDefault();
            }

            var receivers = await ResolveUCDailyReceiversAsync(projectId, razorKey, headerAll ? "All" : (locCodeFromHeader == "WS" || locCodeFromHeader == "FW" || locCodeFromHeader == "TH" ? locCodeFromHeader : "All"));
            if (receivers.To.Count == 0 && receivers.Cc.Count == 0)
            {
                return Json(new { success = false, message = "No receivers configured" });
            }

            string title = actionLabel ?? string.Empty;
            string FormatDec(decimal? v) => v.HasValue ? string.Format(CultureInfo.InvariantCulture, "{0:0.##}", v.Value) : "-";

            var projectEntity = await _context.Projects_tbl.AsNoTracking().FirstOrDefaultAsync(p => p.Project_ID == projectId);
            var projectDisplay = projectEntity != null ? ($"{projectEntity.Project_ID} - {projectEntity.Project_Name}") : projectId.ToString(CultureInfo.InvariantCulture);
            var refDay = day;

            var bodyBuilder = new System.Text.StringBuilder();
            bodyBuilder.Append("<div style='color:#176d8a;font-size:14px;line-height:1.5'>");
            bodyBuilder.Append("<p style='margin:0 0 12px'>This is to inform you that the following action has been recorded in PMS:</p>");
            bodyBuilder.Append("<table style='border-collapse:collapse;width:100%;max-width:600px'>");
            void Row(string k, string v, string? firstWidth = null, string? secondWidth = null) => bodyBuilder.Append($"<tr><td style='padding:6px 8px;border:1px solid #e3e9eb;background:#f9fbfc;{(string.IsNullOrEmpty(firstWidth) ? "width:40%;" : $"width:{firstWidth};")}'><strong>{System.Net.WebUtility.HtmlEncode(k)}</strong></td><td style='padding:6px 8px;border:1px solid #e3e9eb;{(string.IsNullOrEmpty(secondWidth) ? string.Empty : $"width:{secondWidth};")}'>{System.Net.WebUtility.HtmlEncode(v)}</td></tr>");
            Row("Project No.", projectDisplay, "18%", "82%");
            Row("Daily Welding Date", refDay.ToString("dd-MMM-yyyy"));

            if (headerAll)
            {
                var ucWs = ucRows.FirstOrDefault(r => r.U_C_Location == "WS");
                var ucFw = ucRows.FirstOrDefault(r => r.U_C_Location == "FW");
                var ucTh = ucRows.FirstOrDefault(r => r.U_C_Location == "TH");

                string Name(string code) => code switch { "WS" => "Workshop", "FW" => "Field", "TH" => "Threaded", _ => code };

                var locationNames = new System.Collections.Generic.List<string>();
                if (ucWs != null) locationNames.Add(Name("WS"));
                if (ucFw != null) locationNames.Add(Name("FW"));
                if (ucTh != null) locationNames.Add(Name("TH"));
                if (locationNames.Count > 0) Row("Location", string.Join(" / ", locationNames));

                var updatedValues = new System.Collections.Generic.List<string>();
                if (ucWs != null) updatedValues.Add(FormatDec(ucWs.Welding_Total_Dia));
                if (ucFw != null) updatedValues.Add(FormatDec(ucFw.Welding_Total_Dia));
                if (ucTh != null) updatedValues.Add(FormatDec(ucTh.Welding_Total_Dia));
                if (updatedValues.Count > 0) Row("Welding Dia. In. (Updated)", string.Join(" / ", updatedValues));

                var confirmedValues = new System.Collections.Generic.List<string>();
                if (ucWs != null) confirmedValues.Add(FormatDec(ucWs.Welding_Confirmed_Dia));
                if (ucFw != null) confirmedValues.Add(FormatDec(ucFw.Welding_Confirmed_Dia));
                if (ucTh != null) confirmedValues.Add(FormatDec(ucTh.Welding_Confirmed_Dia));
                if (confirmedValues.Count > 0) Row("Welding Dia. In. (Confirmed)", string.Join(" / ", confirmedValues));
            }
            else
            {
                string Name(string code) => code switch { "WS" => "Workshop", "FW" => "Field", "TH" => "Threaded", _ => code };
                var locCode = uc?.U_C_Location ?? locCodeFromHeader;
                Row("Location", Name(locCode));
                if (uc != null)
                {
                    Row("Welding Dia. In. (Updated)", FormatDec(uc.Welding_Total_Dia));
                    Row("Welding Dia. In. (Confirmed)", FormatDec(uc.Welding_Confirmed_Dia));
                }
                else
                {
                    Row("Welding Dia. In. (Updated)", "-");
                    Row("Welding Dia. In. (Confirmed)", "-");
                }
            }

            bodyBuilder.Append("</table>");
            bodyBuilder.Append("<p style='margin:12px 0 0'>Regards,<br/>PMS System</p>");
            bodyBuilder.Append("</div>");

            var (emailHtml, resources) = await BuildBrandEmailShellAsync(title, bodyBuilder.ToString());
            await _emailService.SendEmailAsync(receivers.To, receivers.Cc, title, emailHtml, highImportance: true, inlineResources: resources);

            return Json(new { success = true });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SendDailyWeldingEmail failed for Project={ProjectId}", projectId);
            return Json(new { success = false, message = "Error" });
        }
    }

    // NEW: Welding completion confirmation
    [SessionAuthorization]
    [HttpPost]
    [Route("/Input/Welding/ConfirmCompletion")]
    public async Task<IActionResult> ConfirmWeldingCompletion([FromBody] WeldingCompletionRequest request)
    {
        if (request == null || request.Completions == null || request.Completions.Count == 0)
        {
            return BadRequest("Invalid request data");
        }

        try
        {
            foreach (var completion in request.Completions)
            {
                if (completion.JointId <= 0 || string.IsNullOrWhiteSpace(completion.QrCode))
                {
                    continue;
                }

                // Update DWR with completion QR code
                var dwr = await _context.DWR_tbl.FindAsync(completion.JointId);
                if (dwr != null)
                {
                    dwr.POST_VISUAL_INSPECTION_QR_NO = completion.QrCode;
                    dwr.DWR_Updated_Date = DateTime.UtcNow;
                    // dwr.DWR_Updated_By (unchanged)
                }
            }

            await _context.SaveChangesAsync();

            // NEW: Send email notifications for completions
            var emailSuccess = await SendWeldingCompletionEmails(request.Completions);
            if (!emailSuccess)
            {
                _logger.LogWarning("Failed to send email notifications for welding completions");
            }

            return Ok(new { success = true });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error confirming welding completion");
            return StatusCode(500, "Error confirming welding completion");
        }
    }

    private async Task<bool> SendWeldingCompletionEmails(System.Collections.Generic.List<WeldingCompletionItem> completions)
    {
        try
        {
            // TODO: Implement actual email sending logic
            // This is a placeholder for the email sending code
            foreach (var completion in completions)
            {
                // Example: send email for each completion
                // await _email_service.SendWeldingCompletionEmail(completion);
            }
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error sending welding completion emails");
            return false;
        }
    }

    public class WeldingCompletionRequest
    {
        public System.Collections.Generic.List<WeldingCompletionItem> Completions { get; set; } = new System.Collections.Generic.List<WeldingCompletionItem>();
    }

    public class WeldingCompletionItem
    {
        public int JointId { get; set; }
        public string QrCode { get; set; } = string.Empty;
    }
}