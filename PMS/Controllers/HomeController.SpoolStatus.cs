using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System.Text.Json;
using System.IO;

namespace PMS.Controllers;

public partial class HomeController
{
    private const string PreferredProjectSessionKey = "PreferredProjectId";

    // Reuse these constant arrays to avoid allocating new arrays repeatedly (CA1861)
    private static readonly string[] FabGroupKeys = { "A. PRODUCTION", "B. DISPATCH", "C. PAINTING", "D. QC" };
    private static readonly string[] ShopMinorGroups = { "I. FABRICATION NOT STARTED", "II. HOLD" };
    private static readonly string[] GroupOrderSite = { "I. Erection" };
    private static readonly string[] GroupOrderMain = { "I. FABRICATION NOT STARTED", "II. HOLD", "III. FABRICATION", "Z. UNGROUPED" };

    private static readonly Dictionary<string, string[]> StatusGroupMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["I. Erection"] = new[] { "a. DISPATCHED TO SITE", "b. READY FOR MRI", "c. READY FOR SITE ISSUANCE", "d. ISSUED FOR INSTALLATION", "e. COMPLETED" },
        ["I. FABRICATION NOT STARTED"] = new[] { "a. RANDOM", "b. WELDING" },
        ["II. HOLD"] = new[] { "a. HOLD BY TR", "b. HOLD BY PLANING" },
        ["A. PRODUCTION"] = new[] { "01. PARTIALLY WELDED", "02. SUPPORT BALANCE", "03. REPAIR BALANCE", "04. UNDER ALLOCATION", "05. REJECT FINAL DIMENSION" },
        ["B. DISPATCH"] = new[] { "01. READY FOR DISPATCH", "02. RELEASED FOR DISPATCH" },
        ["C. PAINTING"] = new[] { "01. RELEASED FOR PAINTING" },
        ["D. QC"] = new[]
    {
 "01. RT OF REPAIR BALANCE","02. RT SELECTION BALANCE","03. RT BALANCE","04. RT TRACER","05. WAITING OID RESULT",
 "06. NDE BALANCE","07. PWHT BALANCE","08. PMI BALANCE","09. HT SELECTION BALANCE", "10. HT BALANCE",
 "11. UNDER HYDROTEST ISSUANCE","12. UNDER RTC ISSUANCE", "13. RELEASED FOR PAINTING",
 "13-1. UNDER QC","13-2. UNDER TR APPROVAL","13-3. UNDER SAPID"
 }
    };

    private static readonly HashSet<string> QcFlowUp = new(StringComparer.OrdinalIgnoreCase) { "13-1. UNDER QC", "13-2. UNDER TR APPROVAL", "13-3. UNDER SAPID" };

    private static readonly Dictionary<string, string> ActionByMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["a. RANDOM"] = "SHOP ISSUANCE",
        ["b. WELDING"] = "SHOP ISSUANCE",
        ["01. PARTIALLY WELDED"] = "PRODUCTION",
        ["02. SUPPORT BALANCE"] = "PRODUCTION",
        ["03. REPAIR BALANCE"] = "PRODUCTION",
        ["04. UNDER ALLOCATION"] = "PRODUCTION",
        ["05. REJECT FINAL DIMENSION"] = "PRODUCTION",
        ["01. RT OF REPAIR BALANCE"] = "QC",
        ["02. RT SELECTION BALANCE"] = "QC",
        ["03. RT BALANCE"] = "QC",
        ["04. RT TRACER"] = "QC",
        ["05. WAITING OID RESULT"] = "QC",
        ["06. NDE BALANCE"] = "QC",
        ["07. PWHT BALANCE"] = "QC",
        ["08. PMI BALANCE"] = "QC",
        ["09. HT SELECTION BALANCE"] = "QC",
        ["10. HT BALANCE"] = "QC",
        ["11. UNDER HYDROTEST ISSUANCE"] = "QC",
        ["12. UNDER RTC ISSUANCE"] = "QC",
        ["13. RELEASED FOR PAINTING"] = "QC",
        ["13-1. UNDER QC"] = "QC",
        ["13-2. UNDER TR APPROVAL"] = "QC",
        ["13-3. UNDER SAPID"] = "QC",
        ["01. RELEASED FOR PAINTING"] = "PAINTING",
        ["01. READY FOR DISPATCH"] = "DISPATCH",
        ["02. RELEASED FOR DISPATCH"] = "DISPATCH",
        ["a. DISPATCHED TO SITE"] = "SITE",
        ["b. READY FOR MRI"] = "SITE",
        ["c. READY FOR SITE ISSUANCE"] = "SITE",
        ["d. ISSUED FOR INSTALLATION"] = "SITE",
        ["e. COMPLETED"] = "SITE",
        ["a. HOLD BY TR"] = "HOLD",
        ["b. HOLD BY PLANING"] = "HOLD",
    };

    private int? GetPersistedProjectId()
    {
        var sessionValue = HttpContext.Session.GetInt32(PreferredProjectSessionKey);
        if (sessionValue.HasValue && sessionValue.Value > 0) return sessionValue.Value;

        var userId = HttpContext.Session.GetInt32("UserID");
        if (userId.HasValue && Request.Cookies.TryGetValue($"PMS_DefaultProject_{userId.Value}", out var raw)
            && int.TryParse(raw, out var cookieValue) && cookieValue > 0)
        {
            HttpContext.Session.SetInt32(PreferredProjectSessionKey, cookieValue);
            return cookieValue;
        }

        return null;
    }

    private void PersistProjectSelection(int projectId)
    {
        if (projectId <= 0) return;

        HttpContext.Session.SetInt32(PreferredProjectSessionKey, projectId);
        var userId = HttpContext.Session.GetInt32("UserID");
        if (userId.HasValue)
        {
            Response.Cookies.Append($"PMS_DefaultProject_{userId.Value}", projectId.ToString(), SecureCookieOptions(DateTimeOffset.UtcNow.AddDays(60)));
        }
    }

    private async Task<bool> ProjectExistsAsync(int projectId)
    {
        return await _context.Projects_tbl.AsNoTracking().AnyAsync(p => p.Project_ID == projectId);
    }

    private async Task<int?> GetDefaultProjectIdAsync(int? selectedProjectId = null)
    {
        try
        {
            int? candidate = selectedProjectId;

            if (!candidate.HasValue)
            {
                string[] queryKeys = ["projectId", "project", "project_no", "projectno", "project_welder", "projectwelder"]; // normalized keys
                foreach (var key in queryKeys)
                {
                    var raw = HttpContext?.Request?.Query[key].FirstOrDefault();
                    if (int.TryParse(raw, out var parsed) && parsed > 0)
                    {
                        candidate = parsed;
                        break;
                    }
                }

                if (!candidate.HasValue && HttpContext?.Request?.HasFormContentType == true)
                {
                    foreach (var key in queryKeys)
                    {
                        var raw = HttpContext.Request.Form[key].FirstOrDefault();
                        if (int.TryParse(raw, out var parsed) && parsed > 0)
                        {
                            candidate = parsed;
                            break;
                        }
                    }
                }
            }

            if (candidate.HasValue && await ProjectExistsAsync(candidate.Value))
            {
                PersistProjectSelection(candidate.Value);
                return candidate.Value;
            }

            var persisted = GetPersistedProjectId();
            if (persisted.HasValue && await ProjectExistsAsync(persisted.Value))
            {
                return persisted.Value;
            }

            var defaultProject = await _context.Projects_tbl.AsNoTracking()
                .Where(p => p.Default_P)
                .OrderByDescending(p => p.Project_ID)
                .Select(p => (int?)p.Project_ID)
                .FirstOrDefaultAsync();
            if (defaultProject.HasValue)
            {
                PersistProjectSelection(defaultProject.Value);
                return defaultProject.Value;
            }

            var maxProj = await _context.Projects_tbl.AsNoTracking()
                .OrderByDescending(p => p.Project_ID)
                .Select(p => (int?)p.Project_ID)
                .FirstOrDefaultAsync();
            if (maxProj.HasValue)
            {
                PersistProjectSelection(maxProj.Value);
            }
            return maxProj;
        }
        catch
        {
            return selectedProjectId;
        }
    }

    private static bool NameContainsToken(string? name, string token)
    => !string.IsNullOrWhiteSpace(name) && name.Contains(token, StringComparison.OrdinalIgnoreCase);

    private static string? FindStatusColumn(DataTable? dt)
    {
        if (dt == null) return null;
        foreach (DataColumn c in dt.Columns)
        {
            var name = (c.ColumnName ?? string.Empty).Trim();
            if (NameContainsToken(name, "STATUS")) return c.ColumnName;
        }
        return null;
    }

    private static string? FindCountColumn(DataTable? dt)
    {
        if (dt == null) return null;
        foreach (DataColumn c in dt.Columns)
        {
            var name = (c.ColumnName ?? string.Empty).Trim();
            if (NameContainsToken(name, "COUNT")) return c.ColumnName;
        }
        return null;
    }

    private static string? FindProjectColumn(DataTable? dt)
    {
        if (dt == null) return null;
        foreach (DataColumn c in dt.Columns)
        {
            var name = (c.ColumnName ?? string.Empty).Trim();
            if (NameContainsToken(name, "PROJECT")) return c.ColumnName;
            if (string.Equals(name, "ProjectNo", StringComparison.OrdinalIgnoreCase)) return c.ColumnName;
        }
        return null;
    }

    private static DataTable FilterByProject(DataTable? dt, int? projectId)
    {
        if (dt == null) return new DataTable();
        if (!projectId.HasValue) return dt;
        var projCol = FindProjectColumn(dt);
        if (string.IsNullOrWhiteSpace(projCol)) return dt;

        string target = projectId.Value.ToString();
        var filteredRows = dt.AsEnumerable()
            .Where(r =>
            {
                var raw = Convert.ToString(r[projCol])?.Trim();
                if (string.IsNullOrWhiteSpace(raw)) return false;
                if (int.TryParse(raw, out var val)) return val == projectId.Value;
                // handle numeric values coming as floating strings (e.g., "2453.0000")
                if (double.TryParse(raw, out var dv)) return Math.Abs(dv - projectId.Value) < 0.0001;
                return string.Equals(raw, target, StringComparison.OrdinalIgnoreCase);
            })
            .ToList();
        if (filteredRows.Count == 0 || filteredRows.Count == dt.Rows.Count) return dt;

        var filtered = dt.Clone();
        foreach (var row in filteredRows) filtered.ImportRow(row);
        return filtered;
    }

    private async Task<DataTable> LoadSpoolReleaseLogAsync(int? projectId)
    {
        var dt = new DataTable();
        var pid = await GetDefaultProjectIdAsync(projectId);
        var conn = _context.Database.GetDbConnection();
        if (conn.State != System.Data.ConnectionState.Open) await conn.OpenAsync();

        async Task TryExecAsync(CommandType type, string sql, params (string name, object? val)[] parameters)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.CommandType = type;
            cmd.CommandTimeout = 600;
            foreach (var (name, val) in parameters)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = name;
                p.Value = val ?? DBNull.Value;
                cmd.Parameters.Add(p);
            }
            try
            {
                using var reader = await cmd.ExecuteReaderAsync();
                if (reader != null)
                {
                    DataTable? best = null;
                    int bestRows = -1;
                    do
                    {
                        var tmp = new DataTable();
                        tmp.Load(reader);
                        if (tmp.Columns.Cast<DataColumn>().Any(c => NameContainsToken(c.ColumnName, "STATUS")))
                        {
                            dt = tmp; best = tmp; break;
                        }
                        if (tmp.Rows.Count > bestRows)
                        {
                            best = tmp; bestRows = tmp.Rows.Count;
                        }
                    }
                    while (await reader.NextResultAsync());
                    if (best != null && dt.Rows.Count == 0) dt = best;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "PMS_Spool_Release_Log_Q execution failed for {Type} {Sql}", type, sql);
            }
        }

        await TryExecAsync(CommandType.StoredProcedure, "[dbo].[PMS_Spool_Release_Log_Q]");

        dt = FilterByProject(dt, pid);

        return dt;
    }

    private sealed class SummaryItem
    {
        public string Name { get; set; } = string.Empty;
        public long Count { get; set; }
        public double Percent { get; set; }
        public List<SummaryItem>? Children { get; set; }
    }

    // comparer for pair keys used in shop crosstab
    private sealed class PairComparer : IEqualityComparer<(string grp, string status)>
    {
        private static readonly StringComparer S = StringComparer.OrdinalIgnoreCase;
        public bool Equals((string grp, string status) x, (string grp, string status) y) => S.Equals(x.grp ?? string.Empty, y.grp ?? string.Empty) && S.Equals(x.status ?? string.Empty, y.status ?? string.Empty);
        public int GetHashCode((string grp, string status) obj) { unchecked { int h1 = S.GetHashCode(obj.grp ?? string.Empty); int h2 = S.GetHashCode(obj.status ?? string.Empty); return (h1 * 397) ^ h2; } }
    }

    // comparer for tuple keys used in some methods
    private sealed class TupleKeyComparer : IEqualityComparer<(string area, string group, string status)>
    {
        private static readonly StringComparer S = StringComparer.OrdinalIgnoreCase;
        public bool Equals((string area, string group, string status) x, (string area, string group, string status) y)
        => S.Equals(x.area ?? string.Empty, y.area ?? string.Empty)
        && S.Equals(x.group ?? string.Empty, y.group ?? string.Empty)
        && S.Equals(x.status ?? string.Empty, y.status ?? string.Empty);
        public int GetHashCode((string area, string group, string status) obj)
        {
            int h1 = S.GetHashCode(obj.area ?? string.Empty);
            int h2 = S.GetHashCode(obj.group ?? string.Empty);
            int h3 = S.GetHashCode(obj.status ?? string.Empty);
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + h1;
                hash = hash * 31 + h2;
                hash = hash * 31 + h3;
                return hash;
            }
        }
    }

    private static string NormalizeStatus(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        var s = raw.Trim();
        if (s.Length >= 2 && s[1] == '-') s = s[2..].Trim();
        return s;
    }

    private static (List<SummaryItem> site, List<SummaryItem> shop, long totalSite, long totalShop, long totalOverall) BuildSummary(DataTable dt)
    {
        string? statusCol = FindStatusColumn(dt);
        string? countCol = FindCountColumn(dt);
        if (string.IsNullOrWhiteSpace(statusCol)) return (new List<SummaryItem>(), new List<SummaryItem>(), 0L, 0L, 0L);

        var statusCounts = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (DataRow row in dt.Rows)
        {
            var raw = row[statusCol]?.ToString();
            var s = NormalizeStatus(raw);
            if (string.IsNullOrWhiteSpace(s)) continue;
            long add = 1;
            if (!string.IsNullOrWhiteSpace(countCol))
            {
                var v = row[countCol];
                if (v != null && v != DBNull.Value)
                {
                    try { add = Convert.ToInt64(v); }
                    catch { if (long.TryParse(v.ToString(), out var parsed)) add = parsed; }
                }
            }
            statusCounts.TryGetValue(s, out var n);
            statusCounts[s] = n + add;
        }

        var groupedStatuses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var arr in StatusGroupMap.Values) foreach (var s in arr) groupedStatuses.Add(s);

        var groupTotals = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var k in StatusGroupMap.Keys) groupTotals[k] = 0L;
        groupTotals["Z. UNGROUPED"] = 0L;

        foreach (var kv in statusCounts)
        {
            string group = "Z. UNGROUPED";
            foreach (var g in StatusGroupMap.Keys)
            {
                var arr = StatusGroupMap[g];
                if (arr.Any(x => string.Equals(x, kv.Key, StringComparison.OrdinalIgnoreCase))) { group = g; break; }
            }
            groupTotals[group] = groupTotals.TryGetValue(group, out var vv) ? vv + kv.Value : kv.Value;
        }

        // Use TryGetValue to avoid double dictionary lookup (CA1854)
        groupTotals.TryGetValue("I. Erection", out var totalSite);
        groupTotals.TryGetValue("I. FABRICATION NOT STARTED", out var _fabStart);
        groupTotals.TryGetValue("II. HOLD", out var _hold);
        groupTotals.TryGetValue("A. PRODUCTION", out var _aprod);
        groupTotals.TryGetValue("B. DISPATCH", out var _bdisp);
        groupTotals.TryGetValue("C. PAINTING", out var _cpaint);
        groupTotals.TryGetValue("D. QC", out var _dqc);
        var totalShop = _fabStart + _hold + _aprod + _bdisp + _cpaint + _dqc;
        var totalOverall = totalSite + totalShop;
        double den = totalOverall == 0 ? 1.0 : (double)totalOverall;

        var site = new List<SummaryItem>();
        if (groupTotals.TryGetValue("I. Erection", out var gE) && gE > 0)
        {
            var gname = "I. Erection";
            var items = new List<SummaryItem>();
            foreach (var s in StatusGroupMap[gname]) if (statusCounts.TryGetValue(s, out var c)) items.Add(new SummaryItem { Name = s, Count = c, Percent = ((double)c) / den });
            site.Add(new SummaryItem { Name = gname, Count = gE, Percent = ((double)gE) / den, Children = items });
        }

        var shop = new List<SummaryItem>();
        if (groupTotals.TryGetValue("I. FABRICATION NOT STARTED", out var gFabStart) && gFabStart > 0)
        {
            var items = StatusGroupMap["I. FABRICATION NOT STARTED"].Where(statusCounts.ContainsKey).Select(s => new SummaryItem { Name = s, Count = statusCounts[s], Percent = ((double)statusCounts[s]) / den }).ToList();
            shop.Add(new SummaryItem { Name = "I. FABRICATION NOT STARTED", Count = gFabStart, Percent = ((double)gFabStart) / den, Children = items });
        }
        if (groupTotals.TryGetValue("II. HOLD", out var gHold) && gHold > 0)
        {
            var items = StatusGroupMap["II. HOLD"].Where(statusCounts.ContainsKey).Select(s => new SummaryItem { Name = s, Count = statusCounts[s], Percent = ((double)statusCounts[s]) / den }).ToList();
            shop.Add(new SummaryItem { Name = "II. HOLD", Count = gHold, Percent = ((double)gHold) / den, Children = items });
        }

        long fabTotal = _aprod + _bdisp + _cpaint + _dqc;
        if (fabTotal > 0)
        {
            var fabChildren = new List<SummaryItem>();
            foreach (var sg in FabGroupKeys)
            {
                if (!groupTotals.TryGetValue(sg, out var sgTot) || sgTot <= 0) continue;
                var items = StatusGroupMap[sg].Where(statusCounts.ContainsKey).Select(s => new SummaryItem { Name = s, Count = statusCounts[s], Percent = ((double)statusCounts[s]) / den }).ToList();
                fabChildren.Add(new SummaryItem { Name = sg, Count = sgTot, Percent = ((double)sgTot) / den, Children = items });
            }
            shop.Add(new SummaryItem { Name = "III. FABRICATION", Count = fabTotal, Percent = ((double)fabTotal) / den, Children = fabChildren });
        }

        var ungrouped = statusCounts.Keys.Where(s => !groupedStatuses.Contains(s)).ToList();
        if (ungrouped.Count > 0)
        {
            var items = ungrouped.Select(s => new SummaryItem { Name = s, Count = statusCounts[s], Percent = ((double)statusCounts[s]) / den }).ToList();
            shop.Add(new SummaryItem { Name = "Z. UNGROUPED", Count = items.Sum(x => x.Count), Percent = ((double)items.Sum(x => x.Count)) / den, Children = items });
        }

        return (site, shop, totalSite, totalShop, totalOverall);
    }

    private static object MapSummaryItem(SummaryItem i) => new { name = i.Name, count = i.Count, percent = i.Percent, children = i.Children?.Select(MapSummaryItem).ToList() };
    private static List<object> MapSummaryList(List<SummaryItem> list) => list.Select(MapSummaryItem).ToList();

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetSpoolStatusSummary([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolReleaseLogAsync(projectId);
            dt = FilterByProject(dt, await GetDefaultProjectIdAsync(projectId));
            var (site, shop, totalSite, totalShop, totalOverall) = BuildSummary(dt);
            var payload = new { site = MapSummaryList(site), shop = MapSummaryList(shop), totals = new { totalSite, totalShop, totalOverall } };
            return Json(payload);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetSpoolStatusSummary failed");
            return Json(new { site = Array.Empty<object>(), shop = Array.Empty<object>(), totals = new { totalSite = 0L, totalShop = 0L, totalOverall = 0L } });
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportSpoolReleaseLog([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolReleaseLogAsync(projectId);
            dt = FilterByProject(dt, await GetDefaultProjectIdAsync(projectId));
            string? statusCol = FindStatusColumn(dt); int statusIdx = -1;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var name = (dt.Columns[i].ColumnName ?? string.Empty).Trim();
                if (NameContainsToken(name, "STATUS")) { statusCol = dt.Columns[i].ColumnName; statusIdx = i; break; }
            }
            if (statusIdx >= 0 && !dt.Columns.Contains("ACTION BY"))
            {
                var actionByCol = dt.Columns.Add("ACTION BY", typeof(string));
                actionByCol.SetOrdinal(statusIdx + 1);
                foreach (DataRow row in dt.Rows)
                {
                    var raw = statusCol == null ? string.Empty : (row[statusCol]?.ToString() ?? string.Empty);
                    var s = NormalizeStatus(raw);
                    row["ACTION BY"] = ActionByMap.TryGetValue(s, out var act) ? act : "OTHER";
                }
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Spool Release Log");
            for (int c = 0; c < dt.Columns.Count; c++) ws.Cell(1, c + 1).Value = dt.Columns[c].ColumnName;
            int r = 2;
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dr[c]; var cell = ws.Cell(r, c + 1);
                    if (val == null || val == DBNull.Value) { cell.SetValue(string.Empty); cell.Style.NumberFormat.Format = "@"; }
                    else if (val is DateTime dtv) { cell.SetValue(dtv); cell.Style.DateFormat.Format = "dd-mmm-yyyy"; }
                    else { cell.SetValue(val.ToString() ?? string.Empty); cell.Style.NumberFormat.Format = "@"; }
                }
                ws.Row(r).Height = 17; r++;
            }
            var table = ws.Range(1, 1, dt.Rows.Count + 1, dt.Columns.Count).CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2; ws.Row(1).Height = 30; ws.Columns().AdjustToContents(); ws.SheetView.FreezeRows(1);
            using var ms = new MemoryStream(); wb.SaveAs(ms); var bytes = ms.ToArray(); var fileName = $"SpoolReleaseLog_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportSpoolReleaseLog failed");
            return Content("Failed to export Spool Release Log.");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportSpoolStatusSummary([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolReleaseLogAsync(projectId);
            dt = FilterByProject(dt, await GetDefaultProjectIdAsync(projectId));
            var (siteItems, shopItems, totalSite, totalShop, totalOverall) = BuildSummary(dt);

            using var wb = new XLWorkbook();

            // Fabrication log sheet
            string? statusCol = FindStatusColumn(dt); int statusIdx = -1;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var name = (dt.Columns[i].ColumnName ?? string.Empty).Trim();
                if (NameContainsToken(name, "STATUS")) { statusCol = dt.Columns[i].ColumnName; statusIdx = i; break; }
            }
            if (statusIdx >= 0 && !dt.Columns.Contains("ACTION BY"))
            {
                var actionByCol = dt.Columns.Add("ACTION BY", typeof(string));
                actionByCol.SetOrdinal(statusIdx + 1);
                foreach (DataRow dr in dt.Rows)
                {
                    var raw = statusCol == null ? string.Empty : (dr[statusCol]?.ToString() ?? string.Empty);
                    var s = NormalizeStatus(raw);
                    dr["ACTION BY"] = ActionByMap.TryGetValue(s, out var act) ? act : "OTHER";
                }
            }

            var wsLog = wb.Worksheets.Add("Spool Release Log");
            for (int c = 0; c < dt.Columns.Count; c++) wsLog.Cell(1, c + 1).Value = dt.Columns[c].ColumnName;
            int rr = 2;
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dr[c]; var cell = wsLog.Cell(rr, c + 1);
                    if (val == null || val == DBNull.Value) { cell.SetValue(string.Empty); cell.Style.NumberFormat.Format = "@"; }
                    else if (val is DateTime dtv) { cell.SetValue(dtv); cell.Style.DateFormat.Format = "dd-mmm-yyyy"; }
                    else { cell.SetValue(val.ToString() ?? string.Empty); cell.Style.NumberFormat.Format = "@"; }
                }
                wsLog.Row(rr).Height = 17; rr++;
            }
            var tblLog = wsLog.Range(1, 1, dt.Rows.Count + 1, dt.Columns.Count).CreateTable(); tblLog.Theme = XLTableTheme.TableStyleMedium2; wsLog.Row(1).Height = 30; wsLog.Columns().AdjustToContents(); wsLog.SheetView.FreezeRows(1);

            // Visible summary sheet (static, no pivot objects)
            var wsSummary = wb.Worksheets.Add("Spool Status Summary");
            // Create two table blocks side-by-side: Site (cols1..4) and Shop (cols6..9)
            int siteStartRow = 1, siteStartCol = 1;
            int shopStartRow = 1, shopStartCol = 6;

            // Headers for both blocks
            wsSummary.Cell(siteStartRow, siteStartCol + 0).Value = "Area";
            wsSummary.Cell(siteStartRow, siteStartCol + 1).Value = "Group";
            wsSummary.Cell(siteStartRow, siteStartCol + 2).Value = "Status";
            wsSummary.Cell(siteStartRow, siteStartCol + 3).Value = "Count";

            wsSummary.Cell(shopStartRow, shopStartCol + 0).Value = "Area";
            wsSummary.Cell(shopStartRow, shopStartCol + 1).Value = "Group";
            wsSummary.Cell(shopStartRow, shopStartCol + 2).Value = "Status";
            wsSummary.Cell(shopStartRow, shopStartCol + 3).Value = "Count";

            // Build flattened entries from siteItems and shopItems
            var siteEntries = new List<(string group, string status, long count)>();
            foreach (var sg in siteItems)
            {
                if (sg.Children != null)
                {
                    foreach (var it in sg.Children)
                    {
                        siteEntries.Add((sg.Name, it.Name, it.Count));
                    }
                }
            }
            // Order site entries by GroupOrderSite then group name then status
            siteEntries = siteEntries.OrderBy(e => {
                var gi = Array.IndexOf(GroupOrderSite, e.group);
                return gi >= 0 ? gi : int.MaxValue;
            }).ThenBy(e => e.group, StringComparer.OrdinalIgnoreCase).ThenBy(e => e.status, StringComparer.OrdinalIgnoreCase).ToList();

            int rs = siteStartRow + 1;
            foreach (var e in siteEntries)
            {
                wsSummary.Cell(rs, siteStartCol + 0).Value = "Site";
                wsSummary.Cell(rs, siteStartCol + 1).Value = e.group;
                wsSummary.Cell(rs, siteStartCol + 2).Value = e.status;
                wsSummary.Cell(rs, siteStartCol + 3).Value = e.count;
                rs++;
            }
            var siteRange = wsSummary.Range(siteStartRow, siteStartCol, Math.Max(siteStartRow + 1, rs - 1), siteStartCol + 3);
            siteRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            siteRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Flatten shop entries, handling "III. FABRICATION" subgroups
            var shopEntries = new List<(string group, string status, long count)>();
            foreach (var s in shopItems)
            {
                if (string.Equals(s.Name, "III. FABRICATION", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var sub in s.Children ?? Enumerable.Empty<SummaryItem>())
                    {
                        foreach (var it in sub.Children ?? Enumerable.Empty<SummaryItem>())
                        {
                            shopEntries.Add((sub.Name, it.Name, it.Count));
                        }
                    }
                }
                else
                {
                    foreach (var it in s.Children ?? Enumerable.Empty<SummaryItem>())
                    {
                        shopEntries.Add((s.Name, it.Name, it.Count));
                    }
                }
            }
            // Order shop entries by GroupOrderMain then group then status
            shopEntries = shopEntries.OrderBy(e => {
                var gi = Array.IndexOf(GroupOrderMain, e.group);
                return gi >= 0 ? gi : int.MaxValue;
            }).ThenBy(e => e.group, StringComparer.OrdinalIgnoreCase).ThenBy(e => e.status, StringComparer.OrdinalIgnoreCase).ToList();

            int rsh = shopStartRow + 1;
            foreach (var e in shopEntries)
            {
                wsSummary.Cell(rsh, shopStartCol + 0).Value = "Shop";
                wsSummary.Cell(rsh, shopStartCol + 1).Value = e.group;
                wsSummary.Cell(rsh, shopStartCol + 2).Value = e.status;
                wsSummary.Cell(rsh, shopStartCol + 3).Value = e.count;
                rsh++;
            }
            var shopRange = wsSummary.Range(shopStartRow, shopStartCol, Math.Max(shopStartRow + 1, rsh - 1), shopStartCol + 3);
            shopRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            shopRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            wsSummary.Columns().AdjustToContents();
            wsSummary.SheetView.FreezeRows(1);

            // Hidden data sheet for pivot/chart sources; keep hidden to avoid prompts
            var wsData = wb.Worksheets.Add("SpoolStatus_Data"); wsData.Hide();
            wsData.Cell(1, 1).Value = "Area"; wsData.Cell(1, 2).Value = "Group"; wsData.Cell(1, 3).Value = "Status"; wsData.Cell(1, 4).Value = "Count";
            int dataRowPtr = 2;
            foreach (var sg in siteItems)
            {
                foreach (var it in sg.Children ?? Enumerable.Empty<SummaryItem>())
                {
                    wsData.Cell(dataRowPtr, 1).Value = "Site";
                    wsData.Cell(dataRowPtr, 2).Value = sg.Name;
                    wsData.Cell(dataRowPtr, 3).Value = it.Name;
                    wsData.Cell(dataRowPtr, 4).Value = it.Count;
                    dataRowPtr++;
                }
            }

            var shopPairs = new Dictionary<(string grp, string status), long>(new PairComparer());
            foreach (var s in shopItems)
            {
                if (string.Equals(s.Name, "III. FABRICATION", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var sub in s.Children ?? Enumerable.Empty<SummaryItem>())
                    {
                        foreach (var it in sub.Children ?? Enumerable.Empty<SummaryItem>())
                        {
                            wsData.Cell(dataRowPtr, 1).Value = "Shop";
                            wsData.Cell(dataRowPtr, 2).Value = sub.Name;
                            wsData.Cell(dataRowPtr, 3).Value = it.Name;
                            wsData.Cell(dataRowPtr, 4).Value = it.Count;
                            shopPairs[(sub.Name, it.Name)] = it.Count;
                            dataRowPtr++;
                        }
                    }
                }
                else
                {
                    foreach (var it in s.Children ?? Enumerable.Empty<SummaryItem>())
                    {
                        wsData.Cell(dataRowPtr, 1).Value = "Shop";
                        wsData.Cell(dataRowPtr, 2).Value = s.Name;
                        wsData.Cell(dataRowPtr, 3).Value = it.Name;
                        wsData.Cell(dataRowPtr, 4).Value = it.Count;
                        shopPairs[(s.Name, it.Name)] = it.Count;
                        dataRowPtr++;
                    }
                }
            }

            wsData.Columns().AdjustToContents();

            using var outMs = new MemoryStream(); wb.SaveAs(outMs); outMs.Position = 0;
            using (var doc = SpreadsheetDocument.Open(outMs, true))
            {
                var wbPart = doc.WorkbookPart;
                var sheets = wbPart?.Workbook?.Sheets?.Elements<Sheet>() ?? [];
                var dataSheet = sheets.FirstOrDefault(sh => string.Equals(sh.Name?.Value, wsData.Name, StringComparison.Ordinal));
                if (wbPart != null && dataSheet != null)
                {
                    // simplify fully-qualified EnumValue usage now that Spreadsheet is imported
                    dataSheet.State = new EnumValue<SheetStateValues>(SheetStateValues.VeryHidden);
                    wbPart.Workbook?.Save();
                }
            }

            var bytes = outMs.ToArray(); var fileName = $"Spool_Status_Summary_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportSpoolStatusSummary failed");
            return Content("Failed to export Spool Status Summary.");
        }
    }

    ////////////////////////////////////////////////////////////////////////////////
    // Code modifications start here -- spooling and erection log filtered exports
    ////////////////////////////////////////////////////////////////////////////////

    // Request DTO for filtered export
    public sealed record FilteredExportRequest(List<string>? Columns, List<List<string>>? Rows);

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public IActionResult ExportSpoolReleaseLogFiltered([FromBody] FilteredExportRequest? req)
    {
        try
        {
            if (req == null) return BadRequest("Invalid payload");
            var columns = req.Columns ?? new List<string>();
            var rows = req.Rows ?? new List<List<string>>();

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Spool Release Log");

            if (columns.Count == 0)
            {
                ws.Cell(1, 1).SetValue("No data");
            }
            else
            {
                // headers
                for (int c = 0; c < columns.Count; c++)
                    ws.Cell(1, c + 1).Value = columns[c] ?? string.Empty;

                // rows
                int r = 2;
                foreach (var row in rows)
                {
                    for (int c = 0; c < columns.Count; c++)
                    {
                        var val = (row != null && c < row.Count) ? (row[c] ?? string.Empty) : string.Empty;
                        var cell = ws.Cell(r, c + 1);
                        if (string.IsNullOrWhiteSpace(val))
                        {
                            cell.SetValue(string.Empty); cell.Style.NumberFormat.Format = "@";
                        }
                        else if (DateTime.TryParse(val, out var dtv))
                        {
                            cell.SetValue(dtv);
                            cell.Style.DateFormat.Format = "dd-mmm-yyyy";
                        }
                        else
                        {
                            cell.SetValue(val); cell.Style.NumberFormat.Format = "@";
                        }
                    }
                    ws.Row(r).Height = 17; r++;
                }

                int lastRow = Math.Max(1, rows.Count + 1);
                var fullRange = ws.Range(1, 1, lastRow, columns.Count);
                var table = fullRange.CreateTable();
                table.Theme = XLTableTheme.TableStyleMedium2;
                table.ShowTotalsRow = false;

                ws.Row(1).Height = 30;
                for (int c = 0; c < columns.Count; c++)
                {
                    var name = (columns[c] ?? string.Empty).ToLowerInvariant();
                    var col = ws.Column(c + 1);
                    if (name.Contains("id") || name.Contains("no") || name.Contains("spool") || name.Contains("project"))
                        col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    else
                        col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                }
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            // Build normalized counts by Area/Group/Status (use long)
            int statusColIndex = -1;
            for (int c = 0; c < columns.Count; c++) if (NameContainsToken(columns[c], "STATUS")) { statusColIndex = c; break; }

            var normCounts = new Dictionary<(string area, string group, string status), long>(new TupleKeyComparer());
            if (statusColIndex >= 0)
            {
                foreach (var row in rows)
                {
                    var raw = (row != null && statusColIndex < row.Count) ? (row[statusColIndex] ?? string.Empty) : string.Empty;
                    var s = NormalizeStatus(raw);
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    string group = "Z. UNGROUPED";
                    foreach (var kv in StatusGroupMap)
                    {
                        if (kv.Value.Any(v => string.Equals(v, s, StringComparison.OrdinalIgnoreCase))) { group = kv.Key; break; }
                    }

                    var area = GroupOrderSite.Contains(group, StringComparer.OrdinalIgnoreCase) ? "Site" : "Shop";
                    var key = (area, group, s);
                    normCounts.TryGetValue(key, out var cur);
                    normCounts[key] = cur + 1L;
                }
            }

            // Create summary sheet matching Dashboard grouping
            var wsSum = wb.Worksheets.Add("Spool Status Summary");

            long totalSite = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
            long totalShop = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
            long totalOverall = totalSite + totalShop;
            double den = totalOverall == 0 ? 1.0 : (double)totalOverall;

            int rowPtr = 1;
            // SITE block
            wsSum.Cell(rowPtr, 1).Value = "Site Spool Status"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;
            // Header row for site block
            wsSum.Cell(rowPtr, 1).Value = "Status"; wsSum.Cell(rowPtr, 2).Value = "Count"; wsSum.Cell(rowPtr, 3).Value = "%"; wsSum.Row(rowPtr).Style.Font.SetBold(true);
            // capture site header row index
            int siteHeaderRow = rowPtr;
            rowPtr++;

            var siteGroups = StatusGroupMap.Keys.Where(k => GroupOrderSite.Contains(k, StringComparer.OrdinalIgnoreCase)).ToList();
            if (siteGroups.Count == 0)
                siteGroups = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase)).Select(k => k.Key.group).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            bool wroteAnySite = false;
            foreach (var g in siteGroups)
            {
                var groupCount = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
                if (groupCount <= 0) continue;
                wsSum.Cell(rowPtr, 1).Value = g; wsSum.Cell(rowPtr, 2).Value = groupCount; wsSum.Cell(rowPtr, 3).Value = ((double)groupCount) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                rowPtr++;
                var statuses = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Select(k => (status: k.Key.status, count: k.Value)).OrderBy(s => s.status, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var st in statuses)
                {
                    wsSum.Cell(rowPtr, 1).Value = " " + st.status; wsSum.Cell(rowPtr, 2).Value = st.count; wsSum.Cell(rowPtr, 3).Value = ((double)st.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                    rowPtr++;
                }
                wroteAnySite = true;
            }
            if (!wroteAnySite)
            {
                wsSum.Cell(rowPtr, 1).Value = "No data"; wsSum.Cell(rowPtr, 2).Value = 0; wsSum.Cell(rowPtr, 3).Value = 0.0; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; rowPtr++;
            }
            // Write site total and then mark end of the site table range (include total row)
            wsSum.Cell(rowPtr, 1).Value = "Site Spool Status Total"; wsSum.Cell(rowPtr, 2).Value = totalSite; wsSum.Cell(rowPtr, 3).Value = totalOverall == 0 ? 0.0 : ((double)totalSite) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
            // site table ends at (rowPtr) which we'll include in the table; advance pointer after marking end
            rowPtr += 2;
            // compute site table bounds
            int siteTableEndRow = Math.Max(siteHeaderRow, rowPtr - 2);
            try
            {
                var siteRange = wsSum.Range(siteHeaderRow, 1, siteTableEndRow, 3);
                var siteTable = siteRange.CreateTable();
                siteTable.Theme = XLTableTheme.TableStyleMedium2;
                siteTable.ShowTotalsRow = false;
                // center Count and % columns in the site table
                siteRange.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                siteRange.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            catch { /* safe-fail */ }

            // SHOP block
            wsSum.Cell(rowPtr, 1).Value = "Shop Spool Status (Log)"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;
            // Header row for shop block
            wsSum.Cell(rowPtr, 1).Value = "Status"; wsSum.Cell(rowPtr, 2).Value = "Count"; wsSum.Cell(rowPtr, 3).Value = "%"; wsSum.Row(rowPtr).Style.Font.SetBold(true);
            int shopHeaderRow = rowPtr;
            rowPtr++;

            var presentShopGroups = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase)).Select(k => k.Key.group).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var g in ShopMinorGroups)
            {
                if (!presentShopGroups.Contains(g)) continue;
                var groupCount = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
                if (groupCount <= 0) continue;
                wsSum.Cell(rowPtr, 1).Value = g; wsSum.Cell(rowPtr, 2).Value = groupCount; wsSum.Cell(rowPtr, 3).Value = ((double)groupCount) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                rowPtr++;
                var statuses = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Select(k => (status: k.Key.status, count: k.Value)).OrderBy(s => s.status, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var st in statuses)
                {
                    wsSum.Cell(rowPtr, 1).Value = " " + st.status; wsSum.Cell(rowPtr, 2).Value = st.count; wsSum.Cell(rowPtr, 3).Value = ((double)st.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                    rowPtr++;
                }
            }

            var fabChildren = new List<(string name, long count, List<(string status, long count)> children)>();
            long fabTotal = 0;
            foreach (var fg in FabGroupKeys)
            {
                if (!presentShopGroups.Contains(fg)) continue;
                var fgCount = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, fg, StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
                if (fgCount <= 0) continue;
                var sts = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, fg, StringComparison.OrdinalIgnoreCase)).Select(k => (status: k.Key.status, count: k.Value)).OrderBy(s => s.status, StringComparer.OrdinalIgnoreCase).ToList();
                fabChildren.Add((fg, fgCount, sts));
                fabTotal += fgCount;
            }

            if (fabTotal > 0)
            {
                wsSum.Cell(rowPtr, 1).Value = "III. FABRICATION"; wsSum.Cell(rowPtr, 2).Value = fabTotal; wsSum.Cell(rowPtr, 3).Value = ((double)fabTotal) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                rowPtr++;
                foreach (var sub in fabChildren)
                {
                    wsSum.Cell(rowPtr, 1).Value = " " + sub.name; wsSum.Cell(rowPtr, 2).Value = sub.count; wsSum.Cell(rowPtr, 3).Value = ((double)sub.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                    rowPtr++;
                    foreach (var st in sub.children)
                    {
                        wsSum.Cell(rowPtr, 1).Value = " " + st.status; wsSum.Cell(rowPtr, 2).Value = st.count; wsSum.Cell(rowPtr, 3).Value = ((double)st.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                        rowPtr++;
                    }
                }
            }

            var groupedStatuses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var arr in StatusGroupMap.Values) foreach (var s in arr) groupedStatuses.Add(s);
            var ungrouped = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && !groupedStatuses.Contains(k.Key.status)).GroupBy(k => k.Key.group).Select(g => new { Group = g.Key, Items = g.Select(x => (status: x.Key.status, count: x.Value)).ToList(), Total = g.Sum(x => x.Value) }).ToList();
            if (ungrouped.Count > 0)
            {
                foreach (var ug in ungrouped)
                {
                    wsSum.Cell(rowPtr, 1).Value = ug.Group; wsSum.Cell(rowPtr, 2).Value = ug.Total; wsSum.Cell(rowPtr, 3).Value = ((double)ug.Total) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                    rowPtr++;
                    foreach (var it in ug.Items.OrderBy(i => i.status, StringComparer.OrdinalIgnoreCase))
                    {
                        wsSum.Cell(rowPtr, 1).Value = " " + it.status; wsSum.Cell(rowPtr, 2).Value = it.count; wsSum.Cell(rowPtr, 3).Value = ((double)it.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                        rowPtr++;
                    }
                }
            }

            // Write shop totals and overall total
            wsSum.Cell(rowPtr, 1).Value = "Shop Spool Status Total"; wsSum.Cell(rowPtr, 2).Value = totalShop; wsSum.Cell(rowPtr, 3).Value = totalOverall == 0 ? 0.0 : ((double)totalShop) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;
            wsSum.Cell(rowPtr, 1).Value = "OVERALL TOTAL"; wsSum.Cell(rowPtr, 2).Value = totalOverall; wsSum.Cell(rowPtr, 3).Value = 1.0; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;

            // compute shop table bounds (exclude the final OVERALL TOTAL row)
            int shopTableEndRow = Math.Max(shopHeaderRow, rowPtr - 2);
            try
            {
                var shopRange = wsSum.Range(shopHeaderRow, 1, shopTableEndRow, 3);
                var shopTable = shopRange.CreateTable();
                shopTable.Theme = XLTableTheme.TableStyleMedium2;
                shopTable.ShowTotalsRow = false;
                // center Count and % columns in the shop table
                shopRange.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                shopRange.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            catch { /* safe-fail */ }

            wsSum.Columns().AdjustToContents();
            wsSum.SheetView.FreezeRows(1);

            using var outMs = new MemoryStream();
            wb.SaveAs(outMs);
            var bytes2 = outMs.ToArray();
            var fileName2 = $"SpoolReleaseLog_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(bytes2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName2);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportSpoolReleaseLogFiltered failed");
            return StatusCode(500, "Failed to export filtered Spool Release Log.");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportSpoolErectionLog([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolErectionLogAsync(projectId);
            string? statusCol = FindStatusColumn(dt); int statusIdx = -1;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var name = (dt.Columns[i].ColumnName ?? string.Empty).Trim();
                if (NameContainsToken(name, "STATUS")) { statusCol = dt.Columns[i].ColumnName; statusIdx = i; break; }
            }
            if (statusIdx >= 0 && !dt.Columns.Contains("ACTION BY"))
            {
                var actionByCol = dt.Columns.Add("ACTION BY", typeof(string));
                actionByCol.SetOrdinal(statusIdx + 1);
                foreach (DataRow row in dt.Rows)
                {
                    var raw = statusCol == null ? string.Empty : (row[statusCol]?.ToString() ?? string.Empty);
                    var s = NormalizeStatus(raw);
                    row["ACTION BY"] = ActionByMap.TryGetValue(s, out var act) ? act : "OTHER";
                }
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Spool Erection Log");
            for (int c = 0; c < dt.Columns.Count; c++) ws.Cell(1, c + 1).Value = dt.Columns[c].ColumnName;
            int r = 2;
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dr[c]; var cell = ws.Cell(r, c + 1);
                    if (val == null || val == DBNull.Value) { cell.SetValue(string.Empty); cell.Style.NumberFormat.Format = "@"; }
                    else if (val is DateTime dtv) { cell.SetValue(dtv); cell.Style.DateFormat.Format = "dd-mmm-yyyy"; }
                    else { cell.SetValue(val.ToString() ?? string.Empty); cell.Style.NumberFormat.Format = "@"; }
                }
                ws.Row(r).Height = 17; r++;
            }
            var table = ws.Range(1, 1, dt.Rows.Count + 1, dt.Columns.Count).CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2; ws.Row(1).Height = 30; ws.Columns().AdjustToContents(); ws.SheetView.FreezeRows(1);
            using var ms = new MemoryStream(); wb.SaveAs(ms); var bytes = ms.ToArray(); var fileName = $"SpoolErectionLog_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportSpoolErectionLog failed");
            return Content("Failed to export Spool Erection Log.");
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [IgnoreAntiforgeryToken]
    public IActionResult ExportSpoolErectionLogFiltered([FromBody] FilteredExportRequest? req)
    {
        try
        {
            if (req == null) return BadRequest("Invalid payload");
            var columns = req.Columns ?? new List<string>();
            var rows = req.Rows ?? new List<List<string>>();

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Spool Erection Log");

            if (columns.Count == 0)
            {
                ws.Cell(1, 1).SetValue("No data");
            }
            else
            {
                // headers
                for (int c = 0; c < columns.Count; c++)
                    ws.Cell(1, c + 1).Value = columns[c] ?? string.Empty;

                // rows
                int r = 2;
                foreach (var row in rows)
                {
                    for (int c = 0; c < columns.Count; c++)
                    {
                        var val = (row != null && c < row.Count) ? (row[c] ?? string.Empty) : string.Empty;
                        var cell = ws.Cell(r, c + 1);
                        if (string.IsNullOrWhiteSpace(val))
                        {
                            cell.SetValue(string.Empty); cell.Style.NumberFormat.Format = "@";
                        }
                        else if (DateTime.TryParse(val, out var dtv))
                        {
                            cell.SetValue(dtv);
                            cell.Style.DateFormat.Format = "dd-mmm-yyyy";
                        }
                        else
                        {
                            cell.SetValue(val); cell.Style.NumberFormat.Format = "@";
                        }
                    }
                    ws.Row(r).Height = 17; r++;
                }

                int lastRow = Math.Max(1, rows.Count + 1);
                var fullRange = ws.Range(1, 1, lastRow, columns.Count);
                var table = fullRange.CreateTable();
                table.Theme = XLTableTheme.TableStyleMedium2;
                table.ShowTotalsRow = false;

                ws.Row(1).Height = 30;
                for (int c = 0; c < columns.Count; c++)
                {
                    var name = (columns[c] ?? string.Empty).ToLowerInvariant();
                    var col = ws.Column(c + 1);
                    if (name.Contains("id") || name.Contains("no") || name.Contains("spool") || name.Contains("project"))
                        col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    else
                        col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                }
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            // Build normalized counts by Area/Group/Status (use long)
            int statusColIndex = -1;
            for (int c = 0; c < columns.Count; c++) if (NameContainsToken(columns[c], "STATUS")) { statusColIndex = c; break; }

            var normCounts = new Dictionary<(string area, string group, string status), long>(new TupleKeyComparer());
            if (statusColIndex >= 0)
            {
                foreach (var row in rows)
                {
                    var raw = (row != null && statusColIndex < row.Count) ? (row[statusColIndex] ?? string.Empty) : string.Empty;
                    var s = NormalizeStatus(raw);
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    string group = "Z. UNGROUPED";
                    foreach (var kv in StatusGroupMap)
                    {
                        if (kv.Value.Any(v => string.Equals(v, s, StringComparison.OrdinalIgnoreCase))) { group = kv.Key; break; }
                    }

                    var area = GroupOrderSite.Contains(group, StringComparer.OrdinalIgnoreCase) ? "Site" : "Shop";
                    var key = (area, group, s);
                    normCounts.TryGetValue(key, out var cur);
                    normCounts[key] = cur + 1L;
                }
            }

            // Create summary sheet matching Dashboard grouping
            var wsSum = wb.Worksheets.Add("Spool Status Summary");

            long totalSite = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
            long totalShop = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
            long totalOverall = totalSite + totalShop;
            double den = totalOverall == 0 ? 1.0 : (double)totalOverall;

            int rowPtr = 1;
            // SITE block
            wsSum.Cell(rowPtr, 1).Value = "Site Spool Status"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;
            // Header row for site block
            wsSum.Cell(rowPtr, 1).Value = "Status"; wsSum.Cell(rowPtr, 2).Value = "Count"; wsSum.Cell(rowPtr, 3).Value = "%"; wsSum.Row(rowPtr).Style.Font.SetBold(true);
            // capture site header row index
            int siteHeaderRow = rowPtr;
            rowPtr++;

            var siteGroups = StatusGroupMap.Keys.Where(k => GroupOrderSite.Contains(k, StringComparer.OrdinalIgnoreCase)).ToList();
            if (siteGroups.Count == 0)
                siteGroups = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase)).Select(k => k.Key.group).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            bool wroteAnySite = false;
            foreach (var g in siteGroups)
            {
                var groupCount = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
                if (groupCount <= 0) continue;
                wsSum.Cell(rowPtr, 1).Value = g; wsSum.Cell(rowPtr, 2).Value = groupCount; wsSum.Cell(rowPtr, 3).Value = ((double)groupCount) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                rowPtr++;
                var statuses = normCounts.Where(k => string.Equals(k.Key.area, "Site", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Select(k => (status: k.Key.status, count: k.Value)).OrderBy(s => s.status, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var st in statuses)
                {
                    wsSum.Cell(rowPtr, 1).Value = " " + st.status; wsSum.Cell(rowPtr, 2).Value = st.count; wsSum.Cell(rowPtr, 3).Value = ((double)st.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                    rowPtr++;
                }
                wroteAnySite = true;
            }
            if (!wroteAnySite)
            {
                wsSum.Cell(rowPtr, 1).Value = "No data"; wsSum.Cell(rowPtr, 2).Value = 0; wsSum.Cell(rowPtr, 3).Value = 0.0; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; rowPtr++;
            }
            // Write site total and then mark end of the site table range (include total row)
            wsSum.Cell(rowPtr, 1).Value = "Site Spool Status Total"; wsSum.Cell(rowPtr, 2).Value = totalSite; wsSum.Cell(rowPtr, 3).Value = totalOverall == 0 ? 0.0 : ((double)totalSite) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
            // site table ends at (rowPtr) which we'll include in the table; advance pointer after marking end
            rowPtr += 2;
            // compute site table bounds
            int siteTableEndRow = Math.Max(siteHeaderRow, rowPtr - 2);
            try
            {
                var siteRange = wsSum.Range(siteHeaderRow, 1, siteTableEndRow, 3);
                var siteTable = siteRange.CreateTable();
                siteTable.Theme = XLTableTheme.TableStyleMedium2;
                siteTable.ShowTotalsRow = false;
                // center Count and % columns in the site table
                siteRange.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                siteRange.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            catch { /* safe-fail */ }

            // SHOP block
            wsSum.Cell(rowPtr, 1).Value = "Shop Spool Status (Log)"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;
            // Header row for shop block
            wsSum.Cell(rowPtr, 1).Value = "Status"; wsSum.Cell(rowPtr, 2).Value = "Count"; wsSum.Cell(rowPtr, 3).Value = "%"; wsSum.Row(rowPtr).Style.Font.SetBold(true);
            int shopHeaderRow = rowPtr;
            rowPtr++;

            var presentShopGroups = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase)).Select(k => k.Key.group).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var g in ShopMinorGroups)
            {
                if (!presentShopGroups.Contains(g)) continue;
                var groupCount = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
                if (groupCount <= 0) continue;
                wsSum.Cell(rowPtr, 1).Value = g; wsSum.Cell(rowPtr, 2).Value = groupCount; wsSum.Cell(rowPtr, 3).Value = ((double)groupCount) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                rowPtr++;
                var statuses = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, g, StringComparison.OrdinalIgnoreCase)).Select(k => (status: k.Key.status, count: k.Value)).OrderBy(s => s.status, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var st in statuses)
                {
                    wsSum.Cell(rowPtr, 1).Value = " " + st.status; wsSum.Cell(rowPtr, 2).Value = st.count; wsSum.Cell(rowPtr, 3).Value = ((double)st.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                    rowPtr++;
                }
            }

            var fabChildren = new List<(string name, long count, List<(string status, long count)> children)>();
            long fabTotal = 0;
            foreach (var fg in FabGroupKeys)
            {
                if (!presentShopGroups.Contains(fg)) continue;
                var fgCount = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, fg, StringComparison.OrdinalIgnoreCase)).Sum(k => k.Value);
                if (fgCount <= 0) continue;
                var sts = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && string.Equals(k.Key.group, fg, StringComparison.OrdinalIgnoreCase)).Select(k => (status: k.Key.status, count: k.Value)).OrderBy(s => s.status, StringComparer.OrdinalIgnoreCase).ToList();
                fabChildren.Add((fg, fgCount, sts));
                fabTotal += fgCount;
            }

            if (fabTotal > 0)
            {
                wsSum.Cell(rowPtr, 1).Value = "III. FABRICATION"; wsSum.Cell(rowPtr, 2).Value = fabTotal; wsSum.Cell(rowPtr, 3).Value = ((double)fabTotal) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                rowPtr++;
                foreach (var sub in fabChildren)
                {
                    wsSum.Cell(rowPtr, 1).Value = " " + sub.name; wsSum.Cell(rowPtr, 2).Value = sub.count; wsSum.Cell(rowPtr, 3).Value = ((double)sub.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                    rowPtr++;
                    foreach (var st in sub.children)
                    {
                        wsSum.Cell(rowPtr, 1).Value = " " + st.status; wsSum.Cell(rowPtr, 2).Value = st.count; wsSum.Cell(rowPtr, 3).Value = ((double)st.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                        rowPtr++;
                    }
                }
            }

            var groupedStatuses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var arr in StatusGroupMap.Values) foreach (var s in arr) groupedStatuses.Add(s);
            var ungrouped = normCounts.Where(k => string.Equals(k.Key.area, "Shop", StringComparison.OrdinalIgnoreCase) && !groupedStatuses.Contains(k.Key.status)).GroupBy(k => k.Key.group).Select(g => new { Group = g.Key, Items = g.Select(x => (status: x.Key.status, count: x.Value)).ToList(), Total = g.Sum(x => x.Value) }).ToList();
            if (ungrouped.Count > 0)
            {
                foreach (var ug in ungrouped)
                {
                    wsSum.Cell(rowPtr, 1).Value = ug.Group; wsSum.Cell(rowPtr, 2).Value = ug.Total; wsSum.Cell(rowPtr, 3).Value = ((double)ug.Total) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true);
                    rowPtr++;
                    foreach (var it in ug.Items.OrderBy(i => i.status, StringComparer.OrdinalIgnoreCase))
                    {
                        wsSum.Cell(rowPtr, 1).Value = " " + it.status; wsSum.Cell(rowPtr, 2).Value = it.count; wsSum.Cell(rowPtr, 3).Value = ((double)it.count) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%";
                        rowPtr++;
                    }
                }
            }

            // Write shop totals and overall total
            wsSum.Cell(rowPtr, 1).Value = "Shop Spool Status Total"; wsSum.Cell(rowPtr, 2).Value = totalShop; wsSum.Cell(rowPtr, 3).Value = totalOverall == 0 ? 0.0 : ((double)totalShop) / den; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;
            wsSum.Cell(rowPtr, 1).Value = "OVERALL TOTAL"; wsSum.Cell(rowPtr, 2).Value = totalOverall; wsSum.Cell(rowPtr, 3).Value = 1.0; wsSum.Cell(rowPtr, 3).Style.NumberFormat.Format = "0.00%"; wsSum.Range(rowPtr, 1, rowPtr, 3).Style.Font.SetBold(true); rowPtr++;

            // compute shop table bounds (exclude the final OVERALL TOTAL row)
            int shopTableEndRow = Math.Max(shopHeaderRow, rowPtr - 2);
            try
            {
                var shopRange = wsSum.Range(shopHeaderRow, 1, shopTableEndRow, 3);
                var shopTable = shopRange.CreateTable();
                shopTable.Theme = XLTableTheme.TableStyleMedium2;
                shopTable.ShowTotalsRow = false;
                // center Count and % columns in the shop table
                shopRange.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                shopRange.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            catch { /* safe-fail */ }

            wsSum.Columns().AdjustToContents();
            wsSum.SheetView.FreezeRows(1);

            using var outMs = new MemoryStream();
            wb.SaveAs(outMs);
            var bytes2 = outMs.ToArray();
            var fileName2 = $"SpoolReleaseLog_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(bytes2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName2);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportSpoolErectionLogFiltered failed");
            return StatusCode(500, "Failed to export filtered Spool Erection Log.");
        }
    }

    ////////////////////////////////////////////////////////////////////////////////
    // Code modifications end here
    ////////////////////////////////////////////////////////////////////////////////

    // Helpers: OpenXML chart creation
    private static C.ChartSpace CreateDoughnutChartSpace(string title, string categoryRange, string valuesRange)
    {
        var chart = new C.Chart(
        new C.Title(new C.ChartText(new C.RichText(
        new A.BodyProperties(),
        new A.ListStyle(),
        new A.Paragraph(new A.Run(new A.Text(title))))), new C.Overlay { Val = false }),
        new C.PlotArea(
        new C.Layout(),
        new C.DoughnutChart(
        new C.VaryColors { Val = true },
        new C.PieChartSeries(
        new C.Index { Val = 0u },
        new C.Order { Val = 0u },
        new C.SeriesText(new C.NumericValue { Text = "Count" }),
        new C.CategoryAxisData(new C.StringReference(new C.Formula { Text = categoryRange })),
        new C.Values(new C.NumberReference(new C.Formula { Text = valuesRange }))
        ),
        new C.HoleSize { Val = (byte)40 }
        ),
        new C.DataLabels(new C.ShowLegendKey { Val = false }, new C.ShowValue { Val = true }, new C.ShowCategoryName { Val = false }, new C.ShowPercent { Val = true })
        ),
        new C.PlotVisibleOnly { Val = true }
        );

        return new C.ChartSpace(
        new C.EditingLanguage { Val = "en-US" },
        chart
        );
    }

    private static C.ChartSpace CreateStackedBarChartSpace(string title, string categoryRange, List<(string nameRef, string valuesRef)> seriesRanges)
    {
        var barChart = new C.BarChart(
        new C.BarDirection { Val = C.BarDirectionValues.Bar },
        new C.BarGrouping { Val = C.BarGroupingValues.Stacked },
        new C.VaryColors { Val = false }
        );

        uint idx = 0;
        foreach (var (nameRef, valuesRef) in seriesRanges)
        {
            var ser = new C.BarChartSeries(
            new C.Index { Val = idx },
            new C.Order { Val = idx },
            new C.SeriesText(new C.StringReference(new C.Formula { Text = nameRef })),
            new C.CategoryAxisData(new C.StringReference(new C.Formula { Text = categoryRange })),
            new C.Values(new C.NumberReference(new C.Formula { Text = valuesRef }))
            );
            barChart.Append(ser);
            idx++;
        }
        barChart.Append(new C.DataLabels(new C.ShowLegendKey { Val = false }, new C.ShowValue { Val = true }, new C.ShowCategoryName { Val = false }, new C.ShowPercent { Val = false }));
        barChart.Append(new C.Overlap { Val = 100 });
        barChart.Append(new C.AxisId { Val = 48650112u });
        barChart.Append(new C.AxisId { Val = 48672768u });

        var catAx = new C.CategoryAxis(
        new C.AxisId { Val = 48650112u },
        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
        new C.Delete { Val = false },
        new C.AxisPosition { Val = C.AxisPositionValues.Left },
        new C.MajorTickMark { Val = C.TickMarkValues.Outside },
        new C.MinorTickMark { Val = C.TickMarkValues.None },
        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
        new C.CrossingAxis { Val = 48672768u },
        new C.Crosses { Val = C.CrossesValues.AutoZero },
        new C.AutoLabeled { Val = true },
        new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
        new C.LabelOffset { Val = 100 }
        );
        var valAx = new C.ValueAxis(
        new C.AxisId { Val = 48672768u },
        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
        new C.Delete { Val = false },
        new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
        new C.MajorGridlines(),
        new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
        new C.CrossingAxis { Val = 48650112u },
        new C.Crosses { Val = C.CrossesValues.AutoZero },
        new C.CrossBetween { Val = C.CrossBetweenValues.Between }
        );

        var chart = new C.Chart(
        new C.Title(new C.ChartText(new C.RichText(
        new A.BodyProperties(),
        new A.ListStyle(),
        new A.Paragraph(new A.Run(new A.Text(title))))), new C.Overlay { Val = false }),
        new C.PlotArea(new C.Layout(), barChart, catAx, valAx),
        new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Right }, new C.Overlay { Val = false }),
        new C.PlotVisibleOnly { Val = true }
        );

        return new C.ChartSpace(
        new C.EditingLanguage { Val = "en-US" },
        chart
        );
    }

    private static void AddChartToSheet(DrawingsPart drawingsPart, ChartPart chartPart, int fromCol, int fromRow, int toCol, int toRow)
    {
        var wsDr = drawingsPart.WorksheetDrawing!;
        // Create unique IDs
        uint shapeId = (uint)(wsDr.ChildElements.Count + 1);

        var twoCellAnchor = new Xdr.TwoCellAnchor(
        new Xdr.FromMarker(new Xdr.ColumnId(fromCol.ToString()), new Xdr.ColumnOffset("0"), new Xdr.RowId(fromRow.ToString()), new Xdr.RowOffset("0")),
        new Xdr.ToMarker(new Xdr.ColumnId(toCol.ToString()), new Xdr.ColumnOffset("0"), new Xdr.RowId(toRow.ToString()), new Xdr.RowOffset("0")),
        new Xdr.GraphicFrame(
        new Xdr.NonVisualGraphicFrameProperties(
        new Xdr.NonVisualDrawingProperties { Id = shapeId, Name = $"Chart {shapeId}" },
        new Xdr.NonVisualGraphicFrameDrawingProperties()
        ),
        new Xdr.Transform(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 0, Cy = 0 }),
        new A.Graphic(
        new A.GraphicData(
        new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) }
        )
        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
        )
        ),
        new Xdr.ClientData()
        );
        wsDr.Append(twoCellAnchor);
    }

    private static string GetColumnLetter(int columnNumber)
    {
        //1 -> A,2 -> B, ...
        var dividend = columnNumber;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> SpoolReleaseLog([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolReleaseLogAsync(projectId);

            // Determine project id used (default to max project id when not provided)
            var pid = await GetDefaultProjectIdAsync(projectId);
            // Re-filter in case stored procedure returns multiple projects
            dt = FilterByProject(dt, pid);
            ViewBag.ProjectId = pid;

            // Project options: reuse same logic as SpoolReleaseLog
            var projList = await _context.Projects_tbl.AsNoTracking()
            .Join(_context.SP_Release_tbl.AsNoTracking(), p => p.Project_ID, sp => sp.SP_Project_No, (p, sp) => new { p.Project_ID, p.Project_Name })
            .Distinct()
            .OrderBy(x => x.Project_ID)
            .ToListAsync();
            // Format as simple strings "ID - Name" for the view
            ViewBag.ProjectOptions = projList.Select(p => string.Concat(p.Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty)).ToList();
            // Provide a structured list and the default project id (pid) for the view so it can set the selected item
            ViewBag.ProjectsSelect = projList.Select(p => new { Id = p.Project_ID, Label = string.Concat(p.Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty) }).ToList();
            ViewBag.DefaultProjectId = pid; // pid is the project actually used (may be default when projectId was null)

            //2) Dwg / Sheet / Spool options driven from stored procedure result (already filtered by project)
            var dwgOptions = new List<string>();
            var sheetOptions = new List<string>();
            var spoolOptions = new List<string>();
            var spoolMap = new List<object>();
            try
            {
                if (pid.HasValue)
                {
                    var spQuery = _context.SP_Release_tbl.AsNoTracking().Where(sp => sp.SP_Project_No == pid.Value);
                    dwgOptions = await spQuery.Where(sp => sp.SP_LAYOUT_NUMBER != null && sp.SP_LAYOUT_NUMBER != "").Select(sp => sp.SP_LAYOUT_NUMBER!).Distinct().OrderBy(s => s).ToListAsync();
                    sheetOptions = await spQuery.Where(sp => sp.SP_SHEET != null && sp.SP_SHEET != "").Select(sp => sp.SP_SHEET!).Distinct().OrderBy(s => s).ToListAsync();
                    spoolOptions = await spQuery.Where(sp => sp.SP_SPOOL_NUMBER != null && sp.SP_SPOOL_NUMBER != "").Select(sp => sp.SP_SPOOL_NUMBER!).Distinct().OrderBy(s => s).ToListAsync();

                    var spoolEntries = await spQuery
                        .Where(sp => sp.SP_SPOOL_NUMBER != null && sp.SP_SPOOL_NUMBER != "")
                        .Select(sp => new { Dwg = sp.SP_LAYOUT_NUMBER ?? string.Empty, Sheet = sp.SP_SHEET ?? string.Empty, Spool = sp.SP_SPOOL_NUMBER ?? string.Empty })
                        .ToListAsync();

                    spoolMap = spoolEntries
                        .GroupBy(e => e.Dwg ?? string.Empty)
                        .Select(g => new
                        {
                            Dwg = g.Key ?? string.Empty,
                            Sheets = g.GroupBy(x => x.Sheet ?? string.Empty)
                                .Where(gg => !string.IsNullOrWhiteSpace(gg.Key))
                                .Select(gg => new
                                {
                                    Sheet = gg.Key,
                                    Spools = gg.Select(x => x.Spool).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList()
                                }).ToList(),
                            Spools = g.Select(x => x.Spool).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList()
                        })
                        .Where(x => !string.IsNullOrWhiteSpace(x.Dwg))
                        .Cast<object>()
                        .ToList();
                }
                else
                {
                    // fallback to DataTable-derived options when no project is selected
                    if (dt != null && dt.Columns.Count > 0)
                    {
                        string? dwgCol = dt.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "DWG", StringComparison.OrdinalIgnoreCase))?.ColumnName;
                        string? sheetCol = dt.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "SHEET", StringComparison.OrdinalIgnoreCase))?.ColumnName;
                        string? spoolCol = dt.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "SPOOL", StringComparison.OrdinalIgnoreCase))?.ColumnName;

                        var rows = dt.Rows.Cast<DataRow>();

                        if (!string.IsNullOrWhiteSpace(dwgCol))
                        {
                            dwgOptions = rows
                                .Select(r => r[dwgCol]?.ToString()?.Trim())
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                .ToList()!;
                        }

                        if (!string.IsNullOrWhiteSpace(sheetCol))
                        {
                            sheetOptions = rows
                                .Select(r => r[sheetCol]?.ToString()?.Trim())
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                .ToList()!;
                        }

                        if (!string.IsNullOrWhiteSpace(spoolCol))
                        {
                            spoolOptions = rows
                                .Select(r => r[spoolCol]?.ToString()?.Trim())
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                .ToList()!;
                        }

                        if (!string.IsNullOrWhiteSpace(dwgCol))
                        {
                            spoolMap = rows
                                .Select(r => new
                                {
                                    Dwg = r[dwgCol]?.ToString()?.Trim() ?? string.Empty,
                                    Sheet = string.IsNullOrWhiteSpace(sheetCol) ? string.Empty : (r[sheetCol!]?.ToString()?.Trim() ?? string.Empty),
                                    Spool = string.IsNullOrWhiteSpace(spoolCol) ? string.Empty : (r[spoolCol!]?.ToString()?.Trim() ?? string.Empty)
                                })
                                .Where(x => !string.IsNullOrWhiteSpace(x.Dwg))
                                .GroupBy(x => x.Dwg, StringComparer.OrdinalIgnoreCase)
                                .Select(g => new
                                {
                                    Dwg = g.Key,
                                    Sheets = string.IsNullOrWhiteSpace(sheetCol)
                                        ? new List<object>()
                                        : g.GroupBy(x => x.Sheet, StringComparer.OrdinalIgnoreCase)
                                            .Where(gg => !string.IsNullOrWhiteSpace(gg.Key))
                                            .Select(gg => new
                                            {
                                                Sheet = gg.Key,
                                                Spools = gg.Select(x => x.Spool)
                                                    .Where(s => !string.IsNullOrWhiteSpace(s))
                                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                                    .ToList()
                                            })
                                            .Cast<object>()
                                            .ToList(),
                                    Spools = g.Select(x => x.Spool)
                                        .Where(s => !string.IsNullOrWhiteSpace(s))
                                        .Distinct(StringComparer.OrdinalIgnoreCase)
                                        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                        .ToList()
                                })
                                .Cast<object>()
                                .ToList();
                        }
                    }
                }
            }
            catch
            {
                // safe-fail: leave options empty on error
            }

            ViewBag.DwgOptions = dwgOptions;
            ViewBag.SheetOptions = sheetOptions;
            ViewBag.SpoolOptions = spoolOptions;
            ViewBag.SpoolMap = spoolMap;

            //3) Status options: read distinct STATUS-like column values from returned DataTable if present
            var statusOpts = new List<string>();
            if (dt != null && dt.Columns.Count > 0)
            {
                string? statusCol = FindStatusColumn(dt);
                if (!string.IsNullOrWhiteSpace(statusCol))
                {
                    statusOpts = dt.Rows.Cast<DataRow>()
                    .Select(r => r[statusCol]?.ToString()?.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s)
                    .ToList()!;
                }
            }
            ViewBag.StatusOptions = statusOpts;

            // Build a mapping of status -> short prefix (e.g. "A-", "B-") based on StatusGroupMap so the view
            // can display options like "(A-)01. READY FOR DISPATCH". This keeps grouping logic centralized in the controller.
            try
            {
                var statusPrefixes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var s in statusOpts)
                {
                    if (string.IsNullOrWhiteSpace(s)) continue;
                    string prefix = string.Empty;
                    var grp = StatusGroupMap.Keys.FirstOrDefault(k => StatusGroupMap[k].Any(v => string.Equals(v, s, StringComparison.OrdinalIgnoreCase)));
                    if (!string.IsNullOrWhiteSpace(grp))
                    {
                        // take token before the first '.' (e.g. "A" from "A. PRODUCTION")
                        string token = grp.Trim();
                        int dot = token.IndexOf('.');
                        if (dot >= 0) token = token[..dot].Trim();
                        // Only use a single alphabetic character as a prefix to avoid tokens like "II" producing "II-"
                        if (!string.IsNullOrEmpty(token) && token.Length == 1 && char.IsLetter(token[0]))
                        {
                            prefix = token + "-";
                        }
                    }
                    statusPrefixes[s] = prefix;
                }
                ViewBag.StatusPrefixes = statusPrefixes;
            }
            catch
            {
                ViewBag.StatusPrefixes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            // Ensure the header welcome shows FirstName + LastName when available; fallback to stored FullName
            var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
            var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
            var composed = (fn + " " + ln).Trim();
            if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
            ViewBag.FullName = composed;

            return View("SpoolReleaseLog", dt);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SpoolReleaseLog view failed");
            return Content("Failed to load Spool Release Log.");
        }
    }

    private async Task<DataTable> LoadSpoolErectionLogAsync(int? projectId)
    {
        var dt = new DataTable();
        var pid = await GetDefaultProjectIdAsync(projectId);
        var conn = _context.Database.GetDbConnection();
        if (conn.State != System.Data.ConnectionState.Open) await conn.OpenAsync();

        async Task TryExecAsync(CommandType type, string sql, params (string name, object? val)[] parameters)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.CommandType = type;
            cmd.CommandTimeout = 600;
            foreach (var (name, val) in parameters)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = name;
                p.Value = val ?? DBNull.Value;
                cmd.Parameters.Add(p);
            }
            try
            {
                using var reader = await cmd.ExecuteReaderAsync();
                if (reader != null)
                {
                    DataTable? best = null;
                    int bestRows = -1;
                    do
                    {
                        var tmp = new DataTable();
                        tmp.Load(reader);
                        if (tmp.Columns.Cast<DataColumn>().Any(c => NameContainsToken(c.ColumnName, "STATUS")))
                        {
                            dt = tmp; best = tmp; break;
                        }
                        if (tmp.Rows.Count > bestRows)
                        {
                            best = tmp; bestRows = tmp.Rows.Count;
                        }
                    }
                    while (await reader.NextResultAsync());
                    if (best != null && dt.Rows.Count == 0) dt = best;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Spool_Erection_Log_Q execution failed for {Type} {Sql}", type, sql);
            }
        }

        await TryExecAsync(CommandType.StoredProcedure, "[dbo].[Spool_Erection_Log_Q]", ("@ProjectID", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.StoredProcedure, "[dbo].[Spool_Erection_Log_Q]", ("@Project_No", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.Text, "EXEC [dbo].[Spool_Erection_Log_Q] @ProjectID = @p0", ("@p0", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.Text, "EXEC [dbo].[Spool_Erection_Log_Q] @Project_No = @p0", ("@p0", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.Text, "EXEC [dbo].[Spool_Erection_Log_Q]");

        return dt;
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> SpoolErectionLog([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolErectionLogAsync(projectId);

            // Ensure rows are strictly scoped to the selected/default project even if the proc returns all
            dt = FilterByProject(dt, await GetDefaultProjectIdAsync(projectId));

            // Determine project id used (default to max project id when not provided)
            var pid = await GetDefaultProjectIdAsync(projectId);
            ViewBag.ProjectId = pid;

            // Project options: reuse same logic as SpoolReleaseLog
            var projList = await _context.Projects_tbl.AsNoTracking()
            .Join(_context.SP_Release_tbl.AsNoTracking(), p => p.Project_ID, sp => sp.SP_Project_No, (p, sp) => new { p.Project_ID, p.Project_Name })
            .Distinct()
            .OrderBy(x => x.Project_ID)
            .ToListAsync();

            ViewBag.ProjectOptions = projList.Select(p => string.Concat(p.Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty)).ToList();
            ViewBag.ProjectsSelect = projList.Select(p => new { Id = p.Project_ID, Label = string.Concat(p.Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty) }).ToList();
            ViewBag.DefaultProjectId = pid;

            var dwgOptions = new List<string>();
            var sheetOptions = new List<string>();
            var spoolOptions = new List<string>();
            var spoolMap = new List<object>();
            try
            {
                if (dt != null && dt.Columns.Count > 0)
                {
                    string? dwgCol = dt.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "DWG", StringComparison.OrdinalIgnoreCase))?.ColumnName;
                    string? sheetCol = dt.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "SHEET", StringComparison.OrdinalIgnoreCase))?.ColumnName;
                    string? spoolCol = dt.Columns.Cast<DataColumn>().FirstOrDefault(c => string.Equals(c.ColumnName, "SPOOL", StringComparison.OrdinalIgnoreCase))?.ColumnName;

                    var rows = dt.Rows.Cast<DataRow>();

                    if (!string.IsNullOrWhiteSpace(dwgCol))
                    {
                        dwgOptions = rows
                            .Select(r => r[dwgCol]?.ToString()?.Trim())
                            .Where(s => !string.IsNullOrWhiteSpace(s))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                            .ToList()!;
                    }

                    if (!string.IsNullOrWhiteSpace(sheetCol))
                    {
                        sheetOptions = rows
                            .Select(r => r[sheetCol]?.ToString()?.Trim())
                            .Where(s => !string.IsNullOrWhiteSpace(s))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                            .ToList()!;
                    }

                    if (!string.IsNullOrWhiteSpace(spoolCol))
                    {
                        spoolOptions = rows
                            .Select(r => r[spoolCol]?.ToString()?.Trim())
                            .Where(s => !string.IsNullOrWhiteSpace(s))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                            .ToList()!;
                    }

                    if (!string.IsNullOrWhiteSpace(dwgCol))
                    {
                        spoolMap = rows
                            .Select(r => new
                            {
                                Dwg = r[dwgCol]?.ToString()?.Trim() ?? string.Empty,
                                Sheet = string.IsNullOrWhiteSpace(sheetCol) ? string.Empty : (r[sheetCol!]?.ToString()?.Trim() ?? string.Empty),
                                Spool = string.IsNullOrWhiteSpace(spoolCol) ? string.Empty : (r[spoolCol!]?.ToString()?.Trim() ?? string.Empty)
                            })
                            .Where(x => !string.IsNullOrWhiteSpace(x.Dwg))
                            .GroupBy(x => x.Dwg, StringComparer.OrdinalIgnoreCase)
                            .Select(g => new
                            {
                                Dwg = g.Key,
                                Sheets = string.IsNullOrWhiteSpace(sheetCol)
                                    ? new List<object>()
                                    : g.GroupBy(x => x.Sheet, StringComparer.OrdinalIgnoreCase)
                                        .Where(gg => !string.IsNullOrWhiteSpace(gg.Key))
                                        .Select(gg => new
                                        {
                                            Sheet = gg.Key,
                                            Spools = gg.Select(x => x.Spool)
                                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                                .ToList()
                                        })
                                        .Cast<object>()
                                        .ToList(),
                                Spools = g.Select(x => x.Spool)
                                    .Where(s => !string.IsNullOrWhiteSpace(s))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                    .ToList()
                            })
                            .Cast<object>()
                            .ToList();
                    }
                }
            }
            catch
            {
                // safe-fail: leave options empty on error
            }

            ViewBag.DwgOptions = dwgOptions;
            ViewBag.SheetOptions = sheetOptions;
            ViewBag.SpoolOptions = spoolOptions;
            ViewBag.SpoolMap = spoolMap;

            // Status options
            var statusOpts = new List<string>();
            if (dt != null && dt.Columns.Count > 0)
            {
                string? statusCol = FindStatusColumn(dt);
                if (!string.IsNullOrWhiteSpace(statusCol))
                {
                    statusOpts = dt.Rows.Cast<DataRow>()
                    .Select(r => r[statusCol]?.ToString()?.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s)
                    .ToList()!;
                }
            }
            ViewBag.StatusOptions = statusOpts;

            // Status prefixes
            try
            {
                var statusPrefixes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var s in statusOpts)
                {
                    if (string.IsNullOrWhiteSpace(s)) continue;
                    string prefix = string.Empty;
                    var grp = StatusGroupMap.Keys.FirstOrDefault(k => StatusGroupMap[k].Any(v => string.Equals(v, s, StringComparison.OrdinalIgnoreCase)));
                    if (!string.IsNullOrWhiteSpace(grp))
                    {
                        string token = grp.Trim();
                        int dot = token.IndexOf('.');
                        if (dot >= 0) token = token[..dot].Trim();
                        if (!string.IsNullOrEmpty(token) && token.Length == 1 && char.IsLetter(token[0]))
                        {
                            prefix = token + "-";
                        }
                    }
                    statusPrefixes[s] = prefix;
                }
                ViewBag.StatusPrefixes = statusPrefixes;
            }
            catch
            {
                ViewBag.StatusPrefixes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            // Full name
            var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
            var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
            var composed = (fn + " " + ln).Trim();
            if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
            ViewBag.FullName = composed;

            return View("SpoolErectionLog", dt);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SpoolErectionLog view failed");
            return Content("Failed to load Spool Erection Log.");
        }
    }

    private async Task<DataTable> LoadSpoolCoatingLogAsync(int? projectId)
    {
        var dt = new DataTable();
        var pid = projectId ?? await GetDefaultProjectIdAsync();
        var conn = _context.Database.GetDbConnection();
        if (conn.State != System.Data.ConnectionState.Open) await conn.OpenAsync();

        async Task TryExecAsync(CommandType type, string sql, params (string name, object? val)[] parameters)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.CommandType = type;
            cmd.CommandTimeout = 600;
            foreach (var (name, val) in parameters)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = name;
                p.Value = val ?? DBNull.Value;
                cmd.Parameters.Add(p);
            }
            try
            {
                using var reader = await cmd.ExecuteReaderAsync();
                if (reader != null)
                {
                    DataTable? best = null;
                    int bestRows = -1;
                    do
                    {
                        var tmp = new DataTable();
                        tmp.Load(reader);
                        if (tmp.Columns.Cast<DataColumn>().Any(c => NameContainsToken(c.ColumnName, "STATUS")))
                        {
                            dt = tmp; best = tmp; break;
                        }
                        if (tmp.Rows.Count > bestRows)
                        {
                            best = tmp; bestRows = tmp.Rows.Count;
                        }
                    }
                    while (await reader.NextResultAsync());
                    if (best != null && dt.Rows.Count == 0) dt = best;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Spool_Coating_Log_Q execution failed for {Type} {Sql}", type, sql);
            }
        }

        await TryExecAsync(CommandType.StoredProcedure, "[dbo].[Spool_Coating_Log_Q]", ("@ProjectID", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.StoredProcedure, "[dbo].[Spool_Coating_Log_Q]", ("@Project_No", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.Text, "EXEC [dbo].[Spool_Coating_Log_Q] @ProjectID = @p0", ("@p0", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.Text, "EXEC [dbo].[Spool_Coating_Log_Q] @Project_No = @p0", ("@p0", (object?)pid));
        if (dt.Rows.Count == 0) await TryExecAsync(CommandType.Text, "EXEC [dbo].[Spool_Coating_Log_Q]");

        return dt;
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> PaintingLog([FromQuery] int? projectId)
    {
        try
        {
            var dt = await LoadSpoolCoatingLogAsync(projectId);

            var pid = await GetDefaultProjectIdAsync(projectId);
            ViewBag.ProjectId = pid;

            // Project options: reuse same logic as SpoolReleaseLog
            var projList = await _context.Projects_tbl.AsNoTracking()
            .Join(_context.SP_Release_tbl.AsNoTracking(), p => p.Project_ID, sp => sp.SP_Project_No, (p, sp) => new { p.Project_ID, p.Project_Name })
            .Distinct()
            .OrderBy(x => x.Project_ID)
            .ToListAsync();

            ViewBag.ProjectOptions = projList.Select(p => string.Concat(p.Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty)).ToList();
            ViewBag.ProjectsSelect = projList.Select(p => new { Id = p.Project_ID, Label = string.Concat(p.Project_ID.ToString(), " - ", p.Project_Name ?? string.Empty) }).ToList();
            ViewBag.DefaultProjectId = pid;

            if (pid.HasValue)
            {
                var spQuery = _context.SP_Release_tbl.AsNoTracking().Where(sp => sp.SP_Project_No == pid.Value);
                var dwgOpts = await spQuery.Where(sp => sp.SP_LAYOUT_NUMBER != null && sp.SP_LAYOUT_NUMBER != "").Select(sp => sp.SP_LAYOUT_NUMBER!).Distinct().OrderBy(s => s).ToListAsync();
                var sheetOpts = await spQuery.Where(sp => sp.SP_SHEET != null && sp.SP_SHEET != "").Select(sp => sp.SP_SHEET!).Distinct().OrderBy(s => s).ToListAsync();
                var spoolOpts = await spQuery
                .Where(sp => sp.SP_SPOOL_NUMBER != null && sp.SP_SPOOL_NUMBER != "")
                .Select(sp => sp.SP_SPOOL_NUMBER!)
                .Distinct()
                .OrderBy(s => s)
                .ToListAsync();

                var spoolEntries = await spQuery
                .Where(sp => sp.SP_SPOOL_NUMBER != null && sp.SP_SPOOL_NUMBER != "")
                .Select(sp => new { Dwg = sp.SP_LAYOUT_NUMBER ?? string.Empty, Sheet = sp.SP_SHEET ?? string.Empty, Spool = sp.SP_SPOOL_NUMBER ?? string.Empty })
                .ToListAsync();

                var spoolMap = spoolEntries
                .GroupBy(e => e.Dwg ?? string.Empty)
                .Select(g => new
                {
                    Dwg = g.Key ?? string.Empty,
                    Sheets = g.GroupBy(x => x.Sheet ?? string.Empty)
                .Where(gg => !string.IsNullOrWhiteSpace(gg.Key))
                .Select(gg => new
                {
                    Sheet = gg.Key,
                    Spools = gg.Select(x => x.Spool).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList()
                }).ToList(),
                    Spools = g.Select(x => x.Spool).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList()
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Dwg))
                .ToList();

                ViewBag.DwgOptions = dwgOpts;
                ViewBag.SheetOptions = sheetOpts;
                ViewBag.SpoolOptions = spoolOpts;
                ViewBag.SpoolMap = spoolMap;
            }
            else
            {
                ViewBag.DwgOptions = new List<string>();
                ViewBag.SheetOptions = new List<string>();
                ViewBag.SpoolOptions = new List<string>();
                ViewBag.SpoolMap = new List<object>();
            }

            // Status options
            var statusOpts = new List<string>();
            if (dt != null && dt.Columns.Count > 0)
            {
                string? statusCol = FindStatusColumn(dt);
                if (!string.IsNullOrWhiteSpace(statusCol))
                {
                    statusOpts = dt.Rows.Cast<DataRow>()
                    .Select(r => r[statusCol]?.ToString()?.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s)
                    .ToList()!;
                }
            }
            ViewBag.StatusOptions = statusOpts;

            // Status prefixes
            try
            {
                var statusPrefixes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var s in statusOpts)
                {
                    if (string.IsNullOrWhiteSpace(s)) continue;
                    string prefix = string.Empty;
                    var grp = StatusGroupMap.Keys.FirstOrDefault(k => StatusGroupMap[k].Any(v => string.Equals(v, s, StringComparison.OrdinalIgnoreCase)));
                    if (!string.IsNullOrWhiteSpace(grp))
                    {
                        string token = grp.Trim();
                        int dot = token.IndexOf('.');
                        if (dot >= 0) token = token[..dot].Trim();
                        if (!string.IsNullOrEmpty(token) && token.Length == 1 && char.IsLetter(token[0]))
                        {
                            prefix = token + "-";
                        }
                    }
                    statusPrefixes[s] = prefix;
                }
                ViewBag.StatusPrefixes = statusPrefixes;
            }
            catch
            {
                ViewBag.StatusPrefixes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            // Full name
            var fn = HttpContext.Session.GetString("FirstName") ?? string.Empty;
            var ln = HttpContext.Session.GetString("LastName") ?? string.Empty;
            var composed = (fn + " " + ln).Trim();
            if (string.IsNullOrEmpty(composed)) composed = HttpContext.Session.GetString("FullName") ?? string.Empty;
            ViewBag.FullName = composed;

            return View("PaintingLog", dt);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PaintingLog view failed");
            return Content("Failed to load Painting Log.");
        }
    }
}