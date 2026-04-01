using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;
using System.Data;
using System.IO;
using ClosedXML.Excel;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> JointProgress(int? projectId = null)
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

        var (weldTypes, defaultWeldTypes) = await GetJointProgressWeldTypesAsync(selectedProject);

        var vm = new JointProgressViewModel
        {
            Projects = projects,
            SelectedProjectId = selectedProject,
            SelectedProjectIds = selectedProject > 0 ? new List<int> { selectedProject } : new List<int>(),
            WeldTypeOptions = weldTypes,
            SelectedWeldTypes = defaultWeldTypes.Count > 0 ? defaultWeldTypes : weldTypes,
            DateBasis = "Welding",
            Grouping = "Monthly"
        };

        return View(vm);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetJointProgressWeldTypes([FromQuery] int projectId)
    {
        if (projectId <= 0)
        {
            return Json(new { options = new List<string>(), defaults = new List<string>() });
        }

        var (options, defaults) = await GetJointProgressWeldTypesAsync(projectId);
        return Json(new { options, defaults });
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetJointProgressData(
        [FromQuery] int projectId,
        [FromQuery] List<int>? projectIds,
        [FromQuery] DateTime? startDate,
        [FromQuery] DateTime? endDate,
        [FromQuery] List<string>? weldTypes,
        [FromQuery] string? location = "All",
        [FromQuery] string? dateBasis = "Welding")
    {
        var resolvedIds = (projectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (resolvedIds.Count == 0 && projectId > 0) resolvedIds.Add(projectId);
        if (resolvedIds.Count > 0) projectId = resolvedIds[0];

        if (projectId <= 0)
        {
            return Json(new JointProgressResponse());
        }

        var basis = (dateBasis ?? "Welding").Trim();
        bool useFitupDate = basis.Equals("Fit-up", StringComparison.OrdinalIgnoreCase);

        var rawWtSet = (weldTypes ?? new List<string>())
            .Select(w => (w ?? string.Empty).Trim().ToUpperInvariant())
            .Where(x => x.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var wtSet = (weldTypes ?? new List<string>())
            .Select(NormalizeWeldType)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        bool filterWeldType = rawWtSet.Count > 0 || wtSet.Count > 0;

        static string TrimUpper(string? s) => (s ?? string.Empty).Trim().ToUpperInvariant();

        var start = startDate?.Date;
        var endExclusive = endDate?.Date.AddDays(1);
        bool filterLocation = !string.IsNullOrWhiteSpace(location) && !string.Equals(location, "All", StringComparison.OrdinalIgnoreCase);
        bool wantShop = location != null && location.Equals("Shop", StringComparison.OrdinalIgnoreCase);

        var plannedQuery = _context.PLN_tbl.AsNoTracking()
            .Where(pl => resolvedIds.Contains(pl.PLN_Project_No) && pl.PLN_DATE != null);

        if (start.HasValue)
        {
            plannedQuery = plannedQuery.Where(pl => pl.PLN_DATE >= start.Value);
        }

        if (endExclusive.HasValue)
        {
            plannedQuery = plannedQuery.Where(pl => pl.PLN_DATE < endExclusive.Value);
        }

        if (filterLocation)
        {
            plannedQuery = plannedQuery.Where(pl => wantShop
                ? (pl.PLN_LOCATION != null && pl.PLN_LOCATION.Trim().ToUpper() == "WS")
                : !(pl.PLN_LOCATION != null && pl.PLN_LOCATION.Trim().ToUpper() == "WS"));
        }

        var plannedRaw = await plannedQuery
            .Select(pl => new { Date = pl.PLN_DATE!.Value, LocationRaw = pl.PLN_LOCATION, Dia = pl.PLN_DIA ?? 0 })
            .ToListAsync();

        var baseQuery =
            from p in _context.Projects_tbl.AsNoTracking()
            join dfr in _context.DFR_tbl.AsNoTracking() on p.Project_ID equals dfr.Project_No
            join dwr in _context.DWR_tbl.AsNoTracking() on dfr.Joint_ID equals dwr.Joint_ID_DWR
            where resolvedIds.Contains(p.Project_ID)
                  && (dfr.J_Add == null || !EF.Functions.Like(dfr.J_Add, "R%"))
            let weldDate = useFitupDate ? dfr.FITUP_DATE : dwr.ACTUAL_DATE_WELDED
            where weldDate != null
            select new
            {
                WeldDate = weldDate.Value,
                FitupDate = dfr.FITUP_DATE,
                ActualWeldDate = dwr.ACTUAL_DATE_WELDED,
                LocationRaw = dfr.LOCATION,
                Diameter = dfr.DIAMETER,
                WeldType = dfr.WELD_TYPE
            };

        if (start.HasValue)
        {
            baseQuery = baseQuery.Where(x => x.WeldDate >= start.Value);
        }

        if (endExclusive.HasValue)
        {
            baseQuery = baseQuery.Where(x => x.WeldDate < endExclusive.Value);
        }

        if (filterLocation)
        {
            baseQuery = baseQuery.Where(x => wantShop
                ? (x.LocationRaw != null && x.LocationRaw.Trim().ToUpper() == "WS")
                : !(x.LocationRaw != null && x.LocationRaw.Trim().ToUpper() == "WS"));
        }

        var raw = await baseQuery.ToListAsync();

        if (filterWeldType)
        {
            raw = raw
                .Where(x => x.WeldType != null && (
                    rawWtSet.Contains(TrimUpper(x.WeldType)) ||
                    wtSet.Contains(NormalizeWeldType(x.WeldType))))
                .ToList();
        }

        if (raw.Count == 0)
        {
            return Json(new JointProgressResponse());
        }

        static string MapLoc(string? loc)
        {
            var val = loc?.Trim();
            return (!string.IsNullOrWhiteSpace(val) && val.Equals("WS", StringComparison.OrdinalIgnoreCase)) ? "Shop" : "Field";
        }

        static bool IsThreaded(string? wt)
            => !string.IsNullOrWhiteSpace(wt) && wt.Trim().StartsWith("TH", StringComparison.OrdinalIgnoreCase);

        var actualKeys = raw
            .Select(x =>
            {
                var matchDate = useFitupDate
                    ? (x.FitupDate ?? x.WeldDate)
                    : (x.ActualWeldDate ?? x.WeldDate);
                return (Date: matchDate.Date, Loc: MapLoc(x.LocationRaw));
            })
            .Distinct()
            .ToHashSet();

        var plannedLookup = plannedRaw
            .Select(p => new { Key = (Date: p.Date.Date, Loc: MapLoc(p.LocationRaw)), p.Dia })
            .Where(x => actualKeys.Contains(x.Key))
            .GroupBy(x => x.Key)
            .ToDictionary(g => g.Key, g => g.Sum(x => x.Dia));

        var grouped = raw
            .GroupBy(x => (Date: x.WeldDate.Date, Loc: MapLoc(x.LocationRaw)))
            .Select(g =>
            {
                plannedLookup.TryGetValue(g.Key, out var planVal);
                return new JointProgressRowDto
                {
                    WeldDate = g.Key.Date,
                    Location = g.Key.Loc,
                    WeldedDiaIn = g.Sum(x => !IsThreaded(x.WeldType) ? (x.Diameter ?? 0) : 0),
                    ThreadedDiaIn = g.Sum(x => IsThreaded(x.WeldType) ? (x.Diameter ?? 0) : 0),
                    PlannedDiaIn = planVal
                };
            })
            .ToList();

        foreach (var kvp in plannedLookup)
        {
            bool exists = grouped.Any(r => r.WeldDate == kvp.Key.Date && r.Location == kvp.Key.Loc);
            if (!exists)
            {
                grouped.Add(new JointProgressRowDto
                {
                    WeldDate = kvp.Key.Date,
                    Location = kvp.Key.Loc,
                    PlannedDiaIn = kvp.Value
                });
            }
        }

        grouped = grouped
            .OrderBy(x => x.WeldDate)
            .ThenBy(x => x.Location)
            .ToList();

        double runningTotal = 0;
        foreach (var row in grouped)
        {
            runningTotal += row.TotalDiaIn;
            row.CumulativeDiaIn = runningTotal;
        }

        var dates = grouped.Select(x => x.WeldDate).Distinct().OrderBy(x => x).ToList();
        var chart = new List<JointProgressChartPoint>();
        double cumShop = 0, cumField = 0;

        foreach (var dt in dates)
        {
            var shopDay = grouped.Where(x => x.WeldDate == dt && x.Location == "Shop").Sum(x => x.TotalDiaIn);
            var fieldDay = grouped.Where(x => x.WeldDate == dt && x.Location == "Field").Sum(x => x.TotalDiaIn);
            var plannedDay = grouped.Where(x => x.WeldDate == dt).Sum(x => x.PlannedDiaIn);
            var dayTotal = shopDay + fieldDay;
            cumShop += shopDay;
            cumField += fieldDay;
            chart.Add(new JointProgressChartPoint
            {
                Date = dt.ToString("yyyy-MM-dd"),
                PlannedDailyTotal = Math.Round(plannedDay, 2),
                DailyTotal = Math.Round(dayTotal, 2),
                CumulativeShop = Math.Round(cumShop, 2),
                CumulativeField = Math.Round(cumField, 2),
                CumulativeTotal = Math.Round(cumShop + cumField, 2)
            });
        }

        var resp = new JointProgressResponse
        {
            Rows = grouped,
            Chart = chart,
            TotalWelded = grouped.Sum(x => x.WeldedDiaIn),
            TotalThreaded = grouped.Sum(x => x.ThreadedDiaIn),
            TotalPlanned = grouped.Sum(x => x.PlannedDiaIn)
        };
        resp.Total = resp.TotalWelded + resp.TotalThreaded;

        return Json(resp);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportJointProgressExcel(
        [FromQuery] int projectId,
        [FromQuery] List<int>? projectIds,
        [FromQuery] DateTime? startDate,
        [FromQuery] DateTime? endDate,
        [FromQuery] List<string>? weldTypes,
        [FromQuery] string? location = "All",
        [FromQuery] string? dateBasis = "Welding")
    {
        var resolvedIds = (projectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (resolvedIds.Count == 0 && projectId > 0) resolvedIds.Add(projectId);
        if (resolvedIds.Count == 0)
        {
            return BadRequest("Project required");
        }

        var dt = new DataTable();
        try
        {
            var conn = _context.Database.GetDbConnection();
            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "[PCA_Q]";
            cmd.CommandType = CommandType.StoredProcedure;
            await using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportJointProgressExcel: failed executing stored procedure");
            return Content("Failed to run stored procedure for joint export.");
        }

        var basis = (dateBasis ?? "Welding").Trim();
        bool useFitupDate = basis.Equals("Fit-up", StringComparison.OrdinalIgnoreCase);

        var rawWtSet = (weldTypes ?? new List<string>())
            .Select(w => (w ?? string.Empty).Trim().ToUpperInvariant())
            .Where(x => x.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var wtSet = (weldTypes ?? new List<string>())
            .Select(NormalizeWeldType)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        bool filterWeldType = rawWtSet.Count > 0 || wtSet.Count > 0;
        static string TrimUpper(string? s) => (s ?? string.Empty).Trim().ToUpperInvariant();

        var start = startDate?.Date;
        var endExclusive = endDate?.Date.AddDays(1);
        bool filterLocation = !string.IsNullOrWhiteSpace(location) && !string.Equals(location, "All", StringComparison.OrdinalIgnoreCase);
        bool wantShop = location != null && location.Equals("Shop", StringComparison.OrdinalIgnoreCase);

        static string MapLoc(string? loc)
        {
            var val = loc?.Trim();
            return (!string.IsNullOrWhiteSpace(val) && val.Equals("WS", StringComparison.OrdinalIgnoreCase)) ? "Shop" : "Field";
        }

        var filtered = dt.Clone();
        foreach (DataRow row in dt.Rows)
        {
            if (dt.Columns.Contains("ProjectNo"))
            {
                var raw = row["ProjectNo"]?.ToString();
                if (!int.TryParse(raw, out var rowPid) || !resolvedIds.Contains(rowPid))
                    continue;
            }

            var locRaw = dt.Columns.Contains("Location") ? row["Location"]?.ToString() : null;
            var mappedLoc = MapLoc(locRaw);
            if (filterLocation)
            {
                if (wantShop && !mappedLoc.Equals("Shop", StringComparison.OrdinalIgnoreCase)) continue;
                if (!wantShop && !mappedLoc.Equals("Field", StringComparison.OrdinalIgnoreCase)) continue;
            }

            if (filterWeldType)
            {
                var jt = dt.Columns.Contains("JointType") ? row["JointType"]?.ToString() : null;
                if (!(rawWtSet.Contains(TrimUpper(jt)) || wtSet.Contains(NormalizeWeldType(jt))))
                    continue;
            }

            var basisCol = useFitupDate ? "FitUpDate" : "DateOfWeld";
            DateTime? basisDate = null;
            if (dt.Columns.Contains(basisCol))
            {
                var val = row[basisCol];
                if (val != null && val != DBNull.Value && DateTime.TryParse(val.ToString(), out var dtVal))
                    basisDate = dtVal.Date;
            }

            if (start.HasValue && (!basisDate.HasValue || basisDate.Value < start.Value)) continue;
            if (endExclusive.HasValue && (!basisDate.HasValue || basisDate.Value >= endExclusive.Value)) continue;

            filtered.ImportRow(row);
        }

        dt = filtered;

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Fit-up & Welding");

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

            int rowIdx = 2;
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dr[c];
                    var cell = ws.Cell(rowIdx, c + 1);
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
                ws.Row(rowIdx).Height = 17;
                rowIdx++;
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
                if (name.Contains("id") || name.Contains("no") || name.Contains("number"))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                else if (type == typeof(int) || type == typeof(long) || type == typeof(short))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                else if (type == typeof(double) || type == typeof(decimal) || type == typeof(float))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"JointProgress_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportJointsRecordExcel(
        [FromQuery] int projectId,
        [FromQuery] List<int>? projectIds,
        [FromQuery] DateTime? startDate,
        [FromQuery] DateTime? endDate,
        [FromQuery] List<string>? weldTypes,
        [FromQuery] string? location = "All",
        [FromQuery] string? dateBasis = "Welding")
    {
        var resolvedIds = (projectIds ?? new List<int>()).Where(id => id > 0).ToList();
        if (resolvedIds.Count == 0 && projectId > 0) resolvedIds.Add(projectId);
        if (resolvedIds.Count == 0)
        {
            return BadRequest("Project required");
        }

        var dt = new DataTable();
        try
        {
            var conn = _context.Database.GetDbConnection();
            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "[2008_Q]";
            cmd.CommandType = CommandType.StoredProcedure;
            await using var reader = await cmd.ExecuteReaderAsync();
            dt.Load(reader);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ExportJointsRecordExcel: failed executing stored procedure");
            return Content("Failed to run stored procedure for joints record export.");
        }

        var basis = (dateBasis ?? "Welding").Trim();
        bool useFitupDate = basis.Equals("Fit-up", StringComparison.OrdinalIgnoreCase);

        var rawWtSet = (weldTypes ?? new List<string>())
            .Select(w => (w ?? string.Empty).Trim().ToUpperInvariant())
            .Where(x => x.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var wtSet = (weldTypes ?? new List<string>())
            .Select(NormalizeWeldType)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        bool filterWeldType = rawWtSet.Count > 0 || wtSet.Count > 0;
        static string TrimUpper(string? s) => (s ?? string.Empty).Trim().ToUpperInvariant();

        var start = startDate?.Date;
        var endExclusive = endDate?.Date.AddDays(1);
        bool filterLocation = !string.IsNullOrWhiteSpace(location) && !string.Equals(location, "All", StringComparison.OrdinalIgnoreCase);
        bool wantShop = location != null && location.Equals("Shop", StringComparison.OrdinalIgnoreCase);

        static string MapLoc(string? loc)
        {
            var val = loc?.Trim();
            return (!string.IsNullOrWhiteSpace(val) && val.Equals("WS", StringComparison.OrdinalIgnoreCase)) ? "Shop" : "Field";
        }

        string GetLocation(DataRow row)
        {
            string? locRaw = null;
            if (dt.Columns.Contains("LOCATION MARKERS"))
            {
                locRaw = row["LOCATION MARKERS"]?.ToString();
            }
            if (string.IsNullOrWhiteSpace(locRaw) && dt.Columns.Contains("WELD NO"))
            {
                var weldNo = row["WELD NO"]?.ToString();
                if (!string.IsNullOrWhiteSpace(weldNo))
                {
                    var dashIdx = weldNo.IndexOf('-');
                    locRaw = dashIdx >= 0 ? weldNo[..dashIdx] : weldNo.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                }
            }
            return MapLoc(locRaw);
        }

        var filtered = dt.Clone();
        foreach (DataRow row in dt.Rows)
        {
            if (dt.Columns.Contains("Project No"))
            {
                var raw = row["Project No"]?.ToString();
                if (!int.TryParse(raw, out var rowPid) || !resolvedIds.Contains(rowPid))
                    continue;
            }

            var mappedLoc = GetLocation(row);
            if (filterLocation)
            {
                if (wantShop && !mappedLoc.Equals("Shop", StringComparison.OrdinalIgnoreCase)) continue;
                if (!wantShop && !mappedLoc.Equals("Field", StringComparison.OrdinalIgnoreCase)) continue;
            }

            if (filterWeldType)
            {
                var jt = dt.Columns.Contains("WELD TYPE") ? row["WELD TYPE"]?.ToString() : null;
                if (!(rawWtSet.Contains(TrimUpper(jt)) || wtSet.Contains(NormalizeWeldType(jt))))
                    continue;
            }

            var basisCol = useFitupDate ? "FIT-UP DATE" : "DATE WELDED";
            DateTime? basisDate = null;
            if (dt.Columns.Contains(basisCol))
            {
                var val = row[basisCol];
                if (val != null && val != DBNull.Value && DateTime.TryParse(val.ToString(), out var dtVal))
                    basisDate = dtVal.Date;
            }

            if (start.HasValue && (!basisDate.HasValue || basisDate.Value < start.Value)) continue;
            if (endExclusive.HasValue && (!basisDate.HasValue || basisDate.Value >= endExclusive.Value)) continue;

            filtered.ImportRow(row);
        }

        dt = filtered;

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Joints Record");

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

            int rowIdx = 2;
            foreach (DataRow dr in dt.Rows)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dr[c];
                    var cell = ws.Cell(rowIdx, c + 1);
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
                ws.Row(rowIdx).Height = 17;
                rowIdx++;
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
                if (name.Contains("id") || name.Contains("no") || name.Contains("number"))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                else if (type == typeof(int) || type == typeof(long) || type == typeof(short))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                else if (type == typeof(double) || type == typeof(decimal) || type == typeof(float))
                    col.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var bytes = ms.ToArray();
        var fileName = $"JointsRecord_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    private async Task<(List<string> Options, List<string> DefaultOptions)> GetJointProgressWeldTypesAsync(int projectId)
    {
        var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);
        var weldTypeRows = await _context.PMS_Weld_Type_tbl
            .AsNoTracking()
            .Where(wt => wt.W_Project_No == weldersProjectId && wt.W_Weld_Type != null && wt.W_Weld_Type != "")
            .Select(wt => new { wt.W_Type_ID, wt.W_Weld_Type, wt.PROG_Default_Value })
            .ToListAsync();

        var weldTypeGroups = weldTypeRows
            .Select(wt => new
            {
                Name = (wt.W_Weld_Type ?? string.Empty).Trim(),
                wt.W_Type_ID,
                wt.PROG_Default_Value,
                Norm = NormalizeWeldType(wt.W_Weld_Type)
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Name))
            .GroupBy(x => x.Norm)
            .Select(g => new
            {
                Name = g.OrderBy(x => x.W_Type_ID).Select(x => x.Name).First(),
                MinId = g.Min(x => x.W_Type_ID),
                HasProgDefault = g.Any(x => x.PROG_Default_Value)
            })
            .OrderBy(x => x.MinId)
            .ToList();

        var weldTypes = weldTypeGroups.Select(x => x.Name).ToList();
        var defaultWeldTypes = weldTypeGroups.Where(x => x.HasProgDefault).Select(x => x.Name).ToList();

        return (weldTypes, defaultWeldTypes);
    }
}
