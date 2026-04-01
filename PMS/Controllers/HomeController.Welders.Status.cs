using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> WelderStatus([FromQuery] string? week, [FromQuery] List<int>? projectId)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrWhiteSpace(fullName)) return RedirectToAction("Login");
        ViewBag.FullName = fullName;

        // Build project dropdown: Welders_Project_ID - Project_Name (grouped/distinct, matches Control Lookups)
        var projects = await _context.Projects_tbl.AsNoTracking()
            .Where(p => p.Welders_Project_ID != null)
            .GroupBy(p => p.Welders_Project_ID)
            .Select(g => new ProjectOption
            {
                Id = g.Key!.Value,
                Name = g.OrderBy(p => p.Project_ID).First().Project_Name ?? string.Empty
            })
            .OrderBy(p => p.Id)
            .ToListAsync();

        var selectedIds = (projectId ?? []).Where(id => id > 0).ToList();
        if (selectedIds.Count == 0 && projects.Count > 0)
            selectedIds = [projects[0].Id];

        var today = DateTime.Today;
        var availableWeeks = await _context.Lot_No_tbl.AsNoTracking()
            .Where(l => !string.IsNullOrWhiteSpace(l.Lot_No) && l.From_Date <= today)
            .OrderBy(l => l.From_Date)
            .ThenBy(l => l.Lot_No)
            .Select(l => new LotOption { Id = l.Lot_ID, LotNo = l.Lot_No, From = l.From_Date, To = l.To_Date })
            .ToListAsync();

        var selectedWeek = (week ?? string.Empty).Trim();
        var shouldLoad = !string.IsNullOrWhiteSpace(selectedWeek);

        var dt = new DataTable();
        if (shouldLoad)
        {
            try
            {
                await using var conn = _context.Database.GetDbConnection();
                if (conn.State != ConnectionState.Open)
                    await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = "[dbo].[Welders_Categories_Q]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 300;

                var weekParam = cmd.CreateParameter();
                weekParam.ParameterName = "@WEEK";
                weekParam.DbType = DbType.String;
                weekParam.Value = selectedWeek;
                cmd.Parameters.Add(weekParam);

                await using var reader = await cmd.ExecuteReaderAsync();
                dt.Load(reader);

                // Filter by selected projects if a recognisable project column exists
                if (selectedIds.Count > 0 && dt.Rows.Count > 0)
                {
                    DataColumn? projCol = null;
                    foreach (DataColumn col in dt.Columns)
                    {
                        var name = col.ColumnName.Replace("_", "").Replace(" ", "").ToUpperInvariant();
                        if (name is "PROJECTNO" or "PROJECTID" or "PROJECTNUMBER" or "WELDERSPROJECTID")
                        {
                            projCol = col;
                            break;
                        }
                    }

                    if (projCol != null)
                    {
                        var allowed = new HashSet<string>(selectedIds.Select(id => id.ToString()), StringComparer.OrdinalIgnoreCase);
                        for (int i = dt.Rows.Count - 1; i >= 0; i--)
                        {
                            var val = dt.Rows[i][projCol]?.ToString()?.Trim() ?? string.Empty;
                            if (!allowed.Contains(val))
                                dt.Rows[i].Delete();
                        }
                        dt.AcceptChanges();
                    }
                }
            }
            catch (SqlException ex) when (ex.Number == -2)
            {
                _logger.LogWarning(ex, "WelderStatus timed out for week {Week}", selectedWeek);
                ViewBag.ErrorMessage = "Loading welder status took too long. Please retry with another week.";
            }
        }

        var vm = new WelderStatusPageViewModel
        {
            Week = string.IsNullOrWhiteSpace(selectedWeek) ? null : selectedWeek,
            AvailableWeeks = availableWeeks,
            SelectedProjectIds = selectedIds,
            Projects = projects,
            Results = dt
        };

        return View(vm);
    }
}
