using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> WrrStatus([FromQuery] List<int>? projectId)
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

        double acceptableRate = 5.0;

        var jointRows = new List<WrrJointRow>();
        var linearRows = new List<WrrLinearRow>();

        if (selectedIds.Count > 0)
        {
            try
            {
                var lotProjectId = await _context.Projects_tbl.AsNoTracking()
                    .Where(p => p.Project_ID == selectedIds[0])
                    .Select(p => p.Welders_Project_ID ?? p.Project_ID)
                    .FirstOrDefaultAsync();

                if (lotProjectId <= 0) lotProjectId = selectedIds[0];

                var today = DateTime.Today;

                var lots = await _context.Lot_No_tbl.AsNoTracking()
                    .Where(l => l.Lot_Project_No == lotProjectId
                                && l.From_Date != null
                                && l.To_Date != null
                                && l.To_Date.Value.Date <= today)
                    .OrderBy(l => l.From_Date)
                    .ThenBy(l => l.Lot_No)
                    .ToListAsync();

                if (lots.Count > 0)
                {
                    var rawData = new List<WrrRawJoint>();

                    await using var conn = _context.Database.GetDbConnection();
                    if (conn.State != ConnectionState.Open)
                        await conn.OpenAsync();

                    await using var cmd = conn.CreateCommand();
                    var inParams = new List<string>();
                    for (int i = 0; i < selectedIds.Count; i++)
                    {
                        var paramName = $"@p{i}";
                        inParams.Add(paramName);
                        var prm = cmd.CreateParameter();
                        prm.ParameterName = paramName;
                        prm.DbType = DbType.Int32;
                        prm.Value = selectedIds[i];
                        cmd.Parameters.Add(prm);
                    }
                    cmd.CommandText = $@"
                        SELECT d.Joint_ID,
                               CAST(ISNULL(d.DIAMETER, 0) AS float)          AS DIAMETER,
                               COALESCE(rt.BSR_RT_DATE, rt.Final_RT_DATE)    AS RT_DATE,
                               CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(rt.BSR_RT_RESULT, ''))))
                                              IN ('REJ','REJECT','REJECTED','FAIL','FAILED')
                                    OR UPPER(LTRIM(RTRIM(ISNULL(rt.Final_RT_RESULT, ''))))
                                              IN ('REJ','REJECT','REJECTED','FAIL','FAILED')
                                    OR (rt.REPAIR_TYPE IS NOT NULL
                                        AND LTRIM(RTRIM(rt.REPAIR_TYPE)) <> '')
                                    OR ISNULL(rt.[LINEAR_REJECTED_LENGTH_(mm)], 0) > 0
                                    OR rt.RS_Date_1 IS NOT NULL
                               THEN CAST(1 AS bit) ELSE CAST(0 AS bit)
                               END                                           AS IsRejected,
                               CAST(ISNULL(rt.[LINEAR_REJECTED_LENGTH_(mm)], 0) AS float) AS LINEAR_REJECTED_LENGTH_mm
                        FROM DFR_tbl d
                        INNER JOIN RT_tbl rt ON rt.Joint_ID_RT = d.Joint_ID
                        WHERE d.Project_No IN ({string.Join(",", inParams)})
                          AND d.Deleted = 0
                          AND d.Cancelled = 0
                          AND COALESCE(rt.BSR_RT_DATE, rt.Final_RT_DATE) IS NOT NULL";
                    cmd.CommandTimeout = 300;

                    await using var reader = await cmd.ExecuteReaderAsync();
                    while (await reader.ReadAsync())
                    {
                        rawData.Add(new WrrRawJoint
                        {
                            JointId = reader.GetInt32(0),
                            Diameter = reader.GetDouble(1),
                            RtDate = reader.GetDateTime(2),
                            IsRejected = reader.GetBoolean(3),
                            LinearRejectedMm = reader.GetDouble(4)
                        });
                    }

                    int cummRtDone = 0, cummRtReject = 0;
                    double cummWeldLength = 0, cummDefectLength = 0;

                    foreach (var lot in lots)
                    {
                        var fromDate = lot.From_Date!.Value.Date;
                        var toDate = lot.To_Date!.Value.Date.AddDays(1);

                        var weekJoints = rawData
                            .Where(j => j.RtDate >= fromDate && j.RtDate < toDate)
                            .ToList();

                        int rtDone = weekJoints.Count;
                        int rtReject = weekJoints
                            .Count(j => j.IsRejected);

                        double weldLength = weekJoints
                            .Sum(j => Math.PI * j.Diameter);
                        double defectLength = weekJoints
                            .Where(j => j.IsRejected)
                            .Sum(j => j.LinearRejectedMm / 25.4);

                        cummRtDone += rtDone;
                        cummRtReject += rtReject;
                        cummWeldLength += weldLength;
                        cummDefectLength += defectLength;

                        double weeklyJointWrr = rtDone > 0
                            ? Math.Round((double)rtReject / rtDone * 100, 2)
                            : 0;
                        double cummJointWrr = cummRtDone > 0
                            ? Math.Round((double)cummRtReject / cummRtDone * 100, 4)
                            : 0;

                        double weeklyLinearWrr = weldLength > 0
                            ? Math.Round(defectLength / weldLength * 100, 3)
                            : 0;
                        double cummLinearWrr = cummWeldLength > 0
                            ? Math.Round(cummDefectLength / cummWeldLength * 100, 3)
                            : 0;

                        jointRows.Add(new WrrJointRow
                        {
                            Week = lot.Lot_No,
                            From = lot.From_Date!.Value.ToString("dd-MMM-yy"),
                            To = lot.To_Date!.Value.ToString("dd-MMM-yy"),
                            RtDone = rtDone,
                            RtReject = rtReject,
                            WeeklyWrr = weeklyJointWrr,
                            CummWrr = cummJointWrr,
                            AcceptableRate = acceptableRate
                        });

                        linearRows.Add(new WrrLinearRow
                        {
                            Week = lot.Lot_No,
                            From = lot.From_Date!.Value.ToString("dd-MMM-yy"),
                            To = lot.To_Date!.Value.ToString("dd-MMM-yy"),
                            WeldLengthRtd = Math.Round(weldLength, 2),
                            LengthWeldDefect = Math.Round(defectLength, 3),
                            WeeklyWrr = weeklyLinearWrr,
                            CummWrr = cummLinearWrr
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "WrrStatus failed for project(s) {ProjectIds}", string.Join(",", selectedIds));
                ViewBag.ErrorMessage = $"Failed to load WRR data: {ex.Message}";
            }
        }

        var vm = new WrrStatusPageViewModel
        {
            Projects = projects,
            SelectedProjectIds = selectedIds,
            ProjectAcceptableRate = acceptableRate,
            AsOfDate = DateTime.Now.ToString("dd-MMM-yyyy"),
            JointRows = jointRows,
            LinearRows = linearRows
        };

        return View(vm);
    }
}

internal sealed class WrrRawJoint
{
    public int JointId { get; init; }
    public double Diameter { get; init; }
    public DateTime RtDate { get; init; }
    public bool IsRejected { get; init; }
    public double LinearRejectedMm { get; init; }
}
