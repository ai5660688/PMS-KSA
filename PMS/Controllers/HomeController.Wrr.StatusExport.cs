using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpPost]
    [RequestFormLimits(ValueLengthLimit = 10_000_000)]
    public async Task<IActionResult> ExportWrrStatusExcel(
        [FromForm] List<int>? projectId,
        [FromForm] string? jointChartImage,
        [FromForm] string? linearChartImage)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrWhiteSpace(fullName)) return RedirectToAction("Login");

        var selectedIds = (projectId ?? []).Where(id => id > 0).ToList();
        if (selectedIds.Count == 0)
            return RedirectToAction(nameof(WrrStatus));

        double acceptableRate = 5.0;
        var jointRows = new List<WrrJointRow>();
        var linearRows = new List<WrrLinearRow>();

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
                    int rtReject = weekJoints.Count(j => j.IsRejected);

                    double weldLength = weekJoints.Sum(j => Math.PI * j.Diameter);
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
            _logger.LogError(ex, "ExportWrrStatusExcel failed for project(s) {ProjectIds}", string.Join(",", selectedIds));
            return RedirectToAction(nameof(WrrStatus), new { projectId });
        }

        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("WRR Status");

        var jointHeaders = new[] { "WEEK", "From", "To", "RT DONE", "RT REJECT",
            "WEEKLY WRR% JOINT METHOD", "CUMM WRR% JOINT METHOD", "Project Acceptable Rate" };
        int linearCol = jointHeaders.Length + 2; // column J (skip column I as gap)
        var linearHeaders = new[] { "WEEK", "From", "To", "WELD LENGTH RT'd (Inches)",
            "Length Weld Defect (Inches)", "WEEKLY WRR% LINEAR METHOD", "CUMM WRR% LINEAR METHOD" };

        // ── Chart images at the top (side-by-side, row 1) ──
        const int chartHeight = 300;
        const int chartWidth = 700;
        int chartRowCount = 16; // approximate row span for a 300px chart

        if (!string.IsNullOrWhiteSpace(jointChartImage))
        {
            try
            {
                var bytes = Convert.FromBase64String(
                    jointChartImage.Contains(',')
                        ? jointChartImage[(jointChartImage.IndexOf(',') + 1)..]
                        : jointChartImage);
                using var imgStream = new MemoryStream(bytes);
                var pic = ws.AddPicture(imgStream, XLPictureFormat.Png, "JointChart");
                pic.MoveTo(ws.Cell(1, 1));
                pic.WithSize(chartWidth, chartHeight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to embed joint chart image in Excel");
            }
        }

        if (!string.IsNullOrWhiteSpace(linearChartImage))
        {
            try
            {
                var bytes = Convert.FromBase64String(
                    linearChartImage.Contains(',')
                        ? linearChartImage[(linearChartImage.IndexOf(',') + 1)..]
                        : linearChartImage);
                using var imgStream = new MemoryStream(bytes);
                var pic = ws.AddPicture(imgStream, XLPictureFormat.Png, "LinearChart");
                pic.MoveTo(ws.Cell(1, linearCol));
                pic.WithSize(chartWidth, chartHeight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to embed linear chart image in Excel");
            }
        }

        // ── Tables start below the charts ──
        int titleRow = chartRowCount + 1;

        // Joint Basis title
        var jtTitleCell = ws.Cell(titleRow, 1);
        jtTitleCell.SetValue("WRR STATUS - Joint Basis (Weekly)");
        jtTitleCell.Style.Font.Bold = true;
        jtTitleCell.Style.Font.FontSize = 13;
        jtTitleCell.Style.Font.FontColor = XLColor.FromHtml("#176d8a");
        ws.Range(titleRow, 1, titleRow, jointHeaders.Length).Merge();

        // Linear Basis title
        var lrTitleCell = ws.Cell(titleRow, linearCol);
        lrTitleCell.SetValue("WRR STATUS - Linear Basis (Weekly)");
        lrTitleCell.Style.Font.Bold = true;
        lrTitleCell.Style.Font.FontSize = 13;
        lrTitleCell.Style.Font.FontColor = XLColor.FromHtml("#176d8a");
        ws.Range(titleRow, linearCol, titleRow, linearCol + linearHeaders.Length - 1).Merge();

        ws.Row(titleRow).Height = 26;

        int headerRow = titleRow + 1;

        // ── Joint Basis table ──
        for (int c = 0; c < jointHeaders.Length; c++)
            ws.Cell(headerRow, c + 1).SetValue(jointHeaders[c]);

        for (int r = 0; r < jointRows.Count; r++)
        {
            var jr = jointRows[r];
            int row = headerRow + r + 1;
            ws.Cell(row, 1).SetValue(jr.Week);
            ws.Cell(row, 2).SetValue(jr.From);
            ws.Cell(row, 3).SetValue(jr.To);
            ws.Cell(row, 4).SetValue(jr.RtDone);
            ws.Cell(row, 5).SetValue(jr.RtReject);
            ws.Cell(row, 6).SetValue(jr.WeeklyWrr);
            ws.Cell(row, 7).SetValue(jr.CummWrr);
            ws.Cell(row, 8).SetValue(jr.AcceptableRate);
            ws.Row(row).Height = 17;
        }

        int jointLastRow = headerRow + jointRows.Count;
        if (jointRows.Count > 0)
        {
            var jtRange = ws.Range(headerRow, 1, jointLastRow, jointHeaders.Length);
            var jtTable = jtRange.CreateTable("JointBasis");
            jtTable.Theme = XLTableTheme.TableStyleMedium2;
            jtTable.ShowTotalsRow = false;
        }
        ws.Row(headerRow).Height = 30;

        for (int c = 4; c <= jointHeaders.Length; c++)
            ws.Column(c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // ── Linear Basis table (beside joint table) ──
        for (int c = 0; c < linearHeaders.Length; c++)
            ws.Cell(headerRow, linearCol + c).SetValue(linearHeaders[c]);

        for (int r = 0; r < linearRows.Count; r++)
        {
            var lr = linearRows[r];
            int row = headerRow + r + 1;
            ws.Cell(row, linearCol).SetValue(lr.Week);
            ws.Cell(row, linearCol + 1).SetValue(lr.From);
            ws.Cell(row, linearCol + 2).SetValue(lr.To);
            ws.Cell(row, linearCol + 3).SetValue(lr.WeldLengthRtd);
            ws.Cell(row, linearCol + 4).SetValue(lr.LengthWeldDefect);
            ws.Cell(row, linearCol + 5).SetValue(lr.WeeklyWrr);
            ws.Cell(row, linearCol + 6).SetValue(lr.CummWrr);
            ws.Row(row).Height = 17;
        }

        int linearLastRow = headerRow + linearRows.Count;
        if (linearRows.Count > 0)
        {
            var lrRange = ws.Range(headerRow, linearCol, linearLastRow, linearCol + linearHeaders.Length - 1);
            var lrTable = lrRange.CreateTable("LinearBasis");
            lrTable.Theme = XLTableTheme.TableStyleMedium2;
            lrTable.ShowTotalsRow = false;
        }

        for (int c = 3; c < linearHeaders.Length; c++)
            ws.Column(linearCol + c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(headerRow);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);

        var fileName = $"WRR_Status_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(stream.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }
}
