using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
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
    public async Task<IActionResult> ExportWelderStatusExcel([FromQuery] string? week, [FromQuery] List<int>? projectId)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrWhiteSpace(fullName)) return RedirectToAction("Login");

        var selectedIds = (projectId ?? []).Where(id => id > 0).ToList();
        var selectedWeek = (week ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(selectedWeek))
            return RedirectToAction(nameof(WelderStatus), new { week, projectId });

        var dt = new DataTable();
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
            _logger.LogWarning(ex, "ExportWelderStatusExcel timed out for week {Week}", selectedWeek);
            return RedirectToAction(nameof(WelderStatus), new { week, projectId });
        }

        if (dt.Rows.Count == 0)
            return RedirectToAction(nameof(WelderStatus), new { week, projectId });

        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Welders Status");

        // Headers
        for (int c = 0; c < dt.Columns.Count; c++)
        {
            ws.Cell(1, c + 1).SetValue(dt.Columns[c].ColumnName);
        }

        // Data rows with typed cell values (matching Welder Register style)
        int row = 2;
        foreach (DataRow dr in dt.Rows)
        {
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var val = dr[c];
                var cell = ws.Cell(row, c + 1);
                if (val == DBNull.Value || val == null)
                {
                    cell.SetValue(string.Empty);
                }
                else if (val is DateTime dtVal)
                {
                    cell.SetValue(dtVal);
                    cell.Style.DateFormat.Format = "yyyy-MM-dd";
                }
                else if (val is int iv)
                {
                    cell.SetValue(iv);
                }
                else if (val is long lv)
                {
                    cell.SetValue(lv);
                }
                else if (val is short sv)
                {
                    cell.SetValue((int)sv);
                }
                else if (val is byte bv)
                {
                    cell.SetValue((int)bv);
                }
                else if (val is double dv)
                {
                    cell.SetValue(dv);
                }
                else if (val is float fv)
                {
                    cell.SetValue((double)fv);
                }
                else if (val is decimal decv)
                {
                    cell.SetValue((double)decv);
                }
                else if (val is bool blv)
                {
                    cell.SetValue(blv);
                }
                else
                {
                    cell.SetValue(val.ToString() ?? string.Empty);
                }
            }
            ws.Row(row).Height = 17;
            row++;
        }

        // Table styling (matching Welder Register)
        int lastRow = dt.Rows.Count + 1;
        if (dt.Rows.Count > 0)
        {
            var fullRange = ws.Range(1, 1, lastRow, dt.Columns.Count);
            var table = fullRange.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowTotalsRow = false;
        }
        ws.Row(1).Height = 30;

        // Column alignment based on type
        for (int c = 0; c < dt.Columns.Count; c++)
        {
            var name = dt.Columns[c].ColumnName.ToLowerInvariant();
            var type = dt.Columns[c].DataType;
            if (name.Contains("id") || name.Contains("no") || name.Contains("batch") || name.Contains("project"))
                ws.Column(c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            else if (type == typeof(int) || type == typeof(long))
                ws.Column(c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            else if (type == typeof(double) || type == typeof(decimal) || type == typeof(float))
                ws.Column(c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        var fileName = $"Welders_Status_{selectedWeek}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(stream.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }
}
