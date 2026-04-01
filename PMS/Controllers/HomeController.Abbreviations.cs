using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;

namespace PMS.Controllers;

public partial class HomeController
{
    // ══════════════════════════════════════════════════════════════
    //  GET  /Home/Abbreviations?projectNo=123
    //  Read-only reference view of all Control Lookup tables
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> Abbreviations([FromQuery] int? projectNo)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");
        ViewBag.FullName = fullName;

        // Build project dropdown (same logic as ControlLookups)
        var projects = await _context.Projects_tbl.AsNoTracking()
            .Where(p => p.Welders_Project_ID != null)
            .GroupBy(p => p.Welders_Project_ID)
            .Select(g => new
            {
                Welders_Project_ID = g.Key,
                Project_Name = g.OrderBy(p => p.Project_ID).First().Project_Name
            })
            .OrderBy(p => p.Welders_Project_ID)
            .ToListAsync();

        ViewBag.Projects = projects
            .Select(p => new { Value = p.Welders_Project_ID!.Value, Text = $"{p.Welders_Project_ID} - {p.Project_Name}" })
            .ToList();

        var selectedNo = projectNo
            ?? projects.FirstOrDefault()?.Welders_Project_ID
            ?? 0;

        ViewBag.IpTList = await _context.PMS_IP_T_tbl.AsNoTracking()
            .Where(r => r.IP_T_Project_No == selectedNo)
            .OrderBy(r => r.IP_T_ID).ToListAsync();

        ViewBag.JAddList = await _context.PMS_J_Add_tbl.AsNoTracking()
            .Where(r => r.Add_Project_No == selectedNo)
            .OrderBy(r => r.Add_ID).ToListAsync();

        ViewBag.LocationList = await _context.PMS_Location_tbl.AsNoTracking()
            .Where(r => r.LO_Project_No == selectedNo)
            .OrderBy(r => r.LO_ID).ToListAsync();

        ViewBag.WeldTypeList = await _context.PMS_Weld_Type_tbl.AsNoTracking()
            .Where(r => r.W_Project_No == selectedNo)
            .OrderBy(r => r.W_Type_ID).ToListAsync();

        ViewBag.ProjectNo = selectedNo;
        return View();
    }

    // ══════════════════════════════════════════════════════════════
    //  GET  /Home/ExportAbbreviationsExcel?projectNo=123
    //  Exports all four abbreviation tables to a single Excel file
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportAbbreviationsExcel([FromQuery] int? projectNo)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");

        var selectedNo = projectNo ?? 0;

        var ipTList = await _context.PMS_IP_T_tbl.AsNoTracking()
            .Where(r => r.IP_T_Project_No == selectedNo)
            .OrderBy(r => r.IP_T_ID).ToListAsync();

        var jAddList = await _context.PMS_J_Add_tbl.AsNoTracking()
            .Where(r => r.Add_Project_No == selectedNo)
            .OrderBy(r => r.Add_ID).ToListAsync();

        var locationList = await _context.PMS_Location_tbl.AsNoTracking()
            .Where(r => r.LO_Project_No == selectedNo)
            .OrderBy(r => r.LO_ID).ToListAsync();

        var weldTypeList = await _context.PMS_Weld_Type_tbl.AsNoTracking()
            .Where(r => r.W_Project_No == selectedNo)
            .OrderBy(r => r.W_Type_ID).ToListAsync();

        using var workbook = new XLWorkbook();

        // Sheet 1 – IP/T Type
        var wsIpT = workbook.Worksheets.Add("IP-T Type");
        wsIpT.Cell(1, 1).Value = "ID";
        wsIpT.Cell(1, 2).Value = "IP/T Type";
        wsIpT.Cell(1, 3).Value = "Project IP/T Type";
        wsIpT.Cell(1, 4).Value = "Notes";
        wsIpT.Cell(1, 5).Value = "Project No";
        for (int i = 0; i < ipTList.Count; i++)
        {
            wsIpT.Cell(i + 2, 1).Value = ipTList[i].IP_T_ID;
            wsIpT.Cell(i + 2, 2).Value = ipTList[i].IP_T_List ?? "";
            wsIpT.Cell(i + 2, 3).Value = ipTList[i].P_IP_T_List ?? "";
            wsIpT.Cell(i + 2, 4).Value = ipTList[i].IP_T_Notes ?? "";
            wsIpT.Cell(i + 2, 5).Value = ipTList[i].IP_T_Project_No;
        }
        wsIpT.RangeUsed()?.SetAutoFilter();
        wsIpT.Columns().AdjustToContents();

        // Sheet 2 – Joint Addition
        var wsJAdd = workbook.Worksheets.Add("Joint Addition");
        wsJAdd.Cell(1, 1).Value = "ID";
        wsJAdd.Cell(1, 2).Value = "Joint Addition";
        wsJAdd.Cell(1, 3).Value = "Project Joint Addition";
        wsJAdd.Cell(1, 4).Value = "Notes";
        wsJAdd.Cell(1, 5).Value = "Project No";
        for (int i = 0; i < jAddList.Count; i++)
        {
            wsJAdd.Cell(i + 2, 1).Value = jAddList[i].Add_ID;
            wsJAdd.Cell(i + 2, 2).Value = jAddList[i].Add_J_Add ?? "";
            wsJAdd.Cell(i + 2, 3).Value = jAddList[i].P_J_Add ?? "";
            wsJAdd.Cell(i + 2, 4).Value = jAddList[i].Add_Notes ?? "";
            wsJAdd.Cell(i + 2, 5).Value = jAddList[i].Add_Project_No;
        }
        wsJAdd.RangeUsed()?.SetAutoFilter();
        wsJAdd.Columns().AdjustToContents();

        // Sheet 3 – Location
        var wsLoc = workbook.Worksheets.Add("Location");
        wsLoc.Cell(1, 1).Value = "ID";
        wsLoc.Cell(1, 2).Value = "Location";
        wsLoc.Cell(1, 3).Value = "Project Location";
        wsLoc.Cell(1, 4).Value = "Notes";
        wsLoc.Cell(1, 5).Value = "Project No";
        for (int i = 0; i < locationList.Count; i++)
        {
            wsLoc.Cell(i + 2, 1).Value = locationList[i].LO_ID;
            wsLoc.Cell(i + 2, 2).Value = locationList[i].LO_Location ?? "";
            wsLoc.Cell(i + 2, 3).Value = locationList[i].P_Location ?? "";
            wsLoc.Cell(i + 2, 4).Value = locationList[i].LO_Notes ?? "";
            wsLoc.Cell(i + 2, 5).Value = locationList[i].LO_Project_No;
        }
        wsLoc.RangeUsed()?.SetAutoFilter();
        wsLoc.Columns().AdjustToContents();

        // Sheet 4 – Weld Type
        var wsWeld = workbook.Worksheets.Add("Weld Type");
        wsWeld.Cell(1, 1).Value = "ID";
        wsWeld.Cell(1, 2).Value = "Weld Type";
        wsWeld.Cell(1, 3).Value = "Project Weld Type";
        wsWeld.Cell(1, 4).Value = "Default";
        wsWeld.Cell(1, 5).Value = "PROG Default";
        wsWeld.Cell(1, 6).Value = "Notes";
        wsWeld.Cell(1, 7).Value = "Project No";
        for (int i = 0; i < weldTypeList.Count; i++)
        {
            wsWeld.Cell(i + 2, 1).Value = weldTypeList[i].W_Type_ID;
            wsWeld.Cell(i + 2, 2).Value = weldTypeList[i].W_Weld_Type ?? "";
            wsWeld.Cell(i + 2, 3).Value = weldTypeList[i].P_Weld_Type ?? "";
            wsWeld.Cell(i + 2, 4).Value = weldTypeList[i].Default_Value ? "Yes" : "No";
            wsWeld.Cell(i + 2, 5).Value = weldTypeList[i].PROG_Default_Value ? "Yes" : "No";
            wsWeld.Cell(i + 2, 6).Value = weldTypeList[i].W_Notes ?? "";
            wsWeld.Cell(i + 2, 7).Value = weldTypeList[i].W_Project_No;
        }
        wsWeld.RangeUsed()?.SetAutoFilter();
        wsWeld.Columns().AdjustToContents();

        // Style headers on all sheets
        foreach (var ws in workbook.Worksheets)
        {
            var headerRow = ws.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.FromHtml("#E6F6FA");
            headerRow.Style.Font.FontColor = XLColor.FromHtml("#176D8A");
        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        var fileName = $"Abbreviations_Project_{selectedNo}.xlsx";
        return File(
            stream.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }
}
