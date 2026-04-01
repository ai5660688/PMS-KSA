using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    // ══════════════════════════════════════════════════════════════
    //  GET  /Home/ControlLookups?projectNo=123
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ControlLookups([FromQuery] int? projectNo)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");

        // Build the project dropdown: Welders_Project_ID - Project_Name (grouped/distinct)
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

        // Default to first project if none selected
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
    //  IP/T Type  (PMS_IP_T_tbl)
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveIpT(PmsIpT input)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var existing = await _context.PMS_IP_T_tbl.FirstOrDefaultAsync(r => r.IP_T_ID == input.IP_T_ID);
            if (existing == null)
            {
                _context.PMS_IP_T_tbl.Add(input);
            }
            else
            {
                existing.IP_T_List = input.IP_T_List;
                    existing.P_IP_T_List = input.P_IP_T_List;
                    existing.IP_T_Notes = input.IP_T_Notes;
                    existing.IP_T_Project_No = input.IP_T_Project_No;
            }
            await _context.SaveChangesAsync();
            TempData["Msg"] = "IP/T Type saved successfully.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveIpT failed");
            TempData["Msg"] = "Failed to save IP/T Type.";
        }
        return RedirectToAction(nameof(ControlLookups), new { projectNo = input.IP_T_Project_No });
    }

    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteIpT([FromForm] int id)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var row = await _context.PMS_IP_T_tbl.FirstOrDefaultAsync(r => r.IP_T_ID == id);
            var pn = row?.IP_T_Project_No ?? 0;
            if (row != null) { _context.PMS_IP_T_tbl.Remove(row); await _context.SaveChangesAsync(); TempData["Msg"] = "IP/T Type deleted."; }
            return RedirectToAction(nameof(ControlLookups), new { projectNo = pn });
        }
        catch (Exception ex) { _logger.LogError(ex, "DeleteIpT failed"); TempData["Msg"] = "Failed to delete IP/T Type."; }
        return RedirectToAction(nameof(ControlLookups));
    }

    // ══════════════════════════════════════════════════════════════
    //  Joint Addition  (PMS_J_Add_tbl)
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveJAdd(PmsJAdd input)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var existing = await _context.PMS_J_Add_tbl.FirstOrDefaultAsync(r => r.Add_ID == input.Add_ID);
            if (existing == null)
            {
                _context.PMS_J_Add_tbl.Add(input);
            }
            else
            {
                existing.Add_J_Add = input.Add_J_Add;
                existing.P_J_Add = input.P_J_Add;
                existing.Add_Notes = input.Add_Notes;
                existing.Add_Project_No = input.Add_Project_No;
            }
            await _context.SaveChangesAsync();
            TempData["Msg"] = "Joint Addition saved successfully.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveJAdd failed");
            TempData["Msg"] = "Failed to save Joint Addition.";
        }
        return RedirectToAction(nameof(ControlLookups), new { projectNo = input.Add_Project_No });
    }

    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteJAdd([FromForm] int id)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var row = await _context.PMS_J_Add_tbl.FirstOrDefaultAsync(r => r.Add_ID == id);
            var pn = row?.Add_Project_No ?? 0;
            if (row != null) { _context.PMS_J_Add_tbl.Remove(row); await _context.SaveChangesAsync(); TempData["Msg"] = "Joint Addition deleted."; }
            return RedirectToAction(nameof(ControlLookups), new { projectNo = pn });
        }
        catch (Exception ex) { _logger.LogError(ex, "DeleteJAdd failed"); TempData["Msg"] = "Failed to delete Joint Addition."; }
        return RedirectToAction(nameof(ControlLookups));
    }

    // ══════════════════════════════════════════════════════════════
    //  Location  (PMS_Location_tbl)
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveLocation(PmsLocation input)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var existing = await _context.PMS_Location_tbl.FirstOrDefaultAsync(r => r.LO_ID == input.LO_ID);
            if (existing == null)
            {
                _context.PMS_Location_tbl.Add(input);
            }
            else
            {
                existing.LO_Location = input.LO_Location;
                existing.P_Location = input.P_Location;
                existing.LO_Notes = input.LO_Notes;
                existing.LO_Project_No = input.LO_Project_No;
            }
            await _context.SaveChangesAsync();
            TempData["Msg"] = "Location saved successfully.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveLocation failed");
            TempData["Msg"] = "Failed to save Location.";
        }
        return RedirectToAction(nameof(ControlLookups), new { projectNo = input.LO_Project_No });
    }

    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteLocation([FromForm] int id)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var row = await _context.PMS_Location_tbl.FirstOrDefaultAsync(r => r.LO_ID == id);
            var pn = row?.LO_Project_No ?? 0;
            if (row != null) { _context.PMS_Location_tbl.Remove(row); await _context.SaveChangesAsync(); TempData["Msg"] = "Location deleted."; }
            return RedirectToAction(nameof(ControlLookups), new { projectNo = pn });
        }
        catch (Exception ex) { _logger.LogError(ex, "DeleteLocation failed"); TempData["Msg"] = "Failed to delete Location."; }
        return RedirectToAction(nameof(ControlLookups));
    }

    // ══════════════════════════════════════════════════════════════
    //  Weld Type  (PMS_Weld_Type_tbl)
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveWeldType(PmsWeldType input)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var existing = await _context.PMS_Weld_Type_tbl.FirstOrDefaultAsync(r => r.W_Type_ID == input.W_Type_ID);
            if (existing == null)
            {
                _context.PMS_Weld_Type_tbl.Add(input);
            }
            else
            {
                existing.W_Weld_Type = input.W_Weld_Type;
                existing.P_Weld_Type = input.P_Weld_Type;
                existing.Default_Value = input.Default_Value;
                existing.PROG_Default_Value = input.PROG_Default_Value;
                existing.W_Notes = input.W_Notes;
                existing.W_Project_No = input.W_Project_No;
            }
            await _context.SaveChangesAsync();
            TempData["Msg"] = "Weld Type saved successfully.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveWeldType failed");
            TempData["Msg"] = "Failed to save Weld Type.";
        }
        return RedirectToAction(nameof(ControlLookups), new { projectNo = input.W_Project_No });
    }

    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteWeldType([FromForm] int id)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");
        try
        {
            var row = await _context.PMS_Weld_Type_tbl.FirstOrDefaultAsync(r => r.W_Type_ID == id);
            var pn = row?.W_Project_No ?? 0;
            if (row != null) { _context.PMS_Weld_Type_tbl.Remove(row); await _context.SaveChangesAsync(); TempData["Msg"] = "Weld Type deleted."; }
            return RedirectToAction(nameof(ControlLookups), new { projectNo = pn });
        }
        catch (Exception ex) { _logger.LogError(ex, "DeleteWeldType failed"); TempData["Msg"] = "Failed to delete Weld Type."; }
        return RedirectToAction(nameof(ControlLookups));
    }

    // ══════════════════════════════════════════════════════════════
    //  Copy all lookup rows from one project to another
    // ══════════════════════════════════════════════════════════════
    [SessionAuthorization, HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> CopyLookupsToProject([FromForm] int sourceProjectNo, [FromForm] int targetProjectNo)
    {
        if (!IsAdmin()) return RedirectToAction("Dashboard");

        if (sourceProjectNo == targetProjectNo)
        {
            TempData["Msg"] = "Source and target project cannot be the same.";
            return RedirectToAction(nameof(ControlLookups), new { projectNo = sourceProjectNo });
        }

        try
        {
            var copied = 0;

            // IP/T Type
            var ipRows = await _context.PMS_IP_T_tbl.AsNoTracking()
                .Where(r => r.IP_T_Project_No == sourceProjectNo).ToListAsync();
            foreach (var r in ipRows)
            {
                _context.PMS_IP_T_tbl.Add(new PmsIpT
                {
                    IP_T_List = r.IP_T_List,
                    P_IP_T_List = r.P_IP_T_List,
                    IP_T_Notes = r.IP_T_Notes,
                    IP_T_Project_No = targetProjectNo
                });
                copied++;
            }

            // Joint Addition
            var jAddRows = await _context.PMS_J_Add_tbl.AsNoTracking()
                .Where(r => r.Add_Project_No == sourceProjectNo).ToListAsync();
            foreach (var r in jAddRows)
            {
                _context.PMS_J_Add_tbl.Add(new PmsJAdd
                {
                    Add_J_Add = r.Add_J_Add,
                    P_J_Add = r.P_J_Add,
                    Add_Notes = r.Add_Notes,
                    Add_Project_No = targetProjectNo
                });
                copied++;
            }

            // Location
            var locRows = await _context.PMS_Location_tbl.AsNoTracking()
                .Where(r => r.LO_Project_No == sourceProjectNo).ToListAsync();
            foreach (var r in locRows)
            {
                _context.PMS_Location_tbl.Add(new PmsLocation
                {
                    LO_Location = r.LO_Location,
                    P_Location = r.P_Location,
                    LO_Notes = r.LO_Notes,
                    LO_Project_No = targetProjectNo
                });
                copied++;
            }

            // Weld Type
            var wtRows = await _context.PMS_Weld_Type_tbl.AsNoTracking()
                .Where(r => r.W_Project_No == sourceProjectNo).ToListAsync();
            foreach (var r in wtRows)
            {
                _context.PMS_Weld_Type_tbl.Add(new PmsWeldType
                {
                    W_Weld_Type = r.W_Weld_Type,
                    P_Weld_Type = r.P_Weld_Type,
                    Default_Value = r.Default_Value,
                    PROG_Default_Value = r.PROG_Default_Value,
                    W_Notes = r.W_Notes,
                    W_Project_No = targetProjectNo
                });
                copied++;
            }

            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Successfully copied {copied} rows to project {targetProjectNo}.";
            return RedirectToAction(nameof(ControlLookups), new { projectNo = targetProjectNo });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CopyLookupsToProject failed from {Src} to {Tgt}", sourceProjectNo, targetProjectNo);
            TempData["Msg"] = "Failed to copy lookups.";
        }
        return RedirectToAction(nameof(ControlLookups), new { projectNo = sourceProjectNo });
    }
}
