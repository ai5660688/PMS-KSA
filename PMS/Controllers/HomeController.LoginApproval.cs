using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Models;
using PMS.Data;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;
using System;
using System.Linq;
using ClosedXML.Excel; // added for Excel export
using System.IO; // for MemoryStream

namespace PMS.Controllers;

public partial class HomeController : Controller
{
    private bool IsAdmin()
        => string.Equals(HttpContext.Session.GetString("Access"), "ADMIN", StringComparison.OrdinalIgnoreCase);

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> LoginApproval()
    {
        try
        {
            if (!IsAdmin())
                return RedirectToAction(nameof(Dashboard));

            var fullName = HttpContext.Session.GetString("FullName") ?? string.Empty;
            ViewBag.FullName = fullName;

            var users = await _context.PMS_Login_tbl
                .AsNoTracking()
                .OrderBy(u => u.Approved)
                .ThenBy(u => u.FirstName)
                .ThenBy(u => u.LastName)
                .ToListAsync();

            return View(users); // View name matches LoginApproval.cshtml
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading LoginApproval page");
            return RedirectToAction(nameof(Dashboard));
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApproveUser(int userId)
    {
        if (!IsAdmin())
            return Json(new { success = false, message = "Not authorized" });

        try
        {
            var user = await _context.PMS_Login_tbl.FirstOrDefaultAsync(u => u.UserID == userId);
            if (user == null)
                return Json(new { success = false, message = "User not found" });

            if (!user.Approved)
            {
                user.Approved = true;
                user.Login_Approval_Date = PMS.Infrastructure.AppClock.Now; // localized UTC+03
                user.Login_Approval_By = HttpContext.Session.GetInt32("UserID");
                await _context.SaveChangesAsync();
            }
            return Json(new { success = true });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error approving user {UserId}", userId);
            return Json(new { success = false, message = "Error" });
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UpdateUser(int UserID, string FirstName, string LastName, string Email, string Position, string Company, string Access, bool Approved, DateTime? demobilisedDate, string? Password)
    {
        if (!IsAdmin())
            return RedirectToAction(nameof(LoginApproval));

        try
        {
            var user = await _context.PMS_Login_tbl.FirstOrDefaultAsync(u => u.UserID == UserID);
            if (user == null)
                return RedirectToAction(nameof(LoginApproval));

            user.FirstName = (FirstName ?? string.Empty).Trim();
            user.LastName = (LastName ?? string.Empty).Trim();
            user.Email = string.IsNullOrWhiteSpace(Email) ? null : Email.Trim();
            user.Position = string.IsNullOrWhiteSpace(Position) ? null : Position.Trim();
            user.Company = string.IsNullOrWhiteSpace(Company) ? null : Company.Trim();
            user.Access = string.IsNullOrWhiteSpace(Access) ? null : Access.Trim();
            user.Approved = Approved;
            user.DemobilisedDate = demobilisedDate;
            if (!string.IsNullOrWhiteSpace(Password))
            {
                user.Password = HashPassword(user, Password.Trim());
                user.SessionToken = null;
                user.SessionIssuedUtc = null;
            }
            if (Approved && user.Login_Approval_Date == null)
            {
                user.Login_Approval_Date = PMS.Infrastructure.AppClock.Now; // localized UTC+03
                user.Login_Approval_By = HttpContext.Session.GetInt32("UserID");
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(LoginApproval));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating user {UserId}", UserID);
            return RedirectToAction(nameof(LoginApproval));
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteUser([FromForm] int userId)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty).Trim();
        if (!string.Equals(access, "ADMIN", StringComparison.OrdinalIgnoreCase))
            return RedirectToAction("Dashboard");

        try
        {
            var user = await _context.PMS_Login_tbl.FirstOrDefaultAsync(u => u.UserID == userId);
            if (user == null)
            {
                TempData["Msg"] = "User not found.";
                return RedirectToAction(nameof(LoginApproval));
            }

            _context.PMS_Login_tbl.Remove(user);
            await _context.SaveChangesAsync();
            TempData["Msg"] = $"Deleted user {userId}.";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteUser failed for {UserId}", userId);
            TempData["Msg"] = "Failed to delete user.";
        }

        return RedirectToAction(nameof(LoginApproval));
    }

    // GET: /Home/ExportUsers?status=pending&q=search
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> ExportUsers([FromQuery] string? status, [FromQuery] string? q)
    {
        if (!IsAdmin())
            return RedirectToAction(nameof(Dashboard));

        var nowUtc = PMS.Infrastructure.AppClock.Now; // project local time reference for filtering
        var usersQry = _context.PMS_Login_tbl.AsNoTracking();

        if (!string.IsNullOrWhiteSpace(status) && !string.Equals(status, "all", StringComparison.OrdinalIgnoreCase))
        {
            var st = status.ToLowerInvariant();
            usersQry = st switch
            {
                "pending" => usersQry.Where(u => !u.Approved && (u.DemobilisedDate == null || u.DemobilisedDate >= nowUtc)),
                "approved" => usersQry.Where(u => u.Approved && (u.DemobilisedDate == null || u.DemobilisedDate >= nowUtc)),
                "demobilized" => usersQry.Where(u => u.DemobilisedDate != null && u.DemobilisedDate < nowUtc),
                _ => usersQry
            };
        }

        if (!string.IsNullOrWhiteSpace(q))
        {
            // Suppress CA1862: Using ToLower for EF translation (translates to SQL LOWER()), avoids client eval and maintains behavior
#pragma warning disable CA1862
            var term = q.Trim().ToLower();
            usersQry = usersQry.Where(u =>
                (u.FirstName + " " + u.LastName).ToLower().Contains(term) ||
                (u.Email != null && u.Email.ToLower().Contains(term))
            );
#pragma warning restore CA1862
        }

        var users = await usersQry
            .OrderBy(u => u.Approved)
            .ThenBy(u => u.FirstName)
            .ThenBy(u => u.LastName)
            .Select(u => new
            {
                u.UserID,
                u.FirstName,
                u.LastName,
                u.UserName,
                u.Email,
                u.Position,
                u.Company,
                u.Access,
                u.Approved,
                u.DemobilisedDate,
                u.CreatedDate
            })
            .ToListAsync();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Users");

        string[] headers = ["User ID", "Name", "Username", "Email", "Position", "Company", "Access Level", "Status", "Created Date", "Demobilized Date"];
        for (int i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];

        int r = 2;
        foreach (var u in users)
        {
            string statusText = u.Approved ? "Approved" : (u.DemobilisedDate.HasValue && u.DemobilisedDate < nowUtc ? "Demobilized" : "Pending");
            ws.Cell(r, 1).Value = u.UserID;
            ws.Cell(r, 2).Value = ($"{u.FirstName} {u.LastName}").Trim();
            ws.Cell(r, 3).Value = u.UserName ?? string.Empty;
            ws.Cell(r, 4).Value = u.Email ?? string.Empty;
            ws.Cell(r, 5).Value = u.Position ?? string.Empty;
            ws.Cell(r, 6).Value = u.Company ?? string.Empty;
            ws.Cell(r, 7).Value = u.Access ?? string.Empty;
            ws.Cell(r, 8).Value = statusText;
            if (u.CreatedDate.HasValue)
            {
                ws.Cell(r, 9).Value = u.CreatedDate.Value;
                ws.Cell(r, 9).Style.DateFormat.Format = "yyyy-MM-dd";
            }
            if (u.DemobilisedDate.HasValue)
            {
                ws.Cell(r,10).Value = u.DemobilisedDate.Value;
                ws.Cell(r,10).Style.DateFormat.Format = "yyyy-MM-dd";
            }
            ws.Row(r).Height = 17;
            r++;
        }

        int lastRow = r - 1;
        var fullRange = ws.Range(1, 1, lastRow, headers.Length);
        var table = fullRange.CreateTable();
        table.Theme = XLTableTheme.TableStyleMedium2;
        table.ShowTotalsRow = false;
        ws.Row(1).Height = 30;
        ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var fileName = $"LoginApproval_{PMS.Infrastructure.AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}
