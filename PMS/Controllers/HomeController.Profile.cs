using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Models;
using System;
using System.Threading.Tasks;

namespace PMS.Controllers;

public partial class HomeController : Controller
{
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> Profile()
    {
        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            if (userId is null)
                return RedirectToAction(nameof(Login));

            var user = await _context.PMS_Login_tbl
                .AsNoTracking()
                .FirstOrDefaultAsync(u => u.UserID == userId.Value);

            if (user is null)
                return RedirectToAction(nameof(Dashboard));

            ViewBag.FullName = HttpContext.Session.GetString("FullName") ?? string.Empty;
            return View(user);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading Profile page");
            return RedirectToAction(nameof(Dashboard));
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ProfileUpdate(string? email, string? position)
    {
        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            if (userId is null)
                return RedirectToAction(nameof(Login));

            var user = await _context.PMS_Login_tbl
                .FirstOrDefaultAsync(u => u.UserID == userId.Value);

            if (user is null)
                return RedirectToAction(nameof(Dashboard));

            user.Email = Clean(email, 50);
            user.Position = Clean(position, 50);

            await _context.SaveChangesAsync();

            HttpContext.Session.SetString("Position", user.Position ?? string.Empty);

            TempData["ProfileMessage"] = "Profile updated successfully.";
            return RedirectToAction(nameof(Profile));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating profile");
            TempData["ProfileError"] = "An error occurred while updating your profile.";
            return RedirectToAction(nameof(Profile));
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ChangePassword(string currentPassword, string newPassword, string confirmPassword)
    {
        try
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            if (userId is null)
                return RedirectToAction(nameof(Login));

            if (string.IsNullOrWhiteSpace(currentPassword) || string.IsNullOrWhiteSpace(newPassword) || string.IsNullOrWhiteSpace(confirmPassword))
            {
                TempData["ProfileError"] = "All password fields are required.";
                return RedirectToAction(nameof(Profile));
            }

            if (newPassword != confirmPassword)
            {
                TempData["ProfileError"] = "New password and confirmation do not match.";
                return RedirectToAction(nameof(Profile));
            }

            if (!StrongPasswordRegex().IsMatch(newPassword))
            {
                TempData["ProfileError"] = "Password must be 8–20 characters with at least one uppercase, one lowercase, one digit, and one special character.";
                return RedirectToAction(nameof(Profile));
            }

            var user = await _context.PMS_Login_tbl
                .FirstOrDefaultAsync(u => u.UserID == userId.Value);

            if (user is null)
                return RedirectToAction(nameof(Dashboard));

            // Verify current password using AuthService
            var verified = await _authService.Authenticate(user.UserName ?? string.Empty, currentPassword);
            if (verified is null)
            {
                TempData["ProfileError"] = "Current password is incorrect.";
                return RedirectToAction(nameof(Profile));
            }

            user.Password = HashPassword(user, newPassword);
            await _context.SaveChangesAsync();

            TempData["ProfileMessage"] = "Password changed successfully.";
            return RedirectToAction(nameof(Profile));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error changing password");
            TempData["ProfileError"] = "An error occurred while changing the password.";
            return RedirectToAction(nameof(Profile));
        }
    }
}
