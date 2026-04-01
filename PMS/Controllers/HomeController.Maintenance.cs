using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PMS.Services;
using System.Globalization;

namespace PMS.Controllers;

public partial class HomeController
{
    /// <summary>
    /// Displays the maintenance page to non-admin users when maintenance mode is active.
    /// </summary>
    [AllowAnonymous]
    [HttpGet]
    public IActionResult Maintenance()
    {
        var maintenance = HttpContext.RequestServices.GetRequiredService<MaintenanceService>();

        // If maintenance mode is not active, redirect to Dashboard (or Login)
        if (!maintenance.IsEnabled)
        {
            var userId = HttpContext.Session.GetInt32("UserID");
            return userId.HasValue ? RedirectToAction(nameof(Dashboard)) : RedirectToAction(nameof(Login));
        }

        ViewBag.MaintenanceMessage = maintenance.Message;
        ViewBag.EstimatedEndUtc = maintenance.EstimatedEndUtc;
        return View();
    }

    /// <summary>
    /// Admin-only POST action to toggle maintenance mode on or off.
    /// Called via AJAX from the dashboard header.
    /// </summary>
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public IActionResult ToggleMaintenance([FromForm] bool enable, [FromForm] string? message, [FromForm] int? durationMinutes)
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty)
            .Trim().Replace(' ', '_').ToUpperInvariant();

        if (access != "ADMIN")
        {
            return Forbid();
        }

        var maintenance = HttpContext.RequestServices.GetRequiredService<MaintenanceService>();

        if (enable)
        {
            DateTime? estimatedEnd = durationMinutes.HasValue && durationMinutes.Value > 0
                ? DateTime.UtcNow.AddMinutes(durationMinutes.Value)
                : null;
            maintenance.Enable(message, estimatedEnd);
        }
        else
        {
            maintenance.Disable();
        }

        _logger.LogInformation("Maintenance mode {State} by {User} (duration: {Duration} min)",
            enable ? "ENABLED" : "DISABLED",
            HttpContext.Session.GetString("FullName") ?? "unknown",
            durationMinutes);

        return Json(new
        {
            enabled = maintenance.IsEnabled,
            estimatedEndUtc = maintenance.EstimatedEndUtc?.ToString("o")
        });
    }

    /// <summary>
    /// Returns the current maintenance mode state (for AJAX polling / header init).
    /// </summary>
    [SessionAuthorization]
    [HttpGet]
    public IActionResult MaintenanceStatus()
    {
        var access = (HttpContext.Session.GetString("Access") ?? string.Empty)
            .Trim().Replace(' ', '_').ToUpperInvariant();

        if (access != "ADMIN")
        {
            return Forbid();
        }

        var maintenance = HttpContext.RequestServices.GetRequiredService<MaintenanceService>();
        return Json(new
        {
            enabled = maintenance.IsEnabled,
            message = maintenance.Message,
            estimatedEndUtc = maintenance.EstimatedEndUtc?.ToString("o")
        });
    }
}
