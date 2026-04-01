using PMS.Services;

namespace PMS.Middleware;

/// <summary>
/// When maintenance mode is active, redirects every non-admin request to the
/// maintenance page. Admin users (session Access == "ADMIN") can still browse
/// normally. Login, logout, maintenance page, toggle endpoint, and static
/// files are always excluded so users aren't permanently locked out and
/// admins can authenticate and disable maintenance.
/// </summary>
public sealed class MaintenanceMiddleware
{
    private readonly RequestDelegate _next;

    public MaintenanceMiddleware(RequestDelegate next) => _next = next;

    public async Task InvokeAsync(HttpContext context, MaintenanceService maintenance)
    {
        if (maintenance.IsEnabled)
        {
            var path = context.Request.Path.Value ?? string.Empty;

            // Always allow: static files, login (GET+POST), logout, maintenance page,
            // toggle/status endpoints, root path (maps to Login via routing), favicon,
            // and public pages linked from maintenance (ContactUs, AboutUs, SignUp, ForgotPassword)
            bool isExcluded = path.Equals("/", StringComparison.Ordinal)
                           || path.StartsWith("/css", StringComparison.OrdinalIgnoreCase)
                           || path.StartsWith("/js", StringComparison.OrdinalIgnoreCase)
                           || path.StartsWith("/img", StringComparison.OrdinalIgnoreCase)
                           || path.StartsWith("/lib", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/Login", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/Logout", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/Maintenance", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/ToggleMaintenance", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/MaintenanceStatus", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/ContactUs", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/AboutUs", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/SignUp", StringComparison.OrdinalIgnoreCase)
                           || path.Equals("/Home/ForgotPassword", StringComparison.OrdinalIgnoreCase)
                           || path.StartsWith("/favicon", StringComparison.OrdinalIgnoreCase);

            if (!isExcluded)
            {
                // Let admin users through
                var access = (context.Session.GetString("Access") ?? string.Empty)
                    .Trim().Replace(' ', '_').ToUpperInvariant();

                if (access != "ADMIN")
                {
                    context.Response.Redirect("/Home/Maintenance");
                    return;
                }
            }
        }

        await _next(context);
    }
}
