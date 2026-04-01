using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using PMS.Data;
using PMS.Services;
using System.Linq;
using PMS.Infrastructure;

namespace PMS.Controllers // changed from PMS.Filters to PMS.Controllers
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method)]
    public sealed class SessionAuthorizationAttribute : Attribute, IAsyncAuthorizationFilter
    {
        public async Task OnAuthorizationAsync(AuthorizationFilterContext context)
        {
            var session = context.HttpContext.Session;
            var fullName = session.GetString("FullName");
            var userId = session.GetInt32("UserID");
            var sessionToken = session.GetString("SessionToken");

            if (string.IsNullOrEmpty(fullName) || !userId.HasValue || string.IsNullOrEmpty(sessionToken))
            {
                session.Clear();
                context.Result = new RedirectToActionResult("Login", "Home", null);
                return;
            }

            var db = context.HttpContext.RequestServices.GetRequiredService<AppDbContext>();
            var userRow = await db.PMS_Login_tbl.AsNoTracking()
                .Where(u => u.UserID == userId.Value)
                .Select(u => new { u.SessionToken, u.Approved, u.DemobilisedDate })
                .FirstOrDefaultAsync();

            var hashedSessionToken = AuthService.HashToken(sessionToken);

            bool invalid = userRow == null
                           || !string.Equals(userRow.SessionToken, hashedSessionToken, StringComparison.Ordinal)
                           || !userRow.Approved
                           || (userRow.DemobilisedDate.HasValue && userRow.DemobilisedDate.Value < AppClock.UtcNow);

            if (invalid)
            {
                session.Clear();
                context.Result = new RedirectToActionResult("Login", "Home", null);
            }
        }
    }
}
