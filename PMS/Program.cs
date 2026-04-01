using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.CookiePolicy;
using PMS;
using PMS.Services;
using QuestPDF.Infrastructure;

namespace PMS
{
    public class Program
    {
        public static void Main(string[] args)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            DisableVisualStudioBrowserLink();

            var builder = WebApplication.CreateBuilder(args);

            // Prevent Visual Studio hosting startup (disables Browser Link script injection and related deprecated unload handlers)
            builder.WebHost.UseSetting(WebHostDefaults.PreventHostingStartupKey, "true");

            // Configure Kestrel server limits only (do not bind endpoints here to avoid overriding ASPNETCORE_URLS/applicationUrl)
            builder.WebHost.ConfigureKestrel(options =>
            {
                // Set a bounded max request body size to reduce DoS surface from oversized uploads
                const long maxRequestBody = 200L * 1024L * 1024L; // 200 MB
                options.Limits.MaxRequestBodySize = maxRequestBody;
            });

            // Enforce Secure and SameSite=Strict for all cookies
            builder.Services.Configure<CookiePolicyOptions>(options =>
            {
                options.Secure = CookieSecurePolicy.Always;
                options.MinimumSameSitePolicy = SameSiteMode.Strict;
                options.HttpOnly = HttpOnlyPolicy.Always;
            });

            // Global FormOptions: allow very large multipart bodies and many fields/files
            builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(o =>
            {
                const long maxMultipart = 200L * 1024L * 1024L; // 200 MB
                o.MultipartBodyLengthLimit = maxMultipart; // total multipart size
                o.ValueCountLimit = 10_000;                // number of form keys
                o.ValueLengthLimit = 100_000;              // length per form value
                o.MultipartHeadersCountLimit = 256;
                o.MultipartHeadersLengthLimit = 32_768;
                o.KeyLengthLimit = 256;
                o.MemoryBufferThreshold = 64 * 1024;       // buffer up to 64 KB in memory before streaming to disk
            });

            builder.Services.AddSession(options =>
            {
                options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
                options.Cookie.SameSite = SameSiteMode.Strict;
            });

            builder.Services.AddAntiforgery(options =>
            {
                options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
                options.Cookie.SameSite = SameSiteMode.Strict;
            });

            // Centralized service registration (adds Razor Pages, MVC, Session + DistributedCache, DbContext, Logging, Email, etc.)
            builder.AddPmsServices();

            var app = builder.Build();

            // Enforce HTTPS redirection
            app.UseHttpsRedirection();

            // Enforce cookie policy
            app.UseCookiePolicy();

            // Centralized HTTP pipeline (exception handling, static files, routing, session, authorization, endpoints)
            app.UsePmsPipeline();

            app.Run();
        }

        private static void DisableVisualStudioBrowserLink()
        {
            var environmentName = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
            if (!string.Equals(environmentName, Environments.Development, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            const string preventKey = "ASPNETCORE_PREVENTHOSTINGSTARTUP";
            Environment.SetEnvironmentVariable(preventKey, "true");

            const string hostingStartupKey = "ASPNETCORE_HOSTINGSTARTUPASSEMBLIES";
            var assemblies = Environment.GetEnvironmentVariable(hostingStartupKey);
            if (string.IsNullOrEmpty(assemblies))
            {
                return;
            }

            var filtered = string.Join(';', assemblies
                .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(a => !a.Equals("Microsoft.VisualStudio.Web.BrowserLink", StringComparison.OrdinalIgnoreCase)));

            if (!string.Equals(filtered, assemblies, StringComparison.Ordinal))
            {
                Environment.SetEnvironmentVariable(hostingStartupKey, filtered);
            }
        }
    }
}