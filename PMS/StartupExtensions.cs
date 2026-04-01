using PMS.Middleware;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PMS.Data;
using PMS.Services;
using PMS.Options;
using PMS.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.CookiePolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.AspNetCore.Mvc.ViewFeatures; // for CookieTempDataProviderOptions
using Microsoft.AspNetCore.HttpsPolicy; // for HSTS options
using Microsoft.Extensions.Primitives; // for StringValues
using Microsoft.EntityFrameworkCore.Diagnostics; // suppress EF warnings in dev
using Microsoft.AspNetCore.Identity;
using HeaderNames = Microsoft.Net.Http.Headers.HeaderNames;

namespace PMS;

public static class StartupExtensions
{
    public static void AddPmsServices(this WebApplicationBuilder builder)
    {
        // Logging
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole();
        builder.Logging.AddDebug();
        builder.Logging.AddEventSourceLogger();

        // Session prerequisites
        builder.Services.AddDistributedMemoryCache();
        builder.Services.AddMemoryCache();

        // MVC + Razor Pages
        builder.Services.AddRazorPages();
        builder.Services.AddControllersWithViews();
        // Enforce HTTPS globally for both MVC and Razor Pages
        builder.Services.Configure<MvcOptions>(options =>
        {
            options.Filters.Add(new RequireHttpsAttribute());
        });

        // Secure TempData provider cookie (.AspNetCore.Mvc.CookieTempDataProvider)
        builder.Services.Configure<CookieTempDataProviderOptions>(options =>
        {
            options.Cookie.Name = ".AspNetCore.Mvc.CookieTempDataProvider"; // keep default name explicit
            options.Cookie.HttpOnly = true;
            options.Cookie.IsEssential = true; // required for framework features
            options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
            options.Cookie.SameSite = SameSiteMode.Strict;
        });

        // Prefer permanent redirects for HTTPS to preserve method and body
        builder.Services.AddHttpsRedirection(options =>
        {
            options.RedirectStatusCode = StatusCodes.Status308PermanentRedirect;
            // Ensure we redirect to the correct dev HTTPS port when needed
            options.HttpsPort = 7254;
        });

        // Configure HSTS for all environments (shorter in Development)
        builder.Services.AddHsts(options =>
        {
            options.IncludeSubDomains = true;
            options.Preload = false;
            options.MaxAge = builder.Environment.IsDevelopment()
                ? TimeSpan.FromHours(1)
                : TimeSpan.FromDays(365);
        });

        builder.Services.AddSession(options =>
        {
            options.IdleTimeout = TimeSpan.FromMinutes(30);
            options.Cookie.HttpOnly = true;
            options.Cookie.IsEssential = true;
            options.Cookie.SameSite = SameSiteMode.Strict;
            // Ensure session cookie is only sent over HTTPS
            options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
        });

        // Ensure antiforgery cookie is marked Secure and SameSite
        builder.Services.AddAntiforgery(opts =>
        {
            opts.Cookie.HttpOnly = true;
            opts.Cookie.SecurePolicy = CookieSecurePolicy.Always;
            opts.Cookie.SameSite = SameSiteMode.Strict;
        });

        // When running behind reverse proxies (IIS/Nginx), respect forwarded headers for correct scheme detection
        builder.Services.Configure<ForwardedHeadersOptions>(options =>
        {
            options.ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto;
            // If behind unknown proxies in dev/test, clear the known networks/proxies to allow forwarded headers
            options.KnownNetworks.Clear();
            options.KnownProxies.Clear();
        });

        // Database context
        builder.Services.AddDbContext<AppDbContext>(options =>
        {
            options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"));
            if (builder.Environment.IsDevelopment())
            {
                options.EnableSensitiveDataLogging();
                // Suppress the EF warning about sensitive data logging being enabled in Development
                options.ConfigureWarnings(w => w.Ignore(CoreEventId.SensitiveDataLoggingEnabledWarning));
            }
        });

        // Maintenance mode – singleton so the state is shared across all requests
        builder.Services.AddSingleton<MaintenanceService>();

        // Application services
        builder.Services.AddScoped<AuthService>();
        builder.Services.AddScoped<EmailService>();
        builder.Services.AddScoped<IWpsSelectorService, WpsSelectorService>(); // register WPS selector
        builder.Services.AddHttpClient<IOcrService, AzureOcrService>();
        builder.Services.AddScoped<IPasswordHasher<PMSLogin>, PasswordHasher<PMSLogin>>();

        // Configure email settings
        builder.Services.Configure<EmailSettings>(builder.Configuration.GetSection("EmailSettings"));
        builder.Services.Configure<AzureDocumentIntelligenceOptions>(builder.Configuration.GetSection("AzureDocumentIntelligence"));
    }

    public static void UsePmsPipeline(this WebApplication app)
    {
        if (!app.Environment.IsDevelopment())
        {
            app.UseExceptionHandler("/Home/Error");
        }
        else
        {
            app.UseDeveloperExceptionPage();
        }

        // Always apply HSTS (shorter duration in Development as configured above)
        app.UseHsts();

        // Honor X-Forwarded-* headers BEFORE redirects so the app sees the original HTTPS scheme
        app.UseForwardedHeaders();

        // Always redirect to HTTPS
        app.UseHttpsRedirection();

        // Upgrade any accidental http subresource references to https and block mixed content outright
        app.Use(async (context, next) =>
        {
            const string upgrade = "upgrade-insecure-requests";
            const string block = "block-all-mixed-content";
            var headerName = HeaderNames.ContentSecurityPolicy;

            var headers = context.Response.Headers;
            if (headers.TryGetValue(headerName, out var existing) && !StringValues.IsNullOrEmpty(existing))
            {
                var cur = existing.ToString();
                bool hasUpgrade = cur.Contains(upgrade, StringComparison.OrdinalIgnoreCase);
                bool hasBlock = cur.Contains(block, StringComparison.OrdinalIgnoreCase);
                if (!hasUpgrade || !hasBlock)
                {
                    string additions = string.Empty;
                    if (!hasUpgrade && !hasBlock)
                    {
                        additions = string.Concat(upgrade, "; ", block);
                    }
                    else if (!hasUpgrade)
                    {
                        additions = upgrade;
                    }
                    else
                    {
                        additions = block;
                    }
                    headers[headerName] = string.Concat(cur, "; ", additions);
                }
            }
            else
            {
                headers[headerName] = string.Concat(upgrade, "; ", block);
            }

            await next();
        });

        // Enforce secure cookie policy to require Secure and strong SameSite
        // Note: HttpOnly is configured per-cookie above (session & antiforgery), keep policy to enforce Secure and SameSite defaults
        app.UseCookiePolicy(new CookiePolicyOptions
        {
            Secure = CookieSecurePolicy.Always,
            MinimumSameSitePolicy = SameSiteMode.Strict
        });

        app.UseStaticFiles();
        app.UseRouting();

        // Enable session before authorization/endpoints
        app.UseSession();

        // Maintenance mode gate – must come after UseSession so session data is available
        app.UseMiddleware<MaintenanceMiddleware>();

        app.UseAuthorization();

        app.MapRazorPages();
        app.MapControllerRoute(
            name: "default",
            pattern: "{controller=Home}/{action=Login}/{id?}");
    }
}