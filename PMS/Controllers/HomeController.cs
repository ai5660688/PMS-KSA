using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using PMS.Data;
using PMS.Models;
using PMS.Services;
using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Mvc.Rendering;
using PMS.Infrastructure;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Diagnostics.CodeAnalysis;
using Microsoft.AspNetCore.Hosting;
using System.IO;

namespace PMS.Controllers;

 public partial class HomeController(
 AuthService authService,
 EmailService emailService,
 AppDbContext context,
 ILogger<HomeController> logger,
 IOptions<PMS.Services.EmailSettings> emailOptions,
 IDataProtectionProvider dataProtectionProvider,
 IWpsSelectorService wpsSelectorService,
 IWebHostEnvironment env,
 IOcrService? ocrService = null,
 IMemoryCache? cache = null
) : Controller
{
 private readonly AuthService _authService = authService;
 private readonly EmailService _emailService = emailService;
 private readonly AppDbContext _context = context;
 private readonly ILogger<HomeController> _logger = logger;
 private readonly EmailSettings _emailSettings = emailOptions.Value;
 private readonly IDataProtector _pwdProtector = dataProtectionProvider.CreateProtector("PMS.PasswordCookie.v1");
 private readonly IWpsSelectorService _wpsSelectorService = wpsSelectorService;
 private readonly IWebHostEnvironment _env = env;
 private readonly IOcrService? _ocrService = ocrService;
  private readonly IMemoryCache _cache = cache ?? new MemoryCache(new MemoryCacheOptions());

  private const int LoginMaxFailures = 5;
  private static readonly TimeSpan LoginLockoutDuration = TimeSpan.FromMinutes(5);
  private static readonly TimeSpan LoginFailureWindow = TimeSpan.FromMinutes(10);

  private const int ForgotPasswordMaxRequests = 3;
  private static readonly TimeSpan ForgotPasswordWindow = TimeSpan.FromMinutes(15);

 // Weld types to exclude from daily/confirmed diameter sums
 private static readonly string[] ExcludedWeldTypesForSum = ["SP", "RP", "DM", "TH", "FJ"]; // simplified (IDE0300)

 private static readonly char[] EmailSplitSeparators = [';', ',', '\n', '\r', '\t'];

 // Source-generated Regex methods
 [GeneratedRegex(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^\da-zA-Z])(?=\S+$).{8,20}$", RegexOptions.CultureInvariant)]
 private static partial Regex StrongPasswordRegex();

 [GeneratedRegex(@"(\d+)(?!.*\d)", RegexOptions.CultureInvariant)]
 private static partial Regex LastDigitsRegex();

 [GeneratedRegex(@"(\d+)[A-Za-z]*\s*$", RegexOptions.CultureInvariant)]
 private static partial Regex StripTrailingNumberAndLettersRegex();

 [GeneratedRegex(@"\d+(\.\d+)?", RegexOptions.CultureInvariant)]
 private static partial Regex NumberInStringRegex();

 [GeneratedRegex(@"\s+", RegexOptions.CultureInvariant)]
 private static partial Regex StripWhitespaceRegex();

 [GeneratedRegex("<.*?>", RegexOptions.CultureInvariant)]
 private static partial Regex HtmlTagsRegex();

 private static string? Clean(string? s, int maxLen)
 {
 s = (s ?? string.Empty).Trim();
 if (s.Length > maxLen) s = s[..maxLen];
 return string.IsNullOrWhiteSpace(s) ? null : s;
 }

 private static string? CleanSelect(string? value, int maxLen)
 {
 var v = (value ?? string.Empty).Trim();
 if (v == "__new__") return null;
 if (v.Length > maxLen) v = v[..maxLen];
 return string.IsNullOrWhiteSpace(v) ? null : v;
 }

 private static string NormalizeWeldType(string? weldType)
 {
 if (string.IsNullOrWhiteSpace(weldType)) return string.Empty;
 var wt = weldType.Trim().ToUpperInvariant();
 return wt.Length >4 ? wt[..4] : wt;
 }

 // Remove all whitespace chars
 private static string StripWhitespace(string? s)
 => StripWhitespaceRegex().Replace((s ?? string.Empty), "");

 private string HashPassword(PMSLogin user, string password)
     => _authService.HashPasswordForStorage(user, password);

 // Helper: set a secure cookie consistently
 private static CookieOptions SecureCookieOptions(DateTimeOffset? expires = null)
     => new()
     {
         Expires = expires?.ToUniversalTime(),
         HttpOnly = true,
         Secure = true,
         SameSite = SameSiteMode.Strict,
         IsEssential = true
     };

 // Minimal stubs used across partials to satisfy compile if not linked here
 [SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "Partial stub method; keep instance signature for parity across partials")]
 private Task LoadQualificationDropdownListsAsync(WelderQualification _)
 => Task.CompletedTask;

 [SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "Partial stub method; keep instance signature for parity across partials")]
 private Task<List<object>> QueryDfrRowsAsync(int _1, string _2, string _3, string _4) => Task.FromResult(new List<object>());

  private sealed record LoginThrottleState(int Failures, DateTimeOffset? LockoutUntil);

  private bool IsLoginLockedOut(string normalizedUser, string remoteIp, out TimeSpan? retryAfter)
  {
      retryAfter = null;
      var key = $"login-throttle:{normalizedUser}:{remoteIp}";
      var now = DateTimeOffset.UtcNow;
      if (_cache.TryGetValue<LoginThrottleState>(key, out var state) && state?.LockoutUntil is DateTimeOffset lockout && lockout > now)
      {
          retryAfter = lockout - now;
          return true;
      }
      return false;
  }

  private void RecordLoginFailure(string normalizedUser, string remoteIp)
  {
      var key = $"login-throttle:{normalizedUser}:{remoteIp}";
      var now = DateTimeOffset.UtcNow;
      var state = _cache.Get<LoginThrottleState>(key);
      int failures = state?.Failures ?? 0;
      DateTimeOffset? lockout = null;

      failures++;
      if (failures >= LoginMaxFailures)
      {
          lockout = now + LoginLockoutDuration;
          failures = 0;
          _logger.LogWarning("Login lockout applied for user {User} from {Ip}", normalizedUser, remoteIp);
      }

      _cache.Set(key, new LoginThrottleState(failures, lockout), new MemoryCacheEntryOptions
      {
          SlidingExpiration = LoginFailureWindow,
          AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(1)
      });
  }

  private void ClearLoginFailures(string normalizedUser, string remoteIp)
  {
      var key = $"login-throttle:{normalizedUser}:{remoteIp}";
      _cache.Remove(key);
  }

  private bool IsForgotPasswordThrottled(string remoteIp)
  {
      var key = $"forgot-pw-throttle:{remoteIp}";
      var count = _cache.Get<int?>(key) ?? 0;
      return count >= ForgotPasswordMaxRequests;
  }

  private void RecordForgotPasswordRequest(string remoteIp)
  {
      var key = $"forgot-pw-throttle:{remoteIp}";
      var count = _cache.Get<int?>(key) ?? 0;
      _cache.Set(key, count + 1, new MemoryCacheEntryOptions
      {
          AbsoluteExpirationRelativeToNow = ForgotPasswordWindow
      });
  }

 // Helper: joint display number e.g. WS-0001 or WS-0001R1
 private async Task<string> GetJointDisplayAsync(int jointId)
 {
 var d = await _context.DFR_tbl.AsNoTracking()
 .Where(x => x.Joint_ID == jointId)
 .Select(x => new { x.LOCATION, x.WELD_NUMBER, x.J_Add })
 .FirstOrDefaultAsync();
 if (d == null) return jointId.ToString();
 var loc = d.LOCATION ?? string.Empty;
 var weld = d.WELD_NUMBER ?? string.Empty;
 var jadd = d.J_Add;
 var suffix = (!string.IsNullOrWhiteSpace(jadd) && !string.Equals(jadd, "NEW", StringComparison.OrdinalIgnoreCase)) ? jadd : string.Empty;
 var sep = (loc.Length >0 && weld.Length >0) ? "-" : string.Empty;
 var joint = string.Concat(loc, sep, weld, suffix);
 return string.IsNullOrWhiteSpace(joint) ? jointId.ToString() : joint;
 }

 // Build the same email shell used by Password Reset, with brand logo and heading
 private async Task<(string Html, Dictionary<string, (byte[] Content, string ContentType)>? Resources)> BuildBrandEmailAsync(string title, string contentHtml)
 {
     string cid = "pmslogo";
     Dictionary<string, (byte[] Content, string ContentType)>? resources = null;
     string logoTag;
     try
     {
         var logoPath = Path.Combine(_env.WebRootPath ?? string.Empty, "img", "PMS_logo_email.png");
         if (System.IO.File.Exists(logoPath))
         {
             resources = new() { [cid] = (await System.IO.File.ReadAllBytesAsync(logoPath), "image/png") };
             logoTag = $"<img src='cid:{cid}' alt='PMS Logo' height='60' style='height:60px;width:auto;display:inline-block;border:0;margin:0;vertical-align:middle;' />";
         }
         else
         {
             var absolute = $"{Request.Scheme}://{Request.Host}{Url.Content("~/img/PMS_logo_email.png")}";
             logoTag = $"<img src='{absolute}' alt='PMS Logo' height='60' style='height:60px;width:auto;display:inline-block;border:0;margin:0;vertical-align:middle;' />";
         }
     }
     catch
     {
         var absolute = $"{Request.Scheme}://{Request.Host}{Url.Content("~/img/PMS_logo_email.png")}";
         logoTag = $"<img src='{absolute}' alt='PMS Logo' height='60' style='height:60px;width:auto;display:inline-block;border:0;margin:0;vertical-align:middle;' />";
     }

     var html = $@"<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'/><title>{System.Net.WebUtility.HtmlEncode(title)} - PMS</title></head><body style='margin:0;padding:24px;background:#ffffff;color:#176d8a;font-family:Calibri,Arial,Helvetica,sans-serif;'>
<div style='max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #ffffff;border-radius:8px;box-shadow:0 2px 8px rgba(31,153,190,0.15);overflow:hidden;'>
  <div style='padding:14px 18px;text-align:center;background:#ffffff;'>{logoTag}</div>
  <div style='padding:20px 22px;'>
    <h2 style='margin:0 0 14px;color:#176d8a;font-weight:700;font-size:20px;'>{System.Net.WebUtility.HtmlEncode(title)}</h2>
    {contentHtml}
  </div>
  <div style='background:#ffffff;border-top:1px solid #e3e9eb;padding:12px 16px;text-align:center;'>
    <p style='margin:0;color:#7a9aa8;font-size:12px;'>Automated message from PMS System. Do not reply.</p>
  </div>
</div></body></html>";

     return (html, resources);
 }

 private static IEnumerable<string> SplitEmails(string? raw)
 {
     if (string.IsNullOrWhiteSpace(raw)) yield break;
    foreach (var part in raw.Split(EmailSplitSeparators, StringSplitOptions.RemoveEmptyEntries))
     {
         var v = part.Trim();
         if (!string.IsNullOrWhiteSpace(v)) yield return v;
     }
 }

 private async Task<(List<string> To, List<string> Cc)> ResolveReceiversAsync(int projectId, string razorKey, string locCode)
 {
     try
     {
         var rows = await _context.Receivers_tbl.AsNoTracking()
             .Where(r => r.Project_Receivers == projectId && r.Razor == razorKey)
             .Where(r => r.Location_Receivers == null || r.Location_Receivers == "" ||
                         r.Location_Receivers!.Equals("All", StringComparison.OrdinalIgnoreCase) ||
                         r.Location_Receivers!.Equals(locCode, StringComparison.OrdinalIgnoreCase))
             .ToListAsync();
         var to = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
         var cc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
         foreach (var row in rows)
         {
             foreach (var e in SplitEmails(row.Receivers)) to.Add(e);
             foreach (var e in SplitEmails(row.Receivers_cc)) cc.Add(e);
         }
         // Remove any addresses that are in both sets from cc
         foreach (var e in to) cc.Remove(e);
         return (to.ToList(), cc.ToList());
     }
     catch (Exception ex)
     {
         _logger.LogWarning(ex, "ResolveReceivers failed for Razor={Razor} Project={Project} Loc={Loc}", razorKey, projectId, locCode);
         return (new List<string>(), new List<string>());
     }
 }

 // Validation helper for Fit-up Date rules
 private async Task<(bool ok, string? message)> ValidateFitupDateAgainstConstraintsAsync(int jointId, DateTime newFitup)
 {
 var data = await (
 from d in _context.DFR_tbl.AsNoTracking()
 join w in _context.DWR_tbl.AsNoTracking() on d.Joint_ID equals w.Joint_ID_DWR into jw
 from w in jw.DefaultIfEmpty()
 join n in _context.Other_NDE_tbl.AsNoTracking() on d.Joint_ID equals n.Joint_ID_NDE into jn
 from n in jn.DefaultIfEmpty()
 where d.Joint_ID == jointId
 select new { Welded = w.DATE_WELDED, Other = n.OTHER_NDE_DATE, BevelPt = n.Bevel_PT_DATE }
 ).FirstOrDefaultAsync();
 if (data == null) return (false, "Joint not found");

 bool hasWelded = data.Welded.HasValue;
 bool hasOther = data.Other.HasValue;
 bool hasBevalPt = data.BevelPt.HasValue;

 // New rule: Fit-up must be on or after Bevel PT Date when available
 if (hasBevalPt && newFitup < data.BevelPt!.Value)
 {
     return (false, $"Fit-up Date must be on or after Bevel PT Date {data.BevelPt!.Value:dd-MMM-yyyy}.");
 }

 if (!hasWelded && !hasOther) return (true, null);

 bool ok = (hasWelded && newFitup <= data.Welded!.Value) || (hasOther && newFitup <= data.Other!.Value);
 if (ok) return (true, null);

 // Simplify collection initialization (C# 12 collection expressions)
 List<string> parts = [];
 if (hasWelded) parts.Add($"Date Welded {data.Welded!.Value:dd-Mmm-yyyy}");
 if (hasOther) parts.Add($"Other NDT Date {data.Other!.Value:dd-Mmm-yyyy}");
 var msg = "Fit-up Date must be on or before " + string.Join(" or ", parts) + ".";
 return (false, msg);
 }

 // Update Line_Sheet_ID_DFR and Spool_ID_DFR per user rules
 [SuppressMessage("Performance", "CA1862", Justification = "EF Core translates Trim/ToUpper server-side; StringComparison overloads are not translated in LINQ to Entities")]
 private async Task UpdateLineSheetAndSpoolRefsAsync(Dfr entity, int? userId = null, DateTime? now = null)
 {
     if (entity == null) return;

     var layout = (entity.LAYOUT_NUMBER ?? string.Empty).Trim();
     var sheet = (entity.SHEET ?? string.Empty).Trim();
     if (string.IsNullOrWhiteSpace(layout))
     {
         _logger.LogDebug("Skip spool link: layout empty for Joint_ID={JointId}", entity.Joint_ID);
         return; // layout required at minimum
     }

     // Normalize location to WS/FW style for decisions
     string locRaw = (entity.LOCATION ?? string.Empty).Trim();
     string locUp = locRaw.ToUpperInvariant();
     bool isShopLoc = locUp == "WS" || locUp.StartsWith("WS") || locUp.Contains("SHOP") || locUp.Contains("WORK");

     _logger.LogDebug("UpdateLineSheetAndSpoolRefs start Joint_ID={JointId} Project={Proj} Layout={Layout} Sheet={Sheet} Location={LocRaw}/{LocUp} IsShop={IsShop} CurrentSpoolId={SpoolId}", entity.Joint_ID, entity.Project_No, layout, sheet, locRaw, locUp, isShopLoc, entity.Spool_ID_DFR);

     int? lineSheetId = null;
     int? lineId = null;

     // Resolve Line_Sheet (when sheet is provided)
     if (!string.IsNullOrWhiteSpace(sheet))
     {
         var ls = await _context.Line_Sheet_tbl.AsNoTracking()
             .Where(x => x.LS_LAYOUT_NO == layout && x.LS_SHEET == sheet)
             .Select(x => new { x.Line_Sheet_ID, x.Line_ID_LS })
             .FirstOrDefaultAsync();
         if (ls != null)
         {
             lineSheetId = ls.Line_Sheet_ID;
             lineId = ls.Line_ID_LS;
             entity.Line_Sheet_ID_DFR = lineSheetId;
             _logger.LogDebug("Resolved Line_Sheet_ID={LineSheetId} Line_ID={LineId} for Joint_ID={JointId}", lineSheetId, lineId, entity.Joint_ID);
         }
         else
         {
             _logger.LogDebug("No Line_Sheet match for Layout={Layout} Sheet={Sheet}", layout, sheet);
         }
     }

     // Determine material via LINE_LIST (used for insert payload only)
     string? material = null;
     if (lineId.HasValue)
     {
         material = await _context.LINE_LIST_tbl.AsNoTracking()
             .Where(l => l.Line_ID == lineId.Value)
             .Select(l => l.Material)
             .FirstOrDefaultAsync();
     }
     else
     {
         // Fallback by layout when no sheet mapping (take first match)
         material = await _context.LINE_LIST_tbl.AsNoTracking()
             .Where(l => l.LAYOUT_NO == layout)
             .Select(l => l.Material)
             .FirstOrDefaultAsync();
     }
     string? matKey = string.IsNullOrWhiteSpace(material) ? null : material!.Trim();
     _logger.LogDebug("Resolved Material={Material} (matKey={MatKey}) for Joint_ID={JointId}", material, matKey, entity.Joint_ID);

     string? spoolNoKey = string.IsNullOrWhiteSpace(entity.SPOOL_NUMBER) ? null : entity.SPOOL_NUMBER!.Trim();
     _logger.LogDebug("SpoolNoKey={SpoolNoKey} for Joint_ID={JointId}", spoolNoKey, entity.Joint_ID);

     // Avoid duplicate creation if already linked, but verify it's still correct
     if (entity.Spool_ID_DFR.HasValue && entity.Spool_ID_DFR.Value == 0)
     {
         // Treat 0 as not linked (legacy default value) and clear
         _logger.LogDebug("Normalizing zero Spool_ID_DFR to null for Joint_ID={JointId}", entity.Joint_ID);
         entity.Spool_ID_DFR = null;
         await _context.SaveChangesAsync();
     }
     if (entity.Spool_ID_DFR.HasValue && entity.Spool_ID_DFR.Value > 0)
     {
         // Validate the currently linked SP_Release still matches the DFR data (project/layout/sheet/spool no)
         var current = await _context.SP_Release_tbl.AsNoTracking()
             .Where(s => s.Spool_ID == entity.Spool_ID_DFR.Value)
             .Select(s => new
             {
                 s.Spool_ID,
                 s.SP_Project_No,
                 s.SP_LAYOUT_NUMBER,
                 s.SP_SHEET,
                 s.SP_SPOOL_NUMBER
             })
             .FirstOrDefaultAsync();

         bool SheetEqual(string? a, string? b)
             => (string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b))
                || string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.Ordinal);

         bool SpoolEqual(string? a, string? b)
             => string.Equals((a ?? string.Empty).Trim(), (b ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);

         if (current != null)
         {
             bool ok = current.SP_Project_No == entity.Project_No
                       && string.Equals((current.SP_LAYOUT_NUMBER ?? string.Empty).Trim(), layout, StringComparison.Ordinal)
                       && SheetEqual(current.SP_SHEET, string.IsNullOrWhiteSpace(sheet) ? null : sheet)
                       && SpoolEqual(current.SP_SPOOL_NUMBER, spoolNoKey);

             if (ok)
             {
                 _logger.LogDebug("Keep existing Spool_ID={SpoolId} for Joint_ID={JointId} (still matches)", current.Spool_ID, entity.Joint_ID);
                 return;
             }

             _logger.LogInformation("Unlinking mismatched Spool_ID={SpoolId} from Joint_ID={JointId} (layout/sheet/spool changed)", current.Spool_ID, entity.Joint_ID);
             entity.Spool_ID_DFR = null;
             await _context.SaveChangesAsync();
         }
         else
         {
             _logger.LogWarning("Spool_ID_DFR points to missing SP_Release row. Clearing link for Joint_ID={JointId}", entity.Joint_ID);
             entity.Spool_ID_DFR = null;
             await _context.SaveChangesAsync();
         }
     }

     // Try to find existing SP_Release row ONLY by exact spool number match
     SpRelease? sp = null;
     if (!string.IsNullOrWhiteSpace(spoolNoKey))
     {
         var spoolNoUpper = spoolNoKey.ToUpper();
         var q1 = _context.SP_Release_tbl.AsNoTracking()
             .Where(s => s.SP_Project_No == entity.Project_No
                      && s.SP_TYPE == "W"
                      && s.SP_LAYOUT_NUMBER == layout
                      && s.SP_SPOOL_NUMBER != null
                      && s.SP_SPOOL_NUMBER.ToUpper() == spoolNoUpper);
         if (!string.IsNullOrWhiteSpace(sheet)) q1 = q1.Where(s => s.SP_SHEET == sheet); else q1 = q1.Where(s => s.SP_SHEET == null || s.SP_SHEET == "");
         sp = await q1.OrderByDescending(s => s.Spool_ID).FirstOrDefaultAsync();
         _logger.LogDebug("Lookup by spool number result Spool_ID={FoundSpoolId} for Joint_ID={JointId}", sp?.Spool_ID, entity.Joint_ID);
     }

     if (sp != null)
     {
         entity.Spool_ID_DFR = sp.Spool_ID;
         await _context.SaveChangesAsync();
         _logger.LogInformation("Linked existing Spool_ID={SpoolId} to Joint_ID={JointId}", sp.Spool_ID, entity.Joint_ID);
         return;
     }

     // Create new SP_Release row only for WS/Shop location
     if (!isShopLoc)
     {
         _logger.LogDebug("Skip creation: LOCATION not WS/Shop for Joint_ID={JointId}. Raw={LocRaw}", entity.Joint_ID, locRaw);
         entity.Spool_ID_DFR = null;
         await _context.SaveChangesAsync();
         return;
     }

     // Compute SP_DIA as Max(DIAMETER) for the same Project/Layout/Sheet/Spool
     var spDiaQuery = _context.DFR_tbl.AsNoTracking()
         .Where(d => d.Project_No == entity.Project_No && d.LAYOUT_NUMBER == layout)
         .Where(d => string.IsNullOrWhiteSpace(sheet) ? (d.SHEET == null || d.SHEET == "") : d.SHEET == sheet)
         .Where(d => d.DIAMETER.HasValue);

     // Filter by spool number so each SP_Release row gets the correct per-spool
     // diameter instead of the max across all spools on the sheet.
     if (!string.IsNullOrWhiteSpace(spoolNoKey))
         spDiaQuery = spDiaQuery.Where(d => d.SPOOL_NUMBER == spoolNoKey);

     double? maxDia = await spDiaQuery.MaxAsync(d => (double?)d.DIAMETER);

     var newSp = new SpRelease
     {
         SP_Project_No = entity.Project_No,
         SP_TYPE = "W",
         SP_DIA = maxDia,
         SP_LAYOUT_NUMBER = layout,
         SP_SHEET = string.IsNullOrWhiteSpace(sheet) ? null : sheet,
         SP_Material = matKey,
         SP_SPOOL_NUMBER = spoolNoKey,
         Line_Sheet_ID_SP = lineSheetId,
         SP_Updated_Date = now ?? AppClock.Now,
         SP_Updated_By = userId ?? HttpContext.Session.GetInt32("UserID")
     };

     _context.SP_Release_tbl.Add(newSp);
     var rows = await _context.SaveChangesAsync();
     _logger.LogInformation("Inserted new SP_Release Spool_ID={SpoolId} (rows={Rows}) for Joint_ID={JointId} Project={Proj} Layout={Layout} Sheet={Sheet}", newSp.Spool_ID, rows, entity.Joint_ID, entity.Project_No, layout, sheet);

     entity.Spool_ID_DFR = newSp.Spool_ID;
     await _context.SaveChangesAsync();
     _logger.LogInformation("Updated DFR Joint_ID={JointId} with new Spool_ID={SpoolId}", entity.Joint_ID, newSp.Spool_ID);
 }

 // New helper: transfer current SP/Coating/Dispatch values into their *_SUPERSEDED columns when FITUP_DATE > SP_Date
 private async Task SupersedeSpoolRelatedDataIfNeededAsync(Dfr dfr)
 {
     if (dfr == null || !dfr.Spool_ID_DFR.HasValue || !dfr.FITUP_DATE.HasValue) return;

     // Apply supersede only for WS / Shop locations based on DFR.LOCATION and project location mapping
     var locRaw = (dfr.LOCATION ?? string.Empty).Trim();
     var locUp = locRaw.ToUpperInvariant();
     var isShopLoc = locUp == "WS" || locUp.StartsWith("WS") || locUp.Contains("SHOP") || locUp.Contains("WORK");
     if (!isShopLoc) return;

     var spId = dfr.Spool_ID_DFR.Value;
     var fitup = dfr.FITUP_DATE.Value;

    static string TrimToNonNull(string? value, int maxLen)
    {
        var t = (value ?? string.Empty).Trim();
        if (t.Length > maxLen) t = t[..maxLen];
        return t;
    }

     // Load SpRelease row (tracked)
     var sp = await _context.SP_Release_tbl.FirstOrDefaultAsync(s => s.Spool_ID == spId);
     if (sp == null) return;
     if (!sp.SP_Date.HasValue || fitup <= sp.SP_Date.Value) return; // only supersede when new fit-up date is after existing SP_Date

     bool changedSp = false;
     bool changedCoating = false;
     bool changedDispatch = false;

     // Detect return-type spools: if SP_TYPE == "R" we will prefix superseded QR fields with "R-" and convert type to "W"
     var isReturnType = string.Equals(sp.SP_TYPE, "R", StringComparison.OrdinalIgnoreCase);

     // 1 & 2: SP_QR_NUMBER -> SP_QR_SUPERSEDED and SP_Date -> SP_Date_SUPERSEDED, then clear original
     if (!string.IsNullOrWhiteSpace(sp.SP_QR_NUMBER))
     {
         var qr = sp.SP_QR_NUMBER!;
         if (isReturnType) qr = "R-" + qr;
        sp.SP_QR_SUPERSEDED = TrimToNonNull(qr, 8); // overwrite even if already set
           sp.SP_QR_NUMBER = string.Empty; // clear original (avoid null constraint issues)
         changedSp = true;
     }
     if (sp.SP_Date.HasValue)
     {
         sp.SP_Date_SUPERSEDED = sp.SP_Date; // overwrite even if already set
         sp.SP_Date = null; // clear original
         changedSp = true;
     }

    // Clear status fields on supersede (use empty string to satisfy non-null columns)
    if (!string.IsNullOrWhiteSpace(sp.SP_STATUS) || sp.SP_STATUS_DATE.HasValue)
    {
        sp.SP_STATUS = string.Empty;
        sp.SP_STATUS_DATE = null;
        changedSp = true;
    }

     // If this spool was of return type, convert it to a working spool now that values have been superseded
     if (isReturnType)
     {
         sp.SP_TYPE = "W";
         changedSp = true;
     }

     // Coating table (key by Spool_ID_PA == Spool_ID)
     var coating = await _context.Coating_tbl.FirstOrDefaultAsync(c => c.Spool_ID_PA == spId);
     if (coating != null)
     {
         bool CopyInt(Func<int?> getSrc, Action<int?> setSrc, Func<int?> getDest, Action<int?> setDest)
         {
             var src = getSrc();
             if (src.HasValue)
             {
                 setDest(src); // overwrite even if already set
                 setSrc(null); // clear original after copying
                 return true;
             }
             return false;
         }
         bool CopyDt(Func<DateTime?> getSrc, Action<DateTime?> setSrc, Func<DateTime?> getDest, Action<DateTime?> setDest)
         {
             var src = getSrc();
             if (src.HasValue)
             {
                 setDest(src); // overwrite even if already set
                 setSrc(null); // clear original after copying
                 return true;
             }
             return false;
         }

         // 3-14 mappings (copy then clear original)
         changedCoating |= CopyInt(() => coating.Surface_Preparation_3_1, v => coating.Surface_Preparation_3_1 = v, () => coating.S_P_3_1_SUPERSEDED, v => coating.S_P_3_1_SUPERSEDED = v);
        changedCoating |= CopyDt(() => coating.Surface_Preparation_3_1_Date, v => coating.Surface_Preparation_3_1_Date = v, () => coating.S_P_3_1_Date_SUPERSEDED, v => coating.S_P_3_1_Date_SUPERSEDED = v);
         changedCoating |= CopyInt(() => coating.Primer_Application_3_2, v => coating.Primer_Application_3_2 = v, () => coating.P_A_3_2_SUPERSEDED, v => coating.P_A_3_2_SUPERSEDED = v);
         changedCoating |= CopyDt(() => coating.Primer_Application_3_2_Date, v => coating.Primer_Application_3_2_Date = v, () => coating.P_A_3_2_Date_SUPERSEDED, v => coating.P_A_3_2_Date_SUPERSEDED = v);
         changedCoating |= CopyInt(() => coating.Primer_Inspection_3_3, v => coating.Primer_Inspection_3_3 = v, () => coating.P_I_3_3_SUPERSEDED, v => coating.P_I_3_3_SUPERSEDED = v);
         changedCoating |= CopyDt(() => coating.Primer_Inspection_3_3_Date, v => coating.Primer_Inspection_3_3_Date = v, () => coating.P_I_3_3_Date_SUPERSEDED, v => coating.P_I_3_3_Date_SUPERSEDED = v);
         changedCoating |= CopyInt(() => coating.Top_Coat_3_4, v => coating.Top_Coat_3_4 = v, () => coating.T_C_3_4_SUPERSEDED, v => coating.T_C_3_4_SUPERSEDED = v);
         changedCoating |= CopyDt(() => coating.Top_Coat_3_4_date, v => coating.Top_Coat_3_4_date = v, () => coating.T_C_3_4_date_SUPERSEDED, v => coating.T_C_3_4_date_SUPERSEDED = v);
         changedCoating |= CopyInt(() => coating.Intermediate_Coat_3_4, v => coating.Intermediate_Coat_3_4 = v, () => coating.I_C_3_4_SUPERSEDED, v => coating.I_C_3_4_SUPERSEDED = v);
         changedCoating |= CopyDt(() => coating.Intermediate_Coat_3_4_date, v => coating.Intermediate_Coat_3_4_date = v, () => coating.I_C_3_4_date_SUPERSEDED, v => coating.I_C_3_4_date_SUPERSEDED = v);
         changedCoating |= CopyInt(() => coating.Final_Coating_3_5, v => coating.Final_Coating_3_5 = v, () => coating.F_C_3_5_SUPERSEDED, v => coating.F_C_3_5_SUPERSEDED = v);
         changedCoating |= CopyDt(() => coating.Final_Coating_3_5_date, v => coating.Final_Coating_3_5_date = v, () => coating.F_C_3_5_date_SUPERSEDED, v => coating.F_C_3_5_date_SUPERSEDED = v);
     }

     // Dispatch Receiving (process all rows where Spool_ID_DR == Spool_ID)
     var dispatchRows = await _context.Dispatch_Receiving_tbl
         .Where(dr => dr.Spool_ID_DR == spId)
         .ToListAsync();
     if (dispatchRows.Count > 0)
     {
         foreach (var dispatch in dispatchRows)
         {
             bool rowChanged = false;
             if (!string.IsNullOrWhiteSpace(dispatch.DR_QR_NUMBER))
             {
                 var drqr = dispatch.DR_QR_NUMBER!;
                 if (isReturnType) drqr = "R-" + drqr;
                dispatch.DR_QR_SUPERSEDED = TrimToNonNull(drqr, 8); // overwrite even if already set
                dispatch.DR_QR_NUMBER = string.Empty; // clear original after copying (avoid null constraint issues)
                 rowChanged = true;
             }
             if (dispatch.DR_Date.HasValue)
             {
                 dispatch.DR_Date_SUPERSEDED = dispatch.DR_Date; // overwrite even if already set
                 dispatch.DR_Date = null; // clear original after copying
                 rowChanged = true;
             }
                if (!string.IsNullOrWhiteSpace(dispatch.Bundle_No))
                {
                    dispatch.Bundle_No = string.Empty; // clear bundle reference when superseding (avoid non-null constraint)
                    rowChanged = true;
                }
             changedDispatch |= rowChanged;
         }
     }

     if (changedSp || changedCoating || changedDispatch)
     {
         await _context.SaveChangesAsync();
         _logger.LogInformation("Supersede transfer applied for Spool_ID={SpoolId} due to FITUP_DATE>{SpDate}", spId, sp.SP_Date_SUPERSEDED);
     }
 }

 // New helper: check if supersede condition would apply
 private async Task<bool> ShouldSupersedeAsync(Dfr dfr)
 {
     if (dfr == null || !dfr.Spool_ID_DFR.HasValue || !dfr.FITUP_DATE.HasValue) return false;

     // Apply supersede only for WS / Shop locations based on DFR.LOCATION and project location mapping
     var locRaw = (dfr.LOCATION ?? string.Empty).Trim();
     var locUp = locRaw.ToUpperInvariant();
     var isShopLoc = locUp == "WS" || locUp.StartsWith("WS") || locUp.Contains("SHOP") || locUp.Contains("WORK");
     if (!isShopLoc) return false;

     var sp = await _context.SP_Release_tbl.AsNoTracking().FirstOrDefaultAsync(s => s.Spool_ID == dfr.Spool_ID_DFR.Value);
     if (sp == null || !sp.SP_Date.HasValue) return false;
     return dfr.FITUP_DATE.Value > sp.SP_Date.Value;
 }
public IActionResult Index() => RedirectToAction("Login");



 [AllowAnonymous]
 [HttpGet]
 public IActionResult Login()
 {
 try
 {
 bool rememberMe = false;
 string? username = null;
 string? password = null;

 if (Request.Cookies.TryGetValue("SavedUsername", out var savedUsername) && !string.IsNullOrWhiteSpace(savedUsername))
 {
 username = savedUsername;
 rememberMe = true;
 _logger.LogInformation("Loaded saved username: {Username}", username);
 }
 if (Request.Cookies.ContainsKey("SavedPwd_v1"))
 {
 Response.Cookies.Delete("SavedPwd_v1");
 }

 var shouldClear = (TempData.Peek("ClearLoginFields") as bool?) == true;
 if (shouldClear)
 {
 username = null;
 password = null;
 rememberMe = false;
 }
 ViewBag.ClearLoginFields = shouldClear;

 ViewBag.Username = username;
 ViewBag.Password = password;
 ViewBag.RememberMe = rememberMe;

 return View();
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Error in GET Login");
 ViewBag.ErrorMessage = "An error occurred while loading the login page";
 return View();
 }
 }

 [AllowAnonymous]
 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> Login(string username, string password, bool rememberMe = false)
 {
 try
 {
 username = (username ?? string.Empty).Trim();
 password = (password ?? string.Empty).Trim();

 _logger.LogInformation("Login attempt for: {Username}", username);

  var normalizedUser = username.ToUpperInvariant();
  var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "unknown";
  if (IsLoginLockedOut(normalizedUser, remoteIp, out var retryAfter))
  {
  ViewBag.ErrorMessage = retryAfter.HasValue
      ? $"Too many attempts. Please try again in {Math.Ceiling(retryAfter.Value.TotalMinutes)} minute(s)."
      : "Too many attempts. Please try again later.";
  ViewBag.Username = username;
  ViewBag.RememberMe = rememberMe;
  return View();
  }

 if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
 {
 ViewBag.ErrorMessage = "Username and password are required";
 ViewBag.Username = username;
 ViewBag.RememberMe = rememberMe;
 return View();
 }

 var user = await _authService.Authenticate(username, password);

 if (user == null)
 {
  RecordLoginFailure(normalizedUser, remoteIp);
 ViewBag.ErrorMessage = "Invalid username or password";
 ViewBag.Username = username;
 ViewBag.RememberMe = rememberMe;
 return View();
 }

 if (!user.Approved)
 {
 ViewBag.ErrorMessage = "Your account is not approved yet. Please contact the administrator.";
 ViewBag.Username = username;
 ViewBag.RememberMe = rememberMe;
  RecordLoginFailure(normalizedUser, remoteIp);
 return View();
 }

 //1) Guard nullable DateTime comparisons
 if (user.DemobilisedDate.HasValue && user.DemobilisedDate.Value < DateTime.UtcNow)
 {
 ViewBag.ErrorMessage = "Your account has been demobilized. Contact the administrator.";
 ViewBag.Username = username;
 ViewBag.RememberMe = rememberMe;
  RecordLoginFailure(normalizedUser, remoteIp);
 return View();
 }

 if (rememberMe)
 {
 // Persist username for "Remember me" with strict SameSite and Secure flag
 Response.Cookies.Append("SavedUsername", username, SecureCookieOptions(DateTimeOffset.UtcNow.AddDays(30)));
 }
 else
 {
 Response.Cookies.Delete("SavedUsername");
 Response.Cookies.Delete("SavedPwd_v1");
 }

  var sessionToken = Convert.ToBase64String(RandomNumberGenerator.GetBytes(32));
  var sessionTokenHash = AuthService.HashToken(sessionToken);
  user.SessionToken = sessionTokenHash;
  user.SessionIssuedUtc = DateTime.UtcNow;
  await _context.SaveChangesAsync();

 HttpContext.Session.SetInt32("UserID", user.UserID);
 HttpContext.Session.SetString("FullName", string.Concat(user.FirstName, " ", user.LastName));
 HttpContext.Session.SetString("UserName", user.UserName ?? string.Empty);
 HttpContext.Session.SetString("Position", user.Position ?? string.Empty);
 HttpContext.Session.SetString("Access", user.Access ?? string.Empty);
  HttpContext.Session.SetString("SessionToken", sessionToken);

  ClearLoginFailures(normalizedUser, remoteIp);

 _logger.LogInformation("Successful login for: {Username}", username);

 // Redirect directly to Dashboard after successful login
 return RedirectToAction(nameof(Dashboard));
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Login failed for: {Username}", username);
 ViewBag.ErrorMessage = "An unexpected error occurred during login";
 ViewBag.Username = username;
 ViewBag.RememberMe = rememberMe;
 return View();
 }
 }

 [HttpGet]
 [ActionName("LoginStatus")]
 public IActionResult Login(string? message, string? returnTo, int delayMs =2000)
 {
 ViewBag.Message = string.IsNullOrWhiteSpace(message) ? "Please wait..." : message;

 var target =!string.IsNullOrWhiteSpace(returnTo) ? Url.Action(returnTo, "Home") : null;
 ViewBag.ReturnUrl = target ?? Url.Action("Login", "Home");

 ViewBag.ReturnText = "Continue";
 ViewBag.DelayMs = delayMs;
 return View("LoadingToLogin");
 }

 [SessionAuthorization]
 public async Task<IActionResult> Dashboard()
 {
    try
    {
        var fullName = HttpContext.Session.GetString("FullName");

        if (string.IsNullOrEmpty(fullName))
        {
            _logger.LogWarning("Dashboard accessed without valid session");
            return RedirectToAction("Login");
        }

        ViewBag.FullName = fullName;
        // Provide project list for status summary selector
        var projects = await _context.Projects_tbl.AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .Select(p => new { p.Project_ID, p.Project_Name })
            .ToListAsync();
        ViewBag.Projects = projects;
        var defaultProjectId = await GetDefaultProjectIdAsync();
        if (!defaultProjectId.HasValue && projects.Count > 0)
        {
            defaultProjectId = projects.Max(p => p.Project_ID);
        }
        ViewBag.DefaultProjectId = defaultProjectId;
        return View();
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Error in Dashboard page");
        return RedirectToAction("Login");
    }
 }

 [HttpGet]
 [AllowAnonymous]
 public IActionResult ForgotPassword() => View();

 [HttpPost]
 [AllowAnonymous]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> ForgotPassword(string email, string recoveryType = "password")
 {
    try
    {
        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "unknown";
        if (IsForgotPasswordThrottled(remoteIp))
        {
            ViewBag.Message = "Too many requests. Please try again later.";
            return View();
        }
        RecordForgotPasswordRequest(remoteIp);

        _logger.LogInformation("Account recovery requested type={Type} for: {Email}", recoveryType, email);
        if (string.IsNullOrWhiteSpace(email)) { ViewBag.Message = "Please enter your registered email address."; return View(); }

        if (string.Equals(recoveryType, "username", StringComparison.OrdinalIgnoreCase))
        {
            var usernames = await _context.PMS_Login_tbl.AsNoTracking()
                .Where(u => u.Email == email && !string.IsNullOrEmpty(u.UserName))
                .Select(u => u.UserName!)
                .Distinct()
                .ToListAsync();
            if (usernames.Count > 0)
            {
                var usernameList = string.Join(", ", usernames.Select(System.Net.WebUtility.HtmlEncode));
                var content = $@"<p style='margin:0 0 10px;'>Hello,</p>
<p style='margin:0 0 10px;'>We received a request to remind you of the username(s) associated with this email address.</p>
<p style='margin:0 0 10px;'><strong>Username(s):</strong> {usernameList}</p>
<p style='margin:0;'>If you did not request this reminder, you can safely ignore this email.</p>";
                var (html, resources) = await BuildBrandEmailAsync("Username Reminder", content);
                try
                {
                    if (resources != null) await _emailService.SendEmailAsync(email, "Username Reminder", html, resources); else await _emailService.SendEmailAsync(email, "Username Reminder", html);
                    _logger.LogInformation("Username reminder sent for {Email} (count={Count})", email, usernames.Count);
                }
                catch (Exception sendEx) { _logger.LogError(sendEx, "Failed sending username reminder for {Email}", email); }
            }
            else { _logger.LogWarning("Username reminder requested for unknown email: {Email}", email); }
            ViewBag.Message = "If the email exists, a message with your username(s) has been sent. If you do not receive an email within a few minutes, please contact support.";
            return View();
        }

        // Password reset flow
        if (string.IsNullOrWhiteSpace(_emailSettings.User) || string.IsNullOrWhiteSpace(_emailSettings.Password))
        { _logger.LogError("Email credentials not configured"); ViewBag.Message = "Email service configuration error. Contact administrator."; return View(); }

        var user = await _authService.GetUserByEmail(email);
        if (user != null)
        {
            var token = _authService.GenerateResetToken();
            await _authService.UpdateResetToken(email, token);
            var resetLink = Url.Action("ResetPassword", "Home", new { token }, Request.Scheme);

            // Embed the PMS email logo at a larger size (60px height) in all clients
            string cid = "pmslogo";
            Dictionary<string, (byte[] Content, string ContentType)>? resources = null;
            string logoTag;
            try
            {
                var logoPath = Path.Combine(_env.WebRootPath ?? string.Empty, "img", "PMS_logo_email.png");
                if (System.IO.File.Exists(logoPath))
                {
                    resources = new() { [cid] = (await System.IO.File.ReadAllBytesAsync(logoPath), "image/png") };
                    logoTag = $"<img src='cid:{cid}' alt='PMS Logo' height='60' style='height:60px;width:auto;display:inline-block;border:0;margin:0;vertical-align:middle;' />";
                }
                else
                {
                    var absolute = $"{Request.Scheme}://{Request.Host}{Url.Content("~/img/PMS_logo_email.png")}";
                    logoTag = $"<img src='{absolute}' alt='PMS Logo' height='60' style='height:60px;width:auto;display:inline-block;border:0;margin:0;vertical-align:middle;' />";
                }
            }
            catch
            {
                var absolute = $"{Request.Scheme}://{Request.Host}{Url.Content("~/img/PMS_logo_email.png")}";
                logoTag = $"<img src='{absolute}' alt='PMS Logo' height='60' style='height:60px;width:auto;display:inline-block;border:0;margin:0;vertical-align:middle;' />";
            }
            
            // Button: remove background, set text color to brand color, remove padding background look
            var emailBody = $@"<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'/><title>Password Reset - PMS</title></head><body style='margin:0;padding:24px;background:#ffffff;color:#176d8a;font-family:Calibri,Arial,Helvetica,sans-serif;'>
<div style='max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #ffffff;border-radius:8px;box-shadow:0 2px 8px rgba(31,153,190,0.15);overflow:hidden;'>
  <div style='padding:14px 18px;text-align:center;background:#ffffff;'>{logoTag}</div>
  <div style='padding:20px 22px;'>
    <h2 style='margin:0 0 14px;color:#176d8a;font-weight:700;font-size:20px;'>Password Reset Request</h2>
    <p style='margin:0 0 10px;'>Hello,</p>
    <p style='margin:0 0 10px;'>We received a request to reset the password associated with this email address. If you made this request, click the link below to set a new password.</p>
    <p style='text-align:center;margin:24px 0;'>
      <a href='{resetLink}' style='color:#176d8a;font-weight:700;font-size:15px;letter-spacing:.5px;text-decoration:underline;background:none;border:none;padding:0;'>Reset Password</a>
    </p>
    <p style='margin:0 0 8px;'>If the link above does not work, copy and paste this URL into your browser:</p>
    <p style='margin:0 0 16px;word-break:break-all;'><a href='{resetLink}' style='color:#1f99be;text-decoration:underline;'>{resetLink}</a></p>
    <p style='margin:0 0 8px;'>This link will expire in 30 minutes for security purposes.</p>
    <p style='margin:0;'>If you did not request a password reset, you can ignore this email and your password will remain unchanged.</p>
  </div>
  <div style='background:#ffffff;border-top:1px solid #e3e9eb;padding:12px 16px;text-align:center;'>
    <p style='margin:0;color:#7a9aa8;font-size:12px;'>Automated message from PMS System. Do not reply.</p>
  </div>
</div></body></html>";
            try
            {
                if (resources != null) await _emailService.SendEmailAsync(email, "Password Reset Request", emailBody, resources); else await _emailService.SendEmailAsync(email, "Password Reset Request", emailBody);
                _logger.LogInformation("Password reset link generated and email queued for {Email}", email);
            }
            catch (Exception sendEx) { _logger.LogError(sendEx, "Failed sending password reset email for {Email}", email); }
        }
        else { _logger.LogWarning("Password reset requested for unknown email: {Email}", email); }

        ViewBag.Message = "If the email exists, a reset link has been generated. If you do not receive an email within a few minutes, please contact support.";
        return View();
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Account recovery error type={Type} for {Email}", recoveryType, email);
        ViewBag.Message = "Failed to process your request. Please try again later.";
        return View();
    }
 }

 [HttpGet]
 public IActionResult ResetPassword(string token)
 {
 if (string.IsNullOrEmpty(token))
 {
 _logger.LogWarning("ResetPassword accessed without token");
 return RedirectToAction("Login");
 }
 ViewBag.Token = token;
 return View();
 }

 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> ResetPassword(string token, string newPassword)
 {
 try
 {
 _logger.LogInformation("Password reset attempt");

 if (string.IsNullOrWhiteSpace(token))
 {
 ViewBag.ErrorMessage = "Invalid or expired token";
 ViewBag.Token = token;
 _logger.LogWarning("ResetPassword called without token");
 return View();
 }

 if (string.IsNullOrWhiteSpace(newPassword))
 {
 ViewBag.ErrorMessage = "Password is required";
 ViewBag.Token = token;
 return View();
 }

 var hashedToken = _authService.HashResetTokenForLookup(token);
 var nowUtc = PMS.Infrastructure.AppClock.UtcNow;

 var user = await _context.PMS_Login_tbl
 .FirstOrDefaultAsync(u => u.ResetToken != null && u.ResetToken == hashedToken && u.ResetExpiry > nowUtc);

 if (user == null)
 {
 ViewBag.ErrorMessage = "Invalid or expired token";
 ViewBag.Token = token;
 _logger.LogWarning("Invalid reset token used");
 return View();
 }

 if (!StrongPasswordRegex().IsMatch(newPassword))
 {
 ViewBag.ErrorMessage = PMS.Infrastructure.Messages.PasswordPolicy;
 ViewBag.Token = token;
 return View();
 }
 user.Password = HashPassword(user, newPassword);
 user.ResetToken = null;
 user.ResetExpiry = null;

 await _context.SaveChangesAsync();

 TempData["SuccessMessage"] = "Password reset successful. Please login with your new password.";
 _logger.LogInformation("Password reset successful for user: {UserName}", user.UserName);
 return RedirectToAction("Login");
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Error resetting password");
 ViewBag.ErrorMessage = "An error occurred while resetting your password";
 ViewBag.Token = token;
 return View();
 }
 }

 [HttpGet]
 public IActionResult SignUp() => View();

 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> SignUp(PMSLogin newUser)
 {
 try
 {
 _logger.LogInformation("Signup attempt for: {UserName}", newUser.UserName);

 if (!ModelState.IsValid) return View(newUser);

 var existingUser = await _context.PMS_Login_tbl
 .AnyAsync(u => u.UserName == newUser.UserName);

 if (existingUser)
 {
 ViewBag.ErrorMessage = "Username already exists";
 _logger.LogWarning("Signup failed - username exists: {UserName}", newUser.UserName);
 return View(newUser);
 }

 var confirmPassword = Request.Form["confirmPassword"].ToString();
 if (!string.Equals(newUser.Password ?? string.Empty, confirmPassword ?? string.Empty, StringComparison.Ordinal))
 {
 ViewBag.ErrorMessage = "Passwords do not match.";
 return View(newUser);
 }

 var pwd = newUser.Password ?? string.Empty;
 if (!StrongPasswordRegex().IsMatch(pwd))
 {
 ViewBag.ErrorMessage = PMS.Infrastructure.Messages.PasswordPolicy;
 return View(newUser);
 }
 newUser.Password = HashPassword(newUser, pwd);

 newUser.Approved = false;
 newUser.CreatedDate = DateTime.UtcNow;

 await _authService.CreateUser(newUser);

 ViewBag.Message = "Your account has been created. It will be activated after approval";
 ViewBag.ReturnUrl = Url.Action("Login", "Home");
 ViewBag.DelayMs =5000;
 _logger.LogInformation("New user created: {UserName}", newUser.UserName);
 return View("LoadingAfterSignUp");
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Error during signup for: {UserName}", newUser.UserName);
 ViewBag.ErrorMessage = "An error occurred during registration";
 return View(newUser);
 }
 }

 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> Logout()
 {
 try
 {
 var userId = HttpContext.Session.GetInt32("UserID");
 if (userId.HasValue)
 {
     var user = await _context.PMS_Login_tbl.FirstOrDefaultAsync(u => u.UserID == userId.Value);
     if (user != null)
     {
         user.SessionToken = null;
         user.SessionIssuedUtc = null;
         await _context.SaveChangesAsync();
     }
 }

 HttpContext.Session.Clear();

 bool hasRemembered =
 Request.Cookies.ContainsKey("SavedUsername") ||
 Request.Cookies.ContainsKey("SavedPwd_v1");

 TempData["ClearLoginFields"] = !hasRemembered;
 _logger.LogInformation("User logged out");
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Error during logout");
 }
 return RedirectToAction("Login");
 }

 [HttpGet]
 public IActionResult ContactUs() => View();

 [HttpGet]
 public IActionResult AboutUs() => View();

 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> ContactUs(ContactModel model)
 {
 try
 {
 if (!ModelState.IsValid) return View(model);

 var adminEmail = _emailSettings.AdminEmail;
 if (string.IsNullOrWhiteSpace(adminEmail))
 {
 ModelState.AddModelError("", "Contact form is currently unavailable. Please try again later.");
 return View(model);
 }

 var safeMessageHtml = model.Message ?? string.Empty; // If allowing HTML, sanitize it with a whitelist here.
 var safeMessageText = HtmlTagsRegex().Replace(safeMessageHtml, string.Empty);
 var emailBody = $@"
 <h3>New Contact Form Submission</h3>
 <p><strong>Name:</strong> {model.Name}</p>
 <p><strong>Email:</strong> {model.Email}</p>
 <p><strong>Subject:</strong> {model.Subject}</p>
 <p><strong>Message:</strong></p>
 <p>{safeMessageHtml}</p>
 <p><strong>Message (plain text):</strong></p>
 <p>{System.Net.WebUtility.HtmlEncode(safeMessageText)}</p>
 ";

 await _emailService.SendEmailAsync(adminEmail!, "PMS Contact Form Submission", emailBody);

 ViewBag.Message = "Your message has been sent successfully. We appreciate your patience and will get back to you as soon as possible.";
 ViewBag.ReturnTo = "Login";
 return View("LoadingConLogin");
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Error sending contact form");
 ModelState.AddModelError("", "An error occurred while sending your message. Please try again.");
 return View(model);
 }
 }

 [SessionAuthorization]
 [HttpGet]
 public IActionResult InContactUs() => View();

 [SessionAuthorization]
 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> InContactUs(ContactModel model)
 {
 try
 {
 if (!ModelState.IsValid) return View(model);

 var internalEmail = _emailSettings.InternalContactEmail;
 if (string.IsNullOrWhiteSpace(internalEmail))
 {
 ModelState.AddModelError("", "Internal contact is currently unavailable.");
 return View(model);
 }

 var safeMessageHtml = model.Message ?? string.Empty; // If allowing HTML, sanitize it with a whitelist here.
 var safeMessageText = HtmlTagsRegex().Replace(safeMessageHtml, string.Empty);
 var emailBody = $@"
 <h3>New Internal Contact Request</h3>
 <p><strong>Name:</strong> {model.Name}</p>
 <p><strong>Email:</strong> {model.Email}</p>
 <p><strong>Subject:</strong> {model.Subject}</p>
 <p><strong>Message:</strong></p>
 <p>{safeMessageHtml}</p>
 <p><strong>Message (plain text):</strong></p>
 <p>{System.Net.WebUtility.HtmlEncode(safeMessageText)}</p>
 ";

 await _emailService.SendEmailAsync(internalEmail!, "Internal Contact Request", emailBody);

 // Success in InContactUs
 TempData["FlashMessage"] = "Your message has been sent successfully. We appreciate your patience and will get back to you as soon as possible.";
 ViewBag.ReturnTo = "Dashboard";
 return View("LoadingInConLogin");
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "Error sending internal contact form");
 ModelState.AddModelError("", "An error occurred while sending your message. Please try again.");
 return View(model);
 }
 }

 [HttpGet]
 [Route("/test-db")]
 public async Task<IActionResult> TestDatabase()
 {
 try
 {
 var users = await _context.PMS_Login_tbl.ToListAsync();
 return Content($"Database connection successful! Found {users.Count} users.");
 }
 catch (Exception ex)
 {
 return Content($"Database connection failed: {ex.Message}");
 }
 }

 [ResponseCache(Duration =0, Location = ResponseCacheLocation.None, NoStore = true)]
 public IActionResult Error() => View(new ErrorViewModel());

 // New: validate Sheet and Spool No changes, update references
 [SuppressMessage("Performance", "CA1862", Justification = "EF Core translates Trim/ToUpper server-side; StringComparison overloads are nottranslated in LINQ to Entities")]
 private async Task<bool> UpdateDfrReferencesIfChangedAsync(Dfr entity, DfrRowUpdateDto dto)
 {
 if (entity == null || dto == null) return false;

 bool updated = false;
 string? layout = entity.LAYOUT_NUMBER?.Trim();
 string? targetSheet = (dto.Sheet ?? entity.SHEET)?.Trim();
 string? spoolNoKey = (entity.SPOOL_NUMBER ?? dto.SpoolNo)?.Trim();

 if (string.IsNullOrWhiteSpace(layout) || string.IsNullOrWhiteSpace(targetSheet))
 return false;

 // Lookup Line_Sheet for target layout/sheet
 var ls = await _context.Line_Sheet_tbl.AsNoTracking()
 .Where(x => x.LS_LAYOUT_NO == layout && x.LS_SHEET == targetSheet)
 .Select(x => new { x.Line_Sheet_ID, x.Line_ID_LS })
 .FirstOrDefaultAsync();

 if (ls != null)
 {
 entity.Line_Sheet_ID_DFR = ls.Line_Sheet_ID;
 updated = true;
 }

 // Determine material via LINE_LIST for this line sheet
 var material = (ls?.Line_ID_LS != null)
 ? await _context.LINE_LIST_tbl.AsNoTracking()
 .Where(l => l.Line_ID == ls.Line_ID_LS)
 .Select(l => l.Material)
 .FirstOrDefaultAsync()
 : null;

 string? matKey = string.IsNullOrWhiteSpace(material) ? null : material!.Trim();

 // Try to find existing SP_Release row
 SpRelease? sp = null;

 // First: match by DWG+SHEET+SPOOL_NUMBER when available (requested SQL behavior)
 if (!string.IsNullOrWhiteSpace(spoolNoKey))
 {
 var spoolNoUpper = spoolNoKey.ToUpper();
 sp = await _context.SP_Release_tbl.AsNoTracking()
 .Where(s => s.SP_Project_No == entity.Project_No
 && s.SP_TYPE == "W"
 && s.SP_LAYOUT_NUMBER == layout
 && s.SP_SPOOL_NUMBER != null
 && s.SP_SPOOL_NUMBER.ToUpper() == spoolNoUpper)
 .OrderByDescending(s => s.Spool_ID)
 .FirstOrDefaultAsync();
 }

 if (sp != null)
 {
 entity.Spool_ID_DFR = sp.Spool_ID;
 updated = true;
 }
 else if (string.Equals(entity.LOCATION?.Trim(), "WS", StringComparison.OrdinalIgnoreCase))
 {
     // Compute SP_DIA as Max(DIAMETER) for the same Project/Layout/Sheet
     double? maxDia = await _context.DFR_tbl.AsNoTracking()
         .Where(d => d.Project_No == entity.Project_No && d.LAYOUT_NUMBER == layout && d.SHEET == targetSheet)
         .Where(d => d.DIAMETER.HasValue)
         .MaxAsync(d => (double?)d.DIAMETER);

     var newSp = new SpRelease
     {
     SP_Project_No = entity.Project_No,
     SP_TYPE = "W",
     SP_DIA = maxDia,
     SP_LAYOUT_NUMBER = layout,
     SP_SHEET = targetSheet,
     SP_Material = matKey,
     SP_SPOOL_NUMBER = spoolNoKey
     };

     _context.SP_Release_tbl.Add(newSp);
     await _context.SaveChangesAsync();

     entity.Spool_ID_DFR = newSp.Spool_ID;
     updated = true;
 }

 return updated;
 }

 [SessionAuthorization]
 [HttpPost]
 [ValidateAntiForgeryToken]
 public async Task<IActionResult> UpdateDfrRow([FromBody] DfrRowUpdateDto dto)
 {
 if (dto == null || dto.JointId <=0) return BadRequest("Invalid data");
 var entity = await _context.DFR_tbl.FirstOrDefaultAsync(d => d.Joint_ID == dto.JointId);
 if (entity == null) return NotFound();

 // Allow unchecking Deleted/Cancelled without locking; otherwise lock when flagged
 if ((entity.Deleted || entity.Cancelled))
 {
     bool isAttemptingToUncheck =
         (entity.Deleted && dto.Deleted.HasValue && dto.Deleted.Value == false) ||
         (entity.Cancelled && dto.Cancelled.HasValue && dto.Cancelled.Value == false);
     if (!isAttemptingToUncheck)
         return Conflict("Row locked");
 }
 if (entity.Fitup_Confirmed && !(dto.FitupConfirmed.HasValue && dto.FitupConfirmed.Value == false))
 return Conflict("Row locked");

 // Validation: Fit-up Date and Fit-up Report must be paired
 if (!string.IsNullOrWhiteSpace(dto.FitupReport) && !dto.FitupDate.HasValue)
 {
 return BadRequest("Fit-up Date is required when Fit-up Report is provided.");
 }
 if (dto.FitupDate.HasValue && string.IsNullOrWhiteSpace(dto.FitupReport))
 {
 return BadRequest("Fit-up Report is required when Fit-up Date is provided.");
 }

 if (entity.Fitup_Confirmed && dto.FitupConfirmed == false)
 {
     entity.Fitup_Confirmed = false;
 }
 else
 {
     entity.WELD_TYPE = NormalizeWeldType(dto.WeldType ?? entity.WELD_TYPE);
     entity.DFR_REV = dto.Rev;
     entity.SPOOL_NUMBER = dto.SpoolNo;
     entity.DIAMETER = dto.DiaIn;
     entity.SCHEDULE = dto.Sch;
     entity.MATERIAL_A = dto.MaterialA;
     entity.MATERIAL_B = dto.MaterialB;
     entity.GRADE_A = dto.GradeA;
     entity.GRADE_B = dto.GradeB;
     entity.HEAT_NUMBER_A = dto.HeatNumberA;
     entity.HEAT_NUMBER_B = dto.HeatNumberB;
     // WPS: set by id, resolve by text, or clear when blank
     if (dto.WpsId.HasValue && dto.WpsId.Value >0)
     {
         entity.WPS_ID_DFR = dto.WpsId.Value;
     }
     else if (!string.IsNullOrWhiteSpace(dto.Wps))
     {
         var wpsId = await _context.WPS_tbl.AsNoTracking()
             .Where(w => w.WPS == dto.Wps)
             .Select(w => (int?)w.WPS_ID)
             .FirstOrDefaultAsync();
         entity.WPS_ID_DFR = wpsId; // null if not found
     }
     else
     {
         entity.WPS_ID_DFR = null;
     }
     if (dto.FitupDate.HasValue)
     {
         var dtLocal = AppClock.ToProjectLocal(dto.FitupDate.Value);
         var (ok, message) = await ValidateFitupDateAgainstConstraintsAsync(dto.JointId, dtLocal);
         if (!ok) return BadRequest($"Joint {await GetJointDisplayAsync(dto.JointId)}: {message}");
         entity.FITUP_DATE = dtLocal;
     }
     else
     {
         entity.FITUP_DATE = null;
     }
     entity.FITUP_INSPECTION_QR_NUMBER = dto.FitupReport;
     entity.TACK_WELDER = dto.TackWelder;
     entity.OL_DIAMETER = dto.OlDia;
     entity.OL_SCHEDULE = dto.OlSch;
     entity.OL_Thick = dto.OlThick;
     entity.DFR_REMARKS = dto.Remarks;
     entity.LOCATION = Clean(dto.Location,4) ?? entity.LOCATION;
     entity.LAYOUT_NUMBER = Clean(dto.LayoutNumber,10) ?? entity.LAYOUT_NUMBER;
     if (!string.IsNullOrWhiteSpace(dto.JAdd)) entity.J_Add = CleanSelect(dto.JAdd,8)?.ToUpperInvariant();
     if (!string.IsNullOrWhiteSpace(dto.Sheet)) entity.SHEET = Clean(dto.Sheet,5) ?? entity.SHEET;
     if (!string.IsNullOrWhiteSpace(dto.WeldNumber)) entity.WELD_NUMBER = Clean(dto.WeldNumber,6);
     if (dto.Deleted.HasValue) entity.Deleted = dto.Deleted.Value;
     if (dto.Cancelled.HasValue) entity.Cancelled = dto.Cancelled.Value;
     if (dto.FitupConfirmed.HasValue) entity.Fitup_Confirmed = dto.FitupConfirmed.Value;
     // RFI: assign directly; null/<=0 clears
     entity.RFI_ID_DFR = (dto.RfiId.HasValue && dto.RfiId.Value >0) ? dto.RfiId.Value : (int?)null;
 }
 // Set Updated By / Updated Date on every successful edit attempt
 entity.DFR_Updated_By = HttpContext.Session.GetInt32("UserID");
 entity.DFR_Updated_Date = AppClock.Now;

 // Write Actual Welding Date and crew to DWR when available
 // Welding persistence handled by dedicated DWR endpoints

 await _context.SaveChangesAsync();

 await UpdateLineSheetAndSpoolRefsAsync(entity, HttpContext.Session.GetInt32("UserID"), AppClock.Now);

 // After linking, prune any orphan W-type spools in the same layout/sheet scope
 if (!string.IsNullOrWhiteSpace(entity.LAYOUT_NUMBER) && !string.IsNullOrWhiteSpace(entity.SHEET))
 {
     _ = await PruneUnusedSpReleaseRowsForScopeAsync(entity.Project_No, entity.LAYOUT_NUMBER.Trim(), entity.SHEET.Trim());
 }

 // Supersede confirmation gate
 bool needSupersede = await ShouldSupersedeAsync(entity);
 if (needSupersede && dto.ConfirmSupersede != true)
 {
     var jointDisp = await GetJointDisplayAsync(entity.Joint_ID);
     return Conflict(new { code = "requireSupersedeConfirm", id = entity.Joint_ID, joint = jointDisp, message = $"You are making modification on released spool, are you sure you want to supersede the previous reports for Joint \"{jointDisp}\"?" });
 }

 // Apply supersede (no-op when not needed)
 await SupersedeSpoolRelatedDataIfNeededAsync(entity);

 return Ok(new { success = true });
 }

 [HttpGet]
 [SuppressMessage("Performance", "CA1862", Justification = "EF Core translates Trim/ToUpper server-side; StringComparison overloads are not translated in LINQ to Entities")]
 public async Task<IActionResult> GetHeatNumbers(
 int projectId,
 [FromQuery(Name = "side")] string _,
 double? diaIn,
 double? olDiaIn,
 string? material,
 string? grade,
 string? sch,
 string? olSch)
 {
 try
 {
        if (projectId <=0) return BadRequest("projectId required");

        var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);
        if (weldersProjectId <= 0) return BadRequest("projectId required");
 material = (material ?? string.Empty).Trim();
 grade = (grade ?? string.Empty).Trim();
 var schKey = (sch ?? string.Empty).Trim().ToUpperInvariant();
 var olSchKey = (olSch ?? string.Empty).Trim().ToUpperInvariant();
 if (string.IsNullOrWhiteSpace(material) || string.IsNullOrWhiteSpace(grade) || !diaIn.HasValue || string.IsNullOrWhiteSpace(schKey))
 {
 return Json(new { thickness = (double?)null, source = "invalid" });
 }

 var matKey = material.Trim().ToUpperInvariant();
 var grd = grade.Trim().ToUpperInvariant();
 bool allowPipePad = string.Equals(matKey, "PIPE", StringComparison.OrdinalIgnoreCase) || string.Equals(matKey, "PAD", StringComparison.OrdinalIgnoreCase);

        var q = _context.MATERIAL_TRACE_tbl.AsNoTracking()
        .Where(m => m.MATR_Project_No == weldersProjectId
 && m.Material_Des != null && m.MATR_GRADE != null
 && m.HEAT_NO != null && m.HEAT_NO != "" );

 q = q.Where(m =>
 (
 allowPipePad
 ? (((m.Material_Des ?? "").Trim().ToUpper() == "PIPE") || ((m.Material_Des ?? "").Trim().ToUpper() == "PAD"))
 : ((m.Material_Des ?? "").Trim().ToUpper() == matKey)
 )
 && ((m.MATR_GRADE ?? "").Trim().ToUpper() == grd)
 );

 var d1 = diaIn.Value;
 double? d2 = olDiaIn;
 bool hasOl = d2.HasValue && d2.Value >0 && !string.IsNullOrWhiteSpace(olSchKey);

 q = q.Where(m =>
 (
 m.Dia_In1 == d1 && !hasOl && m.Dia_In2 == null &&
 ((m.SC1 ?? "").Trim().ToUpper() == schKey) && (m.SC2 == null || (m.SC2 ?? "").Trim() == "")
 )
 || (
 hasOl && m.Dia_In1 == d1 && m.Dia_In2 == d2 &&
 ((m.SC1 ?? "").Trim().ToUpper() == schKey) && ((m.SC2 ?? "").Trim().ToUpper() == olSchKey)
 )
 || (
 hasOl && m.Dia_In2 == d1 && m.Dia_In1 == d2 &&
 ((m.SC1 ?? "").Trim().ToUpper() == schKey) && ((m.SC2 ?? "").Trim().ToUpper() == olSchKey)
 )
 );

 var heats = await q
 .GroupBy(m => m.HEAT_NO!)
 .Select(g => new { heat = g.Key, description = g.Select(x => x.Description).FirstOrDefault() })
 .OrderBy(x => x.heat)
 .Take(300)
 .ToListAsync();

 return Json(heats);
 }
 catch (Exception ex)
 {
 _logger.LogError(ex, "GetHeatNumbers failed");
 return Json(Array.Empty<object>());
 }
    }

    // Daily Fit-up actions moved to HomeController.DailyFitup.cs
}

public class ClientLogEntry
{
    public required string Msg { get; set; }
    public required string Level { get; set; }
    public required string Timestamp { get; set; }
    public string? User { get; set; }
    public string? Ip { get; set; }
    public string? Path { get; set; }
}