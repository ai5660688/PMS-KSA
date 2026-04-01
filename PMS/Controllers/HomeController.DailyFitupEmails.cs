using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Data;
using PMS.Models;
using PMS.Services;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using PMS.Infrastructure;
using System.IO; // added for logo file access

namespace PMS.Controllers
{
    public partial class HomeController : Controller
    {
        // POST: /Home/SendDailyFitupEmail
        [SessionAuthorization]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SendDailyFitupEmail([FromForm] int projectId, [FromForm] string location, [FromForm] string fitupDateIso, [FromForm] string? actionLabel, [FromForm] string kind)
        {
            try
            {
                if (projectId <= 0 || string.IsNullOrWhiteSpace(fitupDateIso))
                    return Json(new { success = false, message = "Missing inputs" });

                var isCompleted = string.Equals(kind, "completed", StringComparison.OrdinalIgnoreCase);
                var razorKey = isCompleted ? "Daily_Fit-up_Completed" : "Daily_Fit-up_Confirmed";

                DateTime fitupDate;
                if (!DateTime.TryParse(fitupDateIso, out fitupDate))
                {
                    if (!DateTime.TryParseExact(fitupDateIso, new[] { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss" }, null, System.Globalization.DateTimeStyles.AssumeLocal, out fitupDate))
                        fitupDate = DateTime.Now;
                }
                var day = fitupDate.Date;

                // Normalize heading text variants
                if (string.Equals(actionLabel?.Trim(), "Update Daily Fit-up", StringComparison.OrdinalIgnoreCase) 
                    || string.Equals(actionLabel?.Trim(), "Updated Daily Fit-up", StringComparison.OrdinalIgnoreCase))
                    actionLabel = "Daily Fit-up has been Updated";
                else if (string.Equals(actionLabel?.Trim(), "Update confirmed Fit-up", StringComparison.OrdinalIgnoreCase))
                    actionLabel = "Updated confirmed Fit-up"; // preserve previous mapping first
                if (string.Equals(actionLabel?.Trim(), "Updated confirmed Fit-up", StringComparison.OrdinalIgnoreCase))
                    actionLabel = "Confirmed Fit-up has been Updated";

                string MapHeaderLocation(string raw)
                {
                    var s = (raw ?? string.Empty).Trim().ToUpperInvariant();
                    if (s.StartsWith("WS") || s.StartsWith("SHOP") || s.StartsWith("WORK")) return "WS";
                    if (s.StartsWith("FW") || s.StartsWith("FIELD")) return "FW";
                    if (s.StartsWith("TH") || s.Contains("THREAD")) return "TH";
                    if (s == "ALL") return "All";
                    return s;
                }
                var locCodeFromHeader = MapHeaderLocation(location);
                bool headerAll = string.Equals(locCodeFromHeader, "All", StringComparison.OrdinalIgnoreCase);

                var ucRows = await _context.PMS_Updated_Confirmed_tbl.AsNoTracking()
                    .Where(x => x.U_C_Project_No == projectId && x.Updated_Confirmed_Date.Date == day)
                    .ToListAsync();

                UpdatedConfirmed? uc = null;
                if (!headerAll)
                {
                    var targetLoc = (locCodeFromHeader == "WS" || locCodeFromHeader == "FW" || locCodeFromHeader == "TH") ? locCodeFromHeader : null;
                    var qry = ucRows.AsQueryable();
                    if (targetLoc != null) qry = qry.Where(r => r.U_C_Location == targetLoc).AsQueryable();
                    uc = qry.OrderByDescending(r => r.Updated_Confirmed_Date).FirstOrDefault();
                }

                var receivers = await ResolveUCDailyReceiversAsync(projectId, razorKey, headerAll ? "All" : (locCodeFromHeader == "WS" || locCodeFromHeader == "FW" || locCodeFromHeader == "TH" ? locCodeFromHeader : "All"));
                if (receivers.To.Count == 0 && receivers.Cc.Count == 0)
                {
                    return Json(new { success = false, message = "No receivers configured" });
                }

                string title = actionLabel ?? string.Empty;
                string FormatDec(decimal? v) => v.HasValue ? string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.##}", v.Value) : "-";

                var projectEntity = await _context.Projects_tbl.AsNoTracking().FirstOrDefaultAsync(p => p.Project_ID == projectId);
                var projectDisplay = projectEntity != null ? ($"{projectEntity.Project_ID} - {projectEntity.Project_Name}") : projectId.ToString();
                var refDay = day;

                var bodyBuilder = new StringBuilder();
                bodyBuilder.Append("<div style='color:#176d8a;font-size:14px;line-height:1.5'>");
                bodyBuilder.Append("<p style='margin:0 0 12px'>This is to inform you that the following action has been recorded in PMS:</p>");
                bodyBuilder.Append("<table style='border-collapse:collapse;width:100%;max-width:600px'>");
                void Row(string k, string v, string? firstWidth = null, string? secondWidth = null) => bodyBuilder.Append($"<tr><td style='padding:6px 8px;border:1px solid #e3e9eb;background:#f9fbfc;{(string.IsNullOrEmpty(firstWidth) ? "width:40%;" : $"width:{firstWidth};")}'><strong>{System.Net.WebUtility.HtmlEncode(k)}</strong></td><td style='padding:6px 8px;border:1px solid #e3e9eb;{(string.IsNullOrEmpty(secondWidth) ? "" : $"width:{secondWidth};")}'>{System.Net.WebUtility.HtmlEncode(v)}</td></tr>");
                Row("Project No.", projectDisplay, "18%", "82%");
                Row("Daily Fit-up Date", refDay.ToString("dd-MMM-yyyy"));

                if (headerAll)
                {
                    var ucWs = ucRows.FirstOrDefault(r => r.U_C_Location == "WS");
                    var ucFw = ucRows.FirstOrDefault(r => r.U_C_Location == "FW");
                    var ucTh = ucRows.FirstOrDefault(r => r.U_C_Location == "TH");

                    string Name(string code) => code switch { "WS" => "Workshop", "FW" => "Field", "TH" => "Threaded", _ => code };

                    var locationNames = new List<string>();
                    if (ucWs != null) locationNames.Add(Name("WS"));
                    if (ucFw != null) locationNames.Add(Name("FW"));
                    if (ucTh != null) locationNames.Add(Name("TH"));
                    if (locationNames.Count > 0) Row("Location", string.Join(" / ", locationNames));

                    var updatedValues = new List<string>();
                    if (ucWs != null) updatedValues.Add(FormatDec(ucWs.Fitup_Dia));
                    if (ucFw != null) updatedValues.Add(FormatDec(ucFw.Fitup_Dia));
                    if (ucTh != null) updatedValues.Add(FormatDec(ucTh.Fitup_Dia));
                    if (updatedValues.Count > 0) Row("Fit-up Dia. In. (Updated)", string.Join(" / ", updatedValues));

                    var confirmedValues = new List<string>();
                    if (ucWs != null) confirmedValues.Add(FormatDec(ucWs.Fitup_Confirmed_Dia));
                    if (ucFw != null) confirmedValues.Add(FormatDec(ucFw.Fitup_Confirmed_Dia));
                    if (ucTh != null) confirmedValues.Add(FormatDec(ucTh.Fitup_Confirmed_Dia));
                    if (confirmedValues.Count > 0) Row("Fit-up Dia. In. (Confirmed)", string.Join(" / ", confirmedValues));
                }
                else
                {
                    string Name(string code) => code switch { "WS" => "Workshop", "FW" => "Field", "TH" => "Threaded", _ => code };
                    var locCode = (uc?.U_C_Location ?? locCodeFromHeader);
                    Row("Location", Name(locCode));
                    if (uc != null)
                    {
                        Row("Fit-up Dia. In. (Updated)", FormatDec(uc.Fitup_Dia));
                        Row("Fit-up Dia. In. (Confirmed)", FormatDec(uc.Fitup_Confirmed_Dia));
                    }
                    else
                    {
                        Row("Fit-up Dia. In. (Updated)", "-");
                        Row("Fit-up Dia. In. (Confirmed)", "-");
                    }
                }

                bodyBuilder.Append("</table>");
                bodyBuilder.Append("<p style='margin:12px 0 0'>Regards,<br/>PMS System</p>");
                bodyBuilder.Append("</div>");

                var (emailHtml, resources) = await BuildBrandEmailShellAsync(title, bodyBuilder.ToString());

                // Use overload that supports inline resources
                await _emailService.SendEmailAsync(receivers.To, receivers.Cc, title, emailHtml, highImportance: true, inlineResources: resources);

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "SendDailyFitupEmail failed for Project={ProjectId}", projectId);
                return Json(new { success = false, message = "Error" });
            }
        }

        private async Task<(List<string> To, List<string> Cc)> ResolveUCDailyReceiversAsync(int projectId, string razorKey, string locCode)
        {
            try
            {
                var rows = await _context.Receivers_tbl.AsNoTracking()
                    .Where(r => r.Project_Receivers == projectId && r.Razor == razorKey)
                    .Where(r => r.Location_Receivers == null || r.Location_Receivers == "" ||
                                r.Location_Receivers == "All" || r.Location_Receivers == locCode)
                    .ToListAsync();
                var to = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var cc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var row in rows)
                {
                    foreach (var e in SplitEmailsList(row.Receivers)) to.Add(e);
                    foreach (var e in SplitEmailsList(row.Receivers_cc)) cc.Add(e);
                }
                foreach (var e in to) cc.Remove(e);
                return (to.ToList(), cc.ToList());
            }
            catch
            {
                return (new List<string>(), new List<string>());
            }
        }

        private static IEnumerable<string> SplitEmailsList(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) yield break;
            foreach (var part in raw.Split(new[] { ';', ',', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var v = part.Trim();
                if (!string.IsNullOrWhiteSpace(v)) yield return v;
            }
        }

        // Modified: returns html + inline resources dictionary (CID image when available)
        private async Task<(string Html, Dictionary<string, (byte[] Content, string ContentType)>? Resources)> BuildBrandEmailShellAsync(string title, string contentHtml)
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

            var encodedTitle = System.Net.WebUtility.HtmlEncode(title ?? string.Empty);
            string h2Inner = encodedTitle;
            if (encodedTitle.StartsWith("Updated ", StringComparison.Ordinal))
            {
                var rest = encodedTitle.Substring("Updated ".Length);
                h2Inner = "<span style='color:#ED1E4D'>Updated</span> " + rest;
            }
            else if (encodedTitle.EndsWith("Updated", StringComparison.Ordinal))
            {
                // Color trailing Updated in the new phrase
                h2Inner = encodedTitle.Replace("Updated", "<span style='color:#ED1E4D'>Updated</span>");
            }

            var html = $@"<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'/><title>{encodedTitle} - PMS</title></head><body style='margin:0;padding:24px;background:#ffffff;color:#176d8a;font-family:Calibri,Arial,Helvetica,sans-serif;'>
<div style='max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #ffffff;border-radius:8px;box-shadow:0 2px 8px rgba(31,153,190,0.15);overflow:hidden;'>
  <div style='padding:14px 18px;text-align:center;background:#ffffff;'>{logoTag}</div>
  <div style='padding:20px 22px;'>
    <h2 style='margin:0 0 14px;color:#176d8a;font-weight:700;font-size:20px;'>{h2Inner}</h2>
    {contentHtml}
  </div>
  <div style='background:#ffffff;border-top:1px solid #e3e9eb;padding:12px 16px;text-align:center;'>
    <p style='margin:0;color:#7a9aa8;font-size:12px;'>Automated message from PMS System. Do not reply.</p>
  </div>
</div></body></html>";

            return (html, resources);
        }
    }
}
