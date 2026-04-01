using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Models;
using PMS.Infrastructure;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMS.Controllers;

public partial class HomeController
{
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkUpdateDwrRows([FromBody] List<DwrRowUpdateDto> items)
    {
        if (items == null || items.Count == 0) return BadRequest("No rows");
        var userId = HttpContext.Session.GetInt32("UserID");
        var updatedIds = new List<int>();
        var pendingIds = new List<int>();
        var errors = new List<object>();

        static string? Clip(string? value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        var validItems = items.Where(i => i != null && i.JointId > 0).ToList();
        if (validItems.Count == 0)
        {
            return BadRequest("No valid rows");
        }

        var jointIds = validItems.Select(i => i.JointId).Distinct().ToList();

        var dfrMap = await _context.DFR_tbl.AsNoTracking()
            .Where(d => jointIds.Contains(d.Joint_ID))
            .ToDictionaryAsync(d => d.Joint_ID);

        var otherList = await _context.Set<OtherNde>().AsNoTracking()
            .Where(o => jointIds.Contains(o.Joint_ID_NDE))
            .ToListAsync();
        var otherMap = otherList
            .GroupBy(o => o.Joint_ID_NDE)
            .ToDictionary(g => g.Key, g => g.FirstOrDefault());

        var rtList = await _context.RT_tbl.AsNoTracking()
            .Where(r => jointIds.Contains(r.Joint_ID_RT))
            .Select(r => new { r.Joint_ID_RT, r.DATE_NDE_WAS_REQUESTED, r.BSR_DATE_NDE_WAS_REQ })
            .ToListAsync();
        var rtMap = rtList
            .GroupBy(r => r.Joint_ID_RT)
            .ToDictionary(g => g.Key, g => g.FirstOrDefault());

        var pwhtList = await _context.PWHT_HT_tbl.AsNoTracking()
            .Where(p => jointIds.Contains(p.Joint_ID_PWHT_HT))
            .Select(p => new { p.Joint_ID_PWHT_HT, p.PWHT_DATE, p.HARDNESS_DATE })
            .ToListAsync();
        var pwhtMap = pwhtList
            .GroupBy(p => p.Joint_ID_PWHT_HT)
            .ToDictionary(g => g.Key, g => g.FirstOrDefault());

        var wpsNames = validItems
            .Where(c => (!c.WpsId.HasValue || c.WpsId.Value <= 0) && !string.IsNullOrWhiteSpace(c.Wps))
            .Select(c => c.Wps!.Trim())
            .Distinct()
            .ToList();
        var wpsLookup = wpsNames.Count == 0
            ? new Dictionary<string, int?>()
            : await _context.WPS_tbl.AsNoTracking()
                .Where(w => wpsNames.Contains(w.WPS))
                .ToDictionaryAsync(w => w.WPS, w => (int?)w.WPS_ID);

        var dwrMap = await _context.DWR_tbl
            .Where(w => jointIds.Contains(w.Joint_ID_DWR))
            .ToDictionaryAsync(w => w.Joint_ID_DWR);

        var autoDetectChanges = _context.ChangeTracker.AutoDetectChangesEnabled;
        _context.ChangeTracker.AutoDetectChangesEnabled = false;

        try
        {
            foreach (var dto in items)
            {
                if (dto == null || dto.JointId <= 0)
                {
                    errors.Add(new { id = dto?.JointId ?? 0, message = "Invalid row" });
                    continue;
                }

                if (!dfrMap.TryGetValue(dto.JointId, out var dfr))
                {
                    errors.Add(new { id = dto.JointId, message = "DFR not found" });
                    continue;
                }

                if (!dto.FitupDate.HasValue && !dto.ActualDate.HasValue)
                {
                    otherMap.TryGetValue(dto.JointId, out var otherCheck);
                    rtMap.TryGetValue(dto.JointId, out var rtCheck);
                    pwhtMap.TryGetValue(dto.JointId, out var pwhtCheck);

                    // NOTE: Do NOT consider the DFR FITUP_DATE here when deciding whether a Welding Date is required.
                    // The requirement is only triggered by related NDE/RT/PWHT/Hardness dates.
                    bool anyRelatedPresent =
                        (otherCheck != null && (otherCheck.OTHER_NDE_DATE.HasValue || otherCheck.PMI_DATE.HasValue))
                        || (rtCheck != null && (rtCheck.DATE_NDE_WAS_REQUESTED.HasValue || rtCheck.BSR_DATE_NDE_WAS_REQ.HasValue))
                        || (pwhtCheck != null && (pwhtCheck.PWHT_DATE.HasValue || pwhtCheck.HARDNESS_DATE.HasValue));

                    if (anyRelatedPresent)
                    {
                        errors.Add(new { id = dto.JointId, message = "Skipped: Welding Date is required when related NDE/RT/PWHT/Hardness dates exist." });
                        continue;
                    }
                }

                // Simplify nullable selection: prefer ActualDate, otherwise FitupDate
                System.DateTime? actualDateCandidate = dto.ActualDate ?? dto.FitupDate;
                if (actualDateCandidate.HasValue)
                {
                    var violations = new List<string>();
                    var actualDateOnly = actualDateCandidate.Value.Date;

                    if (dfr.FITUP_DATE.HasValue && actualDateOnly < dfr.FITUP_DATE.Value.Date)
                    {
                        violations.Add("Actual Date must be on/after Fit-up Date");
                    }

                    // Ensure Actual Date is not after the provided Welding (Fitup) Date
                    if (dto.ActualDate.HasValue && dto.FitupDate.HasValue)
                    {
                        var fitupOnly = dto.FitupDate.Value.Date;
                        if (actualDateOnly > fitupOnly)
                        {
                            violations.Add("Actual Date must be on or before Welding Date");
                        }
                    }

                    if (rtMap.TryGetValue(dto.JointId, out var rtDates) && rtDates != null)
                    {
                        if (rtDates.DATE_NDE_WAS_REQUESTED.HasValue && actualDateOnly > rtDates.DATE_NDE_WAS_REQUESTED.Value.Date)
                        {
                            violations.Add("Actual Date must be on/before NDE Requested Date (RT)");
                        }
                        if (rtDates.BSR_DATE_NDE_WAS_REQ.HasValue && actualDateOnly > rtDates.BSR_DATE_NDE_WAS_REQ.Value.Date)
                        {
                            violations.Add("Actual Date must be on/before BSR NDE Requested Date (RT)");
                        }
                    }

                    if (otherMap.TryGetValue(dto.JointId, out var other) && other != null)
                    {
                        if (other.OTHER_NDE_DATE.HasValue && actualDateOnly > other.OTHER_NDE_DATE.Value.Date)
                        {
                            violations.Add("Actual Date must be on/before Other NDE Date");
                        }
                        if (other.PMI_DATE.HasValue && actualDateOnly > other.PMI_DATE.Value.Date)
                        {
                            violations.Add("Actual Date must be on/before PMI Date");
                        }
                    }

                    if (pwhtMap.TryGetValue(dto.JointId, out var pwht) && pwht != null)
                    {
                        if (pwht.PWHT_DATE.HasValue && actualDateOnly > pwht.PWHT_DATE.Value.Date)
                        {
                            violations.Add("Actual Date must be on/before PWHT Date");
                        }
                        if (pwht.HARDNESS_DATE.HasValue && actualDateOnly > pwht.HARDNESS_DATE.Value.Date)
                        {
                            violations.Add("Actual Date must be on/before Hardness Date");
                        }
                    }

                    if (violations.Count > 0)
                    {
                        errors.Add(new { id = dto.JointId, message = string.Join("; ", violations) });
                        continue;
                    }
                }

                if (!dwrMap.TryGetValue(dto.JointId, out var dwr))
                {
                    dwr = new Dwr { Joint_ID_DWR = dto.JointId };
                    _context.DWR_tbl.Add(dwr);
                    dwrMap[dto.JointId] = dwr;
                }

                if (dto.WpsId.HasValue && dto.WpsId.Value > 0)
                {
                    dwr.WPS_ID_DWR = dto.WpsId.Value;
                }
                else if (!string.IsNullOrWhiteSpace(dto.Wps) && wpsLookup.TryGetValue(dto.Wps.Trim(), out var wpsIdValue))
                {
                    dwr.WPS_ID_DWR = wpsIdValue;
                }
                else
                {
                    // If the DTO explicitly has no WPS (either WpsId == 0 or empty text), clear the stored WPS
                    dwr.WPS_ID_DWR = null;
                }

                dwr.POST_VISUAL_INSPECTION_QR_NO = Clip(dto.FitupReport, 8);
                dwr.RFI_ID_DWR = (dto.RfiId.HasValue && dto.RfiId.Value > 0) ? dto.RfiId.Value : (int?)null;

                // Use nullable assignment patterns to simplify null checks
                dwr.DATE_WELDED = dto.FitupDate;
                dwr.ACTUAL_DATE_WELDED = dto.ActualDate ?? (dwr.ACTUAL_DATE_WELDED == null ? dto.FitupDate : null);

                dwr.ROOT_A = Clip(dto.TackWelder, 9);
                dwr.ROOT_B = Clip(dto.TackWelderB, 9);
                dwr.FILL_A = Clip(dto.TackWelderFillA, 9);
                dwr.FILL_B = Clip(dto.TackWelderFillB, 9);
                dwr.CAP_A = Clip(dto.TackWelderCapA, 9);
                dwr.CAP_B = Clip(dto.TackWelderCapB, 9);
                dwr.PREHEAT_TEMP_C = dto.PreheatTempC;
                if (!string.IsNullOrWhiteSpace(dto.IP_or_T))
                {
                    dwr.IP_or_T = Clip(dto.IP_or_T, 4);
                }
                if (!string.IsNullOrWhiteSpace(dto.Open_Closed))
                {
                    dwr.Open_Closed = Clip(dto.Open_Closed, 2);
                }

                if (dto.FitupConfirmed.HasValue)
                {
                    dwr.Weld_Confirmed = dto.FitupConfirmed.Value;
                }

                if (!string.IsNullOrWhiteSpace(dto.DwrRemarks)) dwr.DWR_REMARKS = Clip(dto.DwrRemarks, 50);

                dwr.DWR_Updated_By = userId;
                dwr.DWR_Updated_Date = AppClock.Now;

                pendingIds.Add(dto.JointId);
            }

            if (pendingIds.Count > 0)
            {
                try
                {
                    _context.ChangeTracker.DetectChanges();
                    await _context.SaveChangesAsync();
                    updatedIds.AddRange(pendingIds);
                }
                catch (System.Exception ex)
                {
                    _logger.LogError(ex, "BulkUpdateDwrRows failed for {JointIds}", pendingIds);
                    foreach (var id in pendingIds)
                    {
                        errors.Add(new { id, message = "Save failed" });
                    }
                }
            }
        }
        finally
        {
            _context.ChangeTracker.AutoDetectChangesEnabled = autoDetectChanges;
        }

        return Json(new { success = errors.Count == 0, updated = updatedIds.Count, skipped = errors.Count, errors, updatedIds });
    }
}
