using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PMS.Models;
using System;
using System.Linq;
using System.Threading.Tasks;
using PMS.Infrastructure;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Data.SqlClient;
using System.Data;
using Microsoft.AspNetCore.Http;
using System.IO;

namespace PMS.Controllers;

public partial class HomeController
{
    private enum QualificationSaveStatus
    {
        Saved,
        Duplicate
    }

    private readonly record struct QualificationSaveEvaluation(QualificationSaveStatus Status, string? ExistingJcc);

    #region Private Helpers

    // Helper to normalize location to Shop/Field buckets
    private static string NormalizeLocation(string? loc)
    {
        var s = (loc ?? string.Empty).Trim().ToLowerInvariant();
        if (s.StartsWith('w') && (s.StartsWith("ws") || s.StartsWith("work") || s.StartsWith("shop"))) return "Shop"; // CA1866: char overload for first check
        if (s.StartsWith('f') && (s.StartsWith("fw") || s.StartsWith("field"))) return "Field"; // CA1866
        if (s.StartsWith('s')) return "Shop"; // CA1866
        if (s.StartsWith('f')) return "Field"; // CA1866
        return string.Empty;
    }

    private async Task<string> SuggestWelderLocationAsync(int? batchNo)
    {
        if (!batchNo.HasValue || batchNo.Value <= 0)
            return "Shop"; // default if no batch

        var byBatch = await _context.Welder_List_tbl
            .AsNoTracking()
            .Where(q => q.Batch_No == batchNo.Value)
            .Where(q => q.Welder != null)
            .Select(q => q.Welder!.Welder_Location)
            .ToListAsync();
        int shop = 0, field = 0;
        foreach (var raw in byBatch)
        {
            var norm = NormalizeLocation(raw);
            if (norm == "Shop") shop++;
            else if (norm == "Field") field++;
        }
        if (shop == 0 && field == 0) return "Shop"; // fallback
        if (shop == field) return "Shop"; // tie -> Shop
        return shop > field ? "Shop" : "Field";
    }

    // Suggest most frequent WQT Agency scoped strictly by Batch No (no global fallback)
    private async Task<string?> SuggestWqtAgencyAsync(int? batchNo)
    {
        // Helper for mode with case-insensitive grouping preserving first seen casing
        static string? Mode(IEnumerable<string?> src)
            => src.Where(s => !string.IsNullOrWhiteSpace(s))
                  .Select(s => s!.Trim())
                  .GroupBy(s => s, StringComparer.OrdinalIgnoreCase)
                  .OrderByDescending(g => g.Count())
                  .Select(g => g.First())
                  .FirstOrDefault();

        if (!batchNo.HasValue || batchNo.Value <= 0)
            return null;

        var slice = await _context.Welder_List_tbl
            .AsNoTracking()
            .Where(q => q.Batch_No == batchNo.Value)
            .Select(q => q.WQT_Agency)
            .ToListAsync();
        return Mode(slice);
    }

    #endregion

    #region Suggestion APIs

    // NEW: Suggest the largest Test_Date for a given Batch No (as yyyy-MM-dd)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> SuggestQualTestDate([FromQuery] int? batchNo)
    {
        try
        {
            DateTime? max;
            if (batchNo.HasValue && batchNo.Value > 0)
            {
                max = await _context.Welder_List_tbl
                    .AsNoTracking()
                    .Where(q => q.Batch_No == batchNo.Value && q.Test_Date.HasValue)
                    .MaxAsync(q => q.Test_Date);
            }
            else
            {
                max = await _context.Welder_List_tbl
                    .AsNoTracking()
                    .Where(q => q.Test_Date.HasValue)
                    .MaxAsync(q => q.Test_Date);
            }
            var iso = max?.ToString("yyyy-MM-dd");
            return Json(new { date = iso });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SuggestQualTestDate failed for batch {Batch}", batchNo);
            return Json(new { date = (string?)null });
        }
    }

    // Suggest location API for client-side batch changes
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> SuggestWelderLocation([FromQuery] int? batchNo)
    {
        var s = await SuggestWelderLocationAsync(batchNo);
        return Json(new { suggested = s });
    }

    // NEW: Suggest WQT Agency API for client-side batch changes
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> SuggestWqtAgency([FromQuery] int? batchNo)
    {
        var agency = await SuggestWqtAgencyAsync(batchNo);
        return Json(new { agency = agency ?? string.Empty });
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> NextJccNo([FromQuery] int? projectId, [FromQuery] string? welderLocation, [FromQuery] int? batchNo = null)
    {
        try
        {
            var next = await GenerateNextJccNoAsync(projectId, welderLocation, batchNo);
            return Json(new { jcc = next });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to calculate NextJccNo for project {Project} / location {Location} / batch {Batch}", projectId, welderLocation, batchNo);
            return BadRequest(new { error = "Failed to generate next JCC No" });
        }
    }

    #endregion

    #region Welder Editing

    // GET: /Home/EditWelder (Add or Edit)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> EditWelder(int? id, string? jcc = null)
    {
        // Fixed location list per requirement
        ViewBag.WelderLocations = new List<string> { "Shop", "Field" };

        // Projects list: Welders_Project_ID - Project_Name (grouped/distinct, matches Control Lookups)
        var projectGroups = await _context.Projects_tbl
            .AsNoTracking()
            .Where(p => p.Welders_Project_ID != null)
            .GroupBy(p => p.Welders_Project_ID)
            .Select(g => new
            {
                Welders_Project_ID = g.Key,
                Project_Name = g.OrderBy(p => p.Project_ID).First().Project_Name,
                Default_P = g.Any(p => p.Default_P)
            })
            .OrderBy(p => p.Welders_Project_ID)
            .ToListAsync();

        var preferredProjectId = await GetDefaultProjectIdAsync();
        int? defaultProjectId = null;
        if (preferredProjectId.HasValue)
        {
            // Map a Project_ID to its Welders_Project_ID
            var mapped = await _context.Projects_tbl.AsNoTracking()
                .Where(p => p.Project_ID == preferredProjectId.Value)
                .Select(p => p.Welders_Project_ID)
                .FirstOrDefaultAsync();
            defaultProjectId = mapped ?? preferredProjectId;
        }
        defaultProjectId ??= projectGroups
            .Where(p => p.Default_P)
            .OrderByDescending(p => p.Welders_Project_ID)
            .Select(p => p.Welders_Project_ID)
            .FirstOrDefault()
            ?? projectGroups.OrderByDescending(p => p.Welders_Project_ID)
                .Select(p => p.Welders_Project_ID)
                .FirstOrDefault();

        ViewBag.Projects = projectGroups
            .Select(p => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
            {
                Value = p.Welders_Project_ID.ToString(),
                Text = string.IsNullOrWhiteSpace(p.Project_Name)
                    ? p.Welders_Project_ID.ToString()
                    : ($"{p.Welders_Project_ID} - {p.Project_Name}"),
                Selected = (defaultProjectId.HasValue && p.Welders_Project_ID == defaultProjectId.Value)
            })
            .ToList();

        // Qualification lookups (global distinct for non-cascading fields)
        var qrows = await _context.Welder_List_tbl
            .AsNoTracking()
            .Select(x => new
            {
                x.WQT_Agency,
                x.Received_from_Aramco,
                x.Batch_No
            })
            .ToListAsync();

        static List<string> DistinctList(IEnumerable<string?> src) => src
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();

        ViewBag.WQTagencies = DistinctList(qrows.Select(r => r.WQT_Agency));
        ViewBag.RecivedFromAramco = DistinctList(qrows.Select(r => r.Received_from_Aramco));
        ViewBag.MaxBatchNo = qrows.Where(r => r.Batch_No.HasValue).Select(r => r.Batch_No!.Value).DefaultIfEmpty(0).Max();

        // Populate Status list from Welders_tbl.Status
        var statusValues = await _context.Welders_tbl
            .AsNoTracking()
            .Select(w => w.Status)
            .ToListAsync();
        var statusList = DistinctList(statusValues);
        // Ensure "Waiting for Approval" appears in dropdown even if not yet in database
        if (!statusList.Any(s => string.Equals(s, "Waiting for Approval", StringComparison.OrdinalIgnoreCase)))
            statusList.Insert(0, "Waiting for Approval");
        ViewBag.Statuses = statusList;

        // Default WQT agency scoped by max Batch No
        var maxBatchForWqt = (int)(ViewBag.MaxBatchNo ?? 0);
        var defaultWqt = maxBatchForWqt > 0
            ? qrows
                .Where(r => !string.IsNullOrWhiteSpace(r.WQT_Agency) && r.Batch_No == maxBatchForWqt)
                .GroupBy(r => r.WQT_Agency!.Trim(), StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Count())
                .Select(g => g.First().WQT_Agency!.Trim())
                .FirstOrDefault()
            : null;
        ViewBag.DefaultWqtAgency = defaultWqt ?? string.Empty;

        // Build model (new or existing)
        var model = new WelderEditViewModel();
        if (id.GetValueOrDefault() > 0)
        {
            var welder = await _context.Welders_tbl
                .Include(w => w.Qualifications)
                .FirstOrDefaultAsync(w => w.Welder_ID == id!.Value);
            if (welder == null)
            {
                TempData["Msg"] = "Welder not found.";
                return RedirectToAction(nameof(WelderList));
            }
            model.Welder = welder;
            // Sort qualifications by updated date desc then date issued/test date desc
            var quals = welder.Qualifications
                .OrderByDescending(q => q.Welder_List_Updated_Date)
                .ThenByDescending(q => q.Date_Issued)
                .ThenByDescending(q => q.Test_Date)
                .ToList();
            model.Qualifications = quals;

            // Resolve selected qualification
            WelderQualification? selected = null;
            if (!string.IsNullOrWhiteSpace(jcc))
            {
                selected = quals.FirstOrDefault(q => string.Equals(q.JCC_No, jcc, StringComparison.OrdinalIgnoreCase));
                model.SelectedJcc = jcc;
            }
            selected ??= quals.FirstOrDefault();
            model.Qualification = selected ?? new WelderQualification();

            // Ensure default row shows as "Editing" when jcc not explicitly provided
            if (string.IsNullOrWhiteSpace(model.SelectedJcc) && selected != null)
            {
                model.SelectedJcc = selected.JCC_No;
            }

            // Populate "Qualification Updated By" full name from PMS_Login_tbl
            if (model.Qualification?.Welder_List_Updated_By.HasValue == true)
            {
                var uid = model.Qualification.Welder_List_Updated_By!.Value;
                model.QualificationUpdatedByName = await _context.PMS_Login_tbl
                    .AsNoTracking()
                    .Where(u => u.UserID == uid)
                    .Select(u => ((u.FirstName ?? "").Trim() + " " + (u.LastName ?? "").Trim()).Trim())
                    .FirstOrDefaultAsync();
            }
            if (model.Qualification?.JCC_Upload_By.HasValue == true)
            {
                var uid = model.Qualification.JCC_Upload_By!.Value;
                model.QualificationUploadedByName = await _context.PMS_Login_tbl
                    .AsNoTracking()
                    .Where(u => u.UserID == uid)
                    .Select(u => ((u.FirstName ?? "").Trim() + " " + (u.LastName ?? "").Trim()).Trim())
                    .FirstOrDefaultAsync();
            }
        }
        else
        {
            // New welder
            model.Welder = new Welder
            {
                Welder_ID = 0,
                Status = "Waiting for Approval", // changed from Active to ensure default selection appears
                Project_Welder = defaultProjectId
            };
            model.Qualification = new WelderQualification();
            model.Qualifications = new List<WelderQualification>();

            // Suggest default location by most common in current max batch
            var maxBatch = (int)(ViewBag.MaxBatchNo ?? 0);
            var suggested = await SuggestWelderLocationAsync(maxBatch);
            ViewBag.SuggestedLocation = suggested;
            if (string.IsNullOrWhiteSpace(model.Welder.Welder_Location))
                model.Welder.Welder_Location = suggested;
        }

        // Batch-aware default for WQT Agency: use current qualification batch if set; otherwise max batch
        int? defaultBatchNo = model.Qualification?.Batch_No ?? (int?)(ViewBag.MaxBatchNo ?? 0);
        var batchDefaultWqt = await SuggestWqtAgencyAsync(defaultBatchNo);
        if (!string.IsNullOrWhiteSpace(batchDefaultWqt))
            ViewBag.DefaultWqtAgency = batchDefaultWqt;

        // Cascading qualification dropdowns seeded with the same logic used by client-side updates
        var qual = model.Qualification ?? new WelderQualification();
        var qualOptions = await BuildQualOptionsAsync(
            qual.Batch_No ?? defaultBatchNo,
            qual.Welding_Process,
            qual.Material_P_No,
            qual.Code_Reference,
            qual.Consumable_Root_F_No,
            qual.Consumable_Root_Spec,
            qual.Consumable_Filling_Cap_F_No,
            qual.Consumable_Filling_Cap_Spec,
            qual.Position_Progression,
            qual.Diameter_Range,
            qual.Max_Thickness,
            allowEmpty: false);

        // If some selections are empty, seed them with effective defaults and rebuild once to mirror client-side cascade on load
        string? Seed(string? current, string key)
        {
            if (!string.IsNullOrWhiteSpace(current)) return current;
            return qualOptions.Effective.TryGetValue(key, out var v) ? v : null;
        }

        var seededProc = Seed(qual.Welding_Process, "weldingProcess");
        var seededMaterial = Seed(qual.Material_P_No, "materialPNo");
        var seededCodeRef = Seed(qual.Code_Reference, "codeReference");
        var seededRootF = Seed(qual.Consumable_Root_F_No, "rootFNo");
        var seededRootSpec = Seed(qual.Consumable_Root_Spec, "rootSpec");
        var seededFillF = Seed(qual.Consumable_Filling_Cap_F_No, "fillCapFNo");
        var seededFillSpec = Seed(qual.Consumable_Filling_Cap_Spec, "fillCapSpec");
        var seededPos = Seed(qual.Position_Progression, "positionProgression");
        var seededDia = Seed(qual.Diameter_Range, "diameterRange");
        var seededThk = Seed(qual.Max_Thickness, "maxThickness");

        if (seededProc != qual.Welding_Process ||
            seededMaterial != qual.Material_P_No ||
            seededCodeRef != qual.Code_Reference ||
            seededRootF != qual.Consumable_Root_F_No ||
            seededRootSpec != qual.Consumable_Root_Spec ||
            seededFillF != qual.Consumable_Filling_Cap_F_No ||
            seededFillSpec != qual.Consumable_Filling_Cap_Spec ||
            seededPos != qual.Position_Progression ||
            seededDia != qual.Diameter_Range ||
            seededThk != qual.Max_Thickness)
        {
            qualOptions = await BuildQualOptionsAsync(
                qual.Batch_No ?? defaultBatchNo,
                seededProc,
                seededMaterial,
                seededCodeRef,
                seededRootF,
                seededRootSpec,
                seededFillF,
                seededFillSpec,
                seededPos,
                seededDia,
                seededThk,
                allowEmpty: false);
        }

        ViewBag.WeldingProcesses = qualOptions.WeldingProcesses;
        ViewBag.MaterialPNos = qualOptions.MaterialPNos;
        ViewBag.CodeReferences = qualOptions.CodeReferences;
        ViewBag.RootFNos = qualOptions.RootFNos;
        ViewBag.RootSpecs = qualOptions.RootSpecs;
        ViewBag.FillCapFNos = qualOptions.FillCapFNos;
        ViewBag.FillCapSpecs = qualOptions.FillCapSpecs;
        ViewBag.PositionProgressions = qualOptions.PositionProgressions;
        ViewBag.DiameterRanges = qualOptions.DiameterRanges;
        ViewBag.MaxThicknesses = qualOptions.MaxThicknesses;

        return View(model);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SaveWelder(
        [FromForm] Welder Welder,
        [FromForm] WelderQualification Qualification,
        [FromForm] bool IsAddFlow = false,
        [FromForm] bool AddNewQualification = false,
        [FromForm] string? SelectedJcc = null // original JCC when editing existing qualification
    )
    {
        static string Clean(string? v, int max)
        {
            if (string.IsNullOrWhiteSpace(v) || v == "__new__") return string.Empty;
            var t = v.Trim();
            return t.Length <= max ? t : t[..max];
        }
        static string StripAllWhitespace(string? s)
            => PMS.Controllers.HomeController.StripWhitespace(s);

        // Resolve the effective Project_Welder for duplicate checks (prefers posted, falls back to DB)
        async Task<int?> ResolveProjectIdAsync(int? welderId, int? projectId)
        {
            if (projectId.HasValue && projectId.Value > 0) return projectId.Value;
            if (welderId.HasValue && welderId.Value > 0)
            {
                return await _context.Welders_tbl
                    .AsNoTracking()
                    .Where(w => w.Welder_ID == welderId.Value)
                    .Select(w => w.Project_Welder)
                    .FirstOrDefaultAsync();
            }
            return null;
        }

        // Duplicate detection that is aware of welder symbol and project (Welders_tbl.Project_Welder)
        async Task<string?> FindDuplicateQualificationAsync(int? welderId, string? symbol, int? projectId, WelderQualification qual, string? excludeJcc = null)
        {
            const string ciCollation = "SQL_Latin1_General_CP1_CI_AS";

            var normalizedSymbol = StripAllWhitespace(symbol);
            var effectiveProjectId = await ResolveProjectIdAsync(welderId, projectId);

            // Normalize core fields the same way we save them (trim, empty instead of null)
            var np = Clean(qual.Welding_Process, 30);
            var nm = Clean(qual.Material_P_No, 100);
            var nd = Clean(qual.Diameter_Range, 60);
            var nt = Clean(qual.Max_Thickness, 60);
            var nr = Clean(qual.Consumable_Root_F_No, 20);
            var nrc = Clean(qual.Consumable_Filling_Cap_F_No, 100);
            var npos = Clean(qual.Position_Progression, 100);
            var nBatch = qual.Batch_No; // keep nullable so missing batches don't collide with non-null
            var nCodeRef = Clean(qual.Code_Reference, 75);
            var nCertRef = Clean(qual.Qualification_Cert_Ref_No, 17);
            var nWqt = Clean(qual.WQT_Agency, 30);
            var nReceived = Clean(qual.Received_from_Aramco, 40);

            // Only treat as potential duplicate when all core fields are present
            if (string.IsNullOrWhiteSpace(np) || string.IsNullOrWhiteSpace(nm) || string.IsNullOrWhiteSpace(nd) ||
                string.IsNullOrWhiteSpace(nt) || string.IsNullOrWhiteSpace(nr) || string.IsNullOrWhiteSpace(nrc) ||
                string.IsNullOrWhiteSpace(npos))
            {
                return null;
            }

            var query = from q in _context.Welder_List_tbl.AsNoTracking()
                        join w in _context.Welders_tbl.AsNoTracking() on q.Welder_ID_WL equals w.Welder_ID
                        select new { q, w };

            if (!string.IsNullOrWhiteSpace(normalizedSymbol))
            {
                query = query.Where(x => x.w.Welder_Symbol != null && EF.Functions.Collate(x.w.Welder_Symbol, ciCollation) == normalizedSymbol);
            }

            if (effectiveProjectId.HasValue && effectiveProjectId.Value > 0)
            {
                query = query.Where(x => x.w.Project_Welder == effectiveProjectId.Value);
            }

            if (welderId.HasValue && welderId.Value > 0)
            {
                query = query.Where(x => x.w.Welder_ID == welderId.Value);
            }

            // Core-field match required to flag duplicate
            query = query
                .Where(x => EF.Functions.Collate(x.q.Welding_Process ?? string.Empty, ciCollation) == np)
                .Where(x => EF.Functions.Collate(x.q.Material_P_No ?? string.Empty, ciCollation) == nm)
                .Where(x => EF.Functions.Collate(x.q.Diameter_Range ?? string.Empty, ciCollation) == nd)
                .Where(x => EF.Functions.Collate(x.q.Max_Thickness ?? string.Empty, ciCollation) == nt)
                .Where(x => EF.Functions.Collate(x.q.Consumable_Root_F_No ?? string.Empty, ciCollation) == nr)
                .Where(x => EF.Functions.Collate(x.q.Consumable_Filling_Cap_F_No ?? string.Empty, ciCollation) == nrc)
                .Where(x => EF.Functions.Collate(x.q.Position_Progression ?? string.Empty, ciCollation) == npos);

            if (nBatch.HasValue)
            {
                query = query.Where(x => x.q.Batch_No.HasValue && x.q.Batch_No.Value == nBatch.Value);
            }
            else
            {
                query = query.Where(x => !x.q.Batch_No.HasValue);
            }

            // Always match optional string fields bidirectionally so that
            // an empty new value only matches records that are also empty.
            query = query.Where(x => EF.Functions.Collate(x.q.Code_Reference ?? string.Empty, ciCollation) == nCodeRef);
            query = query.Where(x => EF.Functions.Collate(x.q.Qualification_Cert_Ref_No ?? string.Empty, ciCollation) == nCertRef);
            query = query.Where(x => EF.Functions.Collate(x.q.WQT_Agency ?? string.Empty, ciCollation) == nWqt);
            query = query.Where(x => EF.Functions.Collate(x.q.Received_from_Aramco ?? string.Empty, ciCollation) == nReceived);

            if (!string.IsNullOrWhiteSpace(excludeJcc))
            {
                var trimmed = excludeJcc.Trim();
                query = query.Where(x => x.q.JCC_No == null || EF.Functions.Collate(x.q.JCC_No, ciCollation) != trimmed);
            }

            return await query.Select(x => x.q.JCC_No).FirstOrDefaultAsync();
        }

        async Task<Welder?> FindSymbolProjectConflictAsync(string? symbol, int? projectId, int excludeWelderId = 0)
        {
            if (string.IsNullOrWhiteSpace(symbol) || !projectId.HasValue || projectId.Value <= 0) return null;
            var normalized = StripAllWhitespace(symbol);
            return await _context.Welders_tbl
                .AsNoTracking()
                .Where(w => w.Welder_ID != excludeWelderId)
                .Where(w => w.Project_Welder == projectId.Value)
                .Where(w => w.Welder_Symbol != null && EF.Functions.Collate(w.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == normalized)
                .FirstOrDefaultAsync();
        }

        // Separate flags: IsAddFlow indicates brand-new welder form, AddNewQualification explicit user intent
        bool isNewWelderFlow = IsAddFlow;
        bool wantsNewQualification = AddNewQualification; // no longer implied by IsAddFlow
        const string DuplicateQualificationMessage = "Duplicate qualification detected. Showing existing record.";
        const string QualificationSavedMessage = "Qualification saved successfully.";

        // Track whether this welder already has any qualifications (used to allow welder-only saves)
        var hasExistingQualification = !isNewWelderFlow
            && Welder.Welder_ID > 0
            && await _context.Welder_List_tbl.AsNoTracking().AnyAsync(q => q.Welder_ID_WL == Welder.Welder_ID);

        async Task<QualificationSaveEvaluation> EvaluateQualificationSaveAsync(Welder? targetWelder, WelderQualification qual, string? excludeJcc = null)
        {
            if (targetWelder == null)
                return new QualificationSaveEvaluation(QualificationSaveStatus.Saved, null);

            var duplicate = await FindDuplicateQualificationAsync(targetWelder.Welder_ID, targetWelder.Welder_Symbol, targetWelder.Project_Welder, qual, excludeJcc);
            return string.IsNullOrEmpty(duplicate)
                ? new QualificationSaveEvaluation(QualificationSaveStatus.Saved, null)
                : new QualificationSaveEvaluation(QualificationSaveStatus.Duplicate, duplicate);
        }

        string ApplyQualificationSaveMessage(QualificationSaveStatus status)
        {
            var message = status switch
            {
                QualificationSaveStatus.Saved => QualificationSavedMessage,
                QualificationSaveStatus.Duplicate => DuplicateQualificationMessage,
                _ => TempData["Msg"] as string ?? string.Empty
            };
            TempData["Msg"] = message;
            return message;
        }

        // Ensure JCC_No is unique (regenerate when collision is detected)
        async Task EnsureUniqueJccAsync(WelderQualification q, int? projectId, string? welderLocation)
        {
            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                if (string.IsNullOrWhiteSpace(q.JCC_No))
                {
                    q.JCC_No = await GenerateNextJccNoAsync(projectId, welderLocation, q.Batch_No);
                }

                var exists = await _context.Welder_List_tbl.AsNoTracking().AnyAsync(x => x.JCC_No == q.JCC_No);
                if (!exists) return;

                // regenerate on collision
                q.JCC_No = await GenerateNextJccNoAsync(projectId, welderLocation, q.Batch_No);
            }
        }

        // Detect AJAX/json requests so we can return JSON to allow partial refresh
        // Use typed headers for Accept to satisfy analyzers and preserve correctness
        var typedHeaders = Request.GetTypedHeaders();
        var accept = typedHeaders?.Accept != null ? string.Join(",", typedHeaders.Accept.Select(m => m.ToString())) : string.Empty;
        // X-Requested-With is commonly sent by libraries; read it from the headers dictionary
        var xRequestedWith = Request.Headers.XRequestedWith.ToString();
        bool isAjax = string.Equals(xRequestedWith, "XMLHttpRequest", StringComparison.OrdinalIgnoreCase)
                      || accept.Contains("application/json", StringComparison.OrdinalIgnoreCase);

        // Local helper to build JSON response for updated qualification area (used when isAjax)
        async Task<IActionResult> AjaxEditResponseAsync(int? welderId, string? jcc, bool isDuplicate = false, string? overrideMessage = null)
        {
            if (!isAjax)
            {
                if (isNewWelderFlow)
                    return RedirectToAction(nameof(EditWelder)); // stay on Add Welder page after save
                return RedirectToAction(nameof(EditWelder), new { id = welderId ?? 0, jcc });
            }
            var msg = overrideMessage ?? TempData["Msg"] as string ?? string.Empty;
            if (overrideMessage != null)
                TempData.Remove("Msg");

            // Fallback for new qualification rows where no message was set upstream
            if (string.IsNullOrWhiteSpace(msg) && !isDuplicate && !string.IsNullOrWhiteSpace(jcc))
            {
                msg = QualificationSavedMessage;
            }

            if (!welderId.HasValue || welderId.Value <= 0)
            {
                var maxBatchNo = await _context.Welder_List_tbl
                    .AsNoTracking()
                    .Where(x => x.Batch_No.HasValue)
                    .Select(x => x.Batch_No!.Value)
                    .DefaultIfEmpty(0)
                    .MaxAsync();

                return Json(new { success = true, msg, duplicate = isDuplicate, welderId = 0, qualifications = Array.Empty<object>(), qualification = (object?)null, maxBatchNo });
            }
            var items = await _context.Welder_List_tbl
                 .AsNoTracking()
                 .Where(q => q.Welder_ID_WL == welderId.Value)
                 .OrderBy(q => q.JCC_No)
                 .Select(q => new
                 {
                     q.JCC_No,
                     q.Welding_Process,
                     q.Material_P_No,
                     q.Diameter_Range,
                     q.Max_Thickness
                 })
                 .ToListAsync();

            WelderQualification? qual = null;
            if (!string.IsNullOrWhiteSpace(jcc))
            {
                qual = await _context.Welder_List_tbl.AsNoTracking().FirstOrDefaultAsync(q => q.JCC_No == jcc && q.Welder_ID_WL == welderId.Value);
            }
            qual ??= await _context.Welder_List_tbl.AsNoTracking()
                    .Where(q => q.Welder_ID_WL == welderId.Value)
                    .OrderByDescending(q => q.Welder_List_Updated_Date)
                    .ThenByDescending(q => q.Date_Issued)
                    .ThenByDescending(q => q.Test_Date)
                    .FirstOrDefaultAsync();

            object? qualDto = null;
            if (qual != null)
            {
                qualDto = new
                {
                    qual.JCC_No,
                    Test_Date = qual.Test_Date?.ToString("yyyy-MM-dd"),
                    qual.Welding_Process,
                    qual.Material_P_No,
                    qual.Diameter_Range,
                    qual.Max_Thickness,
                    qual.Qualification_Cert_Ref_No,
                    qual.WQT_Agency,
                    qual.Batch_No,
                    Date_Issued = qual.Date_Issued?.ToString("yyyy-MM-dd"),
                    qual.Remarks,
                    DATE_OF_LAST_CONTINUITY = qual.DATE_OF_LAST_CONTINUITY?.ToString("yyyy-MM-dd"),
                    qual.RECORDING_THE_CONTINUITY_RECORD,
                    hasFile = !string.IsNullOrWhiteSpace(qual.JCC_BlobName),
                    fileName = qual.JCC_FileName,
                    fileSize = qual.JCC_FileSize,
                    uploadDate = qual.JCC_UploadDate?.ToString("yyyy-MM-dd HH:mm")
                };
            }
            else
            {
                // No existing qualification found - provide sensible defaults for the client
                // Determine batch candidate (prefer posted batch, else use max known batch)
                int? batchCandidate = Qualification?.Batch_No > 0 ? Qualification.Batch_No : null;
                if (!batchCandidate.HasValue)
                {
                    var maxBatchCandidate = await _context.Welder_List_tbl
                        .AsNoTracking()
                        .Where(x => x.Batch_No.HasValue)
                        .Select(x => x.Batch_No!.Value)
                        .DefaultIfEmpty(0)
                        .MaxAsync();
                    if (maxBatchCandidate > 0) batchCandidate = maxBatchCandidate;
                }

                // Suggest test date (reuse action) - returns JsonResult { date = "yyyy-MM-dd" }
                string? suggestedDate = null;
                try
                {
                    var resDate = await SuggestQualTestDate(batchCandidate);
                    if (resDate is JsonResult jrDate)
                    {
                        dynamic v = jrDate.Value!;
                        suggestedDate = (v?.date as string) ?? (v?.Date as string);
                    }
                }
                catch { /* ignore */ }

                // Suggest WQT agency (batch-aware helper)
                string? suggestedWqt = null;
                try { suggestedWqt = await SuggestWqtAgencyAsync(batchCandidate); } catch { }

                // Get effective defaults from GetQualOptionsAdvanced
                string? effMaterial = null;
                string? effDiameter = null;
                string? effMaxThickness = null;
                try
                {
                    var resOpts = await GetQualOptionsAdvanced(batchCandidate, Qualification?.Welding_Process, Qualification?.Material_P_No, Qualification?.Code_Reference, Qualification?.Consumable_Root_F_No, Qualification?.Consumable_Root_Spec, Qualification?.Consumable_Filling_Cap_F_No, Qualification?.Consumable_Filling_Cap_Spec, Qualification?.Position_Progression, Qualification?.Diameter_Range, Qualification?.Max_Thickness, true);
                    if (resOpts is JsonResult jrOpts)
                    {
                        dynamic v = jrOpts.Value!;
                        var eff = v?.effective;
                        effMaterial = eff?.materialPNo as string ?? eff?.materialPNo?.ToString();
                        effDiameter = eff?.diameterRange as string ?? eff?.diameterRange?.ToString();
                        effMaxThickness = eff?.maxThickness as string ?? eff?.maxThickness?.ToString();
                    }
                }
                catch { }

                qualDto = new
                {
                    JCC_No = (string?)null,
                    Test_Date = suggestedDate,
                    Welding_Process = Qualification?.Welding_Process,
                    Material_P_No = effMaterial ?? Qualification?.Material_P_No,
                    Diameter_Range = effDiameter ?? Qualification?.Diameter_Range,
                    Max_Thickness = effMaxThickness ?? Qualification?.Max_Thickness,
                    Qualification_Cert_Ref_No = Qualification?.Qualification_Cert_Ref_No,
                    WQT_Agency = suggestedWqt ?? Qualification?.WQT_Agency,
                    Batch_No = batchCandidate,
                    Date_Issued = Qualification?.Date_Issued?.ToString("yyyy-MM-dd"),
                    Remarks = Qualification?.Remarks,
                    DATE_OF_LAST_CONTINUITY = Qualification?.DATE_OF_LAST_CONTINUITY?.ToString("yyyy-MM-dd"),
                    RECORDING_THE_CONTINUITY_RECORD = Qualification?.RECORDING_THE_CONTINUITY_RECORD,
                    hasFile = false,
                    fileName = (string?)null,
                    fileSize = (int?)null,
                    uploadDate = (string?)null
                };
            }

            var maxBatch = await _context.Welder_List_tbl
                .AsNoTracking()
                .Where(x => x.Batch_No.HasValue)
                .Select(x => x.Batch_No!.Value)
                .DefaultIfEmpty(0)
                .MaxAsync();

            return Json(new { success = true, msg, duplicate = isDuplicate, welderId = welderId.Value, qualifications = items, qualification = qualDto, maxBatchNo = maxBatch });
        }

        Task<IActionResult> DuplicateQualificationResponseAsync(int? welderId, string? jcc)
        {
            var duplicateMsg = ApplyQualificationSaveMessage(QualificationSaveStatus.Duplicate);
            return AjaxEditResponseAsync(welderId, jcc, true, duplicateMsg);
        }

        try
        {
            // Normalize
            Welder.Welder_Symbol = StripAllWhitespace(Welder.Welder_Symbol);
            Welder.Iqama_No = StripAllWhitespace(Welder.Iqama_No);
            Welder.Passport = StripAllWhitespace(Welder.Passport);

            Welder.Welder_Location = Clean(Welder.Welder_Location, 5);
            Qualification.Welding_Process = Clean(Qualification.Welding_Process, 30);
            Qualification.Material_P_No = Clean(Qualification.Material_P_No, 100);
            Qualification.Code_Reference = Clean(Qualification.Code_Reference, 75); // NEW
            Qualification.Consumable_Root_F_No = Clean(Qualification.Consumable_Root_F_No, 20);
            Qualification.Consumable_Root_Spec = Clean(Qualification.Consumable_Root_Spec, 40);
            Qualification.Consumable_Filling_Cap_F_No = Clean(Qualification.Consumable_Filling_Cap_F_No, 100);
            Qualification.Consumable_Filling_Cap_Spec = Clean(Qualification.Consumable_Filling_Cap_Spec, 100);
            Qualification.Position_Progression = Clean(Qualification.Position_Progression, 100);
            Qualification.Diameter_Range = Clean(Qualification.Diameter_Range, 60);
            Qualification.Max_Thickness = Clean(Qualification.Max_Thickness, 60);
            Qualification.WQT_Agency = Clean(Qualification.WQT_Agency, 30);
            Qualification.Received_from_Aramco = Clean(Qualification.Received_from_Aramco, 40);

            var userId = HttpContext.Session.GetInt32("UserID");
            var now = DateTime.UtcNow;

            // Validate required fields for Add Welder flow: Name must be provided
            if (isNewWelderFlow && string.IsNullOrWhiteSpace(Welder.Name))
            {
                TempData["Msg"] = "Welder Name is required.";
                // For AJAX requests return JSON suitable for the client-side form update
                if (isAjax)
                {
                    return Welder.Welder_ID > 0
                        ? await AjaxEditResponseAsync(Welder.Welder_ID, null)
                        : await AjaxEditResponseAsync(0, null);
                }

                // Non-AJAX: redirect back to the edit page (new) so user can correct
                return RedirectToAction(nameof(EditWelder), new { id = Welder.Welder_ID > 0 ? Welder.Welder_ID : (int?)null });
            }

            // Validate Project_Welder for Add Welder flow
            if (isNewWelderFlow && (!Welder.Project_Welder.HasValue || Welder.Project_Welder.GetValueOrDefault() <= 0))
            {
                TempData["Msg"] = "Project Welder is required.";
                if (isAjax)
                {
                    return Welder.Welder_ID > 0
                        ? await AjaxEditResponseAsync(Welder.Welder_ID, null)
                        : await AjaxEditResponseAsync(0, null);
                }

                return RedirectToAction(nameof(EditWelder), new { id = Welder.Welder_ID > 0 ? Welder.Welder_ID : (int?)null });
            }

            // ADD flow: if symbol exists, attach to that welder instead of erroring
            if (Welder.Welder_ID <= 0 && !string.IsNullOrWhiteSpace(Welder.Welder_Symbol))
            {
                var normalizedSymbol = Welder.Welder_Symbol;
                var baseExistingQuery = _context.Welders_tbl.AsNoTracking().Where(x => x.Welder_Symbol != null && EF.Functions.Collate(x.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == normalizedSymbol);
                List<Welder> existingMatches;
                if (Welder.Project_Welder.HasValue && Welder.Project_Welder.Value > 0)
                {
                    existingMatches = await baseExistingQuery.Where(x => x.Project_Welder == Welder.Project_Welder).ToListAsync();
                }
                else
                {
                    existingMatches = await baseExistingQuery.ToListAsync();
                    var distinctProjects = existingMatches.Select(e => e.Project_Welder).Distinct().Count();
                    if (distinctProjects > 1)
                    {
                        TempData["Msg"] = "Welder symbol already exists for this project. Switched to existing welder.";
                        return await AjaxEditResponseAsync(0, null);
                    }
                }
                var existing = existingMatches.FirstOrDefault();
                if (existing != null)
                {
                    // If explicitly adding a qualification, attach to existing welder
                    if (wantsNewQualification)
                    {
                        var evaluation = await EvaluateQualificationSaveAsync(existing, Qualification);
                        if (evaluation.Status == QualificationSaveStatus.Duplicate)
                        {
                            // If the duplicate is the same JCC the user just posted, treat as idempotent success
                            if (!string.IsNullOrWhiteSpace(Qualification.JCC_No)
                                && string.Equals(evaluation.ExistingJcc, Qualification.JCC_No, StringComparison.OrdinalIgnoreCase))
                            {
                                TempData["Msg"] = $"Successfully Saved{(string.IsNullOrWhiteSpace(existing.Welder_Symbol) ? string.Empty : $" - {existing.Welder_Symbol}")}";
                                return await AjaxEditResponseAsync(existing.Welder_ID, evaluation.ExistingJcc);
                            }

                            return await DuplicateQualificationResponseAsync(existing.Welder_ID, evaluation.ExistingJcc);
                        }

                        using var txAdd = await _context.Database.BeginTransactionAsync(IsolationLevel.Serializable);
                        if (string.IsNullOrWhiteSpace(Qualification.JCC_No))
                            Qualification.JCC_No = await GenerateNextJccNoAsync(existing.Project_Welder, existing.Welder_Location, Qualification.Batch_No);

                        await EnsureUniqueJccAsync(Qualification, existing.Project_Welder, existing.Welder_Location);

                        // Allow same qualification data; just proceed to insert with new JCC

                        Qualification.Welder_ID_WL = existing.Welder_ID;
                        Qualification.Welder_List_Updated_By = userId;
                        Qualification.Welder_List_Updated_Date = now;
                        _context.Welder_List_tbl.Add(Qualification);
                        await _context.SaveChangesAsync();
                        await txAdd.CommitAsync();

                        var savedMsgExisting = ApplyQualificationSaveMessage(QualificationSaveStatus.Saved);
                        return await AjaxEditResponseAsync(existing.Welder_ID, Qualification.JCC_No, false, savedMsgExisting);
                    }

                    TempData["Msg"] = "Welder symbol already exists. Switched to existing welder.";
                    return await AjaxEditResponseAsync(existing.Welder_ID, null);
                }
            }

            // Only require qualification fields when explicitly adding a qualification
            if (wantsNewQualification)
            {
                if (string.IsNullOrWhiteSpace(Qualification.Welding_Process) || string.IsNullOrWhiteSpace(Welder.Welder_Symbol))
                {
                    TempData["Msg"] = "Missing required fields for new qualification.";
                    return Welder.Welder_ID > 0
                        ? await AjaxEditResponseAsync(Welder.Welder_ID, null)
                        : await AjaxEditResponseAsync(0, null);
                }
            }

            if (!isNewWelderFlow && !wantsNewQualification && string.IsNullOrWhiteSpace(Qualification.Welding_Process))
            {
                if (hasExistingQualification)
                {
                    TempData["Msg"] = "Select a welding process or use Add New Qualification before saving.";
                    var targetJcc = !string.IsNullOrWhiteSpace(SelectedJcc)
                        ? SelectedJcc.Trim()
                        : (string.IsNullOrWhiteSpace(Qualification.JCC_No) ? null : Qualification.JCC_No);
                    if (isAjax)
                    {
                        return Welder.Welder_ID > 0
                            ? await AjaxEditResponseAsync(Welder.Welder_ID, targetJcc)
                            : await AjaxEditResponseAsync(0, targetJcc);
                    }

                    return RedirectToAction(nameof(EditWelder), new { id = Welder.Welder_ID > 0 ? Welder.Welder_ID : (int?)null, jcc = targetJcc });
                }
                // No existing qualifications: allow welder-only save without forcing a process
            }

            using var tx = await _context.Database.BeginTransactionAsync(IsolationLevel.Serializable);

            // Upsert welder
            Welder? dbWelder = null;
            if (Welder.Welder_ID > 0)
            {
                dbWelder = await _context.Welders_tbl.FirstOrDefaultAsync(x => x.Welder_ID == Welder.Welder_ID);
                if (dbWelder == null)
                {
                    TempData["Msg"] = "Welder not found.";
                    return await AjaxEditResponseAsync(0, null);
                }
            }
            else
            {
                Welder.Welders_Updated_By = userId;
                Welder.Welders_Updated_Date = now;
                _context.Welders_tbl.Add(Welder);
                try
                {
                    await _context.SaveChangesAsync();
                    dbWelder = Welder;
                }
                catch (DbUpdateException dbEx) when (IsDuplicateKey(dbEx))
                {
                    var existingByProject = await _context.Welders_tbl
                        .AsNoTracking()
                        .Where(x => x.Project_Welder == Welder.Project_Welder)
                        .Where(x => x.Welder_Symbol != null && EF.Functions.Collate(x.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == Welder.Welder_Symbol)
                        .FirstOrDefaultAsync();

                    if (existingByProject != null)
                    {
                        TempData["Msg"] = "Welder symbol already exists for this project. Switched to existing welder.";
                        return await AjaxEditResponseAsync(existingByProject.Welder_ID, null);
                    }

                    throw;
                }
            }

            dbWelder ??= Welder;

            // Determine symbol move semantics (edit flow only)
            Welder? targetExistingWelder = null;
            Welder? targetNewWelder = null;
            string? newSymbol = Welder.Welder_Symbol;
            string? oldSymbol = StripAllWhitespace(dbWelder.Welder_Symbol);
            bool symbolChanged = dbWelder.Welder_ID > 0 && !string.IsNullOrWhiteSpace(newSymbol) && !string.Equals(oldSymbol ?? string.Empty, newSymbol, StringComparison.OrdinalIgnoreCase);
            if (symbolChanged)
            {
                targetExistingWelder = await _context.Welders_tbl
                    .AsNoTracking()
                    .Where(x => x.Welder_ID != dbWelder.Welder_ID)
                    .Where(x => x.Welder_Symbol != null && EF.Functions.Collate(x.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == newSymbol)
                    .Where(x => !Welder.Project_Welder.HasValue || Welder.Project_Welder.Value <= 0 || x.Project_Welder == Welder.Project_Welder)
                    .FirstOrDefaultAsync();

                if (targetExistingWelder == null)
                {
                    var conflict = await FindSymbolProjectConflictAsync(newSymbol, Welder.Project_Welder, dbWelder.Welder_ID);
                    if (conflict != null)
                    {
                        TempData["Msg"] = "Welder symbol already exists for this project. Switched to existing welder.";
                        await tx.RollbackAsync();
                        return await AjaxEditResponseAsync(conflict.Welder_ID, null);
                    }

                    // Create a new welder with the new symbol and posted details
                    targetNewWelder = new Welder
                    {
                        Welder_Symbol = newSymbol,
                        Iqama_No = Welder.Iqama_No,
                        Name = Welder.Name,
                        Welder_Location = Welder.Welder_Location,
                        Passport = Welder.Passport,
                        Mobile_No = Welder.Mobile_No,
                        Email = Welder.Email,
                        Mobilization_Date = Welder.Mobilization_Date,
                        Demobilization_Date = Welder.Demobilization_Date,
                        Status = Welder.Status,
                        Project_Welder = Welder.Project_Welder,
                        Welders_Updated_By = userId,
                        Welders_Updated_Date = now
                    };
                    _context.Welders_tbl.Add(targetNewWelder);
                    try
                    {
                        await _context.SaveChangesAsync();
                    }
                    catch (DbUpdateException dbEx) when (IsDuplicateKey(dbEx))
                    {
                        var existingByProject = await _context.Welders_tbl
                            .AsNoTracking()
                            .Where(x => x.Project_Welder == Welder.Project_Welder)
                            .Where(x => x.Welder_Symbol != null && EF.Functions.Collate(x.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == newSymbol)
                            .FirstOrDefaultAsync();

                        if (existingByProject != null)
                        {
                            TempData["Msg"] = "Welder symbol already exists for this project. Switched to existing welder.";
                            return await AjaxEditResponseAsync(existingByProject.Welder_ID, null);
                        }

                        throw;
                    }
                }
            }

            if (!symbolChanged)
            {
                dbWelder.Welder_Symbol = Welder.Welder_Symbol;
            }
            dbWelder.Iqama_No = Welder.Iqama_No;
            dbWelder.Name = Welder.Name;
            dbWelder.Welder_Location = Welder.Welder_Location;
            dbWelder.Passport = Welder.Passport;
            dbWelder.Mobile_No = Welder.Mobile_No;
            dbWelder.Email = Welder.Email;
            dbWelder.Mobilization_Date = Welder.Mobilization_Date;
            dbWelder.Demobilization_Date = Welder.Demobilization_Date;
            dbWelder.Status = Welder.Status;
            dbWelder.Project_Welder = Welder.Project_Welder;
            dbWelder.Welders_Updated_By = userId;
            dbWelder.Welders_Updated_Date = now;

            if (!symbolChanged)
            {
                var conflict = await FindSymbolProjectConflictAsync(dbWelder.Welder_Symbol, dbWelder.Project_Welder, dbWelder.Welder_ID);
                if (conflict != null)
                {
                    TempData["Msg"] = "Welder symbol already exists for this project. Switched to existing welder.";
                    await tx.RollbackAsync();
                    return await AjaxEditResponseAsync(conflict.Welder_ID, null);
                }
            }

            await _context.SaveChangesAsync();

            Welder receiverWelder = targetExistingWelder ?? targetNewWelder ?? dbWelder;

            // Detect stale AddNewQualification flag: if the posted JCC already belongs
            // to the receiver welder this is a re-save, not a new add.
            if (wantsNewQualification
                && !string.IsNullOrWhiteSpace(Qualification.JCC_No)
                && await _context.Welder_List_tbl.AsNoTracking()
                    .AnyAsync(q => q.JCC_No == Qualification.JCC_No && q.Welder_ID_WL == receiverWelder.Welder_ID))
            {
                wantsNewQualification = false;
            }

            bool isBrandNewWelder = isNewWelderFlow && !wantsNewQualification && receiverWelder.Welder_ID > 0 && string.IsNullOrWhiteSpace(Qualification.Welding_Process);

            if (wantsNewQualification)
            {
                if (string.IsNullOrWhiteSpace(Qualification.JCC_No))
                    Qualification.JCC_No = await GenerateNextJccNoAsync(receiverWelder.Project_Welder, receiverWelder.Welder_Location, Qualification.Batch_No);

                await EnsureUniqueJccAsync(Qualification, receiverWelder.Project_Welder, receiverWelder.Welder_Location);

                Qualification.Welder_ID_WL = receiverWelder.Welder_ID;
                Qualification.Welder_List_Updated_By = userId;
                Qualification.Welder_List_Updated_Date = now;

                var evaluation = await EvaluateQualificationSaveAsync(receiverWelder, Qualification);
                if (evaluation.Status == QualificationSaveStatus.Duplicate)
                {
                    // If this is a retry with the same JCC (or auto-generated but not posted), treat as already saved
                    if ((!string.IsNullOrWhiteSpace(Qualification.JCC_No) && string.Equals(evaluation.ExistingJcc, Qualification.JCC_No, StringComparison.OrdinalIgnoreCase))
                        || (string.IsNullOrWhiteSpace(Qualification.JCC_No) && !string.IsNullOrWhiteSpace(evaluation.ExistingJcc)))
                    {
                        var savedMsgDuplicate = ApplyQualificationSaveMessage(QualificationSaveStatus.Saved);
                        return await AjaxEditResponseAsync(receiverWelder.Welder_ID, evaluation.ExistingJcc, false, savedMsgDuplicate);
                    }

                    return await DuplicateQualificationResponseAsync(receiverWelder.Welder_ID, evaluation.ExistingJcc);
                }

                // Insert with basic retry for PK collision (rare race)
                const int maxAttempts = 2;
                for (int attempt = 1; attempt <= maxAttempts; attempt++)
                {
                    try
                    {
                        _context.Welder_List_tbl.Add(Qualification);
                        await _context.SaveChangesAsync();
                        break; // success
                    }
                    catch (DbUpdateException dbEx) when (IsDuplicateKey(dbEx) && attempt < maxAttempts)
                    {
                        // Regenerate and retry
                        _logger.LogWarning(dbEx, "Duplicate JCC_No detected, regenerating (attempt {Attempt})", attempt);
                        Qualification.JCC_No = await GenerateNextJccNoAsync(receiverWelder.Project_Welder, receiverWelder.Welder_Location, Qualification.Batch_No);
                        await EnsureUniqueJccAsync(Qualification, receiverWelder.Project_Welder, receiverWelder.Welder_Location);
                        _context.Entry(Qualification).State = EntityState.Detached; // detach old
                        Qualification.Welder_ID_WL = receiverWelder.Welder_ID; // re-assign just in case
                    }
                }

                await tx.CommitAsync();
                var savedMsg = QualificationSavedMessage;
                TempData["Msg"] = savedMsg;
                return await AjaxEditResponseAsync(receiverWelder.Welder_ID, Qualification.JCC_No, false, savedMsg);
            }
            else
            {
                // If brand new welder and no qualification intent/data -> finalize now
                if (isBrandNewWelder)
                {
                    await tx.CommitAsync();
                    var sym = string.IsNullOrWhiteSpace(receiverWelder.Welder_Symbol) ? "" : $" - {receiverWelder.Welder_Symbol}";
                    TempData["Msg"] = $"Successfully Saved{sym}";
                    return await AjaxEditResponseAsync(receiverWelder.Welder_ID, null);
                }

                var originalJcc = string.IsNullOrWhiteSpace(SelectedJcc) ? Qualification.JCC_No : SelectedJcc.Trim();
                bool jccChanged = !string.IsNullOrWhiteSpace(originalJcc) && !string.IsNullOrWhiteSpace(Qualification.JCC_No) &&
                                  !string.Equals(originalJcc, Qualification.JCC_No, StringComparison.OrdinalIgnoreCase);

                WelderQualification? dbQual = null;
                if (!string.IsNullOrWhiteSpace(originalJcc))
                    dbQual = await _context.Welder_List_tbl.FirstOrDefaultAsync(q => q.JCC_No == originalJcc);
                else if (!string.IsNullOrWhiteSpace(Qualification.JCC_No))
                    dbQual = await _context.Welder_List_tbl.FirstOrDefaultAsync(q => q.JCC_No == Qualification.JCC_No);

                if (dbQual == null && !string.IsNullOrWhiteSpace(Qualification.JCC_No))
                {
                    await EnsureUniqueJccAsync(Qualification, receiverWelder.Project_Welder, receiverWelder.Welder_Location);
                    Qualification.Welder_ID_WL = receiverWelder.Welder_ID;
                    Qualification.Welder_List_Updated_By = userId;
                    Qualification.Welder_List_Updated_Date = now;
                    _context.Welder_List_tbl.Add(Qualification);
                }
                else if (dbQual != null && jccChanged)
                {
                    await EnsureUniqueJccAsync(Qualification, receiverWelder.Project_Welder, receiverWelder.Welder_Location);
                    var newExists = await _context.Welder_List_tbl.AnyAsync(q => q.JCC_No == Qualification.JCC_No);
                    if (newExists)
                    {
                        TempData["Msg"] = "JCC No already exists.";
                        await tx.RollbackAsync();
                        return await AjaxEditResponseAsync(receiverWelder.Welder_ID, originalJcc);
                    }

                    var dupOnReceiver = await FindDuplicateQualificationAsync(receiverWelder.Welder_ID, receiverWelder.Welder_Symbol, receiverWelder.Project_Welder, Qualification, originalJcc);
                    if (!string.IsNullOrEmpty(dupOnReceiver))
                    {
                        await tx.RollbackAsync();
                        return await DuplicateQualificationResponseAsync(receiverWelder.Welder_ID, dupOnReceiver);
                    }

                    var replacement = new WelderQualification
                    {
                        JCC_No = Qualification.JCC_No!,
                        Welder_ID_WL = receiverWelder.Welder_ID,
                        Test_Date = Qualification.Test_Date,
                        Welding_Process = Qualification.Welding_Process,
                        Material_P_No = Qualification.Material_P_No,
                        Code_Reference = Qualification.Code_Reference,
                        Consumable_Root_F_No = Qualification.Consumable_Root_F_No,
                        Consumable_Root_Spec = Qualification.Consumable_Root_Spec,
                        Consumable_Filling_Cap_F_No = Qualification.Consumable_Filling_Cap_F_No,
                        Consumable_Filling_Cap_Spec = Qualification.Consumable_Filling_Cap_Spec,
                        Position_Progression = Qualification.Position_Progression,
                        Diameter_Range = Qualification.Diameter_Range,
                        Max_Thickness = Qualification.Max_Thickness,
                        Date_Issued = Qualification.Date_Issued,
                        Remarks = Qualification.Remarks,
                        Qualification_Cert_Ref_No = Qualification.Qualification_Cert_Ref_No,
                        WQT_Agency = Qualification.WQT_Agency,
                        DATE_OF_LAST_CONTINUITY = Qualification.DATE_OF_LAST_CONTINUITY,
                        RECORDING_THE_CONTINUITY_RECORD = Qualification.RECORDING_THE_CONTINUITY_RECORD,
                        Batch_No = Qualification.Batch_No,
                        Received_from_Aramco = Qualification.Received_from_Aramco,
                        Welder_List_Updated_By = userId,
                        Welder_List_Updated_Date = now
                    };

                    _context.Welder_List_tbl.Add(replacement);
                    if (dbQual != null) _context.Welder_List_tbl.Remove(dbQual);
                }
                else if (dbQual != null)
                {
                    if (receiverWelder.Welder_ID != dbQual.Welder_ID_WL)
                    {
                        // Allow same qualification data; only JCC uniqueness enforced
                    }

                    dbQual.Test_Date = Qualification.Test_Date;
                    dbQual.Welding_Process = Qualification.Welding_Process;
                    dbQual.Material_P_No = Qualification.Material_P_No;
                    dbQual.Code_Reference = Qualification.Code_Reference;
                    dbQual.Consumable_Root_F_No = Qualification.Consumable_Root_F_No;
                    dbQual.Consumable_Root_Spec = Qualification.Consumable_Root_Spec;
                    dbQual.Consumable_Filling_Cap_F_No = Qualification.Consumable_Filling_Cap_F_No;
                    dbQual.Consumable_Filling_Cap_Spec = Qualification.Consumable_Filling_Cap_Spec;
                    dbQual.Position_Progression = Qualification.Position_Progression;
                    dbQual.Diameter_Range = Qualification.Diameter_Range;
                    dbQual.Max_Thickness = Qualification.Max_Thickness;
                    dbQual.Date_Issued = Qualification.Date_Issued;
                    dbQual.Remarks = Qualification.Remarks;
                    dbQual.Qualification_Cert_Ref_No = Qualification.Qualification_Cert_Ref_No;
                    dbQual.WQT_Agency = Qualification.WQT_Agency;
                    dbQual.DATE_OF_LAST_CONTINUITY = Qualification.DATE_OF_LAST_CONTINUITY;
                    dbQual.RECORDING_THE_CONTINUITY_RECORD = Qualification.RECORDING_THE_CONTINUITY_RECORD;
                    dbQual.Batch_No = Qualification.Batch_No;
                    dbQual.Received_from_Aramco = Qualification.Received_from_Aramco;
                    dbQual.Welder_ID_WL = receiverWelder.Welder_ID;
                    dbQual.Welder_List_Updated_By = userId;
                    dbQual.Welder_List_Updated_Date = now;
                }

                await _context.SaveChangesAsync();
                await tx.CommitAsync();

                var savedMsg = QualificationSavedMessage;
                TempData["Msg"] = savedMsg;
                return await AjaxEditResponseAsync(receiverWelder.Welder_ID, Qualification.JCC_No, false, savedMsg);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SaveWelder failed");
            TempData["Msg"] = "Save failed. Please try again.";
            var idStr = Request.Form["Welder.Welder_ID"];
            if (int.TryParse(idStr, out var backId) && backId > 0)
                return await AjaxEditResponseAsync(backId, null);
            return RedirectToAction(nameof(WelderList));
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteQualification([FromForm] int welderId, [FromForm] string jcc)
    {
        try
        {
            if (welderId <= 0 || string.IsNullOrWhiteSpace(jcc))
            {
                TempData["Msg"] = "Invalid request.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId });
            }

            var qual = await _context.Welder_List_tbl
                .FirstOrDefaultAsync(q => q.Welder_ID_WL == welderId && q.JCC_No == jcc);

            if (qual == null)
            {
                TempData["Msg"] = "Qualification not found.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId });
            }

            _context.Welder_List_tbl.Remove(qual);
            await _context.SaveChangesAsync();

            TempData["Msg"] = $"Deleted qualification {jcc}.";

            var nextJcc = await _context.Welder_List_tbl
                .AsNoTracking()
                .Where(q => q.Welder_ID_WL == welderId)
                .OrderByDescending(q => q.Welder_List_Updated_Date)
                .ThenByDescending(q => q.Date_Issued)
                .ThenByDescending(q => q.Test_Date)
                .Select(q => q.JCC_No)
                .FirstOrDefaultAsync();

            return string.IsNullOrEmpty(nextJcc)
                ? RedirectToAction(nameof(EditWelder), new { id = welderId })
                : RedirectToAction(nameof(EditWelder), new { id = welderId, jcc = nextJcc });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteQualification failed for Welder {WelderId}, JCC {Jcc}", welderId, jcc);
            TempData["Msg"] = "Delete failed. Please try again.";
            return RedirectToAction(nameof(EditWelder), new { id = welderId });
        }
    }

    #endregion

    #region Welder Lookups

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> FindWelderBySymbol([FromQuery] string symbol, [FromQuery] int? projectId)
    {
        symbol = StripWhitespace((symbol ?? string.Empty).Trim());
        if (string.IsNullOrWhiteSpace(symbol)) return BadRequest(new { error = "Symbol required" });

        // Normalize project filter (<=0 treated as not provided)
        int? pid = (projectId.HasValue && projectId.Value > 0) ? projectId : null;

        var baseQuery = _context.Welders_tbl
            .AsNoTracking()
            .Where(x => x.Welder_Symbol != null
                && EF.Functions.Collate(x.Welder_Symbol, "SQL_Latin1_General_CP1_CI_AS") == symbol);

        List<Welder> matches;
        if (pid.HasValue)
        {
            matches = await baseQuery.Where(x => x.Project_Welder == pid.Value).ToListAsync();
            if (matches.Count == 0)
            {
                return NotFound(new { error = "Welder not found for project." });
            }
        }
        else
        {
            matches = await baseQuery.ToListAsync();
            // If symbol is used by multiple projects, force caller to specify project
            var distinctProjects = matches
                .Select(m => m.Project_Welder)
                .Distinct()
                .Count();
            if (distinctProjects > 1)
            {
                return BadRequest(new { error = "Multiple welders found for this symbol. Select a project." });
            }
        }

        var w = matches.FirstOrDefault();

        if (w == null) return NotFound();

        return Json(new
        {
            w.Welder_ID,
            w.Welder_Symbol,
            w.Name,
            w.Iqama_No,
            w.Passport,
            w.Welder_Location,
            w.Mobile_No,
            w.Email,
            Mobilization_Date = w.Mobilization_Date?.ToString("yyyy-MM-dd"),
            Demobilization_Date = w.Demobilization_Date?.ToString("yyyy-MM-dd"),
            w.Status,
            w.Project_Welder
        });
    }

    // NEW: qualifications list endpoint used by Add Welder dynamic lookup
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetQualifications([FromQuery] int welderId, [FromQuery] int? projectId = null)
    {
        if (welderId <= 0) return BadRequest(new { error = "Invalid welder id" });
        var query = _context.Welder_List_tbl
            .AsNoTracking()
            .Where(q => q.Welder_ID_WL == welderId)
            .Join(_context.Welders_tbl.AsNoTracking(), q => q.Welder_ID_WL, w => w.Welder_ID, (q, w) => new { q, w });

        if (projectId.HasValue && projectId.Value > 0)
        {
            query = query.Where(x => x.w.Project_Welder == projectId.Value);
        }

        var items = await query
            .OrderBy(x => x.q.JCC_No)
            .Select(x => new
            {
                x.q.JCC_No,
                x.q.Welding_Process,
                x.q.Material_P_No,
                x.q.Diameter_Range,
                x.q.Max_Thickness
            })
            .ToListAsync();
        return Json(new { items });
    }

    #endregion

    #region Qualification Options

    // Helper: normalize and filter empty values (preserving original order)
    static List<string?> NormalizeAndFilter(IEnumerable<string?>? src)
    {
        if (src == null) return new List<string?>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var final = new List<string?>();
        foreach (var s in src)
        {
            var low = s?.Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(low) || seen.Contains(low)) continue;
            seen.Add(low);
            final.Add(s);
        }
        return final;
    }

    private sealed record Row(
        string? Welding_Process,
        string? Material_P_No,
        string? Code_Reference,
        string? Consumable_Root_F_No,
        string? Consumable_Root_Spec,
        string? Consumable_Filling_Cap_F_No,
        string? Consumable_Filling_Cap_Spec,
        string? Position_Progression,
        string? Diameter_Range,
        string? Max_Thickness,
        string? WQT_Agency,
        string? Received_from_Aramco,
        int? Batch_No
    );

    private sealed record QualOptionsResult(
        List<string> WeldingProcesses,
        List<string> MaterialPNos,
        List<string> CodeReferences,
        List<string> RootFNos,
        List<string> RootSpecs,
        List<string> FillCapFNos,
        List<string> FillCapSpecs,
        List<string> PositionProgressions,
        List<string> DiameterRanges,
        List<string> MaxThicknesses,
        IDictionary<string, string?> Effective
    );

    private static string? NormalizeQualInput(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var trimmed = value.Trim();
        return string.Equals(trimmed, "__new__", StringComparison.OrdinalIgnoreCase) ? null : trimmed;
    }

    private static bool EqualsCI(string? a, string? b)
        => !string.IsNullOrWhiteSpace(a)
           && !string.IsNullOrWhiteSpace(b)
           && string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);

    private static List<string> DistinctQualList(IEnumerable<string?> source)
        => source
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();

    private async Task<QualOptionsResult> BuildQualOptionsAsync(
        int? batchNo,
        string? weldingProcess,
        string? materialPNo,
        string? codeReference,
        string? rootFNo,
        string? rootSpec,
        string? fillCapFNo,
        string? fillCapSpec,
        string? positionProgression,
        string? diameterRange,
        string? maxThickness,
        bool allowEmpty = false)
    {
        var selProc = NormalizeQualInput(weldingProcess);
        var selMaterial = NormalizeQualInput(materialPNo);
        var selCodeRef = NormalizeQualInput(codeReference);
        var selRootF = NormalizeQualInput(rootFNo);
        var selRootSpec = NormalizeQualInput(rootSpec);
        var selFillCapF = NormalizeQualInput(fillCapFNo);
        var selFillCapSpec = NormalizeQualInput(fillCapSpec);
        var selPos = NormalizeQualInput(positionProgression);
        var selDia = NormalizeQualInput(diameterRange);
        var selThk = NormalizeQualInput(maxThickness);

        IQueryable<Row> BuildRowQuery(IQueryable<WelderQualification> source)
            => source.Select(q => new Row(
                q.Welding_Process,
                q.Material_P_No,
                q.Code_Reference,
                q.Consumable_Root_F_No,
                q.Consumable_Root_Spec,
                q.Consumable_Filling_Cap_F_No,
                q.Consumable_Filling_Cap_Spec,
                q.Position_Progression,
                q.Diameter_Range,
                q.Max_Thickness,
                q.WQT_Agency,
                q.Received_from_Aramco,
                q.Batch_No
            ));

        var baseQuery = _context.Welder_List_tbl.AsNoTracking();
        if (batchNo.HasValue && batchNo.Value > 0)
        {
            baseQuery = baseQuery.Where(q => q.Batch_No == batchNo.Value);
        }

        var rows = await BuildRowQuery(baseQuery).ToListAsync();

        if (rows.Count == 0 && batchNo.HasValue && batchNo.Value > 0 && !allowEmpty)
        {
            rows = await BuildRowQuery(_context.Welder_List_tbl.AsNoTracking()).ToListAsync();
        }

        var selectionMap = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            [nameof(Row.Welding_Process)] = selProc,
            [nameof(Row.Material_P_No)] = selMaterial,
            [nameof(Row.Code_Reference)] = selCodeRef,
            [nameof(Row.Consumable_Root_F_No)] = selRootF,
            [nameof(Row.Consumable_Root_Spec)] = selRootSpec,
            [nameof(Row.Consumable_Filling_Cap_F_No)] = selFillCapF,
            [nameof(Row.Consumable_Filling_Cap_Spec)] = selFillCapSpec,
            [nameof(Row.Position_Progression)] = selPos,
            [nameof(Row.Diameter_Range)] = selDia,
            [nameof(Row.Max_Thickness)] = selThk
        };

        var orderedFields = new[]
        {
            nameof(Row.Welding_Process),
            nameof(Row.Material_P_No),
            nameof(Row.Code_Reference),
            nameof(Row.Consumable_Root_F_No),
            nameof(Row.Consumable_Root_Spec),
            nameof(Row.Consumable_Filling_Cap_F_No),
            nameof(Row.Consumable_Filling_Cap_Spec),
            nameof(Row.Position_Progression),
            nameof(Row.Diameter_Range),
            nameof(Row.Max_Thickness)
        };

        string? FieldValue(Row row, string field) => field switch
        {
            nameof(Row.Welding_Process) => row.Welding_Process,
            nameof(Row.Material_P_No) => row.Material_P_No,
            nameof(Row.Code_Reference) => row.Code_Reference,
            nameof(Row.Consumable_Root_F_No) => row.Consumable_Root_F_No,
            nameof(Row.Consumable_Root_Spec) => row.Consumable_Root_Spec,
            nameof(Row.Consumable_Filling_Cap_F_No) => row.Consumable_Filling_Cap_F_No,
            nameof(Row.Consumable_Filling_Cap_Spec) => row.Consumable_Filling_Cap_Spec,
            nameof(Row.Position_Progression) => row.Position_Progression,
            nameof(Row.Diameter_Range) => row.Diameter_Range,
            nameof(Row.Max_Thickness) => row.Max_Thickness,
            _ => null
        };

        IEnumerable<Row> FilterForField(string targetField)
        {
            // Apply cascading filtering: each field is constrained only by selections made earlier in the ordered list.
            var applicableKeys = orderedFields.TakeWhile(f => !f.Equals(targetField, StringComparison.Ordinal)).ToArray();

            return rows.Where(row =>
            {
                foreach (var key in applicableKeys)
                {
                    if (!selectionMap.TryGetValue(key, out var selected) || selected == null) continue;
                    if (!EqualsCI(FieldValue(row, key), selected)) return false;
                }
                return true;
            });
        }

        var weldingProcesses = DistinctQualList(FilterForField(nameof(Row.Welding_Process)).Select(r => r.Welding_Process));
        var materialPNos = DistinctQualList(FilterForField(nameof(Row.Material_P_No)).Select(r => r.Material_P_No));
        var codeReferences = DistinctQualList(FilterForField(nameof(Row.Code_Reference)).Select(r => r.Code_Reference));
        var rootFNos = DistinctQualList(FilterForField(nameof(Row.Consumable_Root_F_No)).Select(r => r.Consumable_Root_F_No));
        var rootSpecs = DistinctQualList(FilterForField(nameof(Row.Consumable_Root_Spec)).Select(r => r.Consumable_Root_Spec));
        var fillCapFNos = DistinctQualList(FilterForField(nameof(Row.Consumable_Filling_Cap_F_No)).Select(r => r.Consumable_Filling_Cap_F_No));
        var fillCapSpecs = DistinctQualList(FilterForField(nameof(Row.Consumable_Filling_Cap_Spec)).Select(r => r.Consumable_Filling_Cap_Spec));
        var positionProgressions = DistinctQualList(FilterForField(nameof(Row.Position_Progression)).Select(r => r.Position_Progression));
        var diameterRanges = DistinctQualList(FilterForField(nameof(Row.Diameter_Range)).Select(r => r.Diameter_Range));
        var maxThicknesses = DistinctQualList(FilterForField(nameof(Row.Max_Thickness)).Select(r => r.Max_Thickness));

        Row? bestRow = null;
        int bestScore = -1;
        int bestBatch = int.MinValue;
        foreach (var row in rows)
        {
            int score = 0;
            if (selProc != null)
            {
                if (!EqualsCI(row.Welding_Process, selProc)) continue;
                score++;
            }

            if (selMaterial != null)
            {
                if (!EqualsCI(row.Material_P_No, selMaterial)) continue;
                score++;
            }

            if (selCodeRef != null)
            {
                if (!EqualsCI(row.Code_Reference, selCodeRef)) continue;
                score++;
            }

            if (selRootF != null)
            {
                if (!EqualsCI(row.Consumable_Root_F_No, selRootF)) continue;
                score++;
            }

            if (selRootSpec != null)
            {
                if (!EqualsCI(row.Consumable_Root_Spec, selRootSpec)) continue;
                score++;
            }

            if (selFillCapF != null)
            {
                if (!EqualsCI(row.Consumable_Filling_Cap_F_No, selFillCapF)) continue;
                score++;
            }

            if (selFillCapSpec != null)
            {
                if (!EqualsCI(row.Consumable_Filling_Cap_Spec, selFillCapSpec)) continue;
                score++;
            }

            if (selPos != null)
            {
                if (!EqualsCI(row.Position_Progression, selPos)) continue;
                score++;
            }

            if (selDia != null)
            {
                if (!EqualsCI(row.Diameter_Range, selDia)) continue;
                score++;
            }

            if (selThk != null)
            {
                if (!EqualsCI(row.Max_Thickness, selThk)) continue;
                score++;
            }

            var batchScore = row.Batch_No ?? 0;
            if (score > bestScore || (score == bestScore && batchScore > bestBatch))
            {
                bestRow = row;
                bestScore = score;
                bestBatch = batchScore;
            }
        }

        bestRow ??= rows
            .OrderByDescending(r => r.Batch_No ?? 0)
            .FirstOrDefault();

        var effective = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        void AddEffective(string key, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            effective[key] = value.Trim();
        }

        if (bestRow != null)
        {
            AddEffective("weldingProcess", bestRow.Welding_Process);
            AddEffective("materialPNo", bestRow.Material_P_No);
            AddEffective("codeReference", bestRow.Code_Reference);
            AddEffective("rootFNo", bestRow.Consumable_Root_F_No);
            AddEffective("rootSpec", bestRow.Consumable_Root_Spec);
            AddEffective("fillCapFNo", bestRow.Consumable_Filling_Cap_F_No);
            AddEffective("fillCapSpec", bestRow.Consumable_Filling_Cap_Spec);
            AddEffective("positionProgression", bestRow.Position_Progression);
            AddEffective("diameterRange", bestRow.Diameter_Range);
            AddEffective("maxThickness", bestRow.Max_Thickness);
            AddEffective("wqtAgency", bestRow.WQT_Agency);
            AddEffective("receivedFromAramco", bestRow.Received_from_Aramco);
        }

        return new QualOptionsResult(
            weldingProcesses,
            materialPNos,
            codeReferences,
            rootFNos,
            rootSpecs,
            fillCapFNos,
            fillCapSpecs,
            positionProgressions,
            diameterRanges,
            maxThicknesses,
            effective
        );
    }

    // NEW: advanced qualification options lookup (batch-aware, with effective default determination)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetQualOptionsAdvanced(
        [FromQuery] int? batchNo,
        [FromQuery] string? weldingProcess,
        [FromQuery] string? materialPNo,
        [FromQuery] string? codeReference,
        [FromQuery] string? rootFNo,
        [FromQuery] string? rootSpec,
        [FromQuery] string? fillCapFNo,
        [FromQuery] string? fillCapSpec,
        [FromQuery] string? positionProgression,
        [FromQuery] string? diameterRange,
        [FromQuery] string? maxThickness,
        [FromQuery] bool allowEmpty = false
    )
    {
        try
        {
            var result = await BuildQualOptionsAsync(batchNo, weldingProcess, materialPNo, codeReference, rootFNo, rootSpec, fillCapFNo, fillCapSpec, positionProgression, diameterRange, maxThickness, allowEmpty);
            return Json(new
            {
                weldingProcesses = result.WeldingProcesses,
                materialPNos = result.MaterialPNos,
                codeReferences = result.CodeReferences,
                rootFNos = result.RootFNos,
                rootSpecs = result.RootSpecs,
                fillCapFNos = result.FillCapFNos,
                fillCapSpecs = result.FillCapSpecs,
                positionProgressions = result.PositionProgressions,
                diameterRanges = result.DiameterRanges,
                maxThicknesses = result.MaxThicknesses,
                effective = result.Effective
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetQualOptionsAdvanced failed");
            return BadRequest(new { error = "Failed to determine qualification options" });
        }
    }

    #endregion

    #region Report Helpers

    // Helper: generate next JCC number using Report_No_Form_tbl prefix/suffix (5-digit serial)
    private async Task<string> GenerateNextJccNoAsync(int? projectId, string? welderLocation, int? batchNo = null)
    {
        _ = batchNo;
        // Normalize project id: treat null/<=0 as unspecified
        int? pid = projectId.HasValue && projectId.Value > 0 ? projectId.Value : null;

        // Map location to Report_No_Form.tbl.Report_Location: Field -> FW, Shop/other -> WS
        static string LocationCode(string? loc)
        {
            var norm = NormalizeLocation(loc);
            if (string.Equals(norm, "Field", StringComparison.OrdinalIgnoreCase)) return "FW";
            return "WS"; // default to Shop code
        }

        string locCode = LocationCode(welderLocation);
        string? prefix = null;
        string suffix = string.Empty;

        try
        {
            // Base query constrained to JCC report type with non-empty form values
            var baseReportQuery = _context.Report_No_Form_tbl
                .AsNoTracking()
                .Where(r => !string.IsNullOrWhiteSpace(r.Report_Type) && EF.Functions.Collate(r.Report_Type!, "SQL_Latin1_General_CP1_CI_AS") == "JCC")
                .Where(r => !string.IsNullOrWhiteSpace(r.Report_No_Form));

            // STRICT match: Project + mapped location + Report_Type = JCC
            if (pid.HasValue)
            {
                var strictRow = await baseReportQuery
                    .Where(r => r.Project_Report == pid.Value)
                    .Where(r => r.Report_Location != null && EF.Functions.Collate(r.Report_Location!, "SQL_Latin1_General_CP1_CI_AS") == locCode)
                    .OrderByDescending(r => r.Project_Report)
                    .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                    .FirstOrDefaultAsync();

                if (strictRow != null && !string.IsNullOrWhiteSpace(strictRow.Report_No_Form))
                {
                    prefix = strictRow.Report_No_Form.Trim();
                    if (!string.IsNullOrWhiteSpace(strictRow.Remarks))
                        suffix = strictRow.Remarks.Trim();
                }
            }

            if (string.IsNullOrWhiteSpace(prefix))
            {
                // Prefer project-scoped rows for JCC when project is supplied
                var projectRows = pid.HasValue
                    ? baseReportQuery
                        .Where(r => r.Project_Report == pid.Value)
                    : null;

                var baseReportQueryWithProject = baseReportQuery
                    .Where(r => !pid.HasValue || r.Project_Report == pid.Value);

                // First try location-specific prefix (project rows first when available)
                var reportRow = projectRows != null
                    ? await projectRows
                        .Where(r => r.Report_Location != null && r.Report_Location.Trim().Equals(locCode, StringComparison.OrdinalIgnoreCase))
                        .OrderByDescending(r => r.Project_Report)
                        .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                        .FirstOrDefaultAsync()
                    : null;

                reportRow ??= await baseReportQueryWithProject
                    .Where(r => r.Report_Location != null && r.Report_Location.Trim().Equals(locCode, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(r => pid.HasValue && r.Project_Report == pid.Value)
                    .ThenByDescending(r => r.Project_Report)
                    .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                    .FirstOrDefaultAsync();

                // If no location match, fall back to any JCC report form (project-aware) before legacy default
                reportRow ??= await baseReportQueryWithProject
                    .OrderByDescending(r => pid.HasValue && r.Project_Report == pid.Value)
                    .ThenByDescending(r => r.Project_Report)
                    .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                    .FirstOrDefaultAsync();

                // If still none and project specified, take any project row regardless of type/location
                if (reportRow == null && projectRows != null)
                {
                    reportRow = await projectRows
                        .OrderByDescending(r => r.Project_Report)
                        .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                        .FirstOrDefaultAsync();
                }

                // Project-scoped fallback (ignore location/type) before global
                reportRow ??= await _context.Report_No_Form_tbl
                    .AsNoTracking()
                    .Where(r => !string.IsNullOrWhiteSpace(r.Report_No_Form))
                    .Where(r => !pid.HasValue || r.Project_Report == pid.Value)
                    .OrderByDescending(r => r.Project_Report)
                    .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                    .FirstOrDefaultAsync();

                // Last-chance: allow any Report_No_Form (regardless of Report_Type/location/project)
                reportRow ??= await _context.Report_No_Form_tbl
                    .AsNoTracking()
                    .Where(r => !string.IsNullOrWhiteSpace(r.Report_No_Form))
                    .OrderByDescending(r => pid.HasValue && r.Project_Report == pid.Value)
                    .ThenByDescending(r => r.Project_Report)
                    .Select(r => new { Report_No_Form = (r.Report_No_Form ?? ""), Remarks = (r.Remarks ?? ""), r.Project_Report })
                    .FirstOrDefaultAsync();

                if (reportRow != null && !string.IsNullOrWhiteSpace(reportRow.Report_No_Form))
                {
                    prefix = reportRow.Report_No_Form!.Trim();
                    if (!string.IsNullOrWhiteSpace(reportRow.Remarks))
                        suffix = reportRow.Remarks!.Trim();
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load Report_No_Form prefix for JCC generation");
        }

        var baseQ = _context.Welder_List_tbl
            .AsNoTracking()
            .Where(x => x.JCC_No != null && x.JCC_No != "")
            .Join(
                _context.Welders_tbl.AsNoTracking(),
                q => q.Welder_ID_WL,
                w => w.Welder_ID,
                (q, w) => new { q.JCC_No, q.Batch_No, ProjectId = w.Project_Welder, Location = w.Welder_Location }
            );

        if (pid.HasValue)
            baseQ = baseQ.Where(x => x.ProjectId == pid.Value);

        var list = await baseQ.Select(x => x.JCC_No!).ToListAsync();
        int maxNum = 0; string? sample = null;

        foreach (var s in list)
        {
            var candidate = (s ?? string.Empty).Trim();
            var working = candidate;
            if (!string.IsNullOrEmpty(prefix)
                && working.Length >= prefix.Length
                && working.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                working = working[prefix.Length..];
            if (!string.IsNullOrEmpty(suffix) && working.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                working = working[..^suffix.Length];

            var m = LastDigitsRegex().Match(working);
            if (m.Success && int.TryParse(m.Value, out var n))
            {
                if (n > maxNum)
                {
                    maxNum = n;
                    sample = candidate;
                }
            }
        }

        // Legacy fallback if no prefix match was found
        if (maxNum == 0 && string.IsNullOrEmpty(prefix))
        {
            foreach (var s in list)
            {
                var m = LastDigitsRegex().Match(s);
                if (m.Success && int.TryParse(m.Value, out var n))
                {
                    if (n > maxNum)
                    {
                        maxNum = n;
                        sample = s;
                    }
                }
            }
        }

        var serial = (maxNum + 1).ToString("D5");

        if (!string.IsNullOrEmpty(prefix))
            return $"{prefix}{serial}{suffix}";

        var prefixFallback = sample != null
            ? StripTrailingNumberAndLettersRegex().Replace(sample, "")
            : prefix ?? string.Empty;
        return $"{prefixFallback}{serial}";
    }

    private static bool IsDuplicateKey(DbUpdateException ex)
        => ex.InnerException is SqlException sql && (sql.Number == 2627 || sql.Number == 2601);

    #endregion

    #region Qualification Files

    // NEW: Upload qualification file (PDF or image)
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UploadQualificationFile(string jcc, int welderId, IFormFile? file)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(jcc) || welderId <= 0 || file == null || file.Length == 0)
            {
                TempData["Msg"] = "Invalid upload request.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
            }

            var qual = await _context.Welder_List_tbl.FirstOrDefaultAsync(q => q.JCC_No == jcc && q.Welder_ID_WL == welderId);
            if (qual == null)
            {
                TempData["Msg"] = "Qualification not found.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId });
            }

            var ct = file.ContentType?.ToLowerInvariant() ?? string.Empty;
            bool isPdf = ct.Contains("pdf");
            bool isImage = ct.StartsWith("image/");
            if (!isPdf && !isImage)
            {
                TempData["Msg"] = "Unsupported file type. Only PDF or images allowed.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
            }
            var ext = Path.GetExtension(file.FileName);
            var allowedExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp" };
            if (!allowedExt.Contains(ext))
            {
                TempData["Msg"] = "Unsupported file extension.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
            }

            var root = Path.Combine(_env.ContentRootPath, "App_Data", "WelderQualifications");
            Directory.CreateDirectory(root);
            var blobName = $"{Guid.NewGuid():N}{ext}";
            var fullPath = Path.Combine(root, blobName);
            if (!string.IsNullOrWhiteSpace(qual.JCC_BlobName))
            {
                var oldPath = Path.Combine(root, qual.JCC_BlobName);
                if (System.IO.File.Exists(oldPath))
                {
                    try { System.IO.File.Delete(oldPath); } catch (Exception exDel) { _logger.LogWarning(exDel, "Failed deleting old qualification file {Blob}", qual.JCC_BlobName); }
                }
            }
            using (var fs = System.IO.File.Create(fullPath)) { await file.CopyToAsync(fs); }

            // Update ONLY file metadata (rule #1: do not change qualification updated audit fields on file uploads)
            qual.JCC_FileName = (file.FileName ?? string.Empty).Trim();
            if (qual.JCC_FileName.Length > 100) qual.JCC_FileName = qual.JCC_FileName[..100];
            qual.JCC_FileSize = (int)Math.Min(int.MaxValue, file.Length);
            qual.JCC_UploadDate = DateTime.UtcNow;
            qual.JCC_Upload_By = HttpContext.Session.GetInt32("UserID");
            qual.JCC_BlobName = blobName;
            await _context.SaveChangesAsync();

            TempData["Msg"] = "Qualification file uploaded.";
            return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "UploadQualificationFile failed for {Jcc}/{Welder}", jcc, welderId);
            TempData["Msg"] = "Upload failed.";
            return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
        }
    }

    // NEW: Delete qualification file (remove blob + metadata only)
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteQualificationFile(string jcc, int welderId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(jcc) || welderId <= 0)
            {
                TempData["Msg"] = "Invalid delete request.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId });
            }
            var qual = await _context.Welder_List_tbl.FirstOrDefaultAsync(q => q.JCC_No == jcc && q.Welder_ID_WL == welderId);
            if (qual == null)
            {
                TempData["Msg"] = "Qualification not found.";
                return RedirectToAction(nameof(EditWelder), new { id = welderId });
            }
            if (!string.IsNullOrWhiteSpace(qual.JCC_BlobName))
            {
                var root = Path.Combine(_env.ContentRootPath, "App_Data", "WelderQualifications");
                var fullPath = Path.Combine(root, qual.JCC_BlobName);
                if (System.IO.File.Exists(fullPath))
                {
                    try { System.IO.File.Delete(fullPath); } catch (Exception exDel) { _logger.LogWarning(exDel, "Failed deleting qualification file {Blob}", qual.JCC_BlobName); }
                }
            }
            // Clear metadata (do not touch audit Updated fields per rule #1)
            qual.JCC_FileName = null;
            qual.JCC_FileSize = null;
            qual.JCC_UploadDate = null;
            qual.JCC_Upload_By = null;
            qual.JCC_BlobName = null;
            await _context.SaveChangesAsync();
            TempData["Msg"] = "Qualification file deleted.";
            return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteQualificationFile failed for {Jcc}/{Welder}", jcc, welderId);
            TempData["Msg"] = "Delete file failed.";
            return RedirectToAction(nameof(EditWelder), new { id = welderId, jcc });
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> OpenQualificationFile(string jcc, int welderId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(jcc) || welderId <= 0) return NotFound();
            var qual = await _context.Welder_List_tbl.AsNoTracking().FirstOrDefaultAsync(q => q.JCC_No == jcc && q.Welder_ID_WL == welderId);
            if (qual == null || string.IsNullOrWhiteSpace(qual.JCC_BlobName)) return NotFound();
            var root = Path.Combine(_env.ContentRootPath, "App_Data", "WelderQualifications");
            var fullPath = Path.Combine(root, qual.JCC_BlobName);
            if (!System.IO.File.Exists(fullPath)) return NotFound();
            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            var ct = ext switch
            {
                ".pdf" => "application/pdf",
                ".png" => "image/png",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".webp" => "image/webp",
                _ => "application/octet-stream"
            };
            var bytes = await System.IO.File.ReadAllBytesAsync(fullPath);
            return File(bytes, ct); // inline
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "OpenQualificationFile failed for {Jcc}/{Welder}", jcc, welderId);
            return NotFound();
        }
    }

    // NEW: Download qualification file (forced attachment)
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DownloadQualificationFile(string jcc, int welderId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(jcc) || welderId <= 0) return NotFound();
            var qual = await _context.Welder_List_tbl.AsNoTracking().FirstOrDefaultAsync(q => q.JCC_No == jcc && q.Welder_ID_WL == welderId);
            if (qual == null || string.IsNullOrWhiteSpace(qual.JCC_BlobName)) return NotFound();
            var root = Path.Combine(_env.ContentRootPath, "App_Data", "WelderQualifications");
            var fullPath = Path.Combine(root, qual.JCC_BlobName);
            if (!System.IO.File.Exists(fullPath)) return NotFound();
            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            var ct = ext switch
            {
                ".pdf" => "application/pdf",
                ".png" => "image/png",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".webp" => "image/webp",
                _ => "application/octet-stream"
            };
            var bytes = await System.IO.File.ReadAllBytesAsync(fullPath);
            var downloadName = string.IsNullOrWhiteSpace(qual.JCC_FileName) ? ($"{jcc}{ext}") : qual.JCC_FileName;
            return File(bytes, ct, downloadName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DownloadQualificationFile failed for {Jcc}/{Welder}", jcc, welderId);
            return NotFound();
        }
    }

    #endregion
}