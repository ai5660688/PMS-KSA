using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PMS.Data;
using PMS.Models;

namespace PMS.Services;

public class WpsSelectorService : IWpsSelectorService
{
    private readonly AppDbContext _context;
    private readonly ILogger<WpsSelectorService> _logger;

    public WpsSelectorService(AppDbContext context, ILogger<WpsSelectorService> logger)
    {
        _context = context;
        _logger = logger;
    }

    private static string NormalizeSchKey(string? s) => (s ?? string.Empty).Trim().ToUpperInvariant();

    // Simple per-request cache for schedule lookups
    private readonly ConcurrentDictionary<(string sch, double dia), double?> _scheduleCache = new();

    public async Task<(double? thickness, string? reason)> ResolveOletThicknessAsync(string? material, string? olSchedule, double? olDiameter)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(olSchedule) || !olDiameter.HasValue)
                return (null, "Missing OL schedule / diameter");

            var keySch = NormalizeSchKey(olSchedule);
            var row = await _context.Schedule_tbl.AsNoTracking()
                .Where(s => s.SCH != null && s.NPS.HasValue && s.SCH.ToUpper() == keySch && Math.Abs(s.NPS.Value - olDiameter.Value) < 0.0001)
                .Select(s => new { s.THICK, s.OLET_THICK_SS, s.OLET_THICK_CS, s.OLET_THICK_SSS })
                .FirstOrDefaultAsync();
            if (row == null) return (null, "No schedule match");

            double? value = null;
            var mat = (material ?? string.Empty).Trim().ToUpperInvariant();
            if (mat == "SS") value = row.OLET_THICK_SS;
            else if (mat == "SSS") value = row.OLET_THICK_SSS;
            else if (mat.Contains("CS", StringComparison.OrdinalIgnoreCase)) value = row.OLET_THICK_CS;

            // Fallback to general THICK
            value ??= row.THICK;
            return (value, value.HasValue ? null : "No thickness value in schedule row");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ResolveOletThicknessAsync failed");
            return (null, "Exception");
        }
    }

    public async Task<(double? thickness, string? reason)> ResolveThicknessAsync(
        int projectId,
        string? explicitLineClass,
        double? explicitThickness,
        string? sch,
        double? dia,
        string? olSch,
        double? olDia,
        double? olThick)
    {
        if (explicitThickness.HasValue)
            return (explicitThickness, null);
        if (olThick.HasValue)
            return (olThick, null);

        var pairs = new List<(string sch, double dia)>();
        if (!string.IsNullOrWhiteSpace(olSch) && olDia.HasValue) pairs.Add((NormalizeSchKey(olSch), olDia.Value));
        if (!string.IsNullOrWhiteSpace(sch) && dia.HasValue) pairs.Add((NormalizeSchKey(sch), dia.Value));
        if (pairs.Count == 0) return (null, "No schedule / diameter data");

        // batch query distinct schedules
        var schList = pairs.Select(p => p.sch).Distinct().ToList();
        var schedRows = await _context.Schedule_tbl.AsNoTracking()
            .Where(s => s.SCH != null && schList.Contains(s.SCH))
            .Select(s => new { s.SCH, s.NPS, s.THICK })
            .ToListAsync();

        foreach (var pr in pairs)
        {
            if (_scheduleCache.TryGetValue(pr, out var cached))
            {
                if (cached.HasValue) return (cached, null);
                continue;
            }
            var match = schedRows.FirstOrDefault(r => r.SCH != null && r.NPS.HasValue && NormalizeSchKey(r.SCH) == pr.sch && Math.Abs(r.NPS.Value - pr.dia) < 0.0001);
            _scheduleCache[pr] = match?.THICK;
            if (match?.THICK != null) return (match.THICK, null);
        }
        return (null, "No thickness derived from schedule table");
    }

    public async Task<IReadOnlyList<WpsOption>> GetCandidatesAsync(int projectId, string? lineClass, double? effectiveThickness)
    {
        if (!effectiveThickness.HasValue) return Array.Empty<WpsOption>();
        var tokens = Tokenize(lineClass ?? string.Empty);

        var wpsList = await _context.WPS_tbl.AsNoTracking()
            .Where(w => w.Project_WPS == projectId || w.Project_WPS == null)
            .Select(w => new { w.WPS_ID, w.WPS, w.WPS_Pipe_Class, w.PWHT, w.Thickness_Range, w.Thickness_Range_From, w.Thickness_Range_To })
            .ToListAsync();

        double thk = effectiveThickness.Value;
        var filtered = wpsList.Where(w => w.Thickness_Range_From.HasValue && w.Thickness_Range_To.HasValue && thk >= w.Thickness_Range_From.Value && thk <= w.Thickness_Range_To.Value)
            .Where(w => PipeClassMatches(tokens, w.WPS_Pipe_Class))
            .OrderBy(w => w.PWHT) // keep original ordering logic
            .ThenBy(w => w.Thickness_Range_To)
            .ThenBy(w => w.WPS)
            .Select(w => new WpsOption {
                Id = w.WPS_ID,
                Wps = w.WPS,
                ThicknessRange = !string.IsNullOrWhiteSpace(w.Thickness_Range)
                                    ? w.Thickness_Range
                                    : (w.Thickness_Range_From.HasValue && w.Thickness_Range_To.HasValue
                                        ? $"{w.Thickness_Range_From.Value:0.###}-{w.Thickness_Range_To.Value:0.###}"
                                        : null),
                Pwht = w.PWHT
            })
            .Take(200)
            .ToList();
        return filtered;
    }

    private static HashSet<string> Tokenize(string value) => value.Split(new[] { ',', ';', ' ', '\t', '\r', '\n', '/', '\\', '-', '_' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(p => p.Trim().ToLowerInvariant()).Where(p => p.Length > 0).ToHashSet();
    private static bool PipeClassMatches(HashSet<string> lineTokens, string? cls)
    {
        if (string.IsNullOrWhiteSpace(cls)) return true; // wildcard
        var tokens = Tokenize(cls);
        tokens.IntersectWith(lineTokens);
        return tokens.Count > 0;
    }
}
