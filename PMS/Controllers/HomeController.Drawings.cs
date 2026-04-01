using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using PMS.Infrastructure;
using PMS.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions; // added for parsing bulk upload names

namespace PMS.Controllers;

public partial class HomeController
{
    // Compile-time generated regexes (SYSLIB1045)
    [GeneratedRegex(@"(?:^|[\s_])\d+-(?<layout>[A-Za-z]-\d+)-(?<sheet>\d{2,3})(?:\D|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex SheetFirstPattern();

    [GeneratedRegex(@"(?:^|[\s_])\d+-(?<layout>[A-Za-z]-\d+)(?:\D|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex LineOnlyPattern();

    [GeneratedRegex(@"^(?<num>\d+)(?<suf>[A-Za-z]*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex RevSplitPattern();

    [GeneratedRegex(@"^\d+[A-Za-z]*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex Rev2StartsDigitsPattern();

    [GeneratedRegex(@"\s*\((?<dup>\d+)\)\s*$", RegexOptions.CultureInvariant)]
    private static partial Regex TrailingDupMarkerPattern();

    // FIX: correct character class for 'mid'
    [GeneratedRegex(@"(?:^|[\s_-])Rev(?<r1>[A-Za-z0-9]+)(?:[\s_-]+(?<mid>[A-Za-z0-9]+))?(?:[\s_-]+(?<r2>[A-Za-z0-9]+))?", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex RevHeaderPattern();

    [GeneratedRegex(@"(?:^|[\s_-])copy(?:\s*\(\d+\))?\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex CopySuffixRegex();

    [GeneratedRegex(@"(?:^|[\s_-])rm\d+\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex RmSuffixRegex();

    [GeneratedRegex(@"^(?<n>\d+)(?<s>[A-Za-z]*)$", RegexOptions.CultureInvariant)]
    private static partial Regex TagSplitPattern();

    [GeneratedRegex(@"(?:^|[\s_])\d+-(?<layout>[A-Za-z0-9]+-\d+)-(?<sheet>\d{1,3})(?:\D|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex FlexibleLayoutSheetPattern();

    private sealed record RevisionRow(int Id, string? Tag);

    private static string TrimDuplicateMarker(string name, out int dupSequence)
    {
        dupSequence = 0;
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        var match = TrailingDupMarkerPattern().Match(name);
        if (match.Success)
        {
            if (int.TryParse(match.Groups["dup"].Value, out var seq)) dupSequence = seq;
            return name[..match.Index].TrimEnd();
        }
        return name.Trim();
    }

    // NEW: normalize revision tag for equivalence and ordering
    private static string NormalizeRevisionTagForEquivalence(string? tag)
    {
        var safe = (tag ?? "00_00").Trim();
        var parts = safe.Split('_', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var p1Raw = parts.Length > 0 ? parts[0] : "00";
        var p2Raw = parts.Length > 1 ? parts[1] : "00";
        var tertiaryRaw = parts.Length > 2 ? parts[2] : null;
        var primary = NormalizeRev1(p1Raw);
        // When a tertiary segment exists (e.g., 00_0001_01F), use it as the
        // secondary – the middle segment is a legacy bucket code to skip.
        // This matches FormatRevisionTagForDisplay behaviour.
        var secondary = !string.IsNullOrWhiteSpace(tertiaryRaw)
            ? NormalizeRev2(tertiaryRaw!)
            : NormalizeRev2(p2Raw);
        return $"{primary}_{secondary}";
    }

    // NEW: display-friendly formatter: "00_1001" -> "(00_01)", "00_0001" -> "(00_00)", fallback to second part
    private static string? FormatRevisionTagForDisplay(string? tag)
    {
        if (string.IsNullOrWhiteSpace(tag)) return null;
        var parts = tag.Trim().Split('_');
        var primaryRaw = parts.Length > 0 ? parts[0] : string.Empty;
        var secondaryRaw = parts.Length > 1 ? parts[1] : string.Empty;
        var tertiaryRaw = parts.Length > 2 ? parts[2] : null;
        var dispPrimary = NormalizeRev1(primaryRaw);
        var dispSecondary = !string.IsNullOrWhiteSpace(tertiaryRaw)
            ? NormalizeRev2(tertiaryRaw!)
            : NormalizeRev2(secondaryRaw);
        return $"({dispPrimary}_{dispSecondary})";
    }

    private static string StripOuterParens(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var trimmed = value.Trim();
        if (trimmed.Length >= 2 && trimmed.StartsWith("(") && trimmed.EndsWith(")"))
        {
            trimmed = trimmed[1..^1].Trim();
        }
        return trimmed;
    }

    // NEW: global rank helper for revision tags (higher tags later in ordering)
    private static (int n1, string s1, int n2, string s2) RankFromTag(string? tag)
    {
        // Map special codes for consistent ordering
        tag = NormalizeRevisionTagForEquivalence(tag);
        var p = (tag ?? "00_00").Split('_');
        string a = p.Length > 0 ? p[0] : "00";
        string b = p.Length > 1 ? p[1] : "00";
        int n1 = 0, n2 = 0;
        string s1 = string.Empty, s2 = string.Empty;
        var m1 = TagSplitPattern().Match(a);
        if (m1.Success)
        {
            if (!int.TryParse(m1.Groups["n"].Value, out n1)) n1 = 0;
            s1 = (m1.Groups["s"].Value ?? string.Empty).ToUpperInvariant();
        }
        var m2 = TagSplitPattern().Match(b);
        if (m2.Success)
        {
            if (!int.TryParse(m2.Groups["n"].Value, out n2)) n2 = 0;
            s2 = (m2.Groups["s"].Value ?? string.Empty).ToUpperInvariant();
        }
        return (n1, s1, n2, s2);
    }

    // NEW: ordering rank that treats suffixed numerics (01F, 02F) as newer than plain numeric revisions while base (00) lowest
    private static (int n1, string s1, int cat, int n2, string s2) OrderingRank(string? tag)
    {
        // Use existing normalization for equivalence
        tag = NormalizeRevisionTagForEquivalence(tag);
        var p = (tag ?? "00_00").Split('_');
        string a = p.Length > 0 ? p[0] : "00";
        string b = p.Length > 1 ? p[1] : "00";
        int n1 = 0, n2 = 0; string s1 = string.Empty, s2 = string.Empty;
        var m1 = TagSplitPattern().Match(a);
        if (m1.Success)
        {
            if (!int.TryParse(m1.Groups["n"].Value, out n1))
            {
                n1 = 0;
            }
            s1 = (m1.Groups["s"].Value ?? string.Empty).ToUpperInvariant();
        }
        var m2 = TagSplitPattern().Match(b);
        if (m2.Success)
        {
            if (!int.TryParse(m2.Groups["n"].Value, out n2))
            {
                n2 = 0;
            }
            s2 = (m2.Groups["s"].Value ?? string.Empty).ToUpperInvariant();
        }
        // Category ordering: 0 => base 00, 1 => plain numeric (01,02,...), 2 => suffixed numeric (01F,02F)
        int cat;
        if (n2 == 0 && string.IsNullOrEmpty(s2)) cat = 0;
        else if (string.IsNullOrEmpty(s2)) cat = 1;
        else cat = 2;
        return (n1, s1, cat, n2, s2);
    }

    // NEW: recalc revision orders for a single key (Sheet or Line) based on tag ranking
    private async Task RecalculateRevisionOrdersAsync(int projectId, string mode, int? lineSheetId, int? lineId)
    {
        try
        {
            // First consolidate duplicates (0001 vs 1001 vs 00) keeping preferred file
            await CleanDuplicateRevisionRowsAsync(projectId, mode, lineSheetId, lineId);

            var q = _context.DWG_File_tbl.Where(d => d.Project_ID == projectId && d.Mode == mode);
            if (string.Equals(mode, "Sheet", StringComparison.OrdinalIgnoreCase) && lineSheetId.HasValue)
                q = q.Where(d => d.Line_Sheet_ID == lineSheetId);
            else if (string.Equals(mode, "Line", StringComparison.OrdinalIgnoreCase) && lineId.HasValue)
                q = q.Where(d => d.Line_ID == lineId);
            var rows = await q.ToListAsync();
            if (rows.Count == 0) return;

            var ordered = rows
                .Select(r => (row: r, rank: OrderingRank(r.RevisionTag)))
                .OrderBy(x => x.rank.n1)
                .ThenBy(x => x.rank.s1)
                .ThenBy(x => x.rank.cat)
                .ThenBy(x => x.rank.n2)
                .ThenBy(x => x.rank.s2)
                .Select(x => x.row)
                .ToList();
            int ord = 0;
            foreach (var r in ordered) r.RevisionOrder = ++ord; // ascending rank -> increasing order number (latest highest)
            await _context.SaveChangesAsync();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "RecalculateRevisionOrdersAsync failed for {Project}/{Mode}/{LS}/{Line}", projectId, mode, lineSheetId, lineId);
        }
    }

    private async Task UpdateLatestDesignRevisionAsync(int projectId, string mode, int? lineSheetId, int? lineId)
    {
        var latestTag = await _context.DWG_File_tbl.AsNoTracking()
            .Where(d => d.Project_ID == projectId && d.Mode == mode &&
                ((lineSheetId.HasValue && d.Line_Sheet_ID == lineSheetId) || (lineId.HasValue && d.Line_ID == lineId)))
            .OrderByDescending(d => d.RevisionOrder).ThenByDescending(d => d.Id)
            .Select(d => d.RevisionTag)
            .FirstOrDefaultAsync();

        latestTag = string.IsNullOrWhiteSpace(latestTag) ? null : NormalizeRevisionTagForEquivalence(latestTag.Trim());

        if (string.Equals(mode, "Sheet", StringComparison.OrdinalIgnoreCase) && lineSheetId.HasValue)
        {
            var lsRow = await _context.Line_Sheet_tbl.FirstOrDefaultAsync(x => x.Line_Sheet_ID == lineSheetId.Value);
            if (lsRow != null)
            {
                lsRow.LS_REV = latestTag;
                await _context.SaveChangesAsync();
            }
        }
        else if (string.Equals(mode, "Line", StringComparison.OrdinalIgnoreCase) && lineId.HasValue)
        {
            var lineRow = await _context.LINE_LIST_tbl.FirstOrDefaultAsync(x => x.Line_ID == lineId.Value);
            if (lineRow != null)
            {
                lineRow.L_REV = latestTag;
                await _context.SaveChangesAsync();
            }
        }
    }

    private static string SanitizePathSegment(string? s)
    {
        var value = (s ?? string.Empty).Trim();
        foreach (var c in Path.GetInvalidFileNameChars()) value = value.Replace(c, '_');
        return value;
    }

    // Centralized drawing path resolution to support legacy rows without BlobName
    private string GetDrawingRootFolder(DwgFile row)
    {
        return Path.Combine(
            Directory.GetCurrentDirectory(),
            "wwwroot",
            "drawings",
            row.Project_ID.ToString(),
            row.Mode,
            row.Mode == "Sheet" ? GetPairFolderName(row.Line_Sheet_ID) : GetLineFolderName(row.Line_ID)
        );
    }

    private static bool FileNameMatches(string candidatePath, IEnumerable<string> fileNames)
    {
        try
        {
            var name = Path.GetFileName(candidatePath);
            foreach (var fn in fileNames)
            {
                if (string.IsNullOrWhiteSpace(fn)) continue;
                if (string.Equals(name, fn, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }
        catch { return false; }
    }

    private string? TryResolveDrawingFilePath(DwgFile row)
    {
        // Build candidate roots to support legacy layouts
        //1) Current layout: /wwwroot/drawings/{project}/{mode}/{pair-or-line}
        var rootWithMode = GetDrawingRootFolder(row);
        //2) Legacy without mode segment: /wwwroot/drawings/{project}/{pair-or-line}
        var legacyPairOrLine = row.Mode == "Sheet" ? GetPairFolderName(row.Line_Sheet_ID) : GetLineFolderName(row.Line_ID);
        var rootWithoutMode = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "drawings", row.Project_ID.ToString(), legacyPairOrLine);
        //3) Just project folder: /wwwroot/drawings/{project}
        var projectRoot = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "drawings", row.Project_ID.ToString());

        var roots = new[] { rootWithMode, rootWithoutMode, projectRoot }
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Distinct()
            .ToArray();

        // Helper to probe a set of file names under a root
        static string? Probe(string root, IEnumerable<string> candidates)
        {
            foreach (var name in candidates)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                var p = Path.Combine(root, name);
                if (System.IO.File.Exists(p)) return p;
            }
            return null;
        }

        // Name candidates
        var fileCandidates = new List<string>();
        if (!string.IsNullOrWhiteSpace(row.BlobName)) fileCandidates.Add(row.BlobName);
        if (!string.IsNullOrWhiteSpace(row.FileName))
        {
            var original = Path.GetFileName(row.FileName);
            var safeName = SanitizePathSegment(original);
            fileCandidates.Add(original);
            fileCandidates.Add(safeName);
            if (!original.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) fileCandidates.Add(original + ".pdf");
            if (!safeName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) fileCandidates.Add(safeName + ".pdf");
        }

        //1) Direct probe at known roots
        foreach (var root in roots)
        {
            var found = Probe(root, fileCandidates);
            if (found != null) return found;
        }

        //2) If not found, deep-search under project root for a matching file name (covers unknown legacy layouts)
        try
        {
            if (Directory.Exists(projectRoot))
            {
                var all = Directory.EnumerateFiles(projectRoot, "*", SearchOption.AllDirectories);
                foreach (var path in all)
                {
                    if (FileNameMatches(path, fileCandidates)) return path;
                }
            }
        }
        catch { /* ignore IO errors */ }

        return null;
    }

    // Helper: resolve latest DwgFile for given project + layout (+ optional sheet)
    private async Task<DwgFile?> GetLatestDrawingRowAsync(int projectId, string layout, string? sheet)
    {
        layout = layout?.Trim() ?? string.Empty;
        sheet = sheet?.Trim();
        if (string.IsNullOrWhiteSpace(layout)) return null;

        var sheetModes = new[] { "Sheet", "SHEET", "sheet" };
        var lineModes  = new[] { "Line",  "LINE",  "line"  };

        // Prefer Sheet when a sheet is provided and exists; otherwise fallback to Line
        if (!string.IsNullOrWhiteSpace(sheet))
        {
            // Collect all LS candidates for the layout+sheet, scoped by project
            var lsRows = await _context.Line_Sheet_tbl.AsNoTracking()
                .Where(x => x.LS_LAYOUT_NO == layout && (x.Project_No == null || x.Project_No == projectId))
                .Select(x => new { x.Line_Sheet_ID, x.LS_SHEET })
                .ToListAsync();
            var matchIds = lsRows
                .Where(x => string.Equals((x.LS_SHEET ?? string.Empty).Trim(), sheet, StringComparison.OrdinalIgnoreCase))
                .Select(x => x.Line_Sheet_ID)
                .ToList();
            if (matchIds.Count == 0)
            {
                // numeric-equal fallback: treat '1' and '01' as same
                if (int.TryParse(sheet, out var sheetNum))
                {
                    matchIds = lsRows
                        .Where(x => int.TryParse((x.LS_SHEET ?? string.Empty).Trim(), out var n) && n == sheetNum)
                        .Select(x => x.Line_Sheet_ID)
                        .ToList();
                }
            }
            if (matchIds != null && matchIds.Count > 0)
            {
                var row = await _context.DWG_File_tbl.AsNoTracking()
                    .Where(d => d.Project_ID == projectId && d.Line_Sheet_ID.HasValue && matchIds.Contains(d.Line_Sheet_ID.Value) && sheetModes.Contains(d.Mode))
                    .OrderByDescending(d => d.RevisionOrder).ThenByDescending(d => d.Id)
                    .FirstOrDefaultAsync();
                if (row != null) return row;
            }
        }

        // Fallback to Line mode by layout
        var lineId = await _context.LINE_LIST_tbl.AsNoTracking()
            .Where(x => x.LAYOUT_NO == layout)
            .Select(x => x.Line_ID)
            .FirstOrDefaultAsync();
        if (lineId != 0)
        {
            var row = await _context.DWG_File_tbl.AsNoTracking()
                .Where(d => d.Project_ID == projectId && d.Line_ID == lineId && lineModes.Contains(d.Mode))
                .OrderByDescending(d => d.RevisionOrder).ThenByDescending(d => d.Id)
                .FirstOrDefaultAsync();
            if (row != null) return row;
        }
        return null;
    }

    // GET: open latest drawing inline for given project/layout/sheet
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> OpenLatestDrawing([FromQuery] int projectId, [FromQuery] string layout, [FromQuery] string? sheet)
    {
        try
        {
            var row = await GetLatestDrawingRowAsync(projectId, layout, sheet);
            if (row == null) return NotFound();
            var path = TryResolveDrawingFilePath(row);
            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path)) return NotFound();
            var name = string.IsNullOrWhiteSpace(row.FileName) ? $"{SanitizePathSegment(layout)}{(string.IsNullOrWhiteSpace(sheet) ? string.Empty : ("-" + SanitizePathSegment(sheet)))}.pdf" : row.FileName!;
            Response.Headers.ContentDisposition = $"inline; filename=\"{name}\""; // ASP0015 fix
            return PhysicalFile(path, "application/pdf");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "OpenLatestDrawing failed for {Project}/{Layout}/{Sheet}", projectId, layout, sheet);
            return StatusCode(500, "Error");
        }
    }

    // GET: download latest drawing as attachment for given project/layout/sheet
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DownloadLatestDrawing([FromQuery] int projectId, [FromQuery] string layout, [FromQuery] string? sheet)
    {
        try
        {
            var row = await GetLatestDrawingRowAsync(projectId, layout, sheet);
            if (row == null) return NotFound();
            var path = TryResolveDrawingFilePath(row);
            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path)) return NotFound();
            var name = string.IsNullOrWhiteSpace(row.FileName) ? $"{SanitizePathSegment(layout)}{(string.IsNullOrWhiteSpace(sheet) ? string.Empty : ("-" + SanitizePathSegment(sheet)))}.pdf" : row.FileName!;
            return PhysicalFile(path, "application/pdf", name);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DownloadLatestDrawing failed for {Project}/{Layout}/{Sheet}", projectId, layout, sheet);
            return StatusCode(500, "Error");
        }
    }

    // Request DTOs for bulk download
    public class LatestDrawingsZipRequest
    {
        public int ProjectId { get; set; }
        public List<LatestDrawingsItem> Items { get; set; } = new();
    }
    public class LatestDrawingsItem
    {
        public string? Layout { get; set; }
        public string? Sheet { get; set; }
    }

    // POST: download zip of latest drawings for multiple layout/sheet pairs
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DownloadLatestDrawingsZip([FromBody] LatestDrawingsZipRequest req)
    {
        try
        {
            if (req == null || req.ProjectId <= 0 || req.Items == null || req.Items.Count == 0)
                return BadRequest("Invalid request");

            // Distinct by layout|sheet (normalized)
            var pairs = req.Items
                .Select(i => new { Layout = (i.Layout ?? string.Empty).Trim(), Sheet = (i.Sheet ?? string.Empty).Trim() })
                .Where(i => !string.IsNullOrWhiteSpace(i.Layout))
                .GroupBy(i => (i.Layout ?? string.Empty).Trim().ToUpperInvariant() + "|" + (i.Sheet ?? string.Empty).Trim().ToUpperInvariant())
                .Select(g => g.First())
                .ToList();

            using var ms = new MemoryStream();
            int entriesAdded = 0;
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedRowIds = new HashSet<int>();
                foreach (var p in pairs)
                {
                    DwgFile? row = null;

                    var layoutKey = (p.Layout ?? string.Empty).Trim();
                    var sheetKey = (p.Sheet ?? string.Empty).Trim();

                    if (!string.IsNullOrWhiteSpace(sheetKey))
                    {
                        // Strict: only Sheet-mode for provided sheet, scoped by project
                        var lsRows = await _context.Line_Sheet_tbl.AsNoTracking()
                            .Where(x => x.LS_LAYOUT_NO == layoutKey && (x.Project_No == null || x.Project_No == req.ProjectId))
                            .Select(x => new { x.Line_Sheet_ID, x.LS_SHEET })
                            .ToListAsync();
                        var matchIds = lsRows
                            .Where(x => string.Equals((x.LS_SHEET ?? string.Empty).Trim(), sheetKey, StringComparison.OrdinalIgnoreCase))
                            .Select(x => x.Line_Sheet_ID)
                            .ToList();
                        if (matchIds.Count == 0 && int.TryParse(sheetKey, out var sheetNum))
                        {
                            matchIds = lsRows
                                .Where(x => int.TryParse((x.LS_SHEET ?? string.Empty).Trim(), out var n) && n == sheetNum)
                                .Select(x => x.Line_Sheet_ID)
                                .ToList();
                        }
                        if (matchIds.Count > 0)
                        {
                            row = await _context.DWG_File_tbl.AsNoTracking()
                                .Where(d => d.Project_ID == req.ProjectId && d.Line_Sheet_ID.HasValue && matchIds.Contains(d.Line_Sheet_ID.Value) && d.Mode == "Sheet")
                                .OrderByDescending(d => d.RevisionOrder).ThenByDescending(d => d.Id)
                                .FirstOrDefaultAsync();
                        }
                    }
                    else
                    {
                        // Strict: only Line-mode for layout-only
                        var lineId = await _context.LINE_LIST_tbl.AsNoTracking()
                            .Where(x => x.LAYOUT_NO == layoutKey)
                            .Select(x => x.Line_ID)
                            .FirstOrDefaultAsync();
                        if (lineId != 0)
                        {
                            row = await _context.DWG_File_tbl.AsNoTracking()
                                .Where(d => d.Project_ID == req.ProjectId && d.Line_ID == lineId && d.Mode == "Line")
                                .OrderByDescending(d => d.RevisionOrder).ThenByDescending(d => d.Id)
                                .FirstOrDefaultAsync();
                        }
                    }

                    if (row == null) continue;
                    if (usedRowIds.Contains(row.Id)) continue;

                    var path = TryResolveDrawingFilePath(row);
                    if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path)) continue;
                    if (usedPaths.Contains(path)) continue; // skip duplicate physical file

                    // Use original PDF name only (no Layout/Sheet prefix). Fall back to physical file name if metadata missing.
                    var baseName = string.IsNullOrWhiteSpace(row.FileName) ? Path.GetFileName(path) : Path.GetFileName(row.FileName);
                    if (string.IsNullOrWhiteSpace(baseName)) baseName = "drawing.pdf";
                    baseName = SanitizePathSegment(baseName);

                    // Ensure unique inside ZIP by appending (n)
                    var candidate = baseName;
                    int idx = 1;
                    while (usedNames.Contains(candidate))
                    {
                        var noExt = Path.GetFileNameWithoutExtension(baseName);
                        var ext = Path.GetExtension(baseName);
                        candidate = $"{noExt}({idx++}){ext}";
                    }

                    var entry = zip.CreateEntry(candidate, CompressionLevel.Optimal);
                    using var entryStream = entry.Open();
                    using var fs = System.IO.File.OpenRead(path);
                    await fs.CopyToAsync(entryStream);

                    usedNames.Add(candidate);
                    usedPaths.Add(path);
                    usedRowIds.Add(row.Id);
                    entriesAdded++;
                }
            }
            if (entriesAdded == 0)
            {
                return BadRequest("No drawings found for the selected rows.");
            }
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var fileName = $"Drawings_{req.ProjectId}_{stamp}.zip";
            ms.Position = 0;
            return File(ms.ToArray(), "application/zip", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DownloadLatestDrawingsZip failed");
            return StatusCode(500, "Error");
        }
    }

    // --- BULK UPLOAD SUPPORT ---
    // Parse helpers based on provided assumptions
    private static (string? Layout, string? Sheet) TryParseLayoutSheet(string name)
    {
        // Normalize to the part before any underscore (e.g., strip _Rev... suffixes)
        var core = name;
        var us = core.IndexOf('_');
        if (us > 0) core = core[..us];

        // Try flexible pattern: NNN-<layout-with letters/digits>-<digits sheet>
        var m = FlexibleLayoutSheetPattern().Match(core);
        if (m.Success)
        {
            var layout = m.Groups["layout"].Value.Trim().ToUpperInvariant();
            var sheet = m.Groups["sheet"].Value.Trim();
            if (int.TryParse(sheet, out var sn)) sheet = sn.ToString("000");
            return (layout, sheet);
        }

        // Fallback split approach on the core part only
        var parts = core.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length >= 4)
        {
            var sheetPart = parts[^1];
            if (sheetPart.Length >= 1 && sheetPart.Length <= 3 && sheetPart.All(char.IsDigit))
            {
                var head = parts[1];
                var tail = parts[2];
                if (!string.IsNullOrWhiteSpace(head) && !string.IsNullOrWhiteSpace(tail))
                {
                    var layout = (head + "-" + tail).ToUpperInvariant();
                    var sheet = sheetPart;
                    if (int.TryParse(sheet, out var sn)) sheet = sn.ToString("000");
                    return (layout, sheet);
                }
            }
        }

        // As a last resort, try original strict patterns against full name
        m = SheetFirstPattern().Match(name);
        if (m.Success)
        {
            var layout = m.Groups["layout"].Value.Trim().ToUpperInvariant();
            var sheet = m.Groups["sheet"].Value.Trim();
            if (int.TryParse(sheet, out var sn)) sheet = sn.ToString("000");
            return (layout, sheet);
        }
        m = LineOnlyPattern().Match(name);
        if (m.Success)
        {
            var layout = m.Groups["layout"].Value.Trim().ToUpperInvariant();
            return (layout, null);
        }
        return (null, null);
    }

    // Build a revision tag from parsed segments, keeping raw secondary values
    // (e.g., 0001/1001) for DB storage. Equivalence and display mapping are handled
    // by NormalizeRevisionTagForEquivalence / FormatRevisionTagForDisplay.
    private static string BuildRevisionTag(string? rev1, string? mid, string? rev2)
    {
        var p1 = NormalizeRev1(rev1 ?? string.Empty);
        string secondary = "00";
        if (!string.IsNullOrWhiteSpace(rev2) && Rev2StartsDigitsPattern().IsMatch(rev2!))
        {
            secondary = FormatRev2ForStorage(rev2!);
        }
        else if (!string.IsNullOrWhiteSpace(mid) && Rev2StartsDigitsPattern().IsMatch(mid!))
        {
            secondary = FormatRev2ForStorage(mid!);
        }
        return $"{p1}_{secondary}";
    }

    // Parse entire filename (without extension) into layout, sheet, revision tag
    private static (bool ok, string? layout, string? sheet, string? revTag) TryParseBulkUpload(string originalFileName)
    {
        if (string.IsNullOrWhiteSpace(originalFileName)) return (false, null, null, null);
        var baseNameRaw = Path.GetFileNameWithoutExtension(originalFileName);
        var baseName = TrimDuplicateMarker(baseNameRaw, out _);
        var (layout, sheet) = TryParseLayoutSheet(baseName);
        if (string.IsNullOrWhiteSpace(layout)) return (false, null, null, null);

        var m = RevHeaderPattern().Match(baseName);
        string? r1 = null, mid = null, r2 = null;
        if (m.Success)
        {
            r1 = m.Groups["r1"].Value;
            mid = m.Groups["mid"].Success ? m.Groups["mid"].Value : null;
            r2 = m.Groups["r2"].Success ? m.Groups["r2"].Value : null;
        }
        var tag = BuildRevisionTag(r1 ?? "0", mid, r2);
        return (true, layout, sheet, tag);
    }

    // POST: bulk upload many drawings, inferring keys from filenames
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    [DisableRequestSizeLimit]
    [RequestFormLimits(MultipartBodyLengthLimit = long.MaxValue, ValueCountLimit = int.MaxValue)]
    public async Task<IActionResult> BulkUploadDrawings([FromForm] int projectId)
    {
        try
        {
            var files = Request.Form?.Files;
            if (projectId <= 0 || files == null || files.Count == 0) return BadRequest("No files");

            int ok = 0, skipped = 0, failed = 0;
            var addedFiles = new List<string>();
            var skippedFiles = new List<string>();
            var failedFiles = new List<string>();
            var affectedSheets = new HashSet<int>();
            var affectedLines = new HashSet<int>();

            // Stage valid files first to sort within each key
            var stagedMap = new Dictionary<string, (Microsoft.AspNetCore.Http.IFormFile file, string mode, int? lsId, int? lineId, string folder, string revTag, int quality, string displayName)>(StringComparer.OrdinalIgnoreCase);

            foreach (var file in files)
            {
                try
                {
                    if (file == null || file.Length == 0) { failed++; continue; }
                    var displayName = Path.GetFileName(file.FileName ?? string.Empty);
                    // Heuristic PDF detection: allow missing .pdf extension if file header starts with %PDF
                    static bool IsPdf(Microsoft.AspNetCore.Http.IFormFile f)
                    {
                        try
                        {
                            if (f == null || f.Length < 4) return false;
                            using var s = f.OpenReadStream();
                            Span<byte> hdr = stackalloc byte[4];
                            int read = s.Read(hdr);
                            if (read == 4)
                            {
                                return hdr[0] == (byte)'%' && hdr[1] == (byte)'P' && hdr[2] == (byte)'D' && hdr[3] == (byte)'F';
                            }
                        }
                        catch { }
                        return false;
                    }

                    var hasPdfExt = file.FileName?.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) == true;
                    var isPdf = string.Equals(file.ContentType, "application/pdf", StringComparison.OrdinalIgnoreCase) || hasPdfExt || IsPdf(file);
                    if (!isPdf) { failed++; failedFiles.Add(displayName); continue; }

                    // Prevent null being passed to parser
                    var originalFileName = file.FileName ?? string.Empty;
                    var baseNoExt = Path.GetFileNameWithoutExtension(originalFileName) ?? string.Empty;
                    var trimmedBase = TrimDuplicateMarker(baseNoExt, out var dupSequence);
                    var sanitizedForParser = string.Concat(trimmedBase, Path.GetExtension(originalFileName));
                    if (RmSuffixRegex().IsMatch(trimmedBase)) { skipped++; skippedFiles.Add(displayName); continue; }

                    var parsed = TryParseBulkUpload(sanitizedForParser);
                    if (!parsed.ok || string.IsNullOrWhiteSpace(parsed.layout)) { failed++; failedFiles.Add(displayName); continue; }

                    string mode;
                    int? lsId = null; int? lineId = null;
                    string targetFolderName;

                    if (!string.IsNullOrWhiteSpace(parsed.sheet))
                    {
                        mode = "Sheet";
                        // Relaxed matching: find LS by layout then match sheet by string or numeric equivalence, scoped by project
                        var lsRows = await _context.Line_Sheet_tbl.AsNoTracking()
                            .Where(x => x.LS_LAYOUT_NO == parsed.layout && (x.Project_No == null || x.Project_No == projectId))
                            .Select(x => new { x.Line_Sheet_ID, x.LS_SHEET })
                            .ToListAsync();
                        var target = lsRows
                            .FirstOrDefault(x => string.Equals((x.LS_SHEET ?? string.Empty).Trim(), parsed.sheet, StringComparison.OrdinalIgnoreCase));
                        if (target == null && int.TryParse(parsed.sheet, out var sheetNum))
                        {
                            target = lsRows.FirstOrDefault(x => int.TryParse((x.LS_SHEET ?? string.Empty).Trim(), out var n) && n == sheetNum);
                        }
                        if (target == null) { failed++; failedFiles.Add(displayName); continue; }
                        lsId = target.Line_Sheet_ID;
                        targetFolderName = SanitizePathSegment(parsed.layout) + "-" + SanitizePathSegment(parsed.sheet);
                    }
                    else
                    {
                        mode = "Line";
                        var line = await _context.LINE_LIST_tbl.AsNoTracking()
                            .Where(x => x.LAYOUT_NO == parsed.layout)
                            .Select(x => new { x.Line_ID })
                            .FirstOrDefaultAsync();
                        if (line == null) { failed++; failedFiles.Add(displayName); continue; }
                        lineId = line.Line_ID;
                        targetFolderName = SanitizePathSegment(parsed.layout);
                    }

                    // Skip if same revision tag already exists in DB (using normalized equivalence)
                    var normTag = NormalizeRevisionTagForEquivalence(parsed.revTag);
                    var hasCopyMarker = CopySuffixRegex().IsMatch(trimmedBase);
                    var qualityScore = dupSequence + (hasCopyMarker ? 100 : 0);

                    List<RevisionRow> existingRowsForKey;
                    if (lsId.HasValue)
                    {
                        existingRowsForKey = await _context.DWG_File_tbl.AsNoTracking()
                            .Where(d => d.Project_ID == projectId && d.Mode == mode && d.Line_Sheet_ID == lsId && d.RevisionTag != null)
                            .Select(d => new RevisionRow(d.Id, d.RevisionTag))
                            .ToListAsync();
                    }
                    else
                    {
                        existingRowsForKey = await _context.DWG_File_tbl.AsNoTracking()
                            .Where(d => d.Project_ID == projectId && d.Mode == mode && d.Line_ID == lineId && d.RevisionTag != null)
                            .Select(d => new RevisionRow(d.Id, d.RevisionTag))
                            .ToListAsync();
                    }

                    var equivalentRows = existingRowsForKey
                        .Where(r => r.Tag != null && NormalizeRevisionTagForEquivalence(r.Tag) == normTag)
                        .Select(r => r.Id)
                        .ToList();

                    if (equivalentRows.Count > 0)
                    {
                        if (dupSequence > 0)
                        {
                            try
                            {
                                var rowsToDelete = _context.DWG_File_tbl.Where(d => equivalentRows.Contains(d.Id)).ToList();
                                foreach (var rem in rowsToDelete)
                                {
                                    try
                                    {
                                        var pathDel = TryResolveDrawingFilePath(rem);
                                        if (!string.IsNullOrWhiteSpace(pathDel) && System.IO.File.Exists(pathDel)) System.IO.File.Delete(pathDel);
                                    }
                                    catch { }
                                }
                                _context.DWG_File_tbl.RemoveRange(rowsToDelete);
                                await _context.SaveChangesAsync();
                            }
                            catch { }
                        }
                        else
                        {
                            skipped++; skippedFiles.Add(displayName); continue;
                        }
                    }

                    var stageKey = $"{mode}:{(lsId.HasValue ? $"LS:{lsId}" : $"LN:{lineId}")}:{normTag}";
                    var candidate = (file: file, mode, lsId, lineId, folder: targetFolderName, revTag: parsed.revTag ?? "00_00", quality: qualityScore, displayName);
                    if (stagedMap.TryGetValue(stageKey, out var existingCandidate))
                    {
                        if (qualityScore < existingCandidate.quality)
                        {
                            stagedMap[stageKey] = candidate;
                            skipped++;
                            if (!string.IsNullOrWhiteSpace(existingCandidate.displayName)) skippedFiles.Add(existingCandidate.displayName);
                        }
                        else
                        {
                            skipped++; skippedFiles.Add(displayName); continue;
                        }
                    }
                    else
                    {
                        stagedMap[stageKey] = candidate;
                    }
                }
                catch
                {
                    failed++;
                }
            }

            // Group staged by key and assign RevisionOrder respecting tag rank (include previously existing rows for consistent ordering)
            var staged = stagedMap.Values.ToList();
            var groups = staged.GroupBy(s => s.mode == "Sheet" ? ($"S:{s.lsId}") : ($"L:{s.lineId}"));
            foreach (var g in groups)
            {
                var first = g.First();
                var mode = first.mode;
                var lsId = first.lsId; var lineId = first.lineId;
                // Load existing rows for this key to merge ordering
                var existingRows = await _context.DWG_File_tbl
                    .Where(d => d.Project_ID == projectId && d.Mode == mode &&
                        ((lsId.HasValue && d.Line_Sheet_ID == lsId) || (lineId.HasValue && d.Line_ID == lineId)))
                    .AsNoTracking()
                    .Select(d => new { d.Id, d.RevisionTag })
                    .ToListAsync();

                // Sort new staged items by rank (ascending) so lower tag first
                var orderedNew = g
                    .Select(x => (item: x, rank: RankFromTag(x.revTag)))
                    .OrderBy(x => x.rank.n1)
                    .ThenBy(x => x.rank.s1)
                    .ThenBy(x => x.rank.n2)
                    .ThenBy(x => x.rank.s2)
                    .Select(x => x.item)
                    .ToList();

                foreach (var item in orderedNew)
                {
                    try
                    {
                        var root = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "drawings", projectId.ToString(), mode, item.folder);
                        Directory.CreateDirectory(root);
                        var blob = $"{Guid.NewGuid():N}.pdf";
                        var path = Path.Combine(root, blob);
                        using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
                        {
                            await item.file.CopyToAsync(fs);
                        }

                        var fileName = SanitizePathSegment(Path.GetFileName(item.file.FileName));
                        if (!fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) fileName += ".pdf";

                        var row = new DwgFile
                        {
                            Project_ID = projectId,
                            Mode = mode,
                            Line_Sheet_ID = lsId,
                            Line_ID = lineId,
                            FileName = fileName,
                            FileSize = (int)item.file.Length,
                            UploadDate = AppClock.Now,
                            UploadBy = HttpContext.Session.GetInt32("UserID"),
                            BlobName = blob,
                            ContentType = item.file.ContentType,
                            RevisionTag = item.revTag,
                            RevisionOrder = 0 // will be recalculated
                        };
                        _context.DWG_File_tbl.Add(row);
                        await _context.SaveChangesAsync();

                        if (lsId.HasValue) affectedSheets.Add(lsId.Value); else if (lineId.HasValue) affectedLines.Add(lineId.Value);
                        ok++;
                        if (!string.IsNullOrWhiteSpace(item.displayName)) addedFiles.Add(item.displayName);
                    }
                    catch
                    {
                        failed++;
                        if (!string.IsNullOrWhiteSpace(item.displayName)) failedFiles.Add(item.displayName);
                    }
                }

                // Recalculate ordering including existing + newly added rows
                await RecalculateRevisionOrdersAsync(projectId, mode, lsId, lineId);
            }

            // Sync latest design revisions for affected keys (after recalculation)
            foreach (var lsId in affectedSheets)
                await UpdateLatestDesignRevisionAsync(projectId, "Sheet", lsId, null);
            foreach (var lineId in affectedLines)
                await UpdateLatestDesignRevisionAsync(projectId, "Line", null, lineId);

            var summary = $"Bulk upload: {ok} added, {skipped} skipped (duplicate tag), {failed} failed.";

            // If AJAX request, return JSON summary instead of redirect
            if (Request.Headers.TryGetValue("X-Requested-With", out var xrq) && string.Equals(xrq.ToString(), "XMLHttpRequest", StringComparison.OrdinalIgnoreCase))
            {
                return Json(new { added = ok, skipped, failed, message = summary, addedFiles, skippedFiles, failedFiles });
            }

            TempData["Msg"] = summary;
            return RedirectToAction(nameof(Drawings), new { projectId });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "BulkUploadDrawings failed");
            return StatusCode(500, "Error");
        }
    }

    // GET: /Home/Drawings
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> Drawings([FromQuery] int? projectId, [FromQuery] string? layout, [FromQuery] string? sheet)
    {
        var fullName = HttpContext.Session.GetString("FullName");
        if (string.IsNullOrEmpty(fullName)) return RedirectToAction("Login");
        ViewBag.FullName = fullName;

        // Default to last project id
        var projects = await _context.Projects_tbl.AsNoTracking().OrderBy(p => p.Project_ID)
            .Select(p => new { p.Project_ID, p.Project_Name, p.Line_Sheet })
            .ToListAsync();
        var persistedProjectId = GetPersistedProjectId();
        if (projectId.HasValue && persistedProjectId.HasValue && projectId.Value != persistedProjectId.Value)
        {
            layout = null;
            sheet = null;
        }
        var resolvedProjectId = await GetDefaultProjectIdAsync(projectId);
        int pid = resolvedProjectId ?? (projects.Count == 0 ? 0 : projects.Max(p => p.Project_ID));

        var project = projects.FirstOrDefault(p => p.Project_ID == pid);
        string mode = (project?.Line_Sheet ?? "Sheet").Equals("Line", StringComparison.OrdinalIgnoreCase) ? "Line" : "Sheet";

        var vm = new DrawingsViewModel
        {
            SelectedProjectId = pid,
            Mode = mode,
            Layout = layout?.Trim(),
            Sheet = sheet?.Trim()
        };

        int? selectedLineSheetId = null;

        // Populate keys according to mode
        if (pid > 0)
        {
            // Get project-scoped layouts from DFR_tbl (authoritative project-layout mapping, same as DailyFitup)
            var projectDfrLayouts = await _context.DFR_tbl.AsNoTracking()
                .Where(d => d.Project_No == pid && d.LAYOUT_NUMBER != null && d.LAYOUT_NUMBER != "")
                .Select(d => d.LAYOUT_NUMBER!)
                .Distinct()
                .ToListAsync();

            if (mode == "Sheet")
            {
                IQueryable<LineSheet> sheetQuery;
                if (projectDfrLayouts.Count > 0)
                {
                    // Primary: filter Line_Sheet_tbl to DFR-confirmed layouts for this project
                    sheetQuery = _context.Line_Sheet_tbl.AsNoTracking()
                        .Where(ls => ls.LS_LAYOUT_NO != null && ls.LS_SHEET != null
                            && projectDfrLayouts.Contains(ls.LS_LAYOUT_NO));
                }
                else
                {
                    // Fallback: use Line_Sheet_tbl.Project_No when no DFR records exist yet
                    sheetQuery = _context.Line_Sheet_tbl.AsNoTracking()
                        .Where(ls => ls.Project_No == pid && ls.LS_LAYOUT_NO != null && ls.LS_SHEET != null);
                }

                var sheetKeys = await sheetQuery
                    .Select(ls => new
                    {
                        ls.Line_Sheet_ID,
                        Layout = ls.LS_LAYOUT_NO!.Trim(),
                        Sheet = ls.LS_SHEET!.Trim()
                    })
                    .Where(x => x.Layout != string.Empty && x.Sheet != string.Empty)
                    .GroupBy(x => new { x.Layout, x.Sheet })
                    .Select(g => new
                    {
                        g.Key.Layout,
                        g.Key.Sheet,
                        LineSheetId = g.Min(x => x.Line_Sheet_ID)
                    })
                    .OrderBy(x => x.Layout)
                    .ThenBy(x => x.Sheet)
                    .ToListAsync();

                vm.Keys = sheetKeys
                    .Select(k => new DrawingKeyOption
                    {
                        Key = $"{k.Layout}|{k.Sheet}",
                        Display = $"{k.Layout}-{k.Sheet}"
                    })
                    .ToList();

                if (sheetKeys.Count > 0)
                {
                    if (string.IsNullOrWhiteSpace(vm.Layout) ||
                        !sheetKeys.Any(k => string.Equals(k.Layout, vm.Layout, StringComparison.OrdinalIgnoreCase)))
                    {
                        vm.Layout = sheetKeys[0].Layout;
                        vm.Sheet = sheetKeys[0].Sheet;
                    }

                    var layoutGroup = sheetKeys
                        .Where(k => string.Equals(k.Layout, vm.Layout, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (layoutGroup.Count > 0 &&
                        (string.IsNullOrWhiteSpace(vm.Sheet) ||
                         !layoutGroup.Any(k => string.Equals(k.Sheet, vm.Sheet, StringComparison.OrdinalIgnoreCase))))
                    {
                        vm.Sheet = layoutGroup[0].Sheet;
                    }

                    var selectedKey = sheetKeys.FirstOrDefault(k =>
                        string.Equals(k.Layout, vm.Layout, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(k.Sheet, vm.Sheet, StringComparison.OrdinalIgnoreCase));

                    if (selectedKey != null)
                    {
                        selectedLineSheetId = selectedKey.LineSheetId;
                    }
                }
            }
            else
            {
                List<string> keys;
                if (projectDfrLayouts.Count > 0)
                {
                    // Primary: filter LINE_LIST_tbl to DFR-confirmed layouts for this project
                    keys = await _context.LINE_LIST_tbl.AsNoTracking()
                        .Where(x => x.LAYOUT_NO != null && projectDfrLayouts.Contains(x.LAYOUT_NO))
                        .OrderBy(x => x.LAYOUT_NO!)
                        .Select(x => x.LAYOUT_NO!)
                        .Distinct()
                        .ToListAsync();
                }
                else
                {
                    // Fallback: scope via Line_Sheet_tbl when no DFR records exist yet
                    var projectLayouts = await _context.Line_Sheet_tbl.AsNoTracking()
                        .Where(ls => ls.Project_No == pid && ls.LS_LAYOUT_NO != null)
                        .Select(ls => ls.LS_LAYOUT_NO!.Trim())
                        .Where(l => l != string.Empty)
                        .Distinct()
                        .ToListAsync();

                    keys = await _context.LINE_LIST_tbl.AsNoTracking()
                        .Where(x => x.LAYOUT_NO != null && projectLayouts.Contains(x.LAYOUT_NO))
                        .OrderBy(x => x.LAYOUT_NO!)
                        .Select(x => x.LAYOUT_NO!)
                        .Distinct()
                        .ToListAsync();
                }
                vm.Keys = keys.Select(k => new DrawingKeyOption
                {
                    Key = k,
                    Display = k
                }).ToList();
                if (vm.Layout == null && keys.Count > 0)
                {
                    vm.Layout = keys[0];
                }
            }
        }

        // Load revisions from DWG_File_tbl
        if (pid > 0 && !string.IsNullOrWhiteSpace(vm.Layout))
        {
            var q = _context.DWG_File_tbl.AsNoTracking().Where(d => d.Project_ID == pid);
            bool skipRevisionLoad = false;

            if (mode == "Sheet")
            {
                if (selectedLineSheetId.HasValue)
                {
                    q = q.Where(d => d.Mode == "Sheet" && d.Line_Sheet_ID == selectedLineSheetId.Value);
                    var latestRev = await _context.Line_Sheet_tbl.AsNoTracking()
                        .Where(x => x.Line_Sheet_ID == selectedLineSheetId.Value)
                        .Select(x => x.LS_REV)
                        .FirstOrDefaultAsync();
                    ViewBag.LatestDesignRev = FormatRevisionTagForDisplay(latestRev);
                }
                else
                {
                    skipRevisionLoad = true;
                    ViewBag.LatestDesignRev = null;
                }
            }
            else
            {
                var line = await _context.LINE_LIST_tbl.AsNoTracking()
                    .Where(x => x.LAYOUT_NO == vm.Layout)
                    .Select(x => new { x.Line_ID, x.L_REV })
                    .FirstOrDefaultAsync();
                if (line != null)
                {
                    q = q.Where(d => d.Mode == "Line" && d.Line_ID == line.Line_ID);
                    // Show display-friendly latest design revision
                    ViewBag.LatestDesignRev = FormatRevisionTagForDisplay(line.L_REV);
                }
                else
                {
                    skipRevisionLoad = true;
                    ViewBag.LatestDesignRev = null;
                }
            }

            if (!skipRevisionLoad)
            {
                // include uploader full name
                vm.Revisions = await q
                    .OrderBy(d => d.RevisionOrder)
                    .Select(d => new
                    {
                        d.Id,
                        d.RevisionOrder,
                        d.RevisionTag,
                        d.FileName,
                        d.FileSize,
                        d.UploadDate,
                        d.UploadBy,
                        Uploader = _context.PMS_Login_tbl
                            .Where(u => d.UploadBy.HasValue && u.UserID == d.UploadBy.Value)
                            .Select(u => ((u.FirstName ?? "").Trim() + " " + (u.LastName ?? "").Trim()).Trim())
                            .FirstOrDefault()
                    })
                    .Select(x => new DrawingRevisionVm
                    {
                        Id = x.Id,
                        RevisionOrder = x.RevisionOrder,
                        // Display mapping: e.g., 00_1001 -> (00-01), 00_0001 -> (00-00), 00_02F -> (00-02F)
                        RevisionTag = FormatRevisionTagForDisplay(x.RevisionTag),
                        RawRevisionTag = x.RevisionTag,
                        FileName = x.FileName ?? string.Empty,
                        FileSize = x.FileSize,
                        UploadDate = x.UploadDate.ToString("dd-MMM-yyyy hh:mm tt", CultureInfo.InvariantCulture),
                        UploadedBy = string.IsNullOrWhiteSpace(x.Uploader) ? null : x.Uploader
                    })
                    .ToListAsync();
            }
            else
            {
                vm.Revisions = new List<DrawingRevisionVm>();
            }
        }

        // Dropdown of projects for view
        ViewBag.Projects = projects.Select(p => new SelectListItem
        {
            Value = p.Project_ID.ToString(),
            Text = $"{p.Project_ID} - {p.Project_Name}",
            Selected = p.Project_ID == pid
        }).ToList();

        return View(vm);
    }

    [SessionAuthorization]
    [HttpGet]
    public IActionResult SuggestRevisionTagFromName([FromQuery] string? fileName)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return Json(new { ok = false, revisionTag = string.Empty });
            }

            var safeName = Path.GetFileName(fileName);
            var baseNoExt = Path.GetFileNameWithoutExtension(safeName) ?? string.Empty;
            var trimmedBase = TrimDuplicateMarker(baseNoExt, out _);
            if (RmSuffixRegex().IsMatch(trimmedBase))
            {
                return Json(new { ok = false, revisionTag = string.Empty });
            }

            var sanitized = string.Concat(trimmedBase, Path.GetExtension(safeName));
            var parsed = TryParseBulkUpload(sanitized);
            var tag = parsed.ok && !string.IsNullOrWhiteSpace(parsed.revTag)
                ? parsed.revTag!.Trim()
                : string.Empty;

            return Json(new { ok = !string.IsNullOrWhiteSpace(tag), revisionTag = tag });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "SuggestRevisionTagFromName failed for file {FileName}", fileName);
            return Json(new { ok = false, revisionTag = string.Empty });
        }
    }

    // POST: upload new revision
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UploadDrawing([FromForm] int projectId, [FromForm] string mode, [FromForm] string layout, [FromForm] string? sheet, [FromForm] string? revisionTag)
    {
        try
        {
            var files = Request.Form.Files;
            var file = files.Count > 0 ? files[0] : null; // CA1826 fix
            if (file == null || file.Length == 0) return BadRequest("No file");
            if (!string.Equals(file.ContentType, "application/pdf", StringComparison.OrdinalIgnoreCase) && !file.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                return BadRequest("PDF only");

            var originalFileName = file.FileName ?? string.Empty;
            var baseNoExt = Path.GetFileNameWithoutExtension(originalFileName) ?? string.Empty;
            var trimmedBase = TrimDuplicateMarker(baseNoExt, out var singleDupSequence);
            var sanitizedForParser = string.Concat(trimmedBase, Path.GetExtension(originalFileName));
            if (RmSuffixRegex().IsMatch(trimmedBase)) { TempData["Msg"] = "Upload skipped: RM drawings are ignored."; return RedirectToAction(nameof(Drawings), new { projectId, layout, sheet }); }

            int? lsId = null; int? lineId = null;
            var layoutKey = (layout ?? string.Empty).Trim();
            var sheetKey = (sheet ?? string.Empty).Trim();
            string safeLayout = SanitizePathSegment(layoutKey);
            string safeSheet = SanitizePathSegment(sheetKey);
            string root = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "drawings", projectId.ToString());
            Directory.CreateDirectory(root);

            if (string.Equals(mode, "Sheet", StringComparison.OrdinalIgnoreCase))
            {
                lsId = await _context.Line_Sheet_tbl.AsNoTracking()
                    .Where(x => x.LS_LAYOUT_NO == layoutKey && x.LS_SHEET == sheetKey && (x.Project_No == null || x.Project_No == projectId))
                    .Select(x => x.Line_Sheet_ID)
                    .FirstOrDefaultAsync();
                if (lsId == 0) return BadRequest("Invalid key");
                root = Path.Combine(root, "Sheet", safeLayout + "-" + safeSheet);
            }
            else
            {
                lineId = await _context.LINE_LIST_tbl.AsNoTracking()
                    .Where(x => x.LAYOUT_NO == layoutKey)
                    .Select(x => x.Line_ID)
                    .FirstOrDefaultAsync();
                if (lineId == 0) return BadRequest("Invalid key");
                root = Path.Combine(root, "Line", safeLayout);
            }
            Directory.CreateDirectory(root);

            var fileName = SanitizePathSegment(Path.GetFileName(file.FileName));
            var ext = Path.GetExtension(fileName);
            if (!string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase)) fileName += ".pdf";

            // Derive revision tag from filename when not provided (reuse bulk parse logic)
            string? finalRevisionTag = string.IsNullOrWhiteSpace(revisionTag) ? null : revisionTag!.Trim();
            if (string.IsNullOrWhiteSpace(finalRevisionTag))
            {
                var parsedSingle = TryParseBulkUpload(string.Concat(trimmedBase, Path.GetExtension(originalFileName)));
                if (parsedSingle.ok && !string.IsNullOrWhiteSpace(parsedSingle.revTag))
                {
                    finalRevisionTag = parsedSingle.revTag.Trim();
                }
            }

            // Duplicate / upgrade handling (match bulk logic semantics)
            int? keyLsId = lsId; int? keyLineId = lineId;
            var resolvedMode = string.Equals(mode, "Sheet", StringComparison.OrdinalIgnoreCase) ? "Sheet" : "Line";
            if (!string.IsNullOrWhiteSpace(finalRevisionTag))
            {
                List<RevisionRow> existingRowsForKey;
                if (keyLsId.HasValue)
                {
                    existingRowsForKey = await _context.DWG_File_tbl.AsNoTracking()
                        .Where(d => d.Project_ID == projectId && d.Mode == resolvedMode && d.Line_Sheet_ID == keyLsId && d.RevisionTag != null)
                        .Select(d => new RevisionRow(d.Id, d.RevisionTag))
                        .ToListAsync();
                }
                else if (keyLineId.HasValue)
                {
                    existingRowsForKey = await _context.DWG_File_tbl.AsNoTracking()
                        .Where(d => d.Project_ID == projectId && d.Mode == resolvedMode && d.Line_ID == keyLineId && d.RevisionTag != null)
                        .Select(d => new RevisionRow(d.Id, d.RevisionTag))
                        .ToListAsync();
                }
                else
                {
                    existingRowsForKey = new List<RevisionRow>();
                }

                var normTag = NormalizeRevisionTagForEquivalence(finalRevisionTag);
                var equivalentRows = existingRowsForKey
                    .Where(r => r.Tag != null && NormalizeRevisionTagForEquivalence(r.Tag) == normTag)
                    .Select(r => r.Id)
                    .ToList();

                if (equivalentRows.Count > 0)
                {
                    if (singleDupSequence > 0)
                    {
                        try
                        {
                            var rowsToDelete = _context.DWG_File_tbl.Where(d => equivalentRows.Contains(d.Id)).ToList();
                            foreach (var rem in rowsToDelete)
                            {
                                try
                                {
                                    var pathDel = TryResolveDrawingFilePath(rem);
                                    if (!string.IsNullOrWhiteSpace(pathDel) && System.IO.File.Exists(pathDel)) System.IO.File.Delete(pathDel);
                                }
                                catch { }
                            }
                            _context.DWG_File_tbl.RemoveRange(rowsToDelete);
                            await _context.SaveChangesAsync();
                        }
                        catch { }
                    }
                    else
                    {
                        TempData["Msg"] = "Upload skipped: identical revision already exists (append (n) to replace).";
                        return RedirectToAction(nameof(Drawings), new { projectId, layout, sheet });
                    }
                }
            }

            var blob = $"{Guid.NewGuid():N}.pdf";
            var path = Path.Combine(root, blob);
            using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                await file.CopyToAsync(fs);
            }

            var row = new DwgFile
            {
                Project_ID = projectId,
                Mode = resolvedMode,
                Line_Sheet_ID = lsId,
                Line_ID = lineId,
                FileName = fileName,
                FileSize = (int)file.Length,
                UploadDate = AppClock.Now,
                UploadBy = HttpContext.Session.GetInt32("UserID"),
                BlobName = blob,
                ContentType = file.ContentType,
                RevisionTag = finalRevisionTag,
                RevisionOrder = 0 // will be recalculated
            };
            _context.DWG_File_tbl.Add(row);
            await _context.SaveChangesAsync();

            // Recalculate ordering for this key so tag ranking determines order, irrespective of upload sequence
            await RecalculateRevisionOrdersAsync(projectId, row.Mode, lsId, lineId);

            // Sync design revision to latest tag
            await UpdateLatestDesignRevisionAsync(projectId, resolvedMode, lsId, lineId);

            return RedirectToAction(nameof(Drawings), new { projectId, layout, sheet });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "UploadDrawing failed");
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> EditDrawing(int id)
    {
        var row = await _context.DWG_File_tbl.AsNoTracking().FirstOrDefaultAsync(d => d.Id == id);
        if (row == null) return NotFound();
        string? layout = null; string? sheet = null;
        if (string.Equals(row.Mode, "Sheet", StringComparison.OrdinalIgnoreCase))
        {
            var ls = await _context.Line_Sheet_tbl.AsNoTracking()
                .Where(x => x.Line_Sheet_ID == row.Line_Sheet_ID)
                .Select(x => new { x.LS_LAYOUT_NO, x.LS_SHEET })
                .FirstOrDefaultAsync();
            layout = ls?.LS_LAYOUT_NO; sheet = ls?.LS_SHEET;
        }
        else
        {
            var line = await _context.LINE_LIST_tbl.AsNoTracking()
                .Where(x => x.Line_ID == row.Line_ID)
                .Select(x => new { x.LAYOUT_NO })
                .FirstOrDefaultAsync();
            layout = line?.LAYOUT_NO;
        }
        ViewBag.ProjectId = row.Project_ID;
        ViewBag.Mode = row.Mode;
        ViewBag.Layout = layout;
        ViewBag.Sheet = sheet;
        return View(row);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EditDrawing([FromForm] int id, [FromForm] string? fileName, [FromForm] string? revisionTag)
    {
        var row = await _context.DWG_File_tbl.FirstOrDefaultAsync(d => d.Id == id);
        if (row == null) return NotFound();
        // Update editable fields
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            var safeName = SanitizePathSegment(Path.GetFileName(fileName.Trim()));
            row.FileName = string.IsNullOrWhiteSpace(safeName) ? row.FileName : safeName;
        }
        var cleanedTag = StripOuterParens(revisionTag);
        row.RevisionTag = string.IsNullOrWhiteSpace(cleanedTag) ? null : cleanedTag;
        await _context.SaveChangesAsync();

        await RecalculateRevisionOrdersAsync(row.Project_ID, row.Mode, row.Line_Sheet_ID, row.Line_ID);
        await UpdateLatestDesignRevisionAsync(row.Project_ID, row.Mode, row.Line_Sheet_ID, row.Line_ID);

        // Build redirect params based on mode
        string? layout = null; string? sheet = null;
        if (string.Equals(row.Mode, "Sheet", StringComparison.OrdinalIgnoreCase))
        {
            var ls = await _context.Line_Sheet_tbl.AsNoTracking()
                .Where(x => x.Line_Sheet_ID == row.Line_Sheet_ID)
                .Select(x => new { x.LS_LAYOUT_NO, x.LS_SHEET })
                .FirstOrDefaultAsync();
            layout = ls?.LS_LAYOUT_NO; sheet = ls?.LS_SHEET;
        }
        else
        {
            var line = await _context.LINE_LIST_tbl.AsNoTracking()
                .Where(x => x.Line_ID == row.Line_ID)
                .Select(x => new { x.LAYOUT_NO })
                .FirstOrDefaultAsync();
            layout = line?.LAYOUT_NO;
        }
        return RedirectToAction(nameof(Drawings), new { projectId = row.Project_ID, layout, sheet });
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteDrawing([FromForm] int id, [FromForm] int projectId, [FromForm] string? layout, [FromForm] string? sheet)
    {
        try
        {
            var row = await _context.DWG_File_tbl.FirstOrDefaultAsync(d => d.Id == id);
            if (row == null) return NotFound();

            var mode = row.Mode;
            var lsId = row.Line_Sheet_ID;
            var lineId = row.Line_ID;

            // Resolve path (supports legacy locations) and delete if present
            var path = TryResolveDrawingFilePath(row);
            if (!string.IsNullOrWhiteSpace(path) && System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }

            _context.DWG_File_tbl.Remove(row);
            await _context.SaveChangesAsync();

            await RecalculateRevisionOrdersAsync(projectId, mode, lsId, lineId);
            await UpdateLatestDesignRevisionAsync(projectId, mode, lsId, lineId);

            return RedirectToAction(nameof(Drawings), new { projectId, layout, sheet });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "DeleteDrawing failed for {Id}", id);
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> OpenDrawing(int id)
    {
        var row = await _context.DWG_File_tbl.AsNoTracking().FirstOrDefaultAsync(d => d.Id == id);
        if (row == null) return NotFound();

        var path = TryResolveDrawingFilePath(row);
        if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path)) return NotFound();

        var name = string.IsNullOrWhiteSpace(row.FileName) ? $"Drawing_{row.Id}.pdf" : row.FileName!;
        // Display inline in the browser
        Response.Headers.ContentDisposition = $"inline; filename=\"{name}\""; // ASP0015 fix
        return PhysicalFile(path, "application/pdf");
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> DownloadDrawing(int id)
    {
        var row = await _context.DWG_File_tbl.AsNoTracking().FirstOrDefaultAsync(d => d.Id == id);
        if (row == null) return NotFound();

        var path = TryResolveDrawingFilePath(row);
        if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path)) return NotFound();

        var name = string.IsNullOrWhiteSpace(row.FileName) ? $"Drawing_{row.Id}.pdf" : row.FileName!;
        return PhysicalFile(path, "application/pdf", name);
    }

    // POST: reorder revisions
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ReorderDrawings([FromForm] int projectId, [FromForm] string mode, [FromForm] string layout, [FromForm] string? sheet, [FromForm(Name = "orderedIds")] string orderedIdsCsv)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(orderedIdsCsv)) return BadRequest("Empty order");
            var orderedIds = orderedIdsCsv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(s => int.TryParse(s, out var id) ? id : (int?)null)
                .Where(i => i.HasValue)
                .Select(i => i!.Value)
                .ToArray();
            if (orderedIds.Length == 0) return BadRequest("Empty order");

            int? lsId = null; int? lineId = null;
            if (string.Equals(mode, "Sheet", StringComparison.OrdinalIgnoreCase))
            {
                lsId = await _context.Line_Sheet_tbl.AsNoTracking()
                    .Where(x => x.LS_LAYOUT_NO == layout && x.LS_SHEET == sheet && (x.Project_No == null || x.Project_No == projectId))
                    .Select(x => x.Line_Sheet_ID)
                    .FirstOrDefaultAsync();
                if (lsId == 0) return BadRequest("Invalid key");
            }
            else
            {
                lineId = await _context.LINE_LIST_tbl.AsNoTracking()
                    .Where(x => x.LAYOUT_NO == layout)
                    .Select(x => x.Line_ID)
                    .FirstOrDefaultAsync();
                if (lineId == 0) return BadRequest("Invalid key");
            }

            var rows = await _context.DWG_File_tbl.Where(d => d.Project_ID == projectId && d.Mode == mode &&
                ((lsId.HasValue && d.Line_Sheet_ID == lsId) || (lineId.HasValue && d.Line_ID == lineId))).ToListAsync();

            // Assign highest RevisionOrder to the first item (descending order)
            var orderMap = new Dictionary<int, int>(orderedIds.Length);
            for (int i = 0; i < orderedIds.Length; i++) orderMap[orderedIds[i]] = orderedIds.Length - i;

            foreach (var r in rows)
            {
                if (orderMap.TryGetValue(r.Id, out var ord)) r.RevisionOrder = ord;
            }
            await _context.SaveChangesAsync();

            // After reordering, update design revision to latest tag (highest order)
            await UpdateLatestDesignRevisionAsync(projectId, mode, lsId, lineId);

            return RedirectToAction(nameof(Drawings), new { projectId, layout, sheet });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ReorderDrawings failed");
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> UpdateSheetRevision([FromForm] int projectId, [FromForm] string layout, [FromForm] string sheet, [FromForm] string? lsRev)
    {
        try
        {
            var row = await _context.Line_Sheet_tbl.FirstOrDefaultAsync(x => x.LS_LAYOUT_NO == layout && x.LS_SHEET == sheet && (x.Project_No == null || x.Project_No == projectId));
            if (row == null) return NotFound();
            // Normalize manual input as well for consistency
            row.LS_REV = string.IsNullOrWhiteSpace(lsRev) ? null : NormalizeRevisionTagForEquivalence(lsRev.Trim());
            await _context.SaveChangesAsync();
            return Json(new { ok = true, rev = row.LS_REV });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "UpdateSheetRevision failed for {Layout}-{Sheet}", layout, sheet);
            return StatusCode(500, "Error");
        }
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetSheetRevision([FromQuery] int projectId, [FromQuery] string layout, [FromQuery] string sheet)
    {
        try
        {
            var row = await _context.Line_Sheet_tbl.AsNoTracking().FirstOrDefaultAsync(x => x.LS_LAYOUT_NO == layout && x.LS_SHEET == sheet && (x.Project_No == null || x.Project_No == projectId));
            // Return display-normalized value for UI
            var disp = FormatRevisionTagForDisplay(row?.LS_REV);
            return Json(new { rev = disp });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetSheetRevision failed for {Layout}-{Sheet}", layout, sheet);
            return StatusCode(500, "Error");
        }
    }

    private static string NormalizeRev1(string rev1)
    {
        rev1 = (rev1 ?? string.Empty).Trim();
        if (rev1.Length == 0) return "00";
        var match = RevSplitPattern().Match(rev1);
        if (!match.Success) return rev1.ToUpperInvariant();
        var num = match.Groups["num"].Value;
        var suf = match.Groups["suf"].Value.ToUpperInvariant();
        if (num.Length == 0) return (suf.Length == 0 ? "00" : suf);
        if (num.Length == 1) num = "0" + num;
        return num + suf;
    }

    // Format secondary revision segment for DB storage without bucket mapping.
    // Keeps raw values (e.g., 0001 stays 0001, 1001 stays 1001) so the stored
    // tag preserves the original numbering and 1001 sorts after 0001.
    private static string FormatRev2ForStorage(string rev2)
    {
        rev2 = (rev2 ?? string.Empty).Trim();
        if (rev2.Length == 0) return "00";
        var match = RevSplitPattern().Match(rev2);
        if (!match.Success) return rev2.ToUpperInvariant();
        var num = match.Groups["num"].Value;
        var suf = match.Groups["suf"].Value.ToUpperInvariant();
        if (num.Length == 0) return (suf.Length == 0 ? "00" : suf);
        if (num.Length == 1) num = "0" + num;
        return num + suf;
    }

    // Normalize secondary revision segment for equivalence and ordering.
    // Bucket mapping: 4-digit codes ending in "001" map to bucket numbers
    // (0001 -> 00, 1001 -> 01, 2001 -> 02) so 1001 is latest than 0001.
    private static string NormalizeRev2(string rev2)
    {
        rev2 = (rev2 ?? string.Empty).Trim();
        if (rev2.Length == 0) return "00";

        if (int.TryParse(rev2, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numeric) &&
            rev2.Length == 4 &&
            rev2.EndsWith("001", StringComparison.OrdinalIgnoreCase))
        {
            var bucket = Math.Max(0, numeric / 1000);
            return bucket.ToString("00", CultureInfo.InvariantCulture);
        }

        var match = RevSplitPattern().Match(rev2);
        if (!match.Success) return rev2.ToUpperInvariant();
        var num = match.Groups["num"].Value;
        var suf = match.Groups["suf"].Value.ToUpperInvariant();
        if (num.Length == 0) return suf;
        if (num.Length == 1) num = "0" + num;
        return num + suf;
    }

    private async Task CleanDuplicateRevisionRowsAsync(int projectId, string mode, int? lineSheetId, int? lineId)
    {
        try
        {
            var q = _context.DWG_File_tbl.Where(d => d.Project_ID == projectId && d.Mode == mode);
            if (string.Equals(mode, "Sheet", StringComparison.OrdinalIgnoreCase) && lineSheetId.HasValue)
                q = q.Where(d => d.Line_Sheet_ID == lineSheetId);
            else if (string.Equals(mode, "Line", StringComparison.OrdinalIgnoreCase) && lineId.HasValue)
                q = q.Where(d => d.Line_ID == lineId);

            var rows = await q.AsNoTracking().Select(r => new { r.Id, r.RevisionTag }).ToListAsync();
            if (rows.Count <= 1) return;

            var groups = rows
                .Where(r => !string.IsNullOrWhiteSpace(r.RevisionTag))
                .GroupBy(r => NormalizeRevisionTagForEquivalence(r.RevisionTag))
                .ToList();

            var idsToKeep = new HashSet<int>();
            var idsToRemove = new List<int>();

            foreach (var g in groups)
            {
                var keep = g.OrderByDescending(r => r.Id).FirstOrDefault();
                if (keep == null) continue;
                idsToKeep.Add(keep.Id);
                foreach (var r in g)
                {
                    if (r.Id != keep.Id) idsToRemove.Add(r.Id);
                }
            }
            if (idsToRemove.Count == 0) return;
            var removeRows = _context.DWG_File_tbl.Where(d => idsToRemove.Contains(d.Id)).ToList();
            foreach (var r in removeRows)
            {
                try
                {
                    var physical = TryResolveDrawingFilePath(r);
                    if (!string.IsNullOrWhiteSpace(physical) && System.IO.File.Exists(physical)) System.IO.File.Delete(physical);
                }
                catch { }
            }
            _context.DWG_File_tbl.RemoveRange(removeRows);
            await _context.SaveChangesAsync();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "CleanDuplicateRevisionRowsAsync failed for {Project}/{Mode}/{LS}/{Line}", projectId, mode, lineSheetId, lineId);
        }
    }

    private string GetPairFolderName(int? lineSheetId)
    {
        if (!lineSheetId.HasValue) return "_unknown";
        var pair = _context.Line_Sheet_tbl.AsNoTracking().Where(x => x.Line_Sheet_ID == lineSheetId.Value)
            .Select(x => new { x.LS_LAYOUT_NO, x.LS_SHEET }).FirstOrDefault();
        return SanitizePathSegment(pair?.LS_LAYOUT_NO) + "-" + SanitizePathSegment(pair?.LS_SHEET);
    }

    private string GetLineFolderName(int? lineId)
    {
        if (!lineId.HasValue) return "_unknown";
        var layout = _context.LINE_LIST_tbl
            .AsNoTracking()
            .Where(x => x.Line_ID == lineId.Value)
            .Select(x => x.LAYOUT_NO)
            .FirstOrDefault();
        return SanitizePathSegment(layout);
    }
}
