using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using PMS.Infrastructure;
using PMS.Models;

namespace PMS.Controllers;

public partial class HomeController
{
    private const string ScopeJoints = "Joints";
    private const string ScopeLineList = "LineList";
    private const string ScopeLineSheet = "LineSheet";
    private const string ModeUpload = "upload";
    private const string ModeEdit = "edit";
    private const int MaxPdfPages = 50;
    private const int CommitBatchSize = 50;

    private static readonly string[] JointRecordTemplateColumns =
    {
        "UNIT_NUMBER", "FLUID", "LINE_NO", "LAYOUT_NO", "LINE_CLASS", "ISO",
        "L_REV", "LS_SHEET", "MATERIAL", "LS_DIAMETER", "LOCATION", "WELD_NUMBER",
        "J_ADD", "WELD_TYPE", "SPOOL_NUMBER", "DIAMETER", "SCHEDULE", "SP_DIA",
        "OL_DIAMETER", "OL_SCHEDULE", "MATERIAL_A", "MATERIAL_B", "GRADE_A",
        "GRADE_B", "DELETE_CANCEL", "LS_SCOPE"
    };

    private static readonly string[] JointRecordTemplateHeaders =
    {
        "Unit No", "Fluid", "Line No", "Layout No", "Line Class", "ISO",
        "Revision", "Sheet No.", "Material", "ISO Dia. In.", "Location", "Weld No",
        "J-Add", "Weld Type", "Spool", "Dia. In.", "Schedule", "Spool Dia. In.",
        "OL Dia", "OL Schedule", "Material A", "Material B", "Grade A",
        "Grade B", "Delete / Cancel", "Scope"
    };

    private static readonly string[] JointRecordExportHeaders =
    {
        "Unit No", "Fluid", "Line No", "Layout No", "Line Class", "ISO",
        "Revision", "Sheet No.", "Material", "ISO Dia. In.", "Location", "Weld No",
        "J-Add", "Weld Type", "Spool", "Dia. In.", "Schedule", "Spool Dia. In.",
        "OL Dia", "OL Schedule", "Material A", "Material B", "Grade A",
        "Grade B", "Delete / Cancel", "Scope", "System Remarks"
    };

    private static readonly JsonSerializerOptions JointRecordJsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null
    };

    #region Professional PCF Parsing Classes

    private class PcfComponent
    {
        public string Tag { get; set; } = "";
        public string Type { get; set; } = "";
        public string Description { get; set; } = "";
        public string Material { get; set; } = "";
        public string Grade { get; set; } = "";
        public string Size { get; set; } = "";
        public string Schedule { get; set; } = "";
        public string Spool { get; set; } = "";
    }

    private class PcfWeld
    {
        public string Number { get; set; } = "";
        public string Type { get; set; } = "";
        public string Location { get; set; } = "WS";
        public string ComponentATag { get; set; } = "";
        public string ComponentBTag { get; set; } = "";
        public PcfComponent? ComponentA { get; set; }
        public PcfComponent? ComponentB { get; set; }
        public string Spool { get; set; } = "";
        public string Remarks { get; set; } = "";
    }

    private class PcfHeader
    {
        public string UnitNo { get; set; } = "";
        public string Fluid { get; set; } = "";
        public string LineNo { get; set; } = "";
        public string DrawingNo { get; set; } = "";
        public string Iso { get; set; } = "";
        public string Revision { get; set; } = "";
        public string SheetNo { get; set; } = "1";
        public string Size { get; set; } = "";
        public string Schedule { get; set; } = "";
        public string Material { get; set; } = "";
        public string Specification { get; set; } = "";
        public string LineClass { get; set; } = "";
        public string Location { get; set; } = "";
    }

    #endregion

    #region PCF/SHA Parsing Regex Patterns - Enhanced

    // Enhanced PCF parsing regex patterns
    [GeneratedRegex(@"^\s*WELD\s*,\s*(?<weld>\d+|[A-Z]\d+)\s*,\s*(?<type>BW|SW|FW|BR|LET|TH|SP|FJ)\s*,\s*""?(?<compA>[^""]+?)""?\s*,\s*""?(?<compB>[^""]+?)""?\s*,?\s*(?<spool>\d+)?", RegexOptions.IgnoreCase)]
    private static partial Regex PcfWeldRegex();

    [GeneratedRegex(@"^\s*PIPE\s*,\s*""?(?<tag>[^""]+?)""?\s*,\s*(?<dia>\d+(?:\.\d+)?)\s*,\s*(?<sch>\d+|[A-Z]+)\s*,\s*""?(?<material>[^""]+?)""?(?:\s*,\s*""?(?<grade>[^""]+?)""?)?", RegexOptions.IgnoreCase)]
    private static partial Regex PcfPipeRegex();

    [GeneratedRegex(@"^\s*COMPONENT\s*,\s*""?(?<tag>[^""]+?)""?\s*,\s*(?<type>TEE|ELBOW|REDUCER|FLANGE|VALVE|OLET|NIPPLE|CAP|PIPE|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD BOLT|BOLT|NUT)\s*,\s*(?<dia>\d+(?:\.\d+)?)?\s*,\s*(?<sch>\d+|[A-Z]+)?\s*,?\s*""?(?<material>[^""]*?)""?(?:\s*,\s*""?(?<grade>[^""]*?)""?)?", RegexOptions.IgnoreCase)]
    private static partial Regex PcfComponentRegex();

    [GeneratedRegex(@"^\s*(?:HEADER|END-HEADER|PIPING|END-PIPING|ISOMETRIC|END-ISOMETRIC|MATERIAL|END-MATERIAL)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfSectionRegex();

    [GeneratedRegex(@"^\s*(?<key>[A-Z_]+)\s*[:]\s*(?<value>.+)$", RegexOptions.IgnoreCase)]
    private static partial Regex PcfKeyValueRegex();

    [GeneratedRegex(@"UNIT\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfUnitRegex();

    [GeneratedRegex(@"LINE\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfLineRegex();

    [GeneratedRegex(@"SIZE\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfSizeRegex();

    [GeneratedRegex(@"SPECIFICATION\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfSpecRegex();

    [GeneratedRegex(@"DRAWING\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfDrawingRegex();

    [GeneratedRegex(@"REVISION\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfRevisionRegex();

    [GeneratedRegex(@"SHEET\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfSheetRegex();

    [GeneratedRegex(@"FLUID\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfFluidRegex();

    [GeneratedRegex(@"LINE\s+CLASS\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfLineClassRegex();

    [GeneratedRegex(@"ISO\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfIsoRegex();

    [GeneratedRegex(@"LOCATION\s*[:]\s*(?<value>.+)", RegexOptions.IgnoreCase)]
    private static partial Regex PcfLocationRegex();

    // SHA parsing regex patterns
    [GeneratedRegex(@"^S(?<weld>\d+)\s+(?<type>BW|SW|FW|BR|LET|TH|SP|FJ)\s+(?<spool>\d+)\s+(?<dia>\d+)\s+(?<sch>\d+|[A-Z]+)\s+(?<compA>.+?)\s+(?<compB>.+?)(?:\s+(?<gradeA>.+?)\s+(?<gradeB>.+?))?", RegexOptions.IgnoreCase)]
    private static partial Regex ShaWeldRegex();

    [GeneratedRegex(@"^PIPE\s+(?<tag>.+?)\s+(?<dia>\d+)\s+(?<sch>\d+|[A-Z]+)\s+(?<material>.+?)\s+(?<grade>.+?)$", RegexOptions.IgnoreCase)]
    private static partial Regex ShaPipeRegex();

    [GeneratedRegex(@"^COMP\s+(?<tag>.+?)\s+(?<type>TEE|ELBOW|REDUCER|FLANGE|VALVE|OLET|NIPPLE|CAP|PIPE|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD BOLT|BOLT|NUT)\s+(?<dia>\d+)?\s+(?<sch>\d+|[A-Z]+)?\s+(?<material>.+?)?\s+(?<grade>.+?)?$", RegexOptions.IgnoreCase)]
    private static partial Regex ShaComponentRegex();

    #endregion

    #region Existing Regex Patterns (Keep as is)

    [GeneratedRegex(@"(?<dia>\d{1,3})?[""\u201D\u2033]?\s*[-\u2013\u2014_]?\s*(?<fluid>[A-Z0-9]{1,4})[-\u2013\u2014_]?(?<line>\d{3,5})[-\u2013\u2014_]?(?<cls>[A-Z0-9]{3,12})", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex JointRecordIsoRegex();

    [GeneratedRegex(@"(?<fluid>[A-Z0-9]{1,4})[-_](?<line>\d{3,5})", RegexOptions.IgnoreCase)]
    private static partial Regex FluidLinePatternRegex();

    [GeneratedRegex(@"^[A-Z0-9]+$", RegexOptions.IgnoreCase)]
    private static partial Regex ValidWeldNoRegex();

    [GeneratedRegex(@"^(BW|FW|SW|SOF|BR|LET|TH|SP|FJ)$", RegexOptions.IgnoreCase)]
    private static partial Regex ValidWeldTypeRegex();

    [GeneratedRegex(@"^(NEW|[A-Z][A-Z0-9]*)$", RegexOptions.IgnoreCase)]
    private static partial Regex ValidJAddRegex();

    [GeneratedRegex(@"^([A-Z0-9]{1,4}-\d{3,5})$", RegexOptions.IgnoreCase)]
    private static partial Regex ValidLayoutNoRegex();

    [GeneratedRegex(@"^([A-Z0-9]{1,4}-\d{3,5}-[A-Z0-9]{3,12})$", RegexOptions.IgnoreCase)]
    private static partial Regex ValidIsoRegex();

    // Merge OCR-split weld number + J-Add tokens: "01 CS1 WS" or "01 CS1 BW" → "01CS1 WS" / "01CS1 BW"
    // Pattern: (weld number) SPACE (alpha-numeric J-Add code) SPACE (known location or weld type)
    [GeneratedRegex(@"(?<![A-Z])(?<num>[A-Z]?\d{1,4})\s+(?<jadd>[A-Z][A-Z0-9]{1,5})\s+(?=(?:WS|FW|YFW|YWS|SW|BW|SOF|BR|LET|TH|SP|FJ)\b)", RegexOptions.IgnoreCase)]
    private static partial Regex OcrSplitWeldJAddRegex();

    // Separate weld number from glued location suffix: "01CS1WS" → "01CS1 WS", "10FW" → "10 FW"
    // Fires when the merged token is followed by whitespace + digit (the weld-size column).
    [GeneratedRegex(@"(?<!\w)(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*?)(?<loc>YFW|YWS|FW|WS|SW)(?=\s+\d)", RegexOptions.IgnoreCase)]
    private static partial Regex OcrGluedWeldLocationRegex();

    // Split OCR-concatenated compound weld numbers: "S13T12" → "S13 T12", "T10S05" → "T10 S05"
    // Matches two adjacent [Letter][Digits] groups glued together without a separator.
    // Only fires when followed by whitespace + known location or weld type to avoid false splits.
    [GeneratedRegex(@"(?<!\w)(?<w1>[A-Z]\d{1,4})(?<w2>[A-Z]\d{1,4})(?=\s+(?:WS|FW|YFW|YWS|SW|BW|SOF|BR|LET|TH|SP|FJ)\b)", RegexOptions.IgnoreCase)]
    private static partial Regex OcrCompoundWeldSplitRegex();

    [GeneratedRegex(@"\bSP(?:OOL)?[-_\s]*0*(\d{1,4})\b", RegexOptions.IgnoreCase)]
    private static partial Regex SpoolNumberRegex();

    [GeneratedRegex(@"(?<line>\d{3,5})", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex DrawingLineNumberRegex();

    [GeneratedRegex(@"\d+", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex DigitsRegex();

    [GeneratedRegex(@"SP(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex SpoolTagRegex();

    [GeneratedRegex(@"[^A-Z0-9]", RegexOptions.IgnoreCase)]
    private static partial Regex NonAlphaNumericRegex();

    // Matches a trailing letter-prefix weld number suffix: S01, T16, S04, etc.
    // Used by NormalizeWeldNumber to strip OCR-concatenated noise prefixes.
    [GeneratedRegex(@"[A-Z]\d{2,3}$", RegexOptions.CultureInvariant)]
    private static partial Regex TrailingWeldSuffixRegex();

    [GeneratedRegex(@"^(?:[A-Z]\d{1,4}){2,}$", RegexOptions.IgnoreCase)]
    private static partial Regex CompoundWeldNumberRegex();

    [GeneratedRegex(@"[A-Z]\d{1,4}", RegexOptions.IgnoreCase)]
    private static partial Regex CompoundWeldPartRegex();

    [GeneratedRegex(@"(?<!\w)(?<weld>[A-Z]?\d{1,2}[A-Z0-9]*)\s+(?<type>SOF|SW|BW|FW|TH|BR|LET|SP|FJ)\b", RegexOptions.IgnoreCase)]
    private static partial Regex WeldTypeTableRowRegex();

    [GeneratedRegex(@"^(\d+)\.(\d+)/(\d+)$")]
    private static partial Regex FractionDotRegex();

    [GeneratedRegex(@"^(\d+)\s+(\d+)/(\d+)$")]
    private static partial Regex FractionSpaceRegex();

    [GeneratedRegex(@"^(\d+)/(\d+)$")]
    private static partial Regex FractionSimpleRegex();

    [GeneratedRegex(@"\s*GR(?:ADE)?\.?\s*", RegexOptions.IgnoreCase)]
    private static partial Regex AstmGradeCleanRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();

    [GeneratedRegex(@"-{2,}")]
    private static partial Regex MultiHyphenRegex();

    [GeneratedRegex(@"\S+")]
    private static partial Regex NonWhitespaceTokenRegex();

    private static readonly Dictionary<string, string> JointRecordHeaderAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["UNITNO"] = "UNIT_NUMBER",
        ["UNITNOPLANTNO"] = "UNIT_NUMBER",
        ["PLANTNO"] = "UNIT_NUMBER",
        ["PLANTNUMBER"] = "UNIT_NUMBER",

        ["FLUID"] = "FLUID",

        ["LINENO"] = "LINE_NO",
        ["LINENUMBER"] = "LINE_NO",

        ["LAYOUTNO"] = "LAYOUT_NO",
        ["LAYOUTNODRAWINGNUMBER"] = "LAYOUT_NO",
        ["DRAWINGNO"] = "LAYOUT_NO",
        ["DRAWINGNUMBER"] = "LAYOUT_NO",

        ["LINECLASS"] = "LINE_CLASS",
        ["LINECLASSEG1CS2U01"] = "LINE_CLASS",

        ["ISO"] = "ISO",
        ["ISOEG60C101001CS2U01"] = "ISO",

        ["REVISION"] = "L_REV",
        ["REVISIONLINESHEET"] = "L_REV",
        ["LINEREVISION"] = "L_REV",
        ["SHEETREVISION"] = "L_REV",

        ["SHEETNO"] = "LS_SHEET",
        ["SHEET"] = "LS_SHEET",

        ["MATERIAL"] = "MATERIAL",
        ["MATERIALMATFROMMATERIALTBL"] = "MATERIAL",

        ["ISODIAIN"] = "LS_DIAMETER",
        ["ISODIA"] = "LS_DIAMETER",

        ["LOCATION"] = "LOCATION",
        ["LOCATIONPLOCATION"] = "LOCATION",

        ["WELDNO"] = "WELD_NUMBER",
        ["WELDNUMBER"] = "WELD_NUMBER",

        ["JADD"] = "J_ADD",
        ["JADDNEWORDADDPJADD"] = "J_ADD",

        ["WELDTYPE"] = "WELD_TYPE",

        ["SPOOL"] = "SPOOL_NUMBER",
        ["SPOOLSP0101"] = "SPOOL_NUMBER",
        ["SPOOLNO"] = "SPOOL_NUMBER",

        ["DIAIN"] = "DIAMETER",
        ["DIAINWELDSIZE"] = "DIAMETER",
        ["WELDSIZE"] = "DIAMETER",

        ["SCHEDULE"] = "SCHEDULE",
        ["SCHEDULESCHORMM"] = "SCHEDULE",
        ["SCHEDULESCH"] = "SCHEDULE",

        ["SPOOLDIAIN"] = "SP_DIA",

        ["OLDIA"] = "OL_DIAMETER",
        ["OLDIAFORLETBRWELDS"] = "OL_DIAMETER",

        ["OLSCHEDULE"] = "OL_SCHEDULE",
        ["OLSCHEDULEFORLETBRWELDS"] = "OL_SCHEDULE",

        ["MATERIALA"] = "MATERIAL_A",
        ["MATERIALAMATDDESCRIPTION"] = "MATERIAL_A",

        ["MATERIALB"] = "MATERIAL_B",
        ["MATERIALBMATDDESCRIPTION"] = "MATERIAL_B",

        ["GRADEA"] = "GRADE_A",
        ["GRADEAMATGRADE"] = "GRADE_A",

        ["GRADEB"] = "GRADE_B",
        ["GRADEBMATGRADE"] = "GRADE_B",

        ["SCOPE"] = "LS_SCOPE",
        ["SCOPECOMPANY"] = "LS_SCOPE",
    };

    #endregion

    #region Professional PCF Parser

    private DrawingMetadata? ParsePcfFile(Stream stream, string fileName, List<string> warnings)
    {
        try
        {
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();

            var lines = content.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            // Parse PCF file sections
            var header = ParsePcfHeader(lines, warnings);
            var components = ParsePcfComponents(lines, warnings);
            var welds = ParsePcfWelds(lines, components, warnings);

            // Map to DrawingMetadata
            var metadata = MapPcfToDrawingMetadata(header, welds);

            return metadata;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error parsing PCF file: {FileName}", fileName);
            warnings.Add($"Error parsing PCF file: {ex.Message}");
            return null;
        }
    }

    private static string? InferComponentFromGrade(string? grade)
    {
        if (string.IsNullOrWhiteSpace(grade)) return null;
        var g = grade.ToUpperInvariant();

        if (g.Contains("A106") || g.Contains("API 5L") || g.Contains("A333"))
            return "PIPE";
        if (g.Contains("A105"))
            return "FLANGE";
        if (g.Contains("A193"))
            return "STUD BOLT";
        if (g.Contains("A234") || g.Contains("A420") || g.Contains("A234-WPB") || g.Contains("WPL"))
            return "ELBOW"; // generic fitting side

        return null;
    }

    private static void FillMaterialsFromGrades(List<WeldRecord> welds, Dictionary<string, string> bomGradeMap)
    {
        foreach (var weld in welds)
        {
            // Skip fully populated table rows
            if (HasCompleteMaterialAndGrade(weld))
                continue;

            if (string.IsNullOrEmpty(weld.MaterialA) && !string.IsNullOrEmpty(weld.GradeA))
            {
                var fromBom = bomGradeMap.FirstOrDefault(kv => string.Equals(kv.Value, weld.GradeA, StringComparison.OrdinalIgnoreCase));
                weld.MaterialA = !string.IsNullOrEmpty(fromBom.Key)
                    ? fromBom.Key
                    : InferComponentFromGrade(weld.GradeA) ?? weld.MaterialA;
            }

            if (string.IsNullOrEmpty(weld.MaterialB) && !string.IsNullOrEmpty(weld.GradeB))
            {
                var fromBom = bomGradeMap.FirstOrDefault(kv => string.Equals(kv.Value, weld.GradeB, StringComparison.OrdinalIgnoreCase));
                weld.MaterialB = !string.IsNullOrEmpty(fromBom.Key)
                    ? fromBom.Key
                    : InferComponentFromGrade(weld.GradeB) ?? weld.MaterialB;
            }

            // If both materials are now set and one grade is missing, mirror when components match
            if (!string.IsNullOrEmpty(weld.MaterialA)
                && !string.IsNullOrEmpty(weld.MaterialB)
                && string.Equals(weld.MaterialA, weld.MaterialB, StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(weld.GradeA) && !string.IsNullOrEmpty(weld.GradeB))
                    weld.GradeA = weld.GradeB;
                else if (string.IsNullOrEmpty(weld.GradeB) && !string.IsNullOrEmpty(weld.GradeA))
                    weld.GradeB = weld.GradeA;
            }
        }
    }

    private static PcfHeader ParsePcfHeader(string[] lines, List<string> _)
    {
        var header = new PcfHeader();
        bool inHeaderSection = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();

            // Check for HEADER section
            if (trimmed.Equals("HEADER", StringComparison.OrdinalIgnoreCase))
            {
                inHeaderSection = true;
                continue;
            }

            if (trimmed.Equals("END-HEADER", StringComparison.OrdinalIgnoreCase))
            {
                inHeaderSection = false;
                break;
            }

            if (inHeaderSection)
            {
                // Parse key-value pairs in header
                var kvMatch = PcfKeyValueRegex().Match(trimmed);
                if (kvMatch.Success)
                {
                    var key = kvMatch.Groups["key"].Value.Trim().ToUpperInvariant();
                    var value = kvMatch.Groups["value"].Value.Trim();

                    switch (key)
                    {
                        case "UNIT":
                            header.UnitNo = value;
                            break;
                        case "FLUID":
                            header.Fluid = value;
                            break;
                        case "LINE":
                            header.LineNo = value;
                            break;
                        case "DRAWING":
                        case "DWG_NO":
                        case "DRAWING_NO":
                            header.DrawingNo = value;
                            break;
                        case "ISO":
                        case "ISOMETRIC":
                            header.Iso = value;
                            break;
                        case "REVISION":
                        case "REV":
                            header.Revision = value;
                            break;
                        case "SHEET":
                        case "SHEET_NO":
                            header.SheetNo = value;
                            break;
                        case "SIZE":
                            header.Size = value;
                            break;
                        case "SCHEDULE":
                        case "SCH":
                            header.Schedule = value;
                            break;
                        case "MATERIAL":
                        case "MAT":
                            header.Material = value;
                            break;
                        case "SPECIFICATION":
                        case "SPEC":
                            header.Specification = value;
                            break;
                        case "LINE_CLASS":
                        case "CLASS":
                            header.LineClass = value;
                            break;
                        case "LOCATION":
                            header.Location = value;
                            break;
                    }
                }
            }
        }

        // Extract fluid and line from drawing number if not set
        if (string.IsNullOrEmpty(header.Fluid) && !string.IsNullOrEmpty(header.DrawingNo))
        {
            ExtractFluidLineFromString(header.DrawingNo, out var fluidFromDrawing, out var lineFromDrawing);
            header.Fluid = fluidFromDrawing;
            header.LineNo = lineFromDrawing;
        }

        // Extract fluid and line from ISO if not set
        if (string.IsNullOrEmpty(header.Fluid) && !string.IsNullOrEmpty(header.Iso))
        {
            ExtractFluidLineFromString(header.Iso, out var fluidFromIso, out var lineFromIso);
            header.Fluid = fluidFromIso;
            header.LineNo = lineFromIso;
        }

        // Extract line class from ISO if not set
        if (string.IsNullOrEmpty(header.LineClass) && !string.IsNullOrEmpty(header.Iso))
        {
            var parts = header.Iso.Split('-');
            if (parts.Length >= 3)
            {
                header.LineClass = parts[2];
            }
        }

        return header;
    }

    private static Dictionary<string, PcfComponent> ParsePcfComponents(string[] lines, List<string> _)
    {
        var components = new Dictionary<string, PcfComponent>();
        bool inPipingSection = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();

            // Check for PIPING section
            if (trimmed.Equals("PIPING", StringComparison.OrdinalIgnoreCase))
            {
                inPipingSection = true;
                continue;
            }

            if (trimmed.Equals("END-PIPING", StringComparison.OrdinalIgnoreCase))
            {
                inPipingSection = false;
            }

            if (!inPipingSection && !trimmed.StartsWith("COMPONENT", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Parse PIPE lines
            var pipeMatch = PcfPipeRegex().Match(trimmed);
            if (pipeMatch.Success)
            {
                var component = new PcfComponent
                {
                    Tag = pipeMatch.Groups["tag"].Value.Trim('"', ' '),
                    Type = "PIPE",
                    Size = pipeMatch.Groups["dia"].Value,
                    Schedule = pipeMatch.Groups["sch"].Value,
                    Material = ExtractMaterial(pipeMatch.Groups["material"].Value),
                    Grade = ExtractGrade(pipeMatch.Groups["grade"].Value)
                };

                // Extract spool from tag if present
                var spoolMatch = SpoolTagRegex().Match(component.Tag);
                if (spoolMatch.Success)
                {
                    component.Spool = spoolMatch.Groups[1].Value.PadLeft(2, '0');
                }

                components[component.Tag] = component;
                continue;
            }

            // Parse COMPONENT lines
            var compMatch = PcfComponentRegex().Match(trimmed);
            if (compMatch.Success)
            {
                var component = new PcfComponent
                {
                    Tag = compMatch.Groups["tag"].Value.Trim('"', ' '),
                    Type = compMatch.Groups["type"].Value.ToUpperInvariant(),
                    Size = compMatch.Groups["dia"].Success ? compMatch.Groups["dia"].Value : "",
                    Schedule = compMatch.Groups["sch"].Success ? compMatch.Groups["sch"].Value : "",
                    Material = ExtractMaterial(compMatch.Groups["material"].Value),
                    Grade = ExtractGrade(compMatch.Groups["grade"].Value)
                };

                // Extract spool from tag if present
                var spoolMatch = SpoolTagRegex().Match(component.Tag);
                if (spoolMatch.Success)
                {
                    component.Spool = spoolMatch.Groups[1].Value.PadLeft(2, '0');
                }

                components[component.Tag] = component;
            }
        }

        return components;
    }

    private static List<PcfWeld> ParsePcfWelds(string[] lines, Dictionary<string, PcfComponent> components, List<string> _)
    {
        var welds = new List<PcfWeld>();
        bool inIsometricSection = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();

            // Check for ISOMETRIC section
            if (trimmed.Equals("ISOMETRIC", StringComparison.OrdinalIgnoreCase))
            {
                inIsometricSection = true;
                continue;
            }

            if (trimmed.Equals("END-ISOMETRIC", StringComparison.OrdinalIgnoreCase))
            {
                inIsometricSection = false;
            }

            if (!inIsometricSection && !trimmed.StartsWith("WELD", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Parse WELD lines
            var weldMatch = PcfWeldRegex().Match(trimmed);
            if (weldMatch.Success)
            {
                var weld = new PcfWeld
                {
                    Number = weldMatch.Groups["weld"].Value,
                    Type = weldMatch.Groups["type"].Value.ToUpperInvariant(),
                    ComponentATag = weldMatch.Groups["compA"].Value.Trim('"', ' '),
                    ComponentBTag = weldMatch.Groups["compB"].Value.Trim('"', ' '),
                    Spool = weldMatch.Groups["spool"].Success ?
                           weldMatch.Groups["spool"].Value.PadLeft(2, '0') : "01"
                };

                // Set location based on weld type
                weld.Location = (weld.Type == "FW" || weld.Type == "TH" || weld.Type == "FJ") ? "FW" : "WS";

                // Find components
                if (components.TryGetValue(weld.ComponentATag, out var compA))
                {
                    weld.ComponentA = compA;
                    // Update spool from component if not set
                    if (string.IsNullOrEmpty(weld.Spool) && !string.IsNullOrEmpty(compA.Spool))
                    {
                        weld.Spool = compA.Spool;
                    }
                }

                if (components.TryGetValue(weld.ComponentBTag, out var compB))
                {
                    weld.ComponentB = compB;
                    // Update spool from component if not set
                    if (string.IsNullOrEmpty(weld.Spool) && !string.IsNullOrEmpty(compB.Spool))
                    {
                        weld.Spool = compB.Spool;
                    }
                }

                welds.Add(weld);
            }
        }

        return welds;
    }

    private static DrawingMetadata MapPcfToDrawingMetadata(PcfHeader header, List<PcfWeld> welds)
    {
        var revTag = NormalizeRevisionTagForEquivalence(header.Revision);
        var metadata = new DrawingMetadata
        {
            PlantNo = header.UnitNo,
            Fluid = header.Fluid,
            LineNo = header.LineNo,
            DrawingNo = header.DrawingNo,
            Iso = header.Iso,
            LineClass = header.LineClass,
            LRev = revTag,
            LSRev = revTag,
            SheetNo = header.SheetNo,
            Diameter = header.Size,
            Schedule = header.Schedule,
            Material = header.Material,
            Location = header.Location,
            // Convert PCF welds to WeldRecord
            Welds = welds.Select(weld => new WeldRecord
            {
                Number = NormalizeWeldNumber(weld.Number),
                Type = weld.Type,
                Location = weld.Location,
                SpoolNo = weld.Spool,
                Size = GetWeldSize(weld),
                Schedule = GetWeldSchedule(weld),
                MaterialA = GetComponentDescription(weld.ComponentA, weld.ComponentATag),
                MaterialB = GetComponentDescription(weld.ComponentB, weld.ComponentBTag),
                GradeA = weld.ComponentA?.Grade ?? ExtractGradeFromMaterial(weld.ComponentA?.Material),
                GradeB = weld.ComponentB?.Grade ?? ExtractGradeFromMaterial(weld.ComponentB?.Material),
                OLDiameter = (weld.Type == "BR" || weld.Type == "LET") ?
                            GetBranchSize(weld) : null,
                OLSchedule = (weld.Type == "BR" || weld.Type == "LET") ?
                            GetBranchSchedule(weld) : null,
                Remarks = GetWeldRemarks(weld)
            }).ToList()
        };

        FillMissingMetadata(metadata);

        return metadata;
    }

    #endregion

    #region PCF Helper Methods

    private static void ExtractFluidLineFromString(string input, out string fluid, out string lineNo)
    {
        fluid = "";
        lineNo = "";

        var match = FluidLinePatternRegex().Match(input);
        if (match.Success)
        {
            fluid = match.Groups["fluid"].Value;
            lineNo = match.Groups["line"].Value;
        }
    }

    private static readonly string[] MaterialBlacklist =
    {
        "SHOWN", "NOTED", "LISTED", "TABLE", "LIST", "DESCRIPTION",
        "SPECIFIED", "REFER", "ABOVE", "BELOW", "DETAIL", "DRAWING",
        "MATD_", "MATD "
    };

    private static string ExtractMaterial(string materialText)
    {
        if (string.IsNullOrEmpty(materialText)) return "";

        materialText = materialText.Trim('"', ' ').ToUpperInvariant();

        // Reject non-material words commonly found in drawing notes
        if (MaterialBlacklist.Any(b => materialText.Contains(b, StringComparison.Ordinal)))
            return "";

        if (materialText.Contains("A106")) return "A106-B";
        if (materialText.Contains("A105")) return "A105";
        if (materialText.Contains("A234")) return "A234-WPB";
        if (materialText.Contains("CS")) return "CS";
        if (materialText.Contains("CARBON STEEL")) return "CS";

        return materialText;
    }

    private static string ExtractGrade(string gradeText)
    {
        if (string.IsNullOrEmpty(gradeText)) return "";

        gradeText = gradeText.Trim('"', ' ').ToUpperInvariant();

        if (gradeText.Contains("A106-B")) return "A106-B";
        if (gradeText.Contains("A105")) return "A105";
        if (gradeText.Contains("A234-WPB")) return "A234-WPB";
        if (gradeText.Contains("WPB")) return "A234-WPB";
        if (gradeText.Contains("A193")) return "A193";

        return gradeText;
    }

    private static string ExtractGradeFromMaterial(string? material)
    {
        if (string.IsNullOrEmpty(material)) return "";

        material = material.ToUpperInvariant();

        if (material.Contains("A106")) return "A106-B";
        if (material.Contains("A105")) return "A105";
        if (material.Contains("A234")) return "A234-WPB";
        if (material.Contains("A193")) return "A193";

        return "";
    }

    private static string NormalizeWeldNumber(string weldNo)
    {
        if (string.IsNullOrEmpty(weldNo)) return string.Empty;
        var trimmed = weldNo.Trim().ToUpperInvariant();

        // Strip OCR-noise prefix from letter-prefix weld numbers.
        // OCR spatial grouping may concatenate adjacent cell values (spool numbers,
        // line numbers, line class codes, IDs, etc.) with the weld number:
        //   "100369100S02" → "S02"   (line/ID number prefix)
        //   "08T16"        → "T16"   (spool number prefix)
        //   "1CS2U01S01"   → "S01"   (line class prefix)
        //   "890S01"       → "S01"   (drawing number prefix)
        //   "107763S01"    → "S01"   (line number prefix)
        // Extract the trailing [A-Z]\d{2,3} weld number when the prefix is ≥ 2 chars,
        // but NOT when the full string is a valid compound weld (e.g. "S13T12")
        // that should be split by SplitCompoundWeldNumbers instead.
        if (trimmed.Length >= 4 && !CompoundWeldNumberRegex().IsMatch(trimmed))
        {
            var trailMatch = TrailingWeldSuffixRegex().Match(trimmed);
            if (trailMatch.Success && trailMatch.Index >= 2)
            {
                trimmed = trailMatch.Value;
            }
        }

        // Fallback: strip single OCR-noise leading digit from letter-prefix weld numbers.
        // Handles residual cases like "5S01" → "S01" where prefix is only 1 char.
        if (trimmed.Length >= 4 && trimmed.Length <= 5
            && char.IsDigit(trimmed[0]) && char.IsLetter(trimmed[1]))
        {
            var afterStrip = trimmed[1..];
            if (afterStrip.Length >= 3 && afterStrip[1..].All(char.IsDigit))
            {
                trimmed = afterStrip;
            }
        }

        // Pad pure numeric weld numbers to 2 digits: "1" → "01"
        if (int.TryParse(trimmed, out var num) && num > 0 && num < 100)
        {
            return num.ToString("D2");
        }

        return trimmed;
    }

    /// <summary>
    /// Extracts the numeric portion of a weld number for natural sort ordering.
    /// "01"→1, "09"→9, "T10"→10, "11"→11, "T12"→12, "F01"→1, "09A"→9.
    /// This ensures the successive-WS-group algorithm processes welds in
    /// the same logical sequence as the weld list table on the drawing,
    /// regardless of OCR spatial order.
    /// </summary>
    private static int ExtractWeldSortNumber(string? weldNo)
    {
        if (string.IsNullOrEmpty(weldNo)) return int.MaxValue;
        var m = DigitsRegex().Match(weldNo);
        if (m.Success && int.TryParse(m.Value, out var num))
            return num;
        return int.MaxValue;
    }

    private static string GetWeldSize(PcfWeld weld)
    {
        // Try to get size from components
        if (weld.ComponentA?.Size != null)
            return weld.ComponentA.Size;
        if (weld.ComponentB?.Size != null)
            return weld.ComponentB.Size;

        return ""; // Will be filled from header
    }

    private static string GetWeldSchedule(PcfWeld weld)
    {
        // Try to get schedule from components
        if (weld.ComponentA?.Schedule != null)
            return weld.ComponentA.Schedule;
        if (weld.ComponentB?.Schedule != null)
            return weld.ComponentB.Schedule;

        return ""; // Will be filled from header
    }

    private static string GetComponentDescription(PcfComponent? component, string fallbackTag)
    {
        if (component != null)
        {
            // Map component type to standard MATD_Description from Material_Des_tbl
            return component.Type switch
            {
                "PIPE" => "PIPE",
                "TEE" => "TEE",
                "ELBOW" => "ELBOW",
                "FLANGE" => "FLANGE",
                "REDUCER" => "REDUCER",
                "VALVE" => "VALVE",
                "OLET" => "OLET",
                "NIPPLE" => "NIPPLE",
                "CAP" => "CAP",
                "PAD" => "PAD",
                "PLUG" => "PLUG",
                "COUPLING" => "COUPLING",
                "UNION" => "UNION",
                "SWAGE" => "SWAGE",
                "HOSE" => "HOSE",
                "PLATE" => "PLATE",
                "STUD BOLT" => "STUD BOLT",
                "BOLT" => "STUD BOLT",
                "NUT" => "STUD BOLT",
                _ => component.Type
            };
        }

        // For FJ welds
        if (fallbackTag is "STUD" or "STUD BOLT")
            return "STUD BOLT";
        if (fallbackTag == "BOLT")
            return "STUD BOLT";

        return fallbackTag;
    }

    /// <summary>
    /// Infers MaterialA and MaterialB (MATD_Description values) from the weld type
    /// when they are not already populated. Maps to valid Material_Des_tbl entries:
    /// CAP, COUPLING, ELBOW, FLANGE, HOSE, NIPPLE, OLET, PAD, PIPE, PLATE, PLUG,
    /// REDUCER, STUD BOLT, SWAGE, TEE, UNION, VALVE.
    ///
    /// For BW/FW welds, defaults to PIPE/PIPE (the most common pairing).
    /// The user adjusts Material B to ELBOW, TEE, REDUCER, etc. via the dropdown
    /// when the weld connects pipe to a fitting rather than pipe-to-pipe.
    /// </summary>
    private static (string materialA, string materialB) InferMaterialFromWeldType(string? weldType)
    {
        return (weldType ?? "").ToUpperInvariant() switch
        {
            "LET" => ("PIPE", "OLET"),           // Lethole: pipe-to-olet
            "SP" => ("PIPE", "PAD"),             // Support Pad weld: pipe-to-pad
            _ => ("", "")
        };
    }

    private static string GetBranchSize(PcfWeld weld)
    {
        // For branch welds, typically the smaller component is the branch
        if (weld.ComponentA != null && weld.ComponentB != null)
        {
            if (double.TryParse(weld.ComponentA.Size, out var sizeA) &&
                double.TryParse(weld.ComponentB.Size, out var sizeB))
            {
                return sizeA < sizeB ? weld.ComponentA.Size : weld.ComponentB.Size;
            }
        }

        return weld.ComponentA?.Size ?? weld.ComponentB?.Size ?? "";
    }

    private static string GetBranchSchedule(PcfWeld weld)
    {
        // For branch welds, use schedule from the branch component
        return weld.ComponentA?.Schedule ?? weld.ComponentB?.Schedule ?? "";
    }

    private static string GetWeldRemarks(PcfWeld _)
    {
        // Only include actionable remarks; skip internal processing notes
        // (weld type info is already visible in the Weld Type column)
        return string.Empty;
    }

    #endregion

    #region SHA File Parser - Enhanced

    private DrawingMetadata? ParseShaFile(Stream stream, string fileName, List<string> warnings)
    {
        try
        {
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();

            var metadata = new DrawingMetadata();
            var welds = new List<WeldRecord>();
            var components = new Dictionary<string, PcfComponent>();

            var lines = content.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            // Parse SHA format
            foreach (var line in lines)
            {
                var trimmed = line.Trim();

                // Parse weld records
                var weldMatch = ShaWeldRegex().Match(trimmed);
                if (weldMatch.Success)
                {
                    var typeValue = weldMatch.Groups["type"].Value;
                    var isFwOrTh = string.Equals(typeValue, "FW", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(typeValue, "TH", StringComparison.OrdinalIgnoreCase);
                    var weld = new WeldRecord
                    {
                        Number = weldMatch.Groups["weld"].Value.Replace("S", ""),
                        Type = typeValue.ToUpperInvariant(),
                        SpoolNo = weldMatch.Groups["spool"].Value.PadLeft(2, '0'),
                        Size = weldMatch.Groups["dia"].Value,
                        Schedule = weldMatch.Groups["sch"].Value,
                        MaterialA = CleanComponentName(weldMatch.Groups["compA"].Value),
                        MaterialB = CleanComponentName(weldMatch.Groups["compB"].Value),
                        GradeA = weldMatch.Groups["gradea"].Success ? weldMatch.Groups["gradea"].Value : null,
                        GradeB = weldMatch.Groups["gradeb"].Success ? weldMatch.Groups["gradeb"].Value : null,
                        Location = isFwOrTh ? "FW" : "WS"
                    };

                    welds.Add(weld);
                    continue;
                }

                // Parse pipe information
                var pipeMatch = ShaPipeRegex().Match(trimmed);
                if (pipeMatch.Success)
                {
                    if (string.IsNullOrEmpty(metadata.Diameter))
                        metadata.Diameter = pipeMatch.Groups["dia"].Value;
                    if (string.IsNullOrEmpty(metadata.Schedule))
                        metadata.Schedule = pipeMatch.Groups["sch"].Value;
                    if (string.IsNullOrEmpty(metadata.Material))
                        metadata.Material = pipeMatch.Groups["material"].Value;
                    continue;
                }

                // Parse component information
                var compMatch = ShaComponentRegex().Match(trimmed);
                if (compMatch.Success)
                {
                    // Could store for reference
                    continue;
                }

                // Try to extract metadata from key-value patterns
                ExtractShaMetadata(trimmed, metadata);
            }

            metadata.Welds = welds;

            // Fill missing metadata from filename
            if (string.IsNullOrEmpty(metadata.Fluid) || string.IsNullOrEmpty(metadata.LineNo))
            {
                ExtractMetadataFromFileName(fileName, metadata);
            }

            // Build ISO and DrawingNo if possible
            FillMissingMetadata(metadata);

            return metadata;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error parsing SHA file: {FileName}", fileName);
            warnings.Add($"Error parsing SHA file: {ex.Message}");
            return null;
        }
    }

    private static void ExtractShaMetadata(string line, DrawingMetadata metadata)
    {
        // Common SHA metadata patterns
        if (line.StartsWith("ISO:", StringComparison.OrdinalIgnoreCase))
        {
            metadata.Iso = line[4..].Trim();
        }
        else if (line.StartsWith("LINE:", StringComparison.OrdinalIgnoreCase))
        {
            metadata.LineNo = line[5..].Trim();
        }
        else if (line.StartsWith("DWG:", StringComparison.OrdinalIgnoreCase) ||
                 line.StartsWith("DRAWING:", StringComparison.OrdinalIgnoreCase))
        {
            metadata.DrawingNo = line.Contains(':') ?
                line[(line.IndexOf(':') + 1)..].Trim() :
                line;
        }
        else if (line.StartsWith("REV:", StringComparison.OrdinalIgnoreCase) ||
                 line.StartsWith("REVISION:", StringComparison.OrdinalIgnoreCase))
        {
            var rev = line.Contains(':') ?
                line[(line.IndexOf(':') + 1)..].Trim() :
                line;
            var revTag = NormalizeRevisionTagForEquivalence(rev);
            metadata.LRev = revTag;
            metadata.LSRev = revTag;
        }
        else if (line.StartsWith("SHEET:", StringComparison.OrdinalIgnoreCase))
        {
            metadata.SheetNo = line[6..].Trim();
        }
    }

    private static string CleanComponentName(string component)
    {
        var value = component.Trim();

        // Standardize component names
        if (value.Contains("TEE", StringComparison.OrdinalIgnoreCase)) return "TEE";
        if (value.Contains("ELBOW", StringComparison.OrdinalIgnoreCase)) return "ELBOW";
        if (value.Contains("FLANGE", StringComparison.OrdinalIgnoreCase)) return "FLANGE";
        if (value.Contains("REDUCER", StringComparison.OrdinalIgnoreCase)) return "REDUCER";
        if (value.Contains("VALVE", StringComparison.OrdinalIgnoreCase)) return "VALVE";
        if (value.Contains("OLET", StringComparison.OrdinalIgnoreCase)) return "OLET";
        if (value.Contains("NIPPLE", StringComparison.OrdinalIgnoreCase)) return "NIPPLE";
        if (value.Contains("PIPE", StringComparison.OrdinalIgnoreCase)) return "PIPE";
        if (value.Contains("STUD", StringComparison.OrdinalIgnoreCase)) return "STUD";
        if (value.Contains("BOLT", StringComparison.OrdinalIgnoreCase)) return "BOLT";

        return value.ToUpperInvariant();
    }

    #endregion

    #region Main Analysis Method

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> JointRecord(int? projectId, string? scope, string? mode)
    {
        var projects = await _context.Projects_tbl
            .AsNoTracking()
            .OrderBy(p => p.Project_ID)
            .Select(p => new ProjectOption
            {
                Id = p.Project_ID,
                Name = p.Project_Name ?? $"Project {p.Project_ID}",
                LineSheet = p.Line_Sheet
            })
            .ToListAsync();

        var defaultId = projectId ?? await _context.Projects_tbl
            .AsNoTracking()
            .Where(p => p.Default_P)
            .Select(p => (int?)p.Project_ID)
            .FirstOrDefaultAsync();

        var selectedId = defaultId ?? projects.FirstOrDefault()?.Id ?? 0;
        var weldersProjectId = await ResolveWeldersProjectIdAsync(selectedId);

        var normalizedScope = (scope ?? ScopeJoints).Trim();
        if (!normalizedScope.Equals(ScopeLineList, StringComparison.OrdinalIgnoreCase)
            && !normalizedScope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
            normalizedScope = ScopeJoints;

        var normalizedMode = (mode ?? ModeUpload).Trim();
        if (!normalizedMode.Equals(ModeEdit, StringComparison.OrdinalIgnoreCase))
            normalizedMode = ModeUpload;

        var materialDescriptions = await _context.Material_Des_tbl.AsNoTracking()
            .Where(m => m.MATD_Description != null && m.MATD_Description != "")
            .Select(m => m.MATD_Description)
            .Distinct()
            .OrderBy(d => d)
            .ToListAsync();

        var locations = await _context.PMS_Location_tbl.AsNoTracking()
            .Where(l => l.LO_Project_No == weldersProjectId && l.P_Location != null && l.P_Location != "")
            .Select(l => l.P_Location!)
            .Distinct()
            .OrderBy(l => l)
            .ToListAsync();

        var weldTypes = await _context.PMS_Weld_Type_tbl.AsNoTracking()
            .Where(w => w.W_Project_No == weldersProjectId && w.P_Weld_Type != null && w.P_Weld_Type != "")
            .Select(w => w.P_Weld_Type!)
            .Distinct()
            .OrderBy(w => w)
            .ToListAsync();

        var grades = await _context.Material_tbl.AsNoTracking()
            .Where(m => m.MAT_GRADE != null && m.MAT_GRADE != "")
            .Select(m => m.MAT_GRADE!)
            .Distinct()
            .OrderBy(g => g)
            .ToListAsync();

        var materials = await _context.Material_tbl.AsNoTracking()
            .Where(m => m.MAT != null && m.MAT != "")
            .Select(m => m.MAT!)
            .Distinct()
            .OrderBy(m => m)
            .ToListAsync();

        var jAddOptions = await _context.PMS_J_Add_tbl.AsNoTracking()
            .Where(x => x.Add_Project_No == weldersProjectId)
            .Where(x => x.Add_J_Add != null && x.Add_J_Add != "")
            .GroupBy(x => x.Add_J_Add!)
            .Select(g => new { Name = g.Key, MinId = g.Min(x => x.Add_ID) })
            .OrderBy(x => x.MinId)
            .Select(x => x.Name)
            .ToListAsync();
        if (jAddOptions.Count == 0)
        {
            jAddOptions = ["NEW", "R1", "R2"];
        }

        var model = new JointRecordViewModel
        {
            Projects = projects,
            SelectedProjectId = selectedId,
            Scope = normalizedScope,
            Mode = normalizedMode,
            Rows = new List<JointRecordAnalysisRow>(),
            Warnings = new List<string>(),
            MaterialDescriptions = materialDescriptions,
            Materials = materials,
            Locations = locations,
            WeldTypes = weldTypes,
            Grades = grades,
            JAddOptions = jAddOptions
        };

        return View("JointRecord", model);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    [RequestSizeLimit(50_000_000)]
    public async Task<IActionResult> AnalyzeJointRecord(
        IFormFile? drawingFile,
        IFormFile? excelFile,
        [FromForm] int projectId,
        [FromForm] string? scope = ScopeJoints,
        [FromForm] bool advancedParsing = true)
    {
        _ = advancedParsing;
        var analyzeScope = (scope ?? ScopeJoints).Trim();
        var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);

        var result = new JointRecordAnalysisResult
        {
            ProjectId = projectId,
            AnalysisName = $"JointRecord_{DateTime.UtcNow:yyyyMMddHHmmss}"
        };

        try
        {
            List<JointRecordAnalysisRow> excelRows = new();

            // Step 1: Process Excel template if provided
            if (excelFile != null && excelFile.Length > 0)
            {
                await using var excelStream = excelFile.OpenReadStream();
                excelRows = ReadJointRecordExcelWithValidation(excelStream, result.Warnings);
                if (excelRows.Count > 0)
                {
                    result.Warnings.Add($"Successfully imported {excelRows.Count} rows from Excel template.");
                }
            }

            // Step 2: Process uploaded file based on format
            DrawingMetadata? metadata = null;
            if (drawingFile != null && drawingFile.Length > 0)
            {
                var fileExtension = Path.GetExtension(drawingFile.FileName).ToLowerInvariant();

                using var stream = drawingFile.OpenReadStream();

                // Route to appropriate parser based on file type
                switch (fileExtension)
                {
                    case ".pcf":
                        metadata = ParsePcfFile(stream, drawingFile.FileName, result.Warnings);
                        if (metadata != null && metadata.Welds.Count > 0)
                            result.Warnings.Add($"Parsed PCF file: {metadata.Welds.Count} welds found.");
                        break;

                    case ".sha":
                        metadata = ParseShaFile(stream, drawingFile.FileName, result.Warnings);
                        if (metadata != null && metadata.Welds.Count > 0)
                            result.Warnings.Add($"Parsed SHA file: {metadata.Welds.Count} welds found.");
                        break;

                    case ".pdf":
                        metadata = await ParsePdfDrawingAsync(stream, drawingFile.FileName, result.Warnings);
                        if (metadata != null && metadata.Welds.Count > 0)
                            result.Warnings.Add($"Parsed PDF drawing: {metadata.Welds.Count} welds found.");
                        break;

                    default:
                        result.Warnings.Add($"Unsupported file format: {fileExtension}. Supported: .pcf, .sha, .pdf");
                        break;
                }

                // If no metadata was extracted, try filename parsing
                if (metadata == null || metadata.Welds.Count == 0)
                {
                    metadata ??= new DrawingMetadata();
                    ExtractMetadataFromFileName(drawingFile.FileName, metadata);
                    if (!string.IsNullOrEmpty(metadata.Fluid) || !string.IsNullOrEmpty(metadata.LineNo))
                    {
                        result.Warnings.Add("Extracted metadata from file name only.");
                    }
                }
            }

            // Step 3: Generate final rows
            var generatedRows = GenerateRowsFromTemplateMatch(metadata, excelRows,
                drawingFile?.FileName ?? excelFile?.FileName ?? "Unknown", result.Warnings);

            // Step 4: Apply business rules and validation
            bool hasExcelSource = excelRows.Count > 0;
            ApplyBusinessRules(generatedRows, result.Warnings, hasExcelSource);

            // Step 4a: Split OCR-concatenated compound weld numbers (e.g. "S13T12" → two rows: "S13" and "T12")
            SplitCompoundWeldNumbers(generatedRows, result.Warnings);

            // Step 4a2: Extract J-Add codes embedded in weld numbers (e.g. "01CS1" → Weld "01", J-Add "CS1")
            var jAddOpts = await _context.PMS_J_Add_tbl.AsNoTracking()
                .Where(x => x.Add_Project_No == weldersProjectId)
                .Where(x => x.Add_J_Add != null && x.Add_J_Add != "")
                .Select(x => x.Add_J_Add!)
                .Distinct()
                .ToListAsync();
            if (jAddOpts.Count == 0) jAddOpts = ["NEW", "R1", "R2"];
            ExtractJAddFromWeldNumbers(generatedRows, jAddOpts, result.Warnings);

            // Step 4b: Reassign Spool numbers via successive WS group algorithm
            // on the final rows. This overrides any incorrect proximity-based or
            // default spool assignments from PDF parsing.
            // Skip when Excel rows already have Spool values — preserve user input.
            // Exclude Delete/Cancel rows so they don't distort spool group counting.
            var orderedSpools = metadata?.OrderedSpools ?? [];
            var cuttingListSpoolSizes = metadata?.CuttingListSpoolSizes
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool hasExcelSpools = excelRows.Count > 0 && excelRows.Any(r => !string.IsNullOrWhiteSpace(r.SpoolNumber));
            var activeRows = generatedRows.Where(r => !r.Deleted && !r.Cancelled).ToList();
            if (!hasExcelSpools)
            {
                ReassignSpoolsBySuccessiveWsGroup(activeRows, orderedSpools, result.Warnings);
            }

            // Step 4b2: Populate Spool Dia. In. from cutting list SIZE or spool group
            PopulateSpoolDiameterFromGroup(activeRows, cuttingListSpoolSizes, result.Warnings);

            // Step 4b3: Populate OL Dia/Schedule for BR/LET welds from group max
            PopulateOLFieldsFromGroupMax(activeRows, result.Warnings);

            // Step 4b4: Populate ISO Dia. In. as Max(Dia. In.) per (LayoutNo, Sheet)
            PopulateIsoDiameterFromGroupMax(activeRows);

            // Clear Spool / SpDia on Delete/Cancel rows — those fields are not
            // available in the Excel sheet for deleted or cancelled joints.
            foreach (var r in generatedRows.Where(r => r.Deleted || r.Cancelled))
            {
                r.SpoolNumber = null;
                r.SpDia = null;
            }

            // Step 4c: Determine Material from biggest Dia row's Grade A via Material_tbl
            await PopulateMaterialFromGroupMaxGradeAsync(generatedRows, result.Warnings);

            // Step 5: Process and clean rows
            if (generatedRows.Count > 0)
            {
                NormalizeRevisionFields(generatedRows);
                NormalizeLineNumberFromLayout(generatedRows, hasExcelSource);
                PopulateIsoFromLayoutAndClass(generatedRows);
                BackfillFromIso(generatedRows, hasExcelSource);

                // Validate and clean up
                foreach (var row in generatedRows)
                {
                    ValidateAndCorrectJointRecordRow(row, result.Warnings, hasExcelSource);
                    AutoPopulateMissingFields(row);
                }

                // Step 6: Validate against project lookup tables
                if (projectId > 0)
                {
                    await ValidateAgainstLookupTablesAsync(projectId, generatedRows, result.Warnings);
                }

                // Remove duplicates and empty rows
                result.Rows = CleanAndDeduplicateRows(generatedRows);

                result.Warnings.Add($"Final analysis contains {result.Rows.Count} unique joint records.");
            }
            else
            {
                // Fallback to filename parsing, enriched with any PDF metadata
                if (drawingFile != null && drawingFile.Length > 0)
                {
                    var fileRow = BuildJointRecordFromFileName(drawingFile.FileName);

                    // Enrich fallback row with metadata extracted from PDF text
                    if (metadata != null)
                    {
                        EnrichRowWithMetadata(fileRow, metadata);
                    }

                    if (!IsJointRecordRowEmpty(fileRow))
                    {
                        result.Rows.Add(fileRow);
                        result.Warnings.Add("Extracted data from file name only.");
                    }
                }
            }

            MergeImportAndFinalWarnings(result.Warnings, excelRows.Count, result.Rows.Count);

            // Populate action status remarks (Upload / Update / Cannot update)
            if (projectId > 0 && result.Rows.Count > 0)
            {
                await PopulateActionStatusRemarksAsync(projectId, analyzeScope, result.Rows);
            }

            if (result.Rows.Count == 0)
            {
                result.Warnings.Add("No data extracted. Please check your file format and content.");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error analyzing joint record");
            result.Warnings.Add($"Analysis error: {ex.Message}");
        }

        return new JsonResult(result, JointRecordJsonOptions);
    }

    #endregion

    #region Row Generation and Processing

    private static List<JointRecordAnalysisRow> GenerateRowsFromTemplateMatch(
        DrawingMetadata? metadata,
        List<JointRecordAnalysisRow> excelRows,
        string fileName,
        List<string> warnings)
    {
        _ = warnings;
        var rows = new List<JointRecordAnalysisRow>();

        // If we have Excel template rows, use them as primary source
        if (excelRows.Count > 0)
        {
            foreach (var excelRow in excelRows)
            {
                // Enrich with drawing metadata if available
                if (metadata != null)
                {
                    EnrichRowWithMetadata(excelRow, metadata);
                }

                rows.Add(excelRow);
            }
            return rows;
        }

        // Generate from drawing metadata
        if (metadata != null && metadata.Welds.Count > 0)
        {
            int weldCounter = 1;
            foreach (var weld in metadata.Welds)
            {
                var row = CreateRowFromWeld(metadata, weld, fileName, weldCounter++);
                rows.Add(row);
            }
        }

        return rows;
    }

    private static void EnrichRowWithMetadata(JointRecordAnalysisRow row, DrawingMetadata metadata)
    {
        row.UnitNumber ??= metadata.PlantNo;
        row.Fluid ??= metadata.Fluid;
        row.LineNo ??= metadata.LineNo;
        row.LayoutNo ??= metadata.DrawingNo;
        row.LineClass ??= metadata.LineClass;
        row.ISO ??= metadata.Iso;
        row.LSSheet ??= metadata.SheetNo;
        row.LSDiameter ??= metadata.Diameter;
        row.Location ??= metadata.Location;

        // Copy revision from drawing metadata if not already set
        if (string.IsNullOrEmpty(row.LRev) && !string.IsNullOrEmpty(metadata.LRev))
            row.LRev = metadata.LRev;
        if (string.IsNullOrEmpty(row.LSRev) && !string.IsNullOrEmpty(metadata.LSRev))
            row.LSRev = metadata.LSRev;

        // Copy material if available
        if (string.IsNullOrEmpty(row.Material) && !string.IsNullOrEmpty(metadata.Material))
            row.Material = metadata.Material;
    }

    private static JointRecordAnalysisRow CreateRowFromWeld(DrawingMetadata metadata, WeldRecord weld, string fileName, int counter)
    {
        var location = weld.Location ?? "WS";
        // Pass through the spool assignment from ParseWeldDataFromText (Step D).
        // Do NOT default to "01" — the row-level ReassignSpoolsBySuccessiveWsGroup
        // method will handle spool assignment using the successive WS group algorithm
        // with the authoritative ordered spool list from the cutting list.
        var spoolNo = weld.SpoolNo; // null for field welds and unassigned WS welds

        return new JointRecordAnalysisRow
        {
            UnitNumber = metadata.PlantNo,
            Fluid = metadata.Fluid,
            LineNo = metadata.LineNo,
            LayoutNo = metadata.DrawingNo,
            LineClass = metadata.LineClass,
            ISO = metadata.Iso,
            LRev = metadata.LRev,
            LSRev = metadata.LSRev,
            LSSheet = metadata.SheetNo,
            Material = metadata.Material ?? "CS",
            LSDiameter = metadata.Diameter,
            Location = location,
            WeldNumber = weld.Number ?? counter.ToString("D2"),
            JAdd = "NEW",
            WeldType = weld.Type ?? "BW",
            SpoolNumber = spoolNo,
            Diameter = weld.Size ?? metadata.Diameter,
            Schedule = weld.Schedule ?? (metadata.Schedule ?? "40"),
            SpDia = weld.SpoolDia,
            OLDiameter = weld.OLDiameter,
            OLSchedule = weld.OLSchedule,
            MaterialA = weld.MaterialA ?? "",
            MaterialB = weld.MaterialB ?? "",
            GradeA = weld.GradeA ?? "",
            GradeB = weld.GradeB ?? "",
            LSScope = "AIC",
            Remarks = weld.Remarks,
            SourceFile = fileName
        };
    }

    private static void ApplyBusinessRules(List<JointRecordAnalysisRow> rows, List<string> warnings, bool hasExcelSource = false)
    {
        foreach (var row in rows)
        {
            // Validate required fields
            if (string.IsNullOrEmpty(row.LayoutNo))
                AppendRowRemark(row, warnings, "Layout No is required");

            if (string.IsNullOrEmpty(row.WeldNumber))
                AppendRowRemark(row, warnings, "Weld No is required");

            // J-Add is now mandatory
            if (string.IsNullOrEmpty(row.JAdd))
                AppendRowRemark(row, warnings, "J-Add is required");

            // Weld Type: only default to "BW" when NOT from an Excel template.
            // When the source is Excel, update only non-empty cells — leave
            // user-empty fields as-is so they can be filled manually.
            if (string.IsNullOrEmpty(row.WeldType) && !hasExcelSource)
                row.WeldType = "BW";

            // OL fields for BR/LET welds are populated by PopulateOLFieldsFromGroupMax
            // after ApplyBusinessRules, using the largest Dia. In. in the same
            // (LayoutNo, Sheet, Spool) group as the header/run pipe size.

            // Normalize spool number
            if (!string.IsNullOrEmpty(row.SpoolNumber))
            {
                var spoolMatch = SpoolNumberRegex().Match(row.SpoolNumber);
                if (spoolMatch.Success)
                {
                    row.SpoolNumber = spoolMatch.Groups[1].Value.PadLeft(2, '0');
                }
            }

            // Normalize weld number
            if (!string.IsNullOrEmpty(row.WeldNumber))
            {
                row.WeldNumber = row.WeldNumber.Trim().ToUpperInvariant();
            }

            // Auto-populate MaterialA/MaterialB from weld type when missing
            if (string.IsNullOrEmpty(row.MaterialA) || string.IsNullOrEmpty(row.MaterialB))
            {
                var (inferredA, inferredB) = InferMaterialFromWeldType(row.WeldType);
                if (string.IsNullOrEmpty(row.MaterialA) && !string.IsNullOrEmpty(inferredA))
                    row.MaterialA = inferredA;
                if (string.IsNullOrEmpty(row.MaterialB) && !string.IsNullOrEmpty(inferredB))
                    row.MaterialB = inferredB;
            }

            // For FJ welds, set non-material properties
            if (row.WeldType == "FJ")
            {
                row.Diameter = row.LSDiameter ?? row.Diameter ?? "8";
                row.SpoolNumber = null;
                row.Location = "FW";
            }

            // For TH welds (threaded)
            if (row.WeldType == "TH")
            {
                row.Schedule = "160"; // Default for threaded connections
            }

        }
    }

    /// <summary>
    /// Normalizes the Material A/B convention on a <see cref="JointRecordAnalysisRow"/>.
    /// Material A should be the pipe side (PIPE, NIPPLE, SWAGE) and Material B
    /// should be the fitting side. If Material A is a fitting and Material B is
    /// pipe-like (or empty), swaps both Material and Grade values.
    /// </summary>
    private static void NormalizeMaterialABRow(JointRecordAnalysisRow row)
    {
        var matA = (row.MaterialA ?? "").ToUpperInvariant();
        var matB = (row.MaterialB ?? "").ToUpperInvariant();

        bool aIsFitting = matA is "ELBOW" or "TEE" or "FLANGE" or "REDUCER"
            or "VALVE" or "OLET" or "CAP" or "PLUG" or "COUPLING" or "UNION"
            or "HOSE" or "PLATE" or "PAD";
        bool bIsPipe = matB is "PIPE" or "NIPPLE" or "SWAGE";

        if (aIsFitting && bIsPipe)
        {
            (row.MaterialA, row.MaterialB) = (row.MaterialB, row.MaterialA);
            (row.GradeA, row.GradeB) = (row.GradeB, row.GradeA);
        }
        else if (aIsFitting && string.IsNullOrEmpty(matB))
        {
            row.MaterialB = row.MaterialA;
            row.GradeB = row.GradeA;
            row.MaterialA = "PIPE";
            row.GradeA = InferGradeFromComponent("PIPE");
        }
        else if (aIsFitting && !bIsPipe && matA != matB)
        {
            // Both are fittings — force Material A to PIPE
            row.MaterialA = "PIPE";
            row.GradeA = InferGradeFromComponent("PIPE");
        }
    }

    /// <summary>
    /// Detects J-Add codes embedded in weld numbers and splits them.
    /// For example, weld number "01CS1" where "CS1" is a recognized J-Add option
    /// becomes Weld No "01" with J-Add "CS1".
    /// Checks for J-Add codes appearing as a suffix of the weld number.
    /// Longer J-Add codes are matched first to avoid partial matches.
    /// </summary>
    private static void ExtractJAddFromWeldNumbers(
        List<JointRecordAnalysisRow> rows,
        List<string> jAddOptions,
        List<string> _)
    {
        if (rows.Count == 0 || jAddOptions.Count == 0) return;

        // Sort options by length descending to match longest suffix first
        var sortedOptions = jAddOptions
            .Where(o => !string.IsNullOrEmpty(o))
            .OrderByDescending(o => o.Length)
            .ToList();

        foreach (var row in rows)
        {
            var weldNo = (row.WeldNumber ?? "").Trim().ToUpperInvariant();
            if (string.IsNullOrEmpty(weldNo)) continue;

            // Skip if J-Add is already set to something other than NEW/empty
            if (!string.IsNullOrEmpty(row.JAdd)
                && !string.Equals(row.JAdd, "NEW", StringComparison.OrdinalIgnoreCase))
                continue;

            foreach (var jAdd in sortedOptions)
            {
                var jAddUpper = jAdd.ToUpperInvariant();

                // Only check suffix position (e.g. "01CS1" ends with "CS1")
                if (!weldNo.EndsWith(jAddUpper, StringComparison.Ordinal))
                    continue;

                var cleanWeld = weldNo[..^jAddUpper.Length];

                // Remaining part must be a valid weld number:
                // at least one digit, only letters/digits, reasonable length
                if (cleanWeld.Length == 0 || cleanWeld.Length > 6) continue;
                if (!cleanWeld.Any(char.IsDigit)) continue;
                if (!cleanWeld.All(char.IsLetterOrDigit)) continue;

                row.WeldNumber = NormalizeWeldNumber(cleanWeld);
                row.JAdd = jAdd;
                break;
            }
        }
    }

    /// <summary>
    /// Splits compound weld numbers like "S13T12" into separate rows: "S13" and "T12".
    /// This handles OCR artifacts where two adjacent weld callouts on the drawing
    /// are concatenated into a single token (e.g., "S13" next to "T12" → "S13T12").
    /// The pattern detects sequences of [Letter][Digits] repeated two or more times.
    /// Each resulting row inherits all other properties from the original row.
    /// </summary>
    private static void SplitCompoundWeldNumbers(List<JointRecordAnalysisRow> rows, List<string> warnings)
    {
        if (rows.Count == 0) return;

        var newRows = new List<JointRecordAnalysisRow>();
        var indicesToRemove = new List<int>();

        for (int i = 0; i < rows.Count; i++)
        {
            var weldNo = (rows[i].WeldNumber ?? "").Trim().ToUpperInvariant();
            if (string.IsNullOrEmpty(weldNo)) continue;
            if (!CompoundWeldNumberRegex().IsMatch(weldNo)) continue;

            var parts = CompoundWeldPartRegex().Matches(weldNo)
                .Cast<Match>()
                .Select(m => m.Value)
                .ToList();

            if (parts.Count < 2) continue;

            indicesToRemove.Add(i);

            foreach (var part in parts)
            {
                var splitRow = new JointRecordAnalysisRow
                {
                    UnitNumber = rows[i].UnitNumber,
                    Fluid = rows[i].Fluid,
                    LineNo = rows[i].LineNo,
                    LayoutNo = rows[i].LayoutNo,
                    LineClass = rows[i].LineClass,
                    ISO = rows[i].ISO,
                    LRev = rows[i].LRev,
                    LSRev = rows[i].LSRev,
                    LSSheet = rows[i].LSSheet,
                    Material = rows[i].Material,
                    LSDiameter = rows[i].LSDiameter,
                    Location = rows[i].Location,
                    WeldNumber = NormalizeWeldNumber(part),
                    JAdd = rows[i].JAdd,
                    WeldType = rows[i].WeldType,
                    SpoolNumber = rows[i].SpoolNumber,
                    Diameter = rows[i].Diameter,
                    Schedule = rows[i].Schedule,
                    SpDia = rows[i].SpDia,
                    OLDiameter = rows[i].OLDiameter,
                    OLSchedule = rows[i].OLSchedule,
                    OLThick = rows[i].OLThick,
                    MaterialA = rows[i].MaterialA,
                    MaterialB = rows[i].MaterialB,
                    GradeA = rows[i].GradeA,
                    GradeB = rows[i].GradeB,
                    LSScope = rows[i].LSScope,
                    SourceFile = rows[i].SourceFile,
                };
                newRows.Add(splitRow);
            }

            warnings.Add($"Split compound weld number '{weldNo}' into: {string.Join(", ", parts)}");
        }

        // Remove original compound rows (reverse order to preserve indices)
        for (int i = indicesToRemove.Count - 1; i >= 0; i--)
            rows.RemoveAt(indicesToRemove[i]);

        rows.AddRange(newRows);
    }

    /// <summary>
    /// Reassigns spool numbers on the final <see cref="JointRecordAnalysisRow"/> rows
    /// using the successive WS group algorithm. This is the authoritative, row-level
    /// spool assignment that overrides any incorrect values from PDF proximity-based
    /// assignment or default fallbacks.
    ///
    /// Algorithm:
    /// 1. Sort rows by natural weld number order (01, 02, …, 09, T10, 11, T12, 13, …)
    /// 2. Walk the sorted rows: consecutive WS welds form a group assigned to the
    ///    next spool from <paramref name="orderedSpools"/>. A non-WS weld (YFW, FW,
    ///    etc.) breaks the current group and advances the spool pointer.
    /// 3. Non-WS welds get no spool (null).
    ///
    /// If <paramref name="orderedSpools"/> is empty, falls back to counting WS groups
    /// and generating sequential spool numbers ("01", "02", "03", …).
    /// </summary>
    private static void ReassignSpoolsBySuccessiveWsGroup(
        List<JointRecordAnalysisRow> rows,
        List<string> orderedSpools,
        List<string> warnings)
    {
        if (rows.Count == 0) return;

        // Sort by natural weld number once for both group counting and assignment
        var sortedRows = rows
            .OrderBy(r => ExtractWeldSortNumber(r.WeldNumber))
            .ThenBy(r => r.WeldNumber, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Count WS groups to know how many spools are required
        int wsGroupCount = 0;
        bool prevIsWsForCount = false;
        foreach (var r in sortedRows)
        {
            bool isWs = (r.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase);
            if (isWs && !prevIsWsForCount)
                wsGroupCount++;
            prevIsWsForCount = isWs;
        }

        if (wsGroupCount == 0) return; // No WS welds at all

        // Start from the ordered list if provided, otherwise generate sequential spools
        var spools = orderedSpools.Count > 0
            ? orderedSpools.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            : Enumerable.Range(1, wsGroupCount).Select(i => i.ToString("D2")).ToList();

        // If the provided list is shorter than the WS group count, fill the gaps
        if (spools.Count < wsGroupCount)
        {
            var existing = new HashSet<string>(spools, StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i <= wsGroupCount; i++)
            {
                var candidate = i.ToString("D2");
                if (!existing.Contains(candidate))
                    spools.Add(candidate);
            }

            // Sort numerically to align spool numbering with WS group order
            spools = spools
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => int.TryParse(s, out var n) ? n : int.MaxValue)
                .ThenBy(s => s, StringComparer.OrdinalIgnoreCase)
                .ToList();

            warnings.Add($"Supplemented spool list to cover {wsGroupCount} WS groups: {string.Join(", ", spools.Select(s => $"SP{s}"))}");
        }


        // Walk the sorted rows and assign spools
        int spoolIdx = 0;
        bool prevWasWs = false;
        int unassignedWsCount = 0;

        foreach (var row in sortedRows)
        {
            bool isWs = (row.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase);

            if (isWs)
            {
                if (spoolIdx < spools.Count)
                {
                    row.SpoolNumber = spools[spoolIdx];
                }
                else
                {
                    row.SpoolNumber = null;
                    unassignedWsCount++;
                }
                prevWasWs = true;
            }
            else
            {
                // Non-WS weld: no spool; advance pointer when breaking a WS group
                if (prevWasWs)
                    spoolIdx++;
                row.SpoolNumber = null;
                prevWasWs = false;
            }
        }

        if (unassignedWsCount > 0)
        {
            warnings.Add($"{unassignedWsCount} WS weld(s) could not be assigned to a spool — more WS groups than available spools ({spools.Count}). Review spool assignments.");
        }
    }

    /// <summary>
    /// Populates the "Spool Dia. In." (SpDia) field for every row.
    ///
    /// Priority:
    /// 1. Cutting list SIZE (authoritative): look up the row's SpoolNumber in the
    ///    <paramref name="cuttingListSpoolSizes"/> map built from the CUTTING LIST
    ///    table's SPOOL NO → mode(SIZE) mapping.
    /// 2. If SpDia is already set on the row (e.g. from WeldRecord.SpoolDia during
    ///    PDF parsing), keep it.
    /// 3. Propagate from the spool group: among all WS rows that share the same
    ///    (LayoutNo, LSSheet, SpoolNumber), the largest "Dia. In." value
    ///    represents the spool's header/run pipe size.
    /// 4. Last-resort fallback for WS rows: use the row's own Diameter.
    ///
    /// Non-WS rows (field welds) do not belong to a spool; their SpDia is left null
    /// unless explicitly set by the cutting list.
    ///
    /// Note: "Dia. In." (Diameter) is the weld joint diameter which can differ from
    /// the spool pipe diameter for SOF/BR/LET welds (e.g. weld 01 is 3″ SOF on a
    /// spool whose pipe is 2″). This method ensures SpDia reflects the spool's
    /// pipe size, not the individual weld size.
    /// </summary>
    private static void PopulateSpoolDiameterFromGroup(
        List<JointRecordAnalysisRow> rows,
        Dictionary<string, string> cuttingListSpoolSizes,
        List<string> _)
    {
        if (rows.Count == 0) return;

        // Clear any stale SpDia values before recomputing.
        // SpDia must follow the spool assignment from ReassignSpoolsBySuccessiveWsGroup,
        // not any earlier proximity-based or default spool assignments.
        foreach (var row in rows)
            row.SpDia = null;

        // Step 1: Apply cutting list SIZE directly to every row that has a spool.
        // The cutting list is the authoritative source for spool pipe size.
        if (cuttingListSpoolSizes.Count > 0)
        {
            foreach (var row in rows)
            {
                if (string.IsNullOrEmpty(row.SpoolNumber)) continue;
                if (cuttingListSpoolSizes.TryGetValue(row.SpoolNumber, out var clSize))
                {
                    row.SpDia = clSize;
                }
            }
        }

        // Step 2: Build spool group → max diameter lookup for rows still missing SpDia.
        // Use SpDia values already set (from cutting list or WeldRecord.SpoolDia)
        // to propagate to remaining rows in the same spool group.
        var spoolGroupDia = new Dictionary<(string, string, string), string>(
            StringTupleComparer.Instance);

        var groups = rows
            .Where(r => !string.IsNullOrEmpty(r.SpoolNumber))
            .GroupBy(r => (r.LayoutNo ?? "", r.LSSheet ?? "", r.SpoolNumber ?? ""),
                     StringTupleComparer.Instance);

        foreach (var g in groups)
        {
            // Prefer SpDia values already set from cutting list
            var fromCuttingList = g
                .Where(r => !string.IsNullOrEmpty(r.SpDia))
                .Select(r => r.SpDia!)
                .FirstOrDefault();

            if (!string.IsNullOrEmpty(fromCuttingList))
            {
                spoolGroupDia[g.Key] = fromCuttingList;
                continue;
            }

            // Infer from the largest Diameter among WS rows in this spool group.
            // WS rows belong to the spool; non-WS rows (YFW/FW) are between spools.
            var wsDiameters = g
                .Where(r => (r.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase)
                            && !string.IsNullOrEmpty(r.Diameter))
                .Select(r => r.Diameter!)
                .ToList();

            if (wsDiameters.Count == 0) continue;

            // Max: largest diameter in the spool group
            var maxDia = wsDiameters
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(d => TryParseDiameter(d, out var v) ? v : 0)
                .First();

            spoolGroupDia[g.Key] = maxDia;
        }

        // Step 3: Apply SpDia to every row that still has none
        foreach (var row in rows)
        {
            // Already set from cutting list or Step 1 — keep it
            if (!string.IsNullOrEmpty(row.SpDia)) continue;

            bool isWs = (row.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase);

            if (!string.IsNullOrEmpty(row.SpoolNumber))
            {
                var key = (row.LayoutNo ?? "", row.LSSheet ?? "", row.SpoolNumber ?? "");
                if (spoolGroupDia.TryGetValue(key, out var groupDia))
                {
                    row.SpDia = groupDia;
                    continue;
                }
            }

            // Last-resort fallback for WS rows: use the row's own Diameter
            if (isWs && !string.IsNullOrEmpty(row.Diameter))
            {
                row.SpDia = row.Diameter;
            }
            // Non-WS rows without a spool: SpDia stays null
        }
    }

    /// <summary>
    /// For BR/LET (branch/olet) welds, OL Dia and OL Schedule represent the
    /// header (run) pipe size the branch connects to. When these fields are not
    /// already populated (e.g. from PCF component data), infer them by finding
    /// the largest "Dia. In." among all rows that share the same
    /// (LayoutNo, LSSheet, SpoolNumber) group – that largest size is the run pipe.
    /// </summary>
    private static void PopulateOLFieldsFromGroupMax(List<JointRecordAnalysisRow> rows, List<string> warnings)
    {
        // Build lookup: group key → (maxDiameter numeric, maxDiameter string, schedule of that row)
        var groups = rows
            .Where(r => !string.IsNullOrEmpty(r.Diameter))
            .GroupBy(r => (r.LayoutNo ?? "", r.LSSheet ?? "", r.SpoolNumber ?? ""),
                     StringTupleComparer.Instance);

        var maxDiaLookup = new Dictionary<(string, string, string), (double MaxDia, string MaxDiaStr, string Schedule)>();

        foreach (var g in groups)
        {
            double maxDia = double.MinValue;
            string maxDiaStr = "";
            string maxDiaSch = "";

            foreach (var r in g)
            {
                if (TryParseDiameter(r.Diameter, out var d) && d > maxDia)
                {
                    maxDia = d;
                    maxDiaStr = r.Diameter!;
                    maxDiaSch = r.Schedule ?? "";
                }
            }

            if (maxDia > double.MinValue)
                maxDiaLookup[g.Key] = (maxDia, maxDiaStr, maxDiaSch);
        }

        foreach (var row in rows)
        {
            if (row.WeldType is not ("BR" or "LET"))
                continue;

            if (!string.IsNullOrEmpty(row.OLDiameter) && row.OLDiameter != "##")
                continue; // already populated (e.g. from PCF branch data)

            var key = (row.LayoutNo ?? "", row.LSSheet ?? "", row.SpoolNumber ?? "");

            if (maxDiaLookup.TryGetValue(key, out var best))
            {
                row.OLDiameter = best.MaxDiaStr;
                row.OLSchedule = best.Schedule;
            }
            else
            {
                // Last-resort fallback: use the row's own diameter
                row.OLDiameter = row.Diameter ?? "##";
                row.OLSchedule = row.Schedule ?? "##";
                AppendRowRemark(row, warnings,
                    $"OL Diameter required for {row.WeldType} weld type");
            }
        }
    }

    /// <summary>
    /// Sets the ISO Dia. In. (<see cref="JointRecordAnalysisRow.LSDiameter"/>) for every row
    /// to the maximum Dia. In. value within the same (LayoutNo, LSSheet) group.
    /// </summary>
    private static void PopulateIsoDiameterFromGroupMax(List<JointRecordAnalysisRow> rows)
    {
        // Build lookup: (LayoutNo upper, LSSheet upper) → max Dia string
        var maxDiaLookup = new Dictionary<(string, string), (double MaxDia, string MaxDiaStr)>();

        foreach (var row in rows)
        {
            if (string.IsNullOrEmpty(row.Diameter)) continue;
            if (!TryParseDiameter(row.Diameter, out var d)) continue;

            var key = (
                (row.LayoutNo ?? "").ToUpperInvariant(),
                (row.LSSheet ?? "").ToUpperInvariant()
            );

            if (!maxDiaLookup.TryGetValue(key, out var current) || d > current.MaxDia)
                maxDiaLookup[key] = (d, row.Diameter!);
        }

        foreach (var row in rows)
        {
            var key = (
                (row.LayoutNo ?? "").ToUpperInvariant(),
                (row.LSSheet ?? "").ToUpperInvariant()
            );

            if (maxDiaLookup.TryGetValue(key, out var best))
                row.LSDiameter = best.MaxDiaStr;
        }
    }

    /// <summary>
    /// Determines the Material field for every row by:
    /// 1. Grouping rows by (LayoutNo, LSSheet, SpoolNumber)
    /// 2. Finding the row with the largest "Dia. In." in each group
    /// 3. Taking that row's Grade A value
    /// 4. Looking up Material_tbl.MAT where MAT_GRADE matches
    /// 5. Assigning the resolved MAT value to all rows in the group
    /// </summary>
    private async Task PopulateMaterialFromGroupMaxGradeAsync(
        List<JointRecordAnalysisRow> rows, List<string> _)
    {
        if (rows.Count == 0) return;

        // Collect the Grade A from the largest-diameter row in each group
        var gradesByGroup = new Dictionary<(string, string, string), string>(
            StringTupleComparer.Instance);

        var groups = rows
            .GroupBy(r => (r.LayoutNo ?? "", r.LSSheet ?? "", r.SpoolNumber ?? ""),
                     StringTupleComparer.Instance);

        foreach (var g in groups)
        {
            double maxDia = double.MinValue;
            string bestGradeA = "";

            foreach (var r in g)
            {
                if (TryParseDiameter(r.Diameter, out var d) && d > maxDia)
                {
                    maxDia = d;
                    bestGradeA = r.GradeA ?? "";
                }
            }

            if (!string.IsNullOrEmpty(bestGradeA))
                gradesByGroup[g.Key] = bestGradeA;
        }

        if (gradesByGroup.Count == 0) return;

        // Batch-load MAT values for all distinct grades from Material_tbl
        var distinctGrades = gradesByGroup.Values
            .Where(g => !string.IsNullOrEmpty(g))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        Dictionary<string, string> gradeToMat;
        try
        {
            var matRows = await _context.Material_tbl.AsNoTracking()
                .Where(m => distinctGrades.Contains(m.MAT_GRADE))
                .Select(m => new { m.MAT_GRADE, m.MAT })
                .ToListAsync();

            // First match per grade (case-insensitive)
            gradeToMat = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var mr in matRows)
            {
                gradeToMat.TryAdd(mr.MAT_GRADE, mr.MAT);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to look up Material_tbl for grade-to-MAT mapping");
            return;
        }

        // Apply the resolved material to every row in each group
        foreach (var row in rows)
        {
            var key = (row.LayoutNo ?? "", row.LSSheet ?? "", row.SpoolNumber ?? "");

            if (gradesByGroup.TryGetValue(key, out var grade) &&
                gradeToMat.TryGetValue(grade, out var mat) &&
                !string.IsNullOrEmpty(mat))
            {
                row.Material = mat;
            }
        }
    }

    private static bool TryParseDiameter(string? value, out double result)
    {
        result = 0;
        if (string.IsNullOrWhiteSpace(value)) return false;
        return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
    }

    /// <summary>Equality comparer for the (LayoutNo, Sheet, Spool) group key.</summary>
    private sealed class StringTupleComparer : IEqualityComparer<(string, string, string)>
    {
        public static readonly StringTupleComparer Instance = new();
        public bool Equals((string, string, string) x, (string, string, string) y)
            => string.Equals(x.Item1, y.Item1, StringComparison.OrdinalIgnoreCase)
            && string.Equals(x.Item2, y.Item2, StringComparison.OrdinalIgnoreCase)
            && string.Equals(x.Item3, y.Item3, StringComparison.OrdinalIgnoreCase);
        public int GetHashCode((string, string, string) obj)
            => HashCode.Combine(
                StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item1),
                StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item2),
                StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item3));
    }

    private static List<JointRecordAnalysisRow> CleanAndDeduplicateRows(List<JointRecordAnalysisRow> rows)
    {
        // Remove empty rows
        rows.RemoveAll(IsJointRecordRowEmpty);

        // Group by key fields and take first
        return rows
            .GroupBy(r => new { r.LayoutNo, r.WeldNumber, r.JAdd, r.LSSheet })
            .Select(g => g.First())
            .OrderBy(r => r.LayoutNo)
            .ThenBy(r => r.LSSheet)
            .ThenBy(r => r.WeldNumber, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    #endregion

    #region Helper Methods

    private static void MergeImportAndFinalWarnings(List<string> warnings, int excelCount, int finalCount)
    {
        if (warnings == null || warnings.Count == 0) return;
        if (excelCount <= 0 || finalCount <= 0) return;

        var importMsg = $"Successfully imported {excelCount} rows from Excel template.";
        var finalMsg = $"Final analysis contains {finalCount} unique joint records.";

        var importIndex = warnings.FindIndex(w => string.Equals(w, importMsg, StringComparison.OrdinalIgnoreCase));
        var finalIndex = warnings.FindIndex(w => string.Equals(w, finalMsg, StringComparison.OrdinalIgnoreCase));

        if (importIndex < 0 || finalIndex < 0 || importIndex == finalIndex) return;

        var merged = $"{importMsg} {finalMsg}";

        // Keep the earliest occurrence slot for the merged message
        var targetIndex = Math.Min(importIndex, finalIndex);
        var removeIndex = Math.Max(importIndex, finalIndex);

        warnings[targetIndex] = merged;
        warnings.RemoveAt(removeIndex);
    }

    private static string GetRowLabel(JointRecordAnalysisRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.WeldNumber))
            return $"Row {row.WeldNumber}";
        if (!string.IsNullOrWhiteSpace(row.LayoutNo))
            return $"Layout {row.LayoutNo}";
        return "Row";
    }

    private static void AppendRowRemark(JointRecordAnalysisRow row, List<string> warnings, string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;

        // Add to warnings list (shown in the header messages area)
        // but do NOT append to row.Remarks — validation notes are shown
        // via row/cell styling and tooltips, not in the System Remarks column.
        var label = GetRowLabel(row);
        warnings.Add($"{label}: {message}");
    }

    private static void ExtractMetadataFromFileName(string fileName, DrawingMetadata metadata)
    {
        var name = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;

        // Try pattern: 890-60C-10099-003_Rev00 or 890-P-11645-001_Rev00_0001_01F
        var pattern = @"^(?<unit>\d+)[-_](?<fluid>[A-Z0-9]{1,4})[-_](?<line>\d{4,5})[-_](?<sheet>\d{3})_Rev(?<rev>[A-Za-z0-9]+(?:_[A-Za-z0-9]+)*)";
        var match = Regex.Match(name, pattern, RegexOptions.IgnoreCase);
        if (match.Success)
        {
            metadata.PlantNo = match.Groups["unit"].Value;
            metadata.Fluid = match.Groups["fluid"].Value;
            metadata.LineNo = match.Groups["line"].Value;
            metadata.SheetNo = match.Groups["sheet"].Value;
            var fnRevTag = NormalizeRevisionTagForEquivalence(match.Groups["rev"].Value);
            metadata.LRev = fnRevTag;
            metadata.LSRev = fnRevTag;
            metadata.DrawingNo = $"{metadata.Fluid}-{metadata.LineNo}";
            return;
        }

        // Try pattern without revision: 890-60C-10099-003 or 890-P-11645-001
        pattern = @"^(?<unit>\d+)[-_](?<fluid>[A-Z0-9]{1,4})[-_](?<line>\d{4,5})[-_](?<sheet>\d{3})";
        match = Regex.Match(name, pattern, RegexOptions.IgnoreCase);
        if (match.Success)
        {
            metadata.PlantNo = match.Groups["unit"].Value;
            metadata.Fluid = match.Groups["fluid"].Value;
            metadata.LineNo = match.Groups["line"].Value;
            metadata.SheetNo = match.Groups["sheet"].Value;
            metadata.DrawingNo = $"{metadata.Fluid}-{metadata.LineNo}";
            return;
        }

        // Fallback patterns
        var patterns = new[]
        {
            @"^(?<unit>\d+)[-_](?<fluid>[A-Z0-9]{1,4})[-_](?<line>\d{4,5})",
            @"(?<fluid>[A-Z0-9]{1,4})[-_](?<line>\d{4,5})",
            @"(?<line>\d{4,5})"
        };

        foreach (var pat in patterns)
        {
            match = Regex.Match(name, pat, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                if (match.Groups["unit"].Success)
                    metadata.PlantNo = match.Groups["unit"].Value;
                if (match.Groups["fluid"].Success)
                    metadata.Fluid = match.Groups["fluid"].Value;
                if (match.Groups["line"].Success)
                    metadata.LineNo = match.Groups["line"].Value;

                if (!string.IsNullOrEmpty(metadata.Fluid) && !string.IsNullOrEmpty(metadata.LineNo))
                    metadata.DrawingNo = $"{metadata.Fluid}-{metadata.LineNo}";

                break;
            }
        }
    }

    private static void FillMissingMetadata(DrawingMetadata metadata)
    {
        // Try to infer DrawingNo if we have components
        if (string.IsNullOrEmpty(metadata.DrawingNo) &&
            !string.IsNullOrEmpty(metadata.Fluid) &&
            !string.IsNullOrEmpty(metadata.LineNo))
        {
            metadata.DrawingNo = $"{metadata.Fluid}-{metadata.LineNo}";
        }

        // Try to infer ISO if we have components
        if (string.IsNullOrEmpty(metadata.Iso) &&
            !string.IsNullOrEmpty(metadata.Fluid) &&
            !string.IsNullOrEmpty(metadata.LineNo) &&
            !string.IsNullOrEmpty(metadata.LineClass))
        {
            metadata.Iso = $"{metadata.Fluid}-{metadata.LineNo}-{metadata.LineClass}";
        }

        // Try to infer LineClass from ISO
        if (string.IsNullOrEmpty(metadata.LineClass) && !string.IsNullOrEmpty(metadata.Iso))
        {
            var parts = metadata.Iso.Split('-');
            if (parts.Length >= 3)
            {
                metadata.LineClass = parts[2];
            }
        }

        // Infer material from line class if available
        if (string.IsNullOrEmpty(metadata.Material) && !string.IsNullOrEmpty(metadata.LineClass))
        {
            metadata.Material = InferMaterialFromLineClass(metadata.LineClass);
        }

        // Default material if still empty
        if (string.IsNullOrEmpty(metadata.Material))
        {
            metadata.Material = "CS";
        }
    }

    private static string InferMaterialFromLineClass(string lineClass)
    {
        if (string.IsNullOrEmpty(lineClass)) return "";
        var upper = lineClass.ToUpperInvariant();

        // Common line class patterns: 3CS1P03 → CS, 1SS2U01 → SS
        if (upper.Contains("CS")) return "CS";
        if (upper.Contains("SS")) return "SS";
        if (upper.Contains("AS")) return "AS";

        return "";
    }

    /// <summary>
    /// Normalizes PDF weld type abbreviations to standard DFR weld type codes.
    /// </summary>
    private static string NormalizePdfWeldType(string weldType)
    {
        var upper = weldType.Trim().ToUpperInvariant();
        return upper switch
        {
            "SOW" => "SW",
            _ => upper
        };
    }

    /// <summary>
    /// Converts fractional weld size strings to their decimal equivalents.
    /// Handles patterns: "1.1/2" → "1.5", "3/4" → "0.75", "1/2" → "0.5",
    /// "1/4" → "0.25", "3/8" → "0.375", "2.1/2" → "2.5", "1 ½" → "1.5".
    /// Whole numbers and existing decimals pass through unchanged.
    /// </summary>
    private static string NormalizeFractionalSize(string? size)
    {
        if (string.IsNullOrWhiteSpace(size)) return "";

        var s = size.Trim().TrimEnd('"', '\u201D', '\u2033', '\'', '\u2019');

        // Handle unicode fraction characters: ¼ ½ ¾
        s = s.Replace("\u00BD", ".5").Replace("\u00BC", ".25").Replace("\u00BE", ".75");

        // Pattern: whole.numerator/denominator (e.g. 1.1/2, 2.1/2, 2.3/4)
        var m = FractionDotRegex().Match(s);
        if (m.Success && double.TryParse(m.Groups[3].Value, out var d1) && d1 != 0)
        {
            var whole = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var num = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            return FormatDecimalSize(whole + num / d1);
        }

        // Pattern: whole space numerator/denominator (e.g. 1 1/2, 2 1/2)
        m = FractionSpaceRegex().Match(s);
        if (m.Success && double.TryParse(m.Groups[3].Value, out var d2) && d2 != 0)
        {
            var whole = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var num = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            return FormatDecimalSize(whole + num / d2);
        }

        // Pattern: numerator/denominator only (e.g. 1/2, 3/4, 3/8, 1/4)
        m = FractionSimpleRegex().Match(s);
        if (m.Success && double.TryParse(m.Groups[2].Value, out var d3) && d3 != 0)
        {
            var num = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            return FormatDecimalSize(num / d3);
        }

        // Already a whole number or decimal – pass through
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            return s;

        return s;
    }

    private static string FormatDecimalSize(double value)
    {
        if (value == Math.Floor(value))
            return ((int)value).ToString(CultureInfo.InvariantCulture);
        return value.ToString("G", CultureInfo.InvariantCulture);
    }

    private static List<JointRecordAnalysisRow> ReadJointRecordExcelWithValidation(Stream excelStream, List<string> warnings)
    {
        var rows = new List<JointRecordAnalysisRow>();

        using var workbook = new XLWorkbook(excelStream);
        var ws = workbook.Worksheets.First();

        var headerMap = BuildHeaderMap(ws);
        if (headerMap.Count == 0)
        {
            warnings.Add("Could not detect headers in Excel template.");
            return rows;
        }

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        for (var r = 2; r <= lastRow; r++)
        {
            var row = new JointRecordAnalysisRow
            {
                UnitNumber = GetCell(ws, r, headerMap, "UNIT_NUMBER"),
                Fluid = GetCell(ws, r, headerMap, "FLUID"),
                LineNo = GetCell(ws, r, headerMap, "LINE_NO"),
                LayoutNo = GetCell(ws, r, headerMap, "LAYOUT_NO"),
                LineClass = GetCell(ws, r, headerMap, "LINE_CLASS"),
                ISO = GetCell(ws, r, headerMap, "ISO"),
                LRev = GetCell(ws, r, headerMap, "L_REV"),
                LSRev = GetCell(ws, r, headerMap, "L_REV"),
                LSSheet = GetCell(ws, r, headerMap, "LS_SHEET"),
                Material = GetCell(ws, r, headerMap, "MATERIAL"),
                LSDiameter = GetCell(ws, r, headerMap, "LS_DIAMETER"),
                Location = GetCell(ws, r, headerMap, "LOCATION"),
                WeldNumber = GetCell(ws, r, headerMap, "WELD_NUMBER"),
                JAdd = GetCell(ws, r, headerMap, "J_ADD"),
                WeldType = GetCell(ws, r, headerMap, "WELD_TYPE"),
                SpoolNumber = GetCell(ws, r, headerMap, "SPOOL_NUMBER"),
                Diameter = GetCell(ws, r, headerMap, "DIAMETER"),
                SpDia = GetCell(ws, r, headerMap, "SP_DIA"),
                Schedule = GetCell(ws, r, headerMap, "SCHEDULE"),
                OLDiameter = GetCell(ws, r, headerMap, "OL_DIAMETER"),
                OLSchedule = GetCell(ws, r, headerMap, "OL_SCHEDULE"),
                MaterialA = GetCell(ws, r, headerMap, "MATERIAL_A"),
                MaterialB = GetCell(ws, r, headerMap, "MATERIAL_B"),
                GradeA = GetCell(ws, r, headerMap, "GRADE_A"),
                GradeB = GetCell(ws, r, headerMap, "GRADE_B"),
                LSScope = GetCell(ws, r, headerMap, "LS_SCOPE")
            };

            // Map Delete / Cancel column to boolean flags.
            // Both flags are raised as a "delete-or-cancel requested" signal;
            // CommitJointRecord resolves which one to persist based on FITUP_DATE.
            var deleteCancel = GetCell(ws, r, headerMap, "DELETE_CANCEL");
            if (!string.IsNullOrWhiteSpace(deleteCancel))
            {
                var dcNorm = deleteCancel.Trim();
                if (dcNorm.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
                    || dcNorm.Equals("Deleted", StringComparison.OrdinalIgnoreCase)
                    || dcNorm.Equals("Cancelled", StringComparison.OrdinalIgnoreCase))
                {
                    row.Deleted = true;
                    row.Cancelled = true;
                }
            }

            // Treat "0" Schedule from Excel as blank (numeric-formatted cells
            // return "0" via GetValue<string>() even when visually empty)
            if (row.Schedule == "0") row.Schedule = null;

            if (IsJointRecordRowEmpty(row))
            {
                continue;
            }

            var errors = row.Validate();
            if (errors.Count > 0)
            {
                warnings.AddRange(errors.Select(e => $"Row {r}: {e}"));
            }

            rows.Add(row);
        }

        return rows;
    }

    private static Dictionary<string, int> BuildHeaderMap(IXLWorksheet ws)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headerRow = ws.FirstRowUsed();
        if (headerRow == null) return map;

        foreach (var cell in headerRow.CellsUsed())
        {
            var raw = cell.GetString();
            if (string.IsNullOrWhiteSpace(raw)) continue;

            var normalized = Normalize(raw);
            if (string.IsNullOrWhiteSpace(normalized)) continue;

            // 1) Exact template matches
            foreach (var templateCol in JointRecordTemplateColumns)
            {
                var normalizedTemplate = Normalize(templateCol);
                if (normalized == normalizedTemplate)
                {
                    map[templateCol] = cell.Address.ColumnNumber;
                    break;
                }
            }

            // 2) Exact alias matches
            if (!map.ContainsValue(cell.Address.ColumnNumber) &&
                JointRecordHeaderAliases.TryGetValue(normalized, out var target))
            {
                map[target] = cell.Address.ColumnNumber;
            }
        }

        return map;
    }

    private static string? GetCell(IXLWorksheet ws, int rowNumber, Dictionary<string, int> map, string key)
    {
        if (map.TryGetValue(key, out var col))
        {
            var value = ws.Cell(rowNumber, col).GetValue<string>();
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
        return null;
    }

    private static string Normalize(string? header)
    {
        if (string.IsNullOrWhiteSpace(header)) return string.Empty;
        var value = NonAlphaNumericRegex().Replace(header, string.Empty);
        return value.ToUpperInvariant();
    }

    private static void BackfillFromIso(List<JointRecordAnalysisRow> rows, bool hasExcelSource = false)
    {
        if (rows == null || rows.Count == 0) return;

        foreach (var row in rows)
        {
            var isoVal = row.ISO;
            if (string.IsNullOrWhiteSpace(isoVal)) continue;

            var match = JointRecordIsoRegex().Matches(isoVal)
                .Where(m => m.Success && HasIsoSeparators(m.Value))
                .OrderByDescending(m => m.Groups["cls"].Value.Length)
                .FirstOrDefault();
            if (match == null || !match.Success) continue;

            if (match.Groups["cls"].Success)
            {
                var clsFromIso = match.Groups["cls"].Value;
                if (ShouldReplaceLineClass(row.LineClass, clsFromIso))
                {
                    row.LineClass = clsFromIso;
                }
            }

            if (string.IsNullOrWhiteSpace(row.LSDiameter) && match.Groups["dia"].Success)
                row.LSDiameter = match.Groups["dia"].Value;

            if (string.IsNullOrWhiteSpace(row.Fluid) && match.Groups["fluid"].Success)
                row.Fluid = match.Groups["fluid"].Value;

            // Skip auto-deriving Line No from ISO when the source is an Excel
            // template — update only non-empty cells.
            if (!hasExcelSource && string.IsNullOrWhiteSpace(row.LineNo) && match.Groups["line"].Success)
                row.LineNo = match.Groups["line"].Value;
        }
    }

    /// <summary>
    /// Normalizes revision fields on loaded rows so that both LRev and LSRev
    /// carry the compound revision tag ("primary_secondary") exactly as the
    /// Drawings controller's <c>NormalizeRevisionTagForEquivalence</c> produces.
    /// This keeps the JointRecord display/edit/save cycle identical to Drawings.
    /// </summary>
    private static void NormalizeRevisionFields(List<JointRecordAnalysisRow> rows)
    {
        if (rows == null || rows.Count == 0) return;

        foreach (var row in rows)
        {
            // Build the compound tag from whatever DB values we have.
            // Priority: LSRev (may already be compound from LS_REV / DFR_REV),
            //           then fall back to LRev.
            var rawLsRev = (row.LSRev ?? "").Trim().Trim('-', '_');
            var rawLRev  = (row.LRev  ?? "").Trim().Trim('-', '_');

            // Pick the best source for building the compound tag
            string source;
            if (!string.IsNullOrWhiteSpace(rawLsRev))
                source = rawLsRev;
            else if (!string.IsNullOrWhiteSpace(rawLRev))
                source = rawLRev;
            else
            {
                row.LRev  = null;
                row.LSRev = null;
                continue;
            }

            // Normalize to compound format using the same logic as Drawings
            var compound = NormalizeRevisionTagForEquivalence(source);

            // Store the compound tag in both fields so the UI can display it
            // regardless of whether the project is Line or Sheet mode.
            row.LRev  = compound;
            row.LSRev = compound;
        }
    }

    private static void NormalizeLineNumberFromLayout(List<JointRecordAnalysisRow> rows, bool hasExcelSource = false)
    {
        if (rows == null || rows.Count == 0) return;

        // When the source is an Excel template, do not auto-derive Line No
        // from Layout No — update only non-empty cells.
        if (hasExcelSource) return;

        foreach (var row in rows)
        {
            if (!string.IsNullOrWhiteSpace(row.LineNo)) continue;
            if (string.IsNullOrWhiteSpace(row.LayoutNo)) continue;

            var match = DrawingLineNumberRegex().Match(row.LayoutNo);
            if (match.Success)
            {
                row.LineNo = match.Groups["line"].Value;
                continue;
            }

            var parts = row.LayoutNo.Split('-', '_', StringSplitOptions.RemoveEmptyEntries);
            var tail = parts.LastOrDefault();
            if (!string.IsNullOrWhiteSpace(tail) && int.TryParse(tail, out _))
            {
                row.LineNo = tail;
            }
        }
    }

    private static void PopulateIsoFromLayoutAndClass(List<JointRecordAnalysisRow> rows)
    {
        if (rows == null || rows.Count == 0) return;

        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.LayoutNo) || string.IsNullOrWhiteSpace(row.LineClass)) continue;
            if (!IsPlausibleLineClass(row.LineClass)) continue;

            row.ISO = $"{row.LayoutNo}-{row.LineClass}";
        }
    }

    private static void ValidateAndCorrectJointRecordRow(JointRecordAnalysisRow row, List<string> warnings, bool hasExcelSource = false)
    {
        if (!string.IsNullOrEmpty(row.LayoutNo))
        {
            if (!ValidLayoutNoRegex().IsMatch(row.LayoutNo))
            {
                AppendRowRemark(row, warnings, $"Invalid LayoutNo format: {row.LayoutNo}");
                var parts = row.LayoutNo.Split('-');
                if (parts.Length >= 2)
                {
                    row.LayoutNo = $"{parts[0]}-{parts[1]}";
                }
            }
        }

        if (!string.IsNullOrEmpty(row.WeldNumber))
        {
            row.WeldNumber = row.WeldNumber.Trim().ToUpperInvariant();

            if (!ValidWeldNoRegex().IsMatch(row.WeldNumber))
            {
                AppendRowRemark(row, warnings, $"Invalid weld number format: {row.WeldNumber}");
            }
        }

        // J-Add is now mandatory
        if (string.IsNullOrEmpty(row.JAdd))
        {
            AppendRowRemark(row, warnings, "J-Add is required");
            row.JAdd = "NEW";  // Set default to prevent downstream errors
        }

        // Weld Type: only default to "BW" when NOT from an Excel template.
        // When the source is Excel, update only non-empty cells.
        if (string.IsNullOrEmpty(row.WeldType))
        {
            if (!hasExcelSource)
                row.WeldType = "BW";
        }
        else if (!ValidWeldTypeRegex().IsMatch(row.WeldType))
        {
            AppendRowRemark(row, warnings, $"Invalid weld type: {row.WeldType}. Defaulting to BW");
            row.WeldType = "BW";
        }
    }

    private static void AutoPopulateMissingFields(JointRecordAnalysisRow row)
    {
        // LSDiameter is populated by PopulateIsoDiameterFromGroupMax as
        // Max(Dia. In.) per (LayoutNo, Sheet) — no per-row fallback needed.

        // SpDia is populated by PopulateSpoolDiameterFromGroup using cutting list
        // or spool group inference. Only fall back to Diameter for WS welds that
        // still have no SpDia — field welds (YFW/FW) have no spool, so SpDia stays empty.
        if (string.IsNullOrEmpty(row.SpDia) && !string.IsNullOrEmpty(row.Diameter)
            && (row.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase))
            row.SpDia = row.Diameter;

        if (string.IsNullOrEmpty(row.ISO) || row.ISO == "##")
        {
            if (!string.IsNullOrEmpty(row.Fluid) && !string.IsNullOrEmpty(row.LineNo) && !string.IsNullOrEmpty(row.LineClass))
                row.ISO = $"{row.Fluid}-{row.LineNo}-{row.LineClass}";
            else if (!string.IsNullOrEmpty(row.Fluid) && !string.IsNullOrEmpty(row.LineNo))
                row.ISO = $"{row.Fluid}-{row.LineNo}";
        }

        if (string.IsNullOrEmpty(row.LayoutNo) && !string.IsNullOrEmpty(row.Fluid) && !string.IsNullOrEmpty(row.LineNo))
            row.LayoutNo = $"{row.Fluid}-{row.LineNo}";
    }

    private async Task<DrawingMetadata?> ParsePdfDrawingAsync(Stream stream, string fileName, List<string> warnings)
    {
        var metadata = new DrawingMetadata();

        if (!HasPdfHeader(stream))
        {
            warnings.Add("Uploaded file is not a valid PDF.");
            return null;
        }

        var cancellationToken = HttpContext?.RequestAborted ?? CancellationToken.None;
        var text = await ExtractPdfTextAsync(stream, fileName, warnings, cancellationToken);
        if (string.IsNullOrWhiteSpace(text))
        {
            ExtractMetadataFromFileName(fileName, metadata);
            return metadata;
        }

        // Parse structured weld data from the extracted text
        ParseWeldDataFromText(text, metadata, warnings);

        // Supplement with filename-based metadata for any fields still missing
        ExtractMetadataFromFileName(fileName, metadata);
        FillMissingMetadata(metadata);

        if (metadata.Welds.Count > 0)
        {
            warnings.Add($"Extracted {metadata.Welds.Count} weld records from PDF text.");
        }

        return metadata;
    }

    private static bool HasPdfHeader(Stream stream)
    {
        if (!stream.CanSeek) return false;

        var buffer = new byte[4];
        var origin = stream.Position;
        var read = stream.Read(buffer, 0, buffer.Length);
        stream.Position = origin;
        return read == 4 && buffer[0] == 0x25 && buffer[1] == 0x50 && buffer[2] == 0x44 && buffer[3] == 0x46;
    }

    private async Task<string> ExtractPdfTextAsync(Stream stream, string fileName, List<string> warnings, CancellationToken cancellationToken)
    {
        var sb = new StringBuilder();
        int blankPages = 0;
        int totalPages = 0;
        bool usedSpatialGrouping = false;

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            stream.Position = 0;
            using var pdf = PdfDocument.Open(stream);
            foreach (var page in pdf.GetPages())
            {
                totalPages++;
                if (totalPages > MaxPdfPages)
                {
                    warnings.Add($"PDF exceeds {MaxPdfPages} page limit. Only first {MaxPdfPages} pages were processed.");
                    break;
                }
                cancellationToken.ThrowIfCancellationRequested();

                // For engineering drawings, prefer spatial word grouping
                // over page.Text which often produces garbled output
                var words = page.GetWords().ToList();
                if (words.Count > 0)
                {
                    var spatialText = GroupWordsIntoLines(words);
                    if (!string.IsNullOrWhiteSpace(spatialText))
                    {
                        sb.AppendLine(spatialText);
                        usedSpatialGrouping = true;
                        continue;
                    }
                }

                // Fallback to page.Text
                var text = page.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    blankPages++;
                    continue;
                }

                sb.AppendLine(text);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read PDF text for {File}", fileName);
        }

        var content = sb.ToString();

        // Determine if the extracted text is useful
        // (contains ISO patterns or weld-related keywords)
        bool hasUsefulContent = !string.IsNullOrWhiteSpace(content) &&
            (PdfWeldListHeaderRegex().IsMatch(content) ||
             PdfCuttingListHeaderRegex().IsMatch(content) ||
             PdfWeldLineRegex().IsMatch(content) ||
             PdfWeldTableRowRegex().IsMatch(content) ||
             PdfWeldMapRowRegex().IsMatch(content) ||
             PdfIsoDesignationRegex().IsMatch(content) ||
             JointRecordIsoRegex().IsMatch(content));

        // Try Azure OCR if: no text at all, all pages blank, or text lacks useful patterns
        if (string.IsNullOrWhiteSpace(content) ||
            (totalPages > 0 && blankPages == totalPages) ||
            !hasUsefulContent)
        {
            if (_ocrService != null)
            {
                warnings.Add("PDF text extraction insufficient. Attempting Azure OCR...");
                try
                {
                    stream.Position = 0;
                    var ocrText = await _ocrService.ExtractTextAsync(stream, fileName, cancellationToken);
                    if (!string.IsNullOrWhiteSpace(ocrText))
                    {
                        warnings.Add("Azure OCR text extraction succeeded.");
                        return NormalizePdfText(ocrText);
                    }
                    else
                    {
                        warnings.Add("Azure OCR returned no text. Ensure the PDF contains readable weld map content.");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Azure OCR fallback failed for {File}", fileName);
                    warnings.Add("Azure OCR fallback failed. Provide an Excel template instead.");
                }
            }
            else if (string.IsNullOrWhiteSpace(content) || (totalPages > 0 && blankPages == totalPages))
            {
                warnings.Add("PDF has no extractable text and Azure OCR is not configured. Provide an Excel template.");
            }
        }

        if (usedSpatialGrouping && !string.IsNullOrWhiteSpace(content))
        {
            warnings.Add("Used spatial word grouping for engineering drawing PDF.");
        }

        return string.IsNullOrWhiteSpace(content) ? string.Empty : NormalizePdfText(content);
    }

    /// <summary>
    /// Groups PdfPig words into lines by sorting by Y-coordinate (top to bottom),
    /// then X-coordinate (left to right), grouping words on similar Y positions
    /// into the same line. This reconstructs tabular structure from engineering drawings.
    /// </summary>
    private static string GroupWordsIntoLines(List<Word> words)
    {
        if (words == null || words.Count == 0) return string.Empty;

        // Sort by Y descending (PDF coords: bottom=0, top=max), then X ascending
        var sorted = words
            .OrderByDescending(w => Math.Round(w.BoundingBox.Bottom, 1))
            .ThenBy(w => w.BoundingBox.Left)
            .ToList();

        var lines = new List<List<Word>>();
        List<Word> currentLine = new() { sorted[0] };
        double currentY = sorted[0].BoundingBox.Bottom;

        // Estimate typical word height for line-break threshold
        double avgHeight = sorted.Average(w => w.BoundingBox.Height);
        double yTolerance = Math.Max(avgHeight * 0.6, 2.0);

        for (int i = 1; i < sorted.Count; i++)
        {
            var word = sorted[i];
            if (Math.Abs(word.BoundingBox.Bottom - currentY) <= yTolerance)
            {
                // Same line
                currentLine.Add(word);
            }
            else
            {
                // New line
                lines.Add(currentLine);
                currentLine = new List<Word> { word };
                currentY = word.BoundingBox.Bottom;
            }
        }
        lines.Add(currentLine);

        // Build text lines, inserting appropriate spacing
        var sb = new StringBuilder();
        foreach (var line in lines)
        {
            var lineWords = line.OrderBy(w => w.BoundingBox.Left).ToList();
            for (int i = 0; i < lineWords.Count; i++)
            {
                if (i > 0)
                {
                    // Calculate gap between words
                    double gap = lineWords[i].BoundingBox.Left - lineWords[i - 1].BoundingBox.Right;
                    double avgCharWidth = lineWords[i - 1].BoundingBox.Width /
                        Math.Max(lineWords[i - 1].Text.Length, 1);

                    // Use tab-like separator for large gaps (table columns)
                    if (gap > avgCharWidth * 3)
                        sb.Append("  ");
                    else
                        sb.Append(' ');
                }
                sb.Append(lineWords[i].Text);
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static string NormalizePdfText(string content)
    {
        var lines = content.Split('\n');
        var normalized = lines.Select(l => l.TrimEnd());
        return string.Join('\n', normalized);
    }

    [GeneratedRegex(@"\b(?:WELD|WLD)\s*(?:NO\.?|#|NUM(?:BER)?)?\s*[:.]?\s*(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldLineRegex();

    [GeneratedRegex(@"SPOOL\s*(?:NO\.?|#)?\s*[:.]?\s*(?<spool>\d{1,4})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfSpoolRegex();

    [GeneratedRegex(@"(?:DIA(?:METER)?|SIZE)\s*[:.]?\s*(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)[\s""]", RegexOptions.IgnoreCase)]
    private static partial Regex PdfDiameterRegex();

    [GeneratedRegex(@"SCH(?:EDULE)?\s*[:.]?\s*(?<sch>\d{2,3}S?|[XSXL]{2,3})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfScheduleRegex();

    [GeneratedRegex(@"SHEET\s*(?:NO\.?|#)?\s*[:.]?\s*(?<sheet>\d{1,3})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfSheetRegex();

    [GeneratedRegex(@"REV(?:ISION)?\s*[:.]?\s*(?<rev>[A-Z0-9]{1,4})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfRevisionRegex();

    [GeneratedRegex(@"UNIT\s*(?:NO\.?|#|NUM(?:BER)?)?\s*[:.]?\s*(?<unit>\d{2,5})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfUnitRegex();

    [GeneratedRegex(@"MATERIAL\s*[:.]?\s*(?<mat>[A-Z0-9][-A-Z0-9/]{1,20})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfMaterialRegex();

    // Weld list table row: weld number, type, spool, size, schedule, component names, grades
    // Matches lines like: 01 BW 01 10 40 PIPE ELBOW A106-B A234-WPB, 01CS1 BW 01 10 40 PIPE ELBOW
    [GeneratedRegex(@"(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)\s+(?<spool>[\d/]{1,7})\s+(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s+(?<sch>\d{1,3}S?|[XSXL]{2,3}|0)\s+(?<compA>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)\s+(?<compB>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)?\s*(?<gradeA>A\d{2,4}[-]?[A-Z0-9]*)?\s*(?<gradeB>A\d{2,4}[-]?[A-Z0-9]*)?", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldTableRowRegex();

    // Weld list row with explicit Location column before Dia/Schedule and with optional spool
    [GeneratedRegex(@"(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<loc>WS|FW|YFW|YWS|SW)\s+(?<spool>SP?\d{1,4})?\s*(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s+(?<sch>\d{1,3}S?|[XSXL]{2,3}|0)\s+(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)\s+(?<compA>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)\s+(?<compB>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)\s+(?<gradeA>A\d{2,4}[-]?[A-Z0-9]*)\s*(?<gradeB>A\d{2,4}[-]?[A-Z0-9]*)?", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldTableRowWithLocationRegex();

    // Saudi Aramco weld list table row: NO | LOC | WELD SIZE | WELD TYPE | SPOOL | [DIA] | SCH | MATL A | MATL B | GRADE A | GRADE B
    // Column order has TYPE before SPOOL/SCH (unlike PdfWeldTableRowWithLocationRegex which has TYPE after SCH).
    // Flexibly handles 1-3 numeric columns (spool, optional dia, schedule) between weld type and material columns.
    // Numeric tokens can be digits, fractions, or "--" for empty fields.
    [GeneratedRegex(@"(?<!\w)(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<loc>WS|FW|YFW|YWS|SW)\s+(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s*[""\u201C\u201D\u2033\u2032'`]?\s+(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)\s+(?<nums>(?:(?:[\d/.]+|--)\s+){1,3})(?<compA>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)\s+(?<compB>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)(?:\s+(?<gradeA>[A-Z]\d{2,4}[-]?[A-Z0-9]*))?(?:\s+(?<gradeB>[A-Z]\d{2,4}[-]?[A-Z0-9]*))?(?:\s+(?<oldia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?))?(?:\s+(?<olsch>\d{1,3}S?|[XSXL]{2,3}))?(?:\s+(?<gradeA2>[A-Z]\d{2,4}[-]?[A-Z0-9]*))?(?:\s+(?<gradeB2>[A-Z]\d{2,4}[-]?[A-Z0-9]*))?\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex PdfSaudiWeldListRowRegex();

    // Weld table row without spool field: weld number, type, size, schedule, component names, grades
    // Matches lines like: 01 BW 10 40 PIPE ELBOW A106-B A234-WPB, 01CS1 BW 10 40 PIPE ELBOW
    [GeneratedRegex(@"(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)\s+(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s+(?<sch>\d{1,3}S?|[XSXL]{2,3}|0)\s+(?<compA>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)\s+(?<compB>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?|BOLT)?\s*(?<gradeA>A\d{2,4}[-]?[A-Z0-9]*)?\s*(?<gradeB>A\d{2,4}[-]?[A-Z0-9]*)?", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldNoSpoolTableRowRegex();

    // Simple weld-component row: weld number, type, compA, compB (no dia/sch/spool required)
    // Matches lines like: 01 BW PIPE ELBOW, 02 BW ELBOW PIPE, 01CS1 BW PIPE ELBOW
    [GeneratedRegex(@"(?<!\w)(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)\s+(?<compA>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?)\s+(?<compB>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldCompRowRegex();

    // FJ (Flange Joint) weld row: F01 FJ ... STUD BOLT ... A193
    [GeneratedRegex(@"(?<weld>F\d{1,4})\s+(?<type>FJ)\s+(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s+(?<sch>\d{1,3}|0)\s*.*?(?<grade>A\d{2,4}[-]?[A-Z0-9]*)?", RegexOptions.IgnoreCase)]
    private static partial Regex PdfFjWeldRowRegex();

    // Weld list table header detection
    [GeneratedRegex(@"WELD\s*(?:LIST|TABLE|SCHEDULE|MAP)", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldListHeaderRegex();

    // Location patterns: WS, FW, YFW, etc.
    [GeneratedRegex(@"\b(?<loc>WS|FW|YFW|YWS|SW)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfLocationRegex();

    // ISO designation in title block: e.g., ISO 12"-P-11645-3CS1P03
    [GeneratedRegex(@"ISO\s+(?:(?<isoDia>\d+)[""']?\s*[-])?\s*(?<fluid>[A-Z0-9]{1,4})[-](?<line>\d{3,5})[-](?<cls>[A-Z0-9]{3,12})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfIsoDesignationRegex();

    // Flexible weld row: weld number, optional numeric fields, then weld type
    // Matches: "01 BW", "01 10 BW", "01 10 40 BW", "F01 FJ"
    [GeneratedRegex(@"(?<!\w)(?<weld>F?\d{1,3})\s+(?:\d[\d./]*\s+){0,3}(?<type>BW|FW|SW|SOF|BR|LET|TH|FJ)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfFlexWeldRowRegex();

    // Weld map table row: weld number, location (SHOP/FIELD), weld size (diameter), weld type, optional components and remark
    // Matches lines like: "01 WS 3\" BW PIPE ELBOW", "04 YFW 1.1/2\" SW", "03 WS 3/4 BW", "01CS1 WS 2\" BW", "09A YFW 1.1/2\" SW"
    // The separator between dia and type accepts any combination of whitespace, quotes, inch marks, and degree symbols.
    // Diameter handles: "3", "3/4", "1.1/2", "1.5", "1 1/2" (OCR may use space instead of dot for mixed fractions).
    [GeneratedRegex(@"(?<!\w)(?<weld>[A-Z]?\d{1,4}[A-Z0-9]*)\s+(?<loc>WS|FW|YFW|YWS|SW)\s+(?<dia>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+|\s+\d/\d{1,2})?)[\s""\u201C\u201D\u2033\u2032\u00B0`'',.*]*\s*(?<type>BW|FW|SW|SOF|BR|LET|TH|SP|FJ)(?:\s+(?<compA>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?))?(?:\s+(?<compB>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD(?:\s*BOLT)?))?(?:\s+(?<remark>\S.*))?\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfWeldMapRowRegex();

    // Cutting list table header detection: "CUTTING LIST" or "CUT LIST"
    [GeneratedRegex(@"CUTTING\s+LIST|CUT\s+LIST", RegexOptions.IgnoreCase)]
    private static partial Regex PdfCuttingListHeaderRegex();

    // Cutting list table row:
    // PIECE NO  SPOOL NO  SIZE  LENGTH  END1  END2  IDENT
    // e.g. "1  SP01  2"  1011  PE  PE  I6149471"
    // e.g. "3  SP02  1.1/2"  2168  PE  TH  I5536657"
    [GeneratedRegex(@"(?<!\w)(?<piece>\d{1,4})\s+(?<spool>SP\d{1,4})\s+(?<size>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s*[""'\u2033\u2032]?\s+(?<length>\d{1,6})\s+(?<end1>[A-Z]{2,4})\s+(?<end2>[A-Z]{2,4})(?:\s+(?<ident>[A-Z0-9]+))?", RegexOptions.IgnoreCase)]
    private static partial Regex PdfCuttingListRowRegex();

    [GeneratedRegex(@"REV(?:ISION)?\s*[:.]?\s*(?<rev>\d{1,3}[_]\d{1,3})", RegexOptions.IgnoreCase)]
    private static partial Regex PdfCompoundRevisionRegex();

    [GeneratedRegex(@"\b(\d{1,3})\b")]
    private static partial Regex PdfSmallNumberRegex();

    // BOM / material table row: matches lines like "1 PIPE 4\" CS A106-B" or "3 ELBOW 90 LR"
    // Captures MATD_Description component names found in the drawing's Bill of Materials table.
    // Includes piping component aliases common in Saudi Aramco ISOs (BRANCH, WELDOLET, SOCKOLET, etc.)
    [GeneratedRegex(@"\b(?<desc>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|BRANCH|WELDOLET|SOCKOLET|THREADOLET|NIPOLET|LATROLET|HALF\s*COUPLING)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomComponentRegex();

    // BOM grade extraction: captures grade specifications from BOM description text.
    // Matches ASTM: "A106-B", "A106 Gr.B", "A234-WPB", "A105", "A193", "A420-WPL6"
    // Matches API: "API 5L PSL2 GR.B"
    // Matches EEMUA: "EEMUA 234", "EEMUA234" (Cu-Ni piping)
    // Matches UNS: "UNS C70600", "UNS 7060X", "UNS N08825" (alloy designations)
    // Matches BS EN: "BS EN 12451" (European standards)
    [GeneratedRegex(@"(?:API\s*5L\s*(?:PSL\d)?\s*(?:GR?\.?\s*)?(?<grade>[A-Z](?:\d+)?))|(?<grade4>EEMUA\s*\d{2,4})|(?<grade5>UNS\s*[A-Z]?\d{3,6}[A-Z]?)|(?<grade6>BS\s*EN\s*\d{3,6})|(?<grade2>A\d{2,4}[-\s]?(?:GR?\.?\s*)?[A-Z0-9]{1,8})|(?<grade3>A\d{2,4})\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomGradeRx();

    // BOM description row with component type and grade in same line.
    // Matches: "PIPE ... A106-B", "ELBOW 90 LR A234-WPB", "FLANGE ... A105", "STUD BOLT A193"
    // Also matches: "Pipe 90/10 Cu-Ni UNS 7060X ... EEMUA 234"
    // Includes piping component aliases: BRANCH, WELDOLET, SOCKOLET, THREADOLET, NIPOLET, LATROLET, HALF COUPLING
    [GeneratedRegex(@"\b(?<comp>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD\s*BOLT|BOLT|NUT|BRANCH|WELDOLET|SOCKOLET|THREADOLET|NIPOLET|LATROLET|HALF\s*COUPLING)\b.*?(?:(?:API\s*5L\s*(?:PSL\d)?\s*(?:GR?\.?\s*)?(?<grade>[A-Z]\d*))|(?<grade4>EEMUA\s*\d{2,4})|(?<grade5>UNS\s*[A-Z]?\d{3,6}[A-Z]?)|(?<grade6>BS\s*EN\s*\d{3,6})|(?<grade2>A\d{2,4}[-\s]?(?:GR?\.?\s*)?[A-Z0-9/]{1,12})|(?<grade3>A\d{2,4}))", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomComponentGradeRx();

    // Reverse BOM pattern: grade appears BEFORE component name on the same line.
    // Matches: "A106-B Seamless Pipe", "A234 WPB Elbow 90 LR", "A105 Flange Wld Neck"
    // Also matches: "EEMUA 234 ... Pipe", "UNS 7060X ... Elbow"
    [GeneratedRegex(@"(?:(?:API\s*5L\s*(?:PSL\d)?\s*(?:GR?\.?\s*)?(?<grade>[A-Z]\d*))|(?<grade4>EEMUA\s*\d{2,4})|(?<grade5>UNS\s*[A-Z]?\d{3,6}[A-Z]?)|(?<grade6>BS\s*EN\s*\d{3,6})|(?<grade2>A\d{2,4}[-\s]?(?:GR?\.?\s*)?[A-Z0-9/]{1,12})|(?<grade3>A\d{2,4})).*?\b(?<comp>PIPE|TEE|ELBOW|FLANGE|REDUCER|VALVE|OLET|NIPPLE|CAP|PAD|PLUG|COUPLING|UNION|SWAGE|HOSE|PLATE|STUD\s*BOLT|BOLT|NUT|BRANCH|WELDOLET|SOCKOLET|THREADOLET|NIPOLET|LATROLET|HALF\s*COUPLING)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomGradeComponentRx();

    // BOM section header: "FABRICATION MATERIALS" or "BILL OF MATERIALS" or "MATERIAL LIST"
    [GeneratedRegex(@"FABRICATION\s+MATERIAL|BILL\s+OF\s+MATERIAL|MATERIAL\s+LIST|PARTS?\s+LIST", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomSectionHeaderRegex();

    // BOM sub-category headers within FABRICATION MATERIALS:
    // Standalone lines like "PIPE", "FITTINGS", "FLANGES", "VALVES", "BOLTING", "GASKETS"
    // These indicate the component category for subsequent BOM items.
    // Allows up to 3 trailing characters (":", ".", digits from OCR noise).
    [GeneratedRegex(@"^\s*(?<cat>PIPE|FITTINGS?|FLANGES?|VALVES?|BOLTING|GASKETS?|SUPPORTS?|ERECTION\s+MATERIALS?)\s*\S{0,3}\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomCategoryRegex();

    // BOM item line with PT NO: a small item number followed by description text.
    // Two patterns joined by alternation:
    //   1. Double-space separated (PdfPig spatial grouping): any description content
    //   2. Single-space separated (Azure OCR): description must contain a recognized
    //      piping component word to avoid false positives with standalone numbers
    // e.g. "1  Pipe A333-6 Smls BE Seamless B36.10 - - NACE - -  I1622188  8\"  7555 MM"
    // e.g. "5 90 Elbow LR A420-WPL6 BW - Seamless B16.9 - - -  I1622149  8\"  2"
    [GeneratedRegex(@"^\s*(?<ptno>\d{1,3})\s{2,}(?<desc>.+)|^\s*(?<ptno>\d{1,3})\s+(?<desc>(?=.*\b(?:Pipe|Tee|Elbow|Flange|Reducer|Valve|Gate|Globe|Check|Ball|Olet|Nipple|Cap|Pad|Coupling|Union|Weldolet|Sockolet|Bolt|Gasket|Stud|Blind).+).+)", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomItemLineRegex();

    // Schedule from BOM description: "S-40", "SCH-40", "SCH 80", "S 40", "S40", "S-80"
    [GeneratedRegex(@"\bS(?:CH)?[-.\s]?(?<sch>\d{2,3}S?)\b", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomScheduleRegex();

    // NPS (Nominal Pipe Size) from BOM: 8", 10", 3/4", 1.1/2"
    // Inch-mark characters: " (0x22), ″ (double prime), ′ (prime), ”/“ (smart quotes), ', `
    [GeneratedRegex(@"(?<nps>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)\s*[\x22\u2033\u2032\u201D\u201C'`]", RegexOptions.None)]
    private static partial Regex PdfBomNpsRegex();

    // Fallback NPS extraction: NPS that appears after a CMDTY CODE (e.g., "I1622149  8")
    // When the inch mark is missing (spatial grouping or OCR drops it), look for a
    // pipe-sized number (1-48) following an alphanumeric commodity code pattern.
    [GeneratedRegex(@"[A-Z]\d{5,8}\s+(?<nps>\d{1,3}(?:\.\d/\d{1,2}|/\d{1,2}|\.\d+)?)(?=\s|$)", RegexOptions.None)]
    private static partial Regex PdfBomCmdtyNpsRegex();

    // BOM fitting quantity: a standalone 1-2 digit number (1–50) that appears after
    // the NPS token and is NOT followed by "MM", "M", "FT", "IN" (which indicate
    // pipe lengths, not fitting counts). This captures values like "2", "4", "10"
    // from BOM description text after the NPS has been extracted.
    [GeneratedRegex(@"(?<qty>\d{1,2})(?!\s*(?:MM\b|M\b|FT\b|IN\b|\d))", RegexOptions.IgnoreCase)]
    private static partial Regex PdfBomQtyRegex();

    // PT NO annotation on drawing body: a standalone small number (1–999), often
    // inside a square/parenthesis callout (e.g., "[12]") or prefixed with "PT".
    // Each PT NO references a BOM item. This pattern is tolerant of surrounding
    // brackets to better match OCR output from boxed callouts.
    [GeneratedRegex(@"(?<![A-Za-z0-9])(?:PT\s*(?:NO\.?|#)?\s*)?(?:[\[(]?\s*)?(?<ptno>\d{1,3})(?:\s*[\])}]?)?(?![A-Za-z0-9])", RegexOptions.IgnoreCase)]
    private static partial Regex PdfPtNoAnnotationRegex();

    private static void ParseWeldDataFromText(string text, DrawingMetadata metadata, List<string> warnings)
    {
        var rawLines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        // ---- Preprocess: merge OCR-split weld number + J-Add tokens ----
        // OCR may split "01CS1" into separate words "01 CS1". When followed by
        // a known location (WS, FW, …) or weld type (BW, SW, …), merge them
        // back into a single token so the downstream regexes can match.
        var lines = new string[rawLines.Length];
        for (int i = 0; i < rawLines.Length; i++)
        {
            lines[i] = OcrSplitWeldJAddRegex().Replace(rawLines[i], "${num}${jadd} ");
            lines[i] = OcrGluedWeldLocationRegex().Replace(lines[i], "${weld} ${loc}");
            // Split OCR-concatenated compound weld numbers: "S13T12 YFW" → "S13 T12 YFW"
            // so each weld gets its own match downstream.
            lines[i] = OcrCompoundWeldSplitRegex().Replace(lines[i], "${w1} ${w2}");
        }

        // ---- Extract header-level metadata from text ----

        // Try ISO designation pattern first (e.g., "ISO 12"-P-11645-3CS1P03")
        foreach (var line in lines)
        {
            var isoDesig = PdfIsoDesignationRegex().Match(line);
            if (isoDesig.Success)
            {
                if (string.IsNullOrEmpty(metadata.Fluid))
                    metadata.Fluid = isoDesig.Groups["fluid"].Value;
                if (string.IsNullOrEmpty(metadata.LineNo))
                    metadata.LineNo = isoDesig.Groups["line"].Value;
                if (string.IsNullOrEmpty(metadata.LineClass))
                    metadata.LineClass = isoDesig.Groups["cls"].Value;
                if (string.IsNullOrEmpty(metadata.Iso))
                    metadata.Iso = $"{isoDesig.Groups["fluid"].Value}-{isoDesig.Groups["line"].Value}-{isoDesig.Groups["cls"].Value}";
                if (string.IsNullOrEmpty(metadata.DrawingNo))
                    metadata.DrawingNo = $"{isoDesig.Groups["fluid"].Value}-{isoDesig.Groups["line"].Value}";
                // Capture diameter from ISO designation (e.g., "ISO 12"-P-..." → diameter=12)
                if (string.IsNullOrEmpty(metadata.Diameter) && isoDesig.Groups["isoDia"].Success)
                    metadata.Diameter = isoDesig.Groups["isoDia"].Value;
                break;
            }
        }

        // ISO / Layout from general pattern
        if (string.IsNullOrEmpty(metadata.Iso))
        {
            foreach (var line in lines)
            {
                var isoMatch = JointRecordIsoRegex().Match(line);
                if (isoMatch.Success && HasIsoSeparators(isoMatch.Value))
                {
                    if (string.IsNullOrEmpty(metadata.Fluid) && isoMatch.Groups["fluid"].Success)
                        metadata.Fluid = isoMatch.Groups["fluid"].Value;
                    if (string.IsNullOrEmpty(metadata.LineNo) && isoMatch.Groups["line"].Success)
                        metadata.LineNo = isoMatch.Groups["line"].Value;
                    if (string.IsNullOrEmpty(metadata.LineClass) && isoMatch.Groups["cls"].Success)
                        metadata.LineClass = isoMatch.Groups["cls"].Value;
                    if (string.IsNullOrEmpty(metadata.Iso))
                        metadata.Iso = isoMatch.Value.Trim();
                    break;
                }
            }
        }

        // Unit number
        if (string.IsNullOrEmpty(metadata.PlantNo))
        {
            foreach (var line in lines)
            {
                var m = PdfUnitRegex().Match(line);
                if (m.Success) { metadata.PlantNo = m.Groups["unit"].Value; break; }
            }
        }

        // Sheet number
        if (string.IsNullOrEmpty(metadata.SheetNo))
        {
            foreach (var line in lines)
            {
                var m = PdfSheetRegex().Match(line);
                if (m.Success) { metadata.SheetNo = m.Groups["sheet"].Value; break; }
            }
        }

        // Revision – normalize to compound tag matching Drawings controller
        if (string.IsNullOrEmpty(metadata.LRev))
        {
            foreach (var line in lines)
            {
                // Try compound revision first (e.g., 00_00)
                var compRevMatch = PdfCompoundRevisionRegex().Match(line);
                if (compRevMatch.Success)
                {
                    var pdfRevTag = NormalizeRevisionTagForEquivalence(compRevMatch.Groups["rev"].Value);
                    metadata.LRev = pdfRevTag;
                    metadata.LSRev = pdfRevTag;
                    break;
                }

                var m = PdfRevisionRegex().Match(line);
                if (m.Success)
                {
                    var pdfRevTag = NormalizeRevisionTagForEquivalence(m.Groups["rev"].Value);
                    metadata.LRev = pdfRevTag;
                    metadata.LSRev = pdfRevTag;
                    break;
                }
            }
        }

        // Material
        if (string.IsNullOrEmpty(metadata.Material))
        {
            foreach (var line in lines)
            {
                var m = PdfMaterialRegex().Match(line);
                if (m.Success) { metadata.Material = ExtractMaterial(m.Groups["mat"].Value); break; }
            }
        }

        // Header-level diameter and schedule
        if (string.IsNullOrEmpty(metadata.Diameter))
        {
            foreach (var line in lines)
            {
                var m = PdfDiameterRegex().Match(line);
                if (m.Success) { metadata.Diameter = NormalizeFractionalSize(m.Groups["dia"].Value); break; }
            }
        }

        if (string.IsNullOrEmpty(metadata.Schedule))
        {
            foreach (var line in lines)
            {
                var m = PdfScheduleRegex().Match(line);
                if (m.Success) { metadata.Schedule = m.Groups["sch"].Value; break; }
            }
        }

        // ---- Scan BOM / material table to build available component descriptions ----
        // The isometric drawing's BOM area typically lists component types (PIPE, ELBOW, TEE, etc.)
        // and their ASTM grade specifications (A106-B, A234-WPB, A105, etc.).
        // Collect them so we can use them for smarter Material A/B and Grade A/B inference later.
        //
        // Saudi Aramco ISO drawings have a "FABRICATION MATERIALS" section with sub-categories:
        //   PIPE
        //   1  Pipe A333-6 Smls BE Seamless B36.10 - - NACE - -  I1622188  8"  7555 MM  S-40
        //   FITTINGS
        //   2  90 Elbow LR A420-WPL6 BW - Seamless B16.9 - - -  I1622149  8"  2  S-40
        //   SUPPORTS
        //   4  (C01;D01;G09)  POS-0015  8"  1
        //
        // Each item has: PT NO, DESCRIPTION (with ASTM grade), CMDTY CODE, NPS, QTY
        // The PT NO appears in small squares on the drawing body near welds to indicate
        // which components each weld connects.
        var bomComponents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var bomGradeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var bomItemMap = new List<MaterialInfo>();   // PT NO → full item details
        bool inBomSection = false;
        bool inErectionSection = false;  // Skip ERECTION MATERIALS items
        int bomBlankLines = 0;
        string bomCategory = "";  // Current sub-section: PIPE, FITTINGS, FLANGES, etc.
        foreach (var line in lines)
        {
            var trimmedBom = line.Trim();

            // Detect BOM section header
            if (PdfBomSectionHeaderRegex().IsMatch(trimmedBom))
            {
                inBomSection = true;
                inErectionSection = false;
                bomBlankLines = 0;
                bomCategory = "";
            }

            // Track blank lines within BOM section
            if (inBomSection && string.IsNullOrWhiteSpace(trimmedBom))
            {
                bomBlankLines++;
                if (bomBlankLines > 8) inBomSection = false;
            }
            else
            {
                bomBlankLines = 0;
            }

            // Detect BOM sub-category headers: PIPE, FITTINGS, FLANGES, VALVES, BOLTING, SUPPORTS
            if (inBomSection)
            {
                var catMatch = PdfBomCategoryRegex().Match(trimmedBom);
                if (catMatch.Success)
                {
                    var cat = catMatch.Groups["cat"].Value.ToUpperInvariant();
                    if (cat.StartsWith("ERECTION", StringComparison.Ordinal))
                    {
                        inErectionSection = true;
                        bomCategory = "";
                    }
                    else
                    {
                        inErectionSection = false;
                        bomCategory = cat switch
                        {
                            "PIPE" => "PIPE",
                            "FITTING" or "FITTINGS" => "FITTINGS",
                            "FLANGE" or "FLANGES" => "FLANGE",
                            "VALVE" or "VALVES" => "VALVE",
                            "BOLTING" => "STUD BOLT",
                            "GASKET" or "GASKETS" => "",  // Not a weld component
                            "SUPPORT" or "SUPPORTS" => "PAD",
                            _ => ""
                        };
                    }
                    continue;
                }

                // Parse BOM item lines with PT NO (skip erection materials)
                if (!inErectionSection)
                {
                    var itemMatch = PdfBomItemLineRegex().Match(trimmedBom);
                    if (itemMatch.Success)
                    {
                        var ptNo = itemMatch.Groups["ptno"].Value;
                        var desc = itemMatch.Groups["desc"].Value;

                        // Extract component type from description or category
                        string? itemComp = null;
                        var descCompMatch = PdfBomComponentRegex().Match(desc);
                        if (descCompMatch.Success)
                            itemComp = NormalizeBomComponentName(descCompMatch.Groups["desc"].Value);
                        else if (!string.IsNullOrEmpty(bomCategory))
                            itemComp = bomCategory == "FITTINGS" ? null : bomCategory;

                        // Extract grade from description
                        string? itemGrade = null;
                        var descGradeMatch = PdfBomGradeRx().Match(desc);
                        if (descGradeMatch.Success)
                            itemGrade = ExtractBomGradeValue(descGradeMatch);

                        // Extract NPS from description
                        string? itemNps = null;
                        var npsMatch = PdfBomNpsRegex().Match(desc);
                        if (npsMatch.Success)
                            itemNps = NormalizeFractionalSize(npsMatch.Groups["nps"].Value);

                        // Fallback 1: try NPS after CMDTY code pattern (e.g., "I1622149 8")
                        // When spatial grouping or OCR drops the inch mark, the primary
                        // regex misses the NPS. CMDTY codes are alphanumeric identifiers
                        // like I1622149 that appear before the NPS column in BOM tables.
                        if (string.IsNullOrEmpty(itemNps))
                        {
                            var cmdtyNps = PdfBomCmdtyNpsRegex().Match(desc);
                            if (cmdtyNps.Success)
                            {
                                itemNps = NormalizeFractionalSize(cmdtyNps.Groups["nps"].Value);
                                // Update npsMatch position for downstream QTY extraction
                                npsMatch = cmdtyNps;
                            }
                        }

                        // Fallback 2: for fitting items with no captured NPS, default
                        // to the metadata pipe diameter. Most fittings in a spool are
                        // at the same NPS as the main pipe (elbows, tees, flanges, etc.).
                        // Skip PIPE items (they may have different lengths/sizes) and
                        // small-bore items (OLET, COUPLING) which may be branch-sized.
                        if (string.IsNullOrEmpty(itemNps)
                            && !string.IsNullOrEmpty(metadata.Diameter)
                            && !string.IsNullOrEmpty(itemComp)
                            && !itemComp.Equals("PIPE", StringComparison.OrdinalIgnoreCase)
                            && itemComp is not ("OLET" or "COUPLING" or "PLUG" or "NIPPLE"
                                or "STUD BOLT" or "PAD"))
                        {
                            itemNps = metadata.Diameter;
                        }

                        // Extract schedule from description (S-40, SCH 80, etc.)
                        string? itemSch = null;
                        var schMatch = PdfBomScheduleRegex().Match(desc);
                        if (schMatch.Success)
                            itemSch = schMatch.Groups["sch"].Value;

                        // Extract quantity from the description text after the NPS.
                        // For fittings the QTY is a small count (1-50), e.g. "8\" 2 S-40" → 2.
                        // For pipe the QTY is in MM (e.g. "8\" 7555 MM") — excluded by the
                        // negative lookahead in PdfBomQtyRegex.
                        int itemQty = 1;
                        if (npsMatch.Success)
                        {
                            // Search zone: text between NPS and schedule (or end of desc)
                            int afterNpsStart = npsMatch.Index + npsMatch.Length;
                            int searchEnd = (schMatch.Success && schMatch.Index > afterNpsStart)
                                ? schMatch.Index : desc.Length;
                            var qtyZone = desc[afterNpsStart..searchEnd].Trim();

                            var qtyMatch = PdfBomQtyRegex().Match(qtyZone);
                            if (qtyMatch.Success
                                && int.TryParse(qtyMatch.Groups["qty"].Value, out var parsedQty)
                                && parsedQty >= 1 && parsedQty <= 50)
                            {
                                itemQty = parsedQty;
                            }
                        }

                        // Build MaterialInfo entry for this BOM item
                        if (!string.IsNullOrEmpty(itemComp) || !string.IsNullOrEmpty(itemGrade))
                        {
                            bomItemMap.Add(new MaterialInfo
                            {
                                PartNo = ptNo,
                                Description = desc.Trim(),
                                MaterialType = itemComp ?? "",
                                Grade = itemGrade ?? "",
                                Size = itemNps,
                                Schedule = itemSch,
                                Quantity = itemQty,
                            });
                        }

                        // Populate component set and grade map from item data
                        if (!string.IsNullOrEmpty(itemComp))
                            bomComponents.Add(itemComp);
                        if (!string.IsNullOrEmpty(itemComp) && !string.IsNullOrEmpty(itemGrade))
                            bomGradeMap.TryAdd(itemComp, itemGrade);

                        // Use schedule from BOM as metadata default if not set
                        if (!string.IsNullOrEmpty(itemSch) && string.IsNullOrEmpty(metadata.Schedule))
                            metadata.Schedule = itemSch;
                    }
                }
            }

            // Collect component names (both inside and outside BOM section)
            foreach (Match bomMatch in PdfBomComponentRegex().Matches(line))
            {
                var rawComp = bomMatch.Groups["desc"].Value.ToUpperInvariant();
                var normalizedComp = NormalizeBomComponentName(rawComp);
                bomComponents.Add(normalizedComp);
            }

            // Try to extract component → grade mappings from BOM description lines.
            // Use Matches() to capture ALL component-grade pairs on the same line.
            foreach (Match compGradeMatch in PdfBomComponentGradeRx().Matches(line))
            {
                var comp = NormalizeBomComponentName(compGradeMatch.Groups["comp"].Value);
                var grade = ExtractBomGradeValue(compGradeMatch);
                if (!string.IsNullOrEmpty(comp) && !string.IsNullOrEmpty(grade))
                {
                    // Keep the first occurrence per component type (most likely the primary spec)
                    bomGradeMap.TryAdd(comp, grade);
                }
            }

            // Also try reverse-order pattern (grade BEFORE component name)
            // e.g. "A106-B Seamless Pipe" or "A234 WPB Elbow 90 LR"
            foreach (Match revMatch in PdfBomGradeComponentRx().Matches(line))
            {
                var comp = NormalizeBomComponentName(revMatch.Groups["comp"].Value);
                var grade = ExtractBomGradeValue(revMatch);
                if (!string.IsNullOrEmpty(comp) && !string.IsNullOrEmpty(grade))
                {
                    bomGradeMap.TryAdd(comp, grade);
                }
            }

            // Fallback: if line is in BOM section and has a standalone grade
            // pattern near a component name, try a looser association.
            // This handles OCR output where spacing/ordering is garbled.
            if (inBomSection && !inErectionSection && !string.IsNullOrWhiteSpace(trimmedBom))
            {
                var lineComps = PdfBomComponentRegex().Matches(trimmedBom);
                var lineGrades = PdfBomGradeRx().Matches(trimmedBom);
                if (lineComps.Count == 1 && lineGrades.Count >= 1)
                {
                    var singleComp = NormalizeBomComponentName(lineComps[0].Groups["desc"].Value);
                    var singleGrade = ExtractBomGradeValue(lineGrades[0]);
                    if (!string.IsNullOrEmpty(singleComp) && !string.IsNullOrEmpty(singleGrade))
                    {
                        bomGradeMap.TryAdd(singleComp, singleGrade);
                    }
                }

                // Category-based inference: if line is within a known category
                // and has a grade but no explicit component name, use the category
                if (lineComps.Count == 0 && lineGrades.Count >= 1 && !string.IsNullOrEmpty(bomCategory)
                    && bomCategory != "FITTINGS")
                {
                    var catGrade = ExtractBomGradeValue(lineGrades[0]);
                    if (!string.IsNullOrEmpty(catGrade))
                    {
                        bomGradeMap.TryAdd(bomCategory, catGrade);
                    }
                }
            }
        }
        if (bomComponents.Count > 0)
        {
            metadata.BillOfMaterials = bomComponents.ToList();
        }
        // Store the BOM item details and grade map in Materials for downstream use
        if (bomItemMap.Count > 0)
        {
            metadata.Materials = bomItemMap;
        }
        else if (bomGradeMap.Count > 0)
        {
            metadata.Materials = bomGradeMap
                .Select(kvp => new MaterialInfo { MaterialType = kvp.Key, Grade = kvp.Value })
                .ToList();
        }
        if (bomGradeMap.Count > 0)
        {
            warnings.Add($"BOM grade map: {string.Join(", ", bomGradeMap.Select(kv => $"{kv.Key}={kv.Value}"))}");
        }
        if (bomItemMap.Count > 0)
        {
            warnings.Add($"BOM items: {string.Join(", ", bomItemMap.Where(i => !string.IsNullOrEmpty(i.MaterialType)).Select(i => $"PT{i.PartNo}={i.MaterialType}({i.Grade}){(i.Quantity > 1 ? $"×{i.Quantity}" : "")}"))}");
        }

        // ---- Scan for CUTTING LIST table to build spool → size mapping ----
        // The cutting list provides: PIECE NO, SPOOL NO, SIZE, LENGTH, END1, END2, IDENT.
        // We extract SPOOL NO → SIZE to populate SpoolDia for WS joints.
        var cuttingListEntries = new List<CuttingListEntry>();
        bool inCuttingList = false;
        int cuttingListBlank = 0;
        for (int cli = 0; cli < lines.Length; cli++)
        {
            var clLine = lines[cli].Trim();
            if (PdfCuttingListHeaderRegex().IsMatch(clLine))
            {
                inCuttingList = true;
                cuttingListBlank = 0;
                continue;
            }

            if (!inCuttingList) continue;

            if (string.IsNullOrWhiteSpace(clLine))
            {
                cuttingListBlank++;
                if (cuttingListBlank > 5) inCuttingList = false;
                continue;
            }
            cuttingListBlank = 0;

            var clMatch = PdfCuttingListRowRegex().Match(clLine);
            if (clMatch.Success)
            {
                cuttingListEntries.Add(new CuttingListEntry
                {
                    PieceNo = clMatch.Groups["piece"].Value,
                    SpoolNo = clMatch.Groups["spool"].Value,
                    Size = NormalizeFractionalSize(clMatch.Groups["size"].Value),
                    Length = clMatch.Groups["length"].Value,
                    End1 = clMatch.Groups["end1"].Success ? clMatch.Groups["end1"].Value : null,
                    End2 = clMatch.Groups["end2"].Success ? clMatch.Groups["end2"].Value : null,
                    Ident = clMatch.Groups["ident"].Success ? clMatch.Groups["ident"].Value : null,
                });
            }
        }

        // Build spool → size lookup from cutting list using Max(SIZE) per spool.
        // A spool may contain pieces of different pipe sizes (e.g. reducer transitions);
        // the maximum size represents the spool's header/run pipe diameter.
        var cuttingListSpoolSize = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        {
            var spoolSizes = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in cuttingListEntries)
            {
                if (string.IsNullOrEmpty(entry.SpoolNo) || string.IsNullOrEmpty(entry.Size)) continue;

                // Normalize spool: "SP01" → "01"
                var spoolMatch = SpoolTagRegex().Match(entry.SpoolNo);
                var normalizedSpool = spoolMatch.Success
                    ? spoolMatch.Groups[1].Value.PadLeft(2, '0')
                    : entry.SpoolNo;

                if (!spoolSizes.TryGetValue(normalizedSpool, out var sizes))
                {
                    sizes = new List<string>();
                    spoolSizes[normalizedSpool] = sizes;
                }

                sizes.Add(entry.Size);
            }

            // Pick the largest size per spool (Max)
            foreach (var (spool, sizes) in spoolSizes)
            {
                var maxSize = sizes
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(s => TryParseDiameter(s, out var d) ? d : 0)
                    .First();
                cuttingListSpoolSize[spool] = maxSize;
            }
        }
        metadata.CuttingListEntries = cuttingListEntries;
        metadata.CuttingListSpoolSizes = cuttingListSpoolSize;

        if (cuttingListEntries.Count > 0)
        {
            warnings.Add($"Parsed {cuttingListEntries.Count} cutting list entries. Spool sizes: {string.Join(", ", cuttingListSpoolSize.Select(kv => $"SP{kv.Key}={kv.Value}\""))}");
        }

        // ---- Extract individual weld records ----

        // Pre-scan: collect spool labels (SP01, SP02, SPOOL 03, etc.) with
        // their line indices for proximity-based assignment.
        // IMPORTANT: only include standalone drawing body labels, NOT spool
        // references embedded in cutting list or BOM table rows.  Drawing body
        // labels sit on short lines ("SP01") with minimal surrounding text,
        // while cutting list entries embed the spool number inside longer
        // descriptive rows ("SP01  Pipe 3\" A106-B Seamless ...").  Cutting list
        // entries at the top of the page add noise that makes nearest-spool
        // lookup assign the wrong spool.
        var spoolLabels = new List<(int LineIndex, string SpoolNo)>();
        for (int li = 0; li < lines.Length; li++)
        {
            var lineText = lines[li].Trim();

            string? detectedSpool = null;
            int matchLen = 0;

            var spM = SpoolNumberRegex().Match(lineText);
            if (spM.Success)
            {
                detectedSpool = spM.Groups[1].Value.PadLeft(2, '0');
                matchLen = spM.Length;
            }
            else
            {
                var spM2 = PdfSpoolRegex().Match(lineText);
                if (spM2.Success)
                {
                    detectedSpool = spM2.Groups["spool"].Value.PadLeft(2, '0');
                    matchLen = spM2.Length;
                }
            }

            if (detectedSpool == null) continue;

            // Filter: surrounding text length (line length minus the matched label).
            // Standalone drawing body labels: surrounding text ≈ 0–15 chars.
            // Cutting list / BOM rows: surrounding text > 25 chars.
            var surroundingLen = lineText.Length - matchLen;
            if (surroundingLen <= 25)
            {
                spoolLabels.Add((li, detectedSpool));
            }
        }

        // Fallback: if filtering removed all labels, use unfiltered scan
        if (spoolLabels.Count == 0)
        {
            for (int li = 0; li < lines.Length; li++)
            {
                var spM = SpoolNumberRegex().Match(lines[li]);
                if (spM.Success)
                {
                    spoolLabels.Add((li, spM.Groups[1].Value.PadLeft(2, '0')));
                }
                else
                {
                    var spM2 = PdfSpoolRegex().Match(lines[li]);
                    if (spM2.Success)
                    {
                        spoolLabels.Add((li, spM2.Groups["spool"].Value.PadLeft(2, '0')));
                    }
                }
            }
        }

        // Deduplicate: when the same spool number appears at multiple positions
        // (e.g. SP01 in cutting list line 5 AND drawing body line 50), keep only
        // the occurrence closest to weld annotations.  Drawing body labels are
        // near the weld annotations; cutting list entries are far away.
        if (spoolLabels.GroupBy(s => s.SpoolNo, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
        {
            // Find approximate center of weld annotation area
            var weldCenterLines = new List<int>();
            for (int li = 0; li < lines.Length; li++)
            {
                var lt = lines[li].Trim();
                if (PdfWeldMapRowRegex().IsMatch(lt) || PdfWeldLineRegex().IsMatch(lt))
                    weldCenterLines.Add(li);
            }
            int weldMedian = weldCenterLines.Count > 0
                ? weldCenterLines[weldCenterLines.Count / 2]
                : lines.Length / 2;

            spoolLabels = spoolLabels
                .GroupBy(s => s.SpoolNo, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderBy(s => Math.Abs(s.LineIndex - weldMedian)).First())
                .ToList();
        }

        // Phase 1: Try to parse structured weld list table rows
        //   e.g. "01 BW 01 10 40 PIPE FLANGE A106-B A105"
        //   e.g. "F01 FJ 10 0 STUD BOLT A193"
        bool inWeldListSection = false;
        int weldCounter = 0;

        // Track the line index where each weld was first detected.
        // Used by the PT NO proximity algorithm to associate BOM items with welds.
        var weldLineIndices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Helper: adds a weld to metadata and records its source line index
        void AddWeld(WeldRecord w, int srcLineIdx)
        {
            w.SourceLineIndex = srcLineIdx;
            metadata.Welds.Add(w);
            if (!string.IsNullOrEmpty(w.Number))
                weldLineIndices.TryAdd(w.Number, srcLineIdx);
        }

        for (int lineIdx = 0; lineIdx < lines.Length; lineIdx++)
        {
            var trimmed = lines[lineIdx].Trim();

            // Detect weld list section header
            if (PdfWeldListHeaderRegex().IsMatch(trimmed))
            {
                inWeldListSection = true;
                continue;
            }

            // Try FJ (Flange Joint) weld table row first
            var fjMatch = PdfFjWeldRowRegex().Match(trimmed);
            if (fjMatch.Success)
            {
                weldCounter++;
                var weld = new WeldRecord
                {
                    Number = NormalizeWeldNumber(fjMatch.Groups["weld"].Value),
                    Type = "FJ",
                    Location = "FW",
                    SpoolNo = null,
                    Size = fjMatch.Groups["dia"].Success ? NormalizeFractionalSize(fjMatch.Groups["dia"].Value) : null,
                    Schedule = fjMatch.Groups["sch"].Success ? fjMatch.Groups["sch"].Value : "0",
                    MaterialA = "STUD BOLT",
                    MaterialB = "FLANGE",
                    GradeA = "",
                    GradeB = fjMatch.Groups["grade"].Success ? fjMatch.Groups["grade"].Value : "",
                };
                AddWeld(weld, lineIdx);
                continue;
            }

            // Try weld list row that includes Location + Dia/Schedule + Type + components + grades
            var locTableMatch = PdfWeldTableRowWithLocationRegex().Match(trimmed);
            if (locTableMatch.Success)
            {
                weldCounter++;
                var locVal = locTableMatch.Groups["loc"].Value.ToUpperInvariant();
                string? tableSpool = null;
                if (locVal == "WS" && locTableMatch.Groups["spool"].Success && !string.IsNullOrWhiteSpace(locTableMatch.Groups["spool"].Value))
                {
                    var rawSpool = locTableMatch.Groups["spool"].Value;
                    var sm = SpoolTagRegex().Match(rawSpool);
                    tableSpool = sm.Success ? sm.Groups[1].Value.PadLeft(2, '0') : rawSpool.PadLeft(2, '0');
                }

                var compA = locTableMatch.Groups["compA"].Value.Trim().ToUpperInvariant();
                var compB = locTableMatch.Groups["compB"].Value.Trim().ToUpperInvariant();
                var gradeA = locTableMatch.Groups["gradeA"].Success ? NormalizeAstmGrade(locTableMatch.Groups["gradeA"].Value) : "";
                var gradeB = locTableMatch.Groups["gradeB"].Success ? NormalizeAstmGrade(locTableMatch.Groups["gradeB"].Value) : "";

                // If grade columns were OCR-missed, try trailing text after the match
                var locTail = trimmed[(locTableMatch.Index + locTableMatch.Length)..];
                (gradeA, gradeB) = ExtractGradesFromTail(locTail, gradeA, gradeB);

                AddWeld(new WeldRecord
                {
                    Number = NormalizeWeldNumber(locTableMatch.Groups["weld"].Value),
                    Type = NormalizePdfWeldType(locTableMatch.Groups["type"].Value),
                    Location = locVal,
                    SpoolNo = locVal == "WS" ? tableSpool : null,
                    Size = NormalizeFractionalSize(locTableMatch.Groups["dia"].Value),
                    Schedule = locTableMatch.Groups["sch"].Value,
                    MaterialA = compA,
                    MaterialB = compB,
                    GradeA = gradeA,
                    GradeB = gradeB,
                    ExplicitTableData = !string.IsNullOrEmpty(compA) && !string.IsNullOrEmpty(gradeA),
                }, lineIdx);
                continue;
            }

            // Try Saudi Aramco weld list table row
            // This format has TYPE before SPOOL/SCH, specific to Saudi Aramco ISOs.
            var saudiMatch = PdfSaudiWeldListRowRegex().Match(trimmed);
            if (saudiMatch.Success)
            {
                weldCounter++;
                var saudiLoc = saudiMatch.Groups["loc"].Value.ToUpperInvariant();
                var saudiType = NormalizePdfWeldType(saudiMatch.Groups["type"].Value);
                var saudiCompA = saudiMatch.Groups["compA"].Value.Trim().ToUpperInvariant();
                var saudiCompB = saudiMatch.Groups["compB"].Value.Trim().ToUpperInvariant();
                var saudiGradeA = saudiMatch.Groups["gradeA"].Success
                    ? NormalizeAstmGrade(saudiMatch.Groups["gradeA"].Value) : "";
                var saudiGradeB = saudiMatch.Groups["gradeB"].Success
                    ? NormalizeAstmGrade(saudiMatch.Groups["gradeB"].Value) : "";

                // Extract spool and schedule from the numeric tokens between TYPE and MATL A
                var numTokens = saudiMatch.Groups["nums"].Value.Trim()
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries);
                string? saudiSpool = null;
                string? saudiSch = null;
                if (numTokens.Length >= 1)
                {
                    var first = numTokens[0];
                    if (first != "--" && int.TryParse(first, out var spoolNum) && spoolNum >= 1 && spoolNum <= 99)
                        saudiSpool = spoolNum.ToString("D2");
                }
                if (numTokens.Length >= 2)
                {
                    var last = numTokens[^1];
                    if (last != "--")
                        saudiSch = last;
                }

                // Try trailing text for missed grades
                var saudiTail = trimmed[(saudiMatch.Index + saudiMatch.Length)..];
                (saudiGradeA, saudiGradeB) = ExtractGradesFromTail(saudiTail,
                    string.IsNullOrEmpty(saudiGradeA) ? null : saudiGradeA,
                    string.IsNullOrEmpty(saudiGradeB) ? null : saudiGradeB);

                // Infer grades from component types if still missing
                if (string.IsNullOrEmpty(saudiGradeA))
                    saudiGradeA = InferGradeFromComponent(saudiCompA);
                if (string.IsNullOrEmpty(saudiGradeB))
                    saudiGradeB = InferGradeFromComponent(saudiCompB);

                AddWeld(new WeldRecord
                {
                    Number = NormalizeWeldNumber(saudiMatch.Groups["weld"].Value),
                    Type = saudiType,
                    Location = saudiLoc,
                    SpoolNo = saudiLoc == "WS" ? saudiSpool : null,
                    Size = NormalizeFractionalSize(saudiMatch.Groups["dia"].Value),
                    Schedule = saudiSch,
                    MaterialA = saudiCompA,
                    MaterialB = saudiCompB,
                    GradeA = saudiGradeA ?? "",
                    GradeB = saudiGradeB ?? "",
                    ExplicitTableData = !string.IsNullOrEmpty(saudiCompA),
                }, lineIdx);
                continue;
            }

            // Try weld map table row (WELD NO, SHOP/FIELD, WELD SIZE, WELD TYPE, optional COMP A, COMP B)
            // Use Matches() to capture ALL annotations on the same line (spatial grouping may
            // merge multiple weld callouts, e.g. "T10 YFW 1.1/2\" TH  09A YFW 1.1/2\" SW").
            var weldMapMatches = PdfWeldMapRowRegex().Matches(trimmed);
            if (weldMapMatches.Count > 0)
            {
                foreach (Match weldMapMatch in weldMapMatches)
                {
                    weldCounter++;
                    var mapLoc = weldMapMatch.Groups["loc"].Value.ToUpperInvariant();
                    var mapType = NormalizePdfWeldType(weldMapMatch.Groups["type"].Value);
                    var mapCompA = weldMapMatch.Groups["compA"].Success ? weldMapMatch.Groups["compA"].Value.Trim().ToUpperInvariant() : "";
                    var mapCompB = weldMapMatch.Groups["compB"].Success ? weldMapMatch.Groups["compB"].Value.Trim().ToUpperInvariant() : "";

                    // Extract Material A/B and Grade A/B from trailing remark text.
                    // When the weld list table row matches PdfWeldMapRowRegex instead of
                    // PdfSaudiWeldListRowRegex, the component names and grades end up in
                    // the "remark" capture group. Parse them to populate the correct fields.
                    var mapRemarkText = weldMapMatch.Groups["remark"].Success
                        ? weldMapMatch.Groups["remark"].Value.Trim() : "";
                    string? mapGradeA = null;
                    string? mapGradeB = null;
                    string? cleanRemark = null;

                    if (!string.IsNullOrEmpty(mapRemarkText))
                    {
                        // Extract additional component names from remark when compA or compB are missing
                        if (string.IsNullOrEmpty(mapCompA) || string.IsNullOrEmpty(mapCompB))
                        {
                            var remarkComps = PdfBomComponentRegex().Matches(mapRemarkText);
                            if (remarkComps.Count > 0 && string.IsNullOrEmpty(mapCompA))
                                mapCompA = NormalizeBomComponentName(remarkComps[0].Groups["desc"].Value);
                            if (remarkComps.Count > 1 && string.IsNullOrEmpty(mapCompB))
                                mapCompB = NormalizeBomComponentName(remarkComps[1].Groups["desc"].Value);
                        }

                        // Extract grades from remark text
                        (mapGradeA, mapGradeB) = ExtractGradesFromTail(mapRemarkText, null, null);

                        // Preserve remark only if it contained no parseable grade data
                        cleanRemark = (mapGradeA != null || mapGradeB != null) ? null : mapRemarkText;
                    }

                    // Assign spool based on location:
                    // WS (circle on drawing) → find the nearest spool label by line proximity
                    // FW/YFW/YWS (square/triangle/rhombus) → field weld between spools, no spool
                    var mapSpool = mapLoc == "WS" ? FindNearestSpool(lineIdx, spoolLabels) : null;
                    var weld = new WeldRecord
                    {
                        Number = NormalizeWeldNumber(weldMapMatch.Groups["weld"].Value),
                        Type = mapType,
                        Location = mapLoc,
                        SpoolNo = mapSpool,
                        Size = NormalizeFractionalSize(weldMapMatch.Groups["dia"].Value),
                        MaterialA = mapCompA,
                        MaterialB = mapCompB,
                        GradeA = mapGradeA ?? (!string.IsNullOrEmpty(mapCompA) ? InferGradeFromComponent(mapCompA) : ""),
                        GradeB = mapGradeB ?? (!string.IsNullOrEmpty(mapCompB) ? InferGradeFromComponent(mapCompB) : ""),
                        Remarks = cleanRemark,
                    };
                    AddWeld(weld, lineIdx);
                }
                continue;
            }

            // Try structured weld list table row
            var tableMatch = PdfWeldTableRowRegex().Match(trimmed);
            if (tableMatch.Success)
            {
                weldCounter++;
                var weldNo = tableMatch.Groups["weld"].Value.ToUpperInvariant();
                var weldType = NormalizePdfWeldType(tableMatch.Groups["type"].Value);
                var spool = tableMatch.Groups["spool"].Value;
                var compA = tableMatch.Groups["compA"].Value.Trim().ToUpperInvariant();
                var compB = tableMatch.Groups["compB"].Success ? tableMatch.Groups["compB"].Value.Trim().ToUpperInvariant() : "";
                var gradeA = tableMatch.Groups["gradeA"].Success ? NormalizeAstmGrade(tableMatch.Groups["gradeA"].Value) : "";
                var gradeB = tableMatch.Groups["gradeB"].Success ? NormalizeAstmGrade(tableMatch.Groups["gradeB"].Value) : "";

                // If OCR skipped the grade columns, scrape trailing text
                var tableTail = trimmed[(tableMatch.Index + tableMatch.Length)..];
                (gradeA, gradeB) = ExtractGradesFromTail(tableTail, gradeA, gradeB);

                // Determine location from line context or weld type
                string location;
                var locMatch = PdfLocationRegex().Match(trimmed);
                if (locMatch.Success)
                {
                    location = locMatch.Groups["loc"].Value.ToUpperInvariant();
                }
                else
                {
                    location = weldType is "FW" or "TH" or "FJ" ? "FW" : "WS";
                }

                // Infer grades from component types if not explicitly provided
                if (string.IsNullOrEmpty(gradeA))
                    gradeA = InferGradeFromComponent(compA);
                if (string.IsNullOrEmpty(gradeB))
                    gradeB = InferGradeFromComponent(compB);

                // Assign spool: for WS welds use the table's explicit spool value, or
                // fall back to nearest spool label; field welds get no spool.
                var tableSpool = location == "WS"
                    ? (!string.IsNullOrEmpty(spool) ? spool : FindNearestSpool(lineIdx, spoolLabels))
                    : null;

                var weld = new WeldRecord
                {
                    Number = NormalizeWeldNumber(weldNo),
                    Type = weldType,
                    Location = location,
                    SpoolNo = tableSpool,
                    Size = NormalizeFractionalSize(tableMatch.Groups["dia"].Value),
                    Schedule = tableMatch.Groups["sch"].Value,
                    MaterialA = compA,
                    MaterialB = compB,
                    GradeA = gradeA,
                    GradeB = gradeB,
                    ExplicitTableData = !string.IsNullOrEmpty(compA) && !string.IsNullOrEmpty(gradeA),
                };

                AddWeld(weld, lineIdx);
                continue;
            }

            // Try weld table row without spool field
            var noSpoolMatch = PdfWeldNoSpoolTableRowRegex().Match(trimmed);
            if (noSpoolMatch.Success)
            {
                weldCounter++;
                var nsWeldNo = noSpoolMatch.Groups["weld"].Value.ToUpperInvariant();
                var nsType = NormalizePdfWeldType(noSpoolMatch.Groups["type"].Value);
                var nsCompA = noSpoolMatch.Groups["compA"].Value.Trim().ToUpperInvariant();
                var nsCompB = noSpoolMatch.Groups["compB"].Success ? noSpoolMatch.Groups["compB"].Value.Trim().ToUpperInvariant() : "";
                var nsGradeA = noSpoolMatch.Groups["gradeA"].Success ? NormalizeAstmGrade(noSpoolMatch.Groups["gradeA"].Value) : "";
                var nsGradeB = noSpoolMatch.Groups["gradeB"].Success ? NormalizeAstmGrade(noSpoolMatch.Groups["gradeB"].Value) : "";

                var noSpoolTail = trimmed[(noSpoolMatch.Index + noSpoolMatch.Length)..];
                (nsGradeA, nsGradeB) = ExtractGradesFromTail(noSpoolTail, nsGradeA, nsGradeB);

                if (string.IsNullOrEmpty(nsGradeA))
                    nsGradeA = InferGradeFromComponent(nsCompA);
                if (string.IsNullOrEmpty(nsGradeB))
                    nsGradeB = InferGradeFromComponent(nsCompB);

                string nsLoc;
                var nsLocM = PdfLocationRegex().Match(trimmed);
                if (nsLocM.Success) nsLoc = nsLocM.Groups["loc"].Value.ToUpperInvariant();
                else nsLoc = nsType is "FW" or "TH" or "FJ" ? "FW" : "WS";

                AddWeld(new WeldRecord
                {
                    Number = NormalizeWeldNumber(nsWeldNo),
                    Type = nsType,
                    Location = nsLoc,
                    SpoolNo = nsLoc == "WS" ? FindNearestSpool(lineIdx, spoolLabels) : null,
                    Size = NormalizeFractionalSize(noSpoolMatch.Groups["dia"].Value),
                    Schedule = noSpoolMatch.Groups["sch"].Value,
                    MaterialA = nsCompA,
                    MaterialB = nsCompB,
                    GradeA = nsGradeA,
                    GradeB = nsGradeB,
                    ExplicitTableData = !string.IsNullOrEmpty(nsCompA) && !string.IsNullOrEmpty(nsGradeA),
                }, lineIdx);
                continue;
            }

            // Try simple weld-component row (weld TYPE compA compB — no dia/sch/spool)
                var compRowMatch = PdfWeldCompRowRegex().Match(trimmed);
            if (compRowMatch.Success)
            {
                weldCounter++;
                var crType = NormalizePdfWeldType(compRowMatch.Groups["type"].Value);
                var crCompA = compRowMatch.Groups["compA"].Value.Trim().ToUpperInvariant();
                var crCompB = compRowMatch.Groups["compB"].Value.Trim().ToUpperInvariant();

                    var compTail = trimmed[(compRowMatch.Index + compRowMatch.Length)..];
                    var (crGradeA, crGradeB) = ExtractGradesFromTail(compTail, null, null);

                string crLoc;
                var crLocM = PdfLocationRegex().Match(trimmed);
                if (crLocM.Success) crLoc = crLocM.Groups["loc"].Value.ToUpperInvariant();
                else crLoc = crType is "FW" or "TH" or "FJ" ? "FW" : "WS";

                AddWeld(new WeldRecord
                {
                    Number = NormalizeWeldNumber(compRowMatch.Groups["weld"].Value),
                    Type = crType,
                    Location = crLoc,
                    SpoolNo = crLoc == "WS" ? FindNearestSpool(lineIdx, spoolLabels) : null,
                    MaterialA = crCompA,
                    MaterialB = crCompB,
                        GradeA = crGradeA ?? InferGradeFromComponent(crCompA),
                        GradeB = crGradeB ?? InferGradeFromComponent(crCompB),
                }, lineIdx);
                continue;
            }

            // Try flexible weld row (simpler format: number + type) within weld list section
            if (inWeldListSection)
            {
                var flexMatch = PdfFlexWeldRowRegex().Match(trimmed);
                if (flexMatch.Success)
                {
                    weldCounter++;
                    var flexWeldNo = flexMatch.Groups["weld"].Value;
                    var flexType = NormalizePdfWeldType(flexMatch.Groups["type"].Value);

                    // Extract fields from the rest of the line
                    var afterMatch = trimmed[(flexMatch.Index + flexMatch.Length)..];
                    var nums = PdfSmallNumberRegex().Matches(afterMatch)
                        .Cast<Match>().Select(m => m.Groups[1].Value).ToList();

                    string? flexDia = nums.Count > 0 ? nums[0] : null;
                    string? flexSch = nums.Count > 1 ? nums[1] : null;

                    string flexLoc;
                    var flexLocMatch = PdfLocationRegex().Match(trimmed);
                    if (flexLocMatch.Success)
                        flexLoc = flexLocMatch.Groups["loc"].Value.ToUpperInvariant();
                    else
                        flexLoc = flexType is "FW" or "TH" or "FJ" ? "FW" : "WS";

                    // Try component names from the line, ordered by position in text
                    string[] compKeywords = { "PIPE", "TEE", "ELBOW", "FLANGE", "REDUCER", "VALVE", "OLET", "NIPPLE", "CAP", "PAD", "PLUG", "COUPLING", "UNION", "SWAGE", "HOSE", "PLATE", "STUD BOLT" };
                    var comps = compKeywords
                        .Select(c => new { Keyword = c, Match = Regex.Match(trimmed, @"\b" + Regex.Escape(c) + @"\b", RegexOptions.IgnoreCase) })
                        .Where(x => x.Match.Success)
                        .OrderBy(x => x.Match.Index)
                        .Select(x => x.Keyword)
                        .ToList();

                    var flexTail = trimmed[(flexMatch.Index + flexMatch.Length)..];
                    var (flexGradeA, flexGradeB) = ExtractGradesFromTail(flexTail, null, null);

                    var flexWeld = new WeldRecord
                    {
                        Number = NormalizeWeldNumber(flexWeldNo),
                        Type = flexType,
                        Location = flexLoc,
                        SpoolNo = flexLoc == "WS" ? FindNearestSpool(lineIdx, spoolLabels) : null,
                        Size = flexDia,
                        Schedule = flexSch,
                        MaterialA = comps.Count > 0 ? comps[0] : "",
                        MaterialB = comps.Count > 1 ? comps[1] : "",
                        GradeA = flexGradeA ?? (comps.Count > 0 ? InferGradeFromComponent(comps[0]) : ""),
                        GradeB = flexGradeB ?? (comps.Count > 1 ? InferGradeFromComponent(comps[1]) : ""),
                    };

                    if (flexType == "FJ")
                    {
                        flexWeld.MaterialA = "STUD BOLT";
                        flexWeld.MaterialB = "FLANGE";
                        flexWeld.GradeA = "";
                        flexWeld.GradeB = comps.Any(c => c == "A193") ? "A193" : "";
                        flexWeld.SpoolNo = null;
                        flexWeld.Location = "FW";
                        flexWeld.Schedule ??= "0";
                    }

                    AddWeld(flexWeld, lineIdx);
                    continue;
                }
            }
        }

        // Phase 2: If no table rows found, fall back to simple weld line pattern
        if (weldCounter == 0)
        {
            for (int lineIdx2 = 0; lineIdx2 < lines.Length; lineIdx2++)
            {
                var trimmed = lines[lineIdx2].Trim();

                // Parse weld lines
                var weldMatch = PdfWeldLineRegex().Match(trimmed);
                if (weldMatch.Success)
                {
                    weldCounter++;
                    var weldNo = weldMatch.Groups["weld"].Value.ToUpperInvariant();
                    var weldType = NormalizePdfWeldType(weldMatch.Groups["type"].Value);

                    // Try to extract inline diameter and schedule from same line
                    string? lineDia = null;
                    string? lineSch = null;
                    var diaMatch = PdfDiameterRegex().Match(trimmed);
                    if (diaMatch.Success) lineDia = NormalizeFractionalSize(diaMatch.Groups["dia"].Value);
                    var schMatch = PdfScheduleRegex().Match(trimmed);
                    if (schMatch.Success) lineSch = schMatch.Groups["sch"].Value;

                    // Try to extract location
                    string location;
                    var locMatch = PdfLocationRegex().Match(trimmed);
                    if (locMatch.Success)
                    {
                        location = locMatch.Groups["loc"].Value.ToUpperInvariant();
                    }
                    else
                    {
                        location = weldType is "FW" or "TH" or "FJ" ? "FW" : "WS";
                    }

                    // Assign spool via nearest-label proximity
                    var weld = new WeldRecord
                    {
                        Number = NormalizeWeldNumber(weldNo),
                        Type = weldType,
                        Location = location,
                        SpoolNo = location == "WS" ? (FindNearestSpool(lineIdx2, spoolLabels) ?? "01") : null,
                        Size = lineDia,
                        Schedule = lineSch,
                    };

                    // For FJ welds set standard material values
                    if (weldType == "FJ")
                    {
                        weld.MaterialA = "STUD BOLT";
                        weld.MaterialB = "FLANGE";
                        weld.GradeA = "";
                        weld.GradeB = "";
                        weld.Schedule ??= "0";
                        weld.SpoolNo = null;
                        weld.Location = "FW";
                    }

                    AddWeld(weld, lineIdx2);
                }
            }
        }

        // Phase 3: If still no welds, try flexible pattern globally
        // Only accept results if at least 2 candidates found (reduces false positives)
        if (weldCounter == 0)
        {
            var flexCandidates = new List<(Match match, string line, int lineIndex)>();
            for (int li = 0; li < lines.Length; li++)
            {
                var trimmed3 = lines[li].Trim();
                var fm = PdfFlexWeldRowRegex().Match(trimmed3);
                if (fm.Success)
                {
                    flexCandidates.Add((fm, trimmed3, li));
                }
            }

            if (flexCandidates.Count >= 2)
            {
                foreach (var (fm, candidateLine, candIdx) in flexCandidates)
                {
                    weldCounter++;
                    var cWeldNo = fm.Groups["weld"].Value;
                    var cType = NormalizePdfWeldType(fm.Groups["type"].Value);

                    var afterFm = candidateLine[(fm.Index + fm.Length)..];
                    var cNums = PdfSmallNumberRegex().Matches(afterFm)
                        .Cast<Match>().Select(m => m.Groups[1].Value).ToList();

                    string cLoc;
                    var cLocM = PdfLocationRegex().Match(candidateLine);
                    if (cLocM.Success) cLoc = cLocM.Groups["loc"].Value.ToUpperInvariant();
                    else cLoc = cType is "FW" or "TH" or "FJ" ? "FW" : "WS";

                    // Find nearest spool label for WS welds
                    var cSpool = cLoc == "WS" ? FindNearestSpool(candIdx, spoolLabels) : null;

                    string[] cCompKeys = { "PIPE", "TEE", "ELBOW", "FLANGE", "REDUCER", "VALVE", "OLET", "NIPPLE", "CAP", "PAD", "PLUG", "COUPLING", "UNION", "SWAGE", "HOSE", "PLATE", "STUD BOLT" };
                    var cComps = cCompKeys
                        .Select(c => new { Keyword = c, Match = Regex.Match(candidateLine, @"\b" + Regex.Escape(c) + @"\b", RegexOptions.IgnoreCase) })
                        .Where(x => x.Match.Success)
                        .OrderBy(x => x.Match.Index)
                        .Select(x => x.Keyword)
                        .ToList();

                    var flexTail = candidateLine[(fm.Index + fm.Length)..];
                    var (cGradeA, cGradeB) = ExtractGradesFromTail(flexTail, null, null);

                    var cWeld = new WeldRecord
                    {
                        Number = NormalizeWeldNumber(cWeldNo),
                        Type = cType,
                        Location = cLoc,
                        SpoolNo = cSpool,
                        Size = cNums.Count > 0 ? cNums[0] : null,
                        Schedule = cNums.Count > 1 ? cNums[1] : null,
                        MaterialA = cComps.Count > 0 ? cComps[0] : "",
                        MaterialB = cComps.Count > 1 ? cComps[1] : "",
                        GradeA = cGradeA ?? (cComps.Count > 0 ? InferGradeFromComponent(cComps[0]) : ""),
                        GradeB = cGradeB ?? (cComps.Count > 1 ? InferGradeFromComponent(cComps[1]) : ""),
                    };

                    if (cType == "FJ")
                    {
                        cWeld.MaterialA = "STUD BOLT";
                        cWeld.MaterialB = "FLANGE";
                        cWeld.GradeA = "";
                        cWeld.GradeB = "";
                        cWeld.SpoolNo = null;
                        cWeld.Location = "FW";
                        cWeld.Schedule ??= "0";
                    }

                    AddWeld(cWeld, candIdx);
                }
            }
        }

        if (weldCounter == 0)
        {
            warnings.Add("No weld records could be parsed from the PDF text. Consider using an Excel template.");
        }

        // Phase 4: Supplementary scan for weld map annotations not captured by Phase 1.
        // Drawing body annotations like "09A YFW 1.1/2\" SW" may share OCR lines
        // with weld table rows that were matched by PdfWeldTableRowRegex / PdfFjWeldRowRegex.
        // The 'continue' after those matches skips PdfWeldMapRowRegex for the same line.
        // Re-scan every line and add any new weld map hits not already captured.
        {
            var capturedNumbers = new HashSet<string>(
                metadata.Welds.Select(w => w.Number ?? ""),
                StringComparer.OrdinalIgnoreCase);

            for (int lineIdx4 = 0; lineIdx4 < lines.Length; lineIdx4++)
            {
                var trimmedLine = lines[lineIdx4].Trim();

                foreach (Match wmm in PdfWeldMapRowRegex().Matches(trimmedLine))
                {
                    var suppNo = NormalizeWeldNumber(wmm.Groups["weld"].Value);
                    if (capturedNumbers.Contains(suppNo)) continue;

                    var suppLoc = wmm.Groups["loc"].Value.ToUpperInvariant();
                    var suppType = NormalizePdfWeldType(wmm.Groups["type"].Value);
                    var suppCompA = wmm.Groups["compA"].Success ? wmm.Groups["compA"].Value.Trim().ToUpperInvariant() : "";
                    var suppCompB = wmm.Groups["compB"].Success ? wmm.Groups["compB"].Value.Trim().ToUpperInvariant() : "";

                    AddWeld(new WeldRecord
                    {
                        Number = suppNo,
                        Type = suppType,
                        Location = suppLoc,
                        SpoolNo = suppLoc == "WS" ? FindNearestSpool(lineIdx4, spoolLabels) : null,
                        Size = NormalizeFractionalSize(wmm.Groups["dia"].Value),
                        MaterialA = suppCompA,
                        MaterialB = suppCompB,
                        GradeA = !string.IsNullOrEmpty(suppCompA) ? InferGradeFromComponent(suppCompA) : "",
                        GradeB = !string.IsNullOrEmpty(suppCompB) ? InferGradeFromComponent(suppCompB) : "",
                        Remarks = wmm.Groups["remark"].Success
                            ? wmm.Groups["remark"].Value.Trim() : null,
                    }, lineIdx4);
                    capturedNumbers.Add(suppNo);
                    weldCounter++;
                }
            }
        }

        // ---- Post-process: reconcile spool assignments from weld list table ----
        // The weld list table in the drawing is the authoritative source for
        // weld → spool mapping.  Extract it and override any proximity-inferred
        // spools which may be wrong due to 2D→1D text flattening.
        if (metadata.Welds.Count > 0)
        {
            var knownWeldNumbers = new HashSet<string>(
                metadata.Welds.Where(w => !string.IsNullOrEmpty(w.Number)).Select(w => w.Number!),
                StringComparer.OrdinalIgnoreCase);
            var knownSpoolNumbers = new HashSet<string>(
                spoolLabels.Select(s => s.SpoolNo),
                StringComparer.OrdinalIgnoreCase);

            var tableSpools = ExtractWeldListTableSpools(lines, knownWeldNumbers, knownSpoolNumbers);

            if (tableSpools.Count > 0)
            {
                foreach (var w in metadata.Welds)
                {
                    if (!string.IsNullOrEmpty(w.Number)
                        && (w.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase)
                        && tableSpools.TryGetValue(w.Number, out var tableSpool))
                    {
                        w.SpoolNo = tableSpool;
                    }
                }
            }

            // Propagate best spool to any remaining WS welds without a spool
            var bestSpoolPerWeld = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var w in metadata.Welds)
            {
                if (string.IsNullOrEmpty(w.Number) || string.IsNullOrEmpty(w.SpoolNo)) continue;
                bestSpoolPerWeld.TryAdd(w.Number, w.SpoolNo);
            }

            foreach (var w in metadata.Welds)
            {
                if (!string.IsNullOrEmpty(w.Number) && string.IsNullOrEmpty(w.SpoolNo)
                    && (w.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase))
                {
                    if (bestSpoolPerWeld.TryGetValue(w.Number, out var sp))
                        w.SpoolNo = sp;
                }
            }
        }

        // ---- Post-process: deduplicate, sort, and assign spools ----
        // Deduplication and natural sort always run so downstream consumers
        // see a clean, ordered weld list regardless of spool-assignment source.
        // Spool assignment uses the successive-WS-group algorithm with the best
        // available ordered spool list:
        //   1. Cutting list (authoritative) — preserves first-appearance order
        //   2. Spool labels from drawing body (fallback) — sorted numerically
        // This overrides any proximity-based or weld-list-table-based spool
        // assignments made earlier, which are unreliable due to OCR spatial noise.
        if (metadata.Welds.Count > 0)
        {
            // Step A: Deduplicate welds by number — the same weld may appear
            // twice (once from a drawing-body annotation via PdfWeldMapRowRegex,
            // once from a structured table row via PdfWeldTableRowRegex).
            // Merge fields: take the richest entry as base, but override its
            // Location with the most specific value from any duplicate, because
            // PdfWeldMapRowRegex captures the explicit SHOP/FIELD column
            // (WS/YFW/FW/YWS/SW) while PdfWeldTableRowRegex infers location
            // from weld type (defaulting to WS), which can be wrong.
            var deduped = metadata.Welds
                .GroupBy(w => w.Number ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    var entries = g.ToList();
                    if (entries.Count == 1) return entries[0];

                    // Pick the entry with the most populated fields as base
                    var merged = entries.OrderByDescending(w =>
                        (string.IsNullOrEmpty(w.MaterialA) ? 0 : 1) +
                        (string.IsNullOrEmpty(w.Size) ? 0 : 1) +
                        (string.IsNullOrEmpty(w.Schedule) ? 0 : 1) +
                        (string.IsNullOrEmpty(w.GradeA) ? 0 : 1))
                        .First();

                    // Override location with the most specific value from any entry.
                    // YFW/YWS/SW are exclusively from the weld map's SHOP/FIELD column;
                    // FW may come from weld map or type inference;
                    // WS is the default fallback when location is inferred from type.
                    string? bestLoc = null;
                    foreach (var e in entries)
                    {
                        var loc = (e.Location ?? "").ToUpperInvariant();
                        if (loc is "YFW" or "YWS" or "SW")
                        {
                            bestLoc = e.Location;
                            break; // Most specific — always from weld map
                        }
                        if (loc == "FW" && bestLoc == null)
                            bestLoc = e.Location;
                    }
                    if (bestLoc != null)
                        merged.Location = bestLoc;

                    // Fill in any missing fields from other entries
                    foreach (var e in entries)
                    {
                        if (ReferenceEquals(e, merged)) continue;
                        // Propagate ExplicitTableData flag from any entry
                        if (e.ExplicitTableData) merged.ExplicitTableData = true;
                        if (string.IsNullOrEmpty(merged.Size) && !string.IsNullOrEmpty(e.Size))
                            merged.Size = e.Size;
                        if (string.IsNullOrEmpty(merged.Schedule) && !string.IsNullOrEmpty(e.Schedule))
                            merged.Schedule = e.Schedule;
                        if (string.IsNullOrEmpty(merged.MaterialA) && !string.IsNullOrEmpty(e.MaterialA))
                            merged.MaterialA = e.MaterialA;
                        if (string.IsNullOrEmpty(merged.MaterialB) && !string.IsNullOrEmpty(e.MaterialB))
                            merged.MaterialB = e.MaterialB;
                        if (string.IsNullOrEmpty(merged.GradeA) && !string.IsNullOrEmpty(e.GradeA))
                            merged.GradeA = e.GradeA;
                        if (string.IsNullOrEmpty(merged.GradeB) && !string.IsNullOrEmpty(e.GradeB))
                            merged.GradeB = e.GradeB;
                    }

                    return merged;
                })
                .ToList();

            // Step B: Sort welds by natural weld-number order so the
            // successive-WS-group algorithm sees them in the same sequence
            // as the weld list table (01, 02, …, 09, T10, 11, T12, 13, …).
            // OCR spatial order is unreliable because annotations are
            // scattered across the drawing body.
            deduped.Sort((a, b) =>
            {
                var numA = ExtractWeldSortNumber(a.Number);
                var numB = ExtractWeldSortNumber(b.Number);
                if (numA != numB) return numA.CompareTo(numB);
                return string.Compare(a.Number, b.Number, StringComparison.OrdinalIgnoreCase);
            });

            // Replace metadata.Welds with the deduped + sorted list
            metadata.Welds.Clear();
            metadata.Welds.AddRange(deduped);

            // Step C: Build ordered spool list from the best available source.
            //   1. Cutting list (authoritative): preserves first-appearance order
            //   2. Spool labels detected in drawing body: sorted numerically
            //   3. Count WS groups and generate sequential spools (last resort)
            var orderedSpools = new List<string>();

            if (cuttingListSpoolSize.Count > 0)
            {
                // From cutting list (preserving first-appearance order)
                foreach (var entry in cuttingListEntries)
                {
                    if (string.IsNullOrEmpty(entry.SpoolNo)) continue;
                    var clSpoolMatch = SpoolTagRegex().Match(entry.SpoolNo);
                    var normalizedCl = clSpoolMatch.Success
                        ? clSpoolMatch.Groups[1].Value.PadLeft(2, '0')
                        : entry.SpoolNo;
                    if (!orderedSpools.Any(s => s.Equals(normalizedCl, StringComparison.OrdinalIgnoreCase)))
                        orderedSpools.Add(normalizedCl);
                }
            }
            else if (spoolLabels.Count > 0)
            {
                // No cutting list available — use spool labels from the drawing body,
                // sorted numerically (SP01, SP02, SP03, …).
                orderedSpools = spoolLabels
                    .Select(s => s.SpoolNo)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => int.TryParse(s, out var n) ? n : int.MaxValue)
                    .ThenBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else
            {
                // No cutting list or spool labels — count the distinct WS groups
                // in the sorted weld list and generate sequential spool numbers.
                // A WS group is a consecutive run of WS welds separated by non-WS
                // welds (YFW, FW, etc.). This mirrors the physical spool arrangement
                // on the drawing: each run of shop welds belongs to one spool.
                int wsGroupCount = 0;
                bool prevIsWs = false;
                foreach (var w in metadata.Welds)
                {
                    bool isWs = (w.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase);
                    if (isWs && !prevIsWs)
                        wsGroupCount++;
                    prevIsWs = isWs;
                }
                for (int g = 1; g <= wsGroupCount; g++)
                    orderedSpools.Add(g.ToString("D2"));
            }

            if (orderedSpools.Count > 0)
            {
                // Clear all stale spool assignments before reassigning.
                // Phase 1 proximity-based and reconciliation-based spools may be wrong;
                // the successive WS group algorithm is the authoritative source.
                foreach (var w in metadata.Welds)
                {
                    w.SpoolNo = null;
                    w.SpoolDia = null;
                }

                // Step D: Group successive WS welds and assign to spools in order.
                // A non-WS weld (YFW, FW, etc.) breaks the current WS group and
                // advances the spool pointer to the next spool.
                int spoolIdx = 0;
                bool prevWasWs = false;

                int unassignedWsCount = 0;
                foreach (var weld in metadata.Welds)
                {
                    bool isWs = (weld.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase);

                    if (isWs)
                    {
                        if (spoolIdx < orderedSpools.Count)
                        {
                            weld.SpoolNo = orderedSpools[spoolIdx];
                            if (cuttingListSpoolSize.TryGetValue(weld.SpoolNo, out var spSize))
                                weld.SpoolDia = spSize;
                        }
                        else
                        {
                            unassignedWsCount++;
                        }
                        prevWasWs = true;
                    }
                    else
                    {
                        // Non-WS weld: clear spool and advance pointer when breaking a WS group
                        if (prevWasWs)
                            spoolIdx++;
                        weld.SpoolNo = null;
                        weld.SpoolDia = null;
                        prevWasWs = false;
                    }
                }

                if (unassignedWsCount > 0)
                {
                    warnings.Add($"{unassignedWsCount} WS weld(s) could not be assigned to a spool — more WS groups than available spools ({orderedSpools.Count}). Review spool assignments.");
                }

                var source = cuttingListSpoolSize.Count > 0 ? "cutting list"
                           : spoolLabels.Count > 0 ? "spool labels"
                           : "WS group count";
                warnings.Add($"Assigned spools to WS joints via {source}: {string.Join(", ", orderedSpools.Select(s => $"SP{s}"))}");
            }

            // Persist ordered spools for downstream row-level reassignment
            metadata.OrderedSpools = orderedSpools;
        }
    }

    /// <summary>
    /// Finds the nearest spool label by line-index proximity for a weld at
    /// the given <paramref name="lineIndex"/>. On the drawing each spool label
    /// (SP01, SP02, …) is placed near the welds it contains, so the closest
    /// label in the spatially-grouped text is the most likely match.
    /// Returns <c>null</c> if no spool labels were detected.
    /// </summary>
    private static string? FindNearestSpool(int lineIndex, List<(int LineIndex, string SpoolNo)> spoolLabels)
    {
        if (spoolLabels.Count == 0) return null;

        string? best = null;
        int bestDist = int.MaxValue;
        foreach (var (li, sp) in spoolLabels)
        {
            var dist = Math.Abs(li - lineIndex);
            if (dist < bestDist)
            {
                bestDist = dist;
                best = sp;
            }
        }
        return best;
    }

    /// <summary>
    /// Scans the extracted text for a WELD LIST / WELD MAP section and extracts
    /// the weld number → spool number mapping from the table rows.
    /// The weld list table typically has columns: NO, TYPE, SPOOL, DIA, SCH, …
    /// Spool numbers are identified as zero-padded 2-digit values (01, 02, 03)
    /// that match the <paramref name="knownSpoolNumbers"/> set.
    /// Returns a map of normalized weld number → spool number.
    /// </summary>
    private static Dictionary<string, string> ExtractWeldListTableSpools(
        string[] lines,
        HashSet<string> knownWeldNumbers,
        HashSet<string> knownSpoolNumbers)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (knownWeldNumbers.Count == 0 || knownSpoolNumbers.Count == 0) return result;

        bool inWeldListSection = false;
        int blankLineCount = 0;

        for (int li = 0; li < lines.Length; li++)
        {
            var trimmed = lines[li].Trim();

            // Detect WELD LIST / WELD MAP section header
            if (PdfWeldListHeaderRegex().IsMatch(trimmed))
            {
                inWeldListSection = true;
                blankLineCount = 0;
                continue;
            }

            if (!inWeldListSection) continue;

            // Stop after several blank / non-matching lines (end of section)
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                blankLineCount++;
                if (blankLineCount > 5) inWeldListSection = false;
                continue;
            }
            blankLineCount = 0;

            // Try to match a weld number + weld type anywhere on the line
            var weldMatch = WeldTypeTableRowRegex().Match(trimmed);
            if (!weldMatch.Success) continue;

            var normalizedWeld = NormalizeWeldNumber(weldMatch.Groups["weld"].Value);

            // Only process if this is a weld number we've already extracted
            if (!knownWeldNumbers.Contains(normalizedWeld)) continue;

            // Extract all tokens after the weld type
            var afterType = trimmed[(weldMatch.Index + weldMatch.Length)..].Trim();
            var tokens = NonWhitespaceTokenRegex().Matches(afterType)
                .Cast<Match>()
                .Select(m => m.Value)
                .ToList();

            // Find the first token that is a zero-padded 2-digit number matching
            // a known spool.  Spool values in the weld list table are typically
            // "01", "02", "03" while diameters are "2", "3", "1.5" (un-padded)
            // and schedules are "40", "80", "160" (not matching small spools).
            foreach (var token in tokens)
            {
                if (token.Length == 2 && char.IsDigit(token[0]) && char.IsDigit(token[1])
                    && knownSpoolNumbers.Contains(token))
                {
                    result.TryAdd(normalizedWeld, token);
                    break;
                }
            }

            // Fallback: try un-padded single-digit tokens ("1", "2", "3").
            // Only accept if the token differs from the raw diameter or schedule
            // on the same row — diameter is usually a larger or fractional number,
            // schedule is 40/80/160, so single-digit values ≤ max-spool are safe.
            if (!result.ContainsKey(normalizedWeld))
            {
                foreach (var token in tokens)
                {
                    if (token.Length == 1 && char.IsDigit(token[0]))
                    {
                        var padded = token.PadLeft(2, '0');
                        if (knownSpoolNumbers.Contains(padded))
                        {
                            result.TryAdd(normalizedWeld, padded);
                            break;
                        }
                    }
                }
            }
        }

        // Second pass: scan ALL lines (not just the WELD LIST section) for weld
        // table rows that provide explicit spool numbers.  Some drawings embed
        // the weld list in a table area without a clear "WELD LIST" header.
        if (result.Count < knownWeldNumbers.Count)
        {
            for (int li = 0; li < lines.Length; li++)
            {
                var trimmed = lines[li].Trim();
                var weldMatch = WeldTypeTableRowRegex().Match(trimmed);
                if (!weldMatch.Success) continue;

                var normalizedWeld = NormalizeWeldNumber(weldMatch.Groups["weld"].Value);
                if (!knownWeldNumbers.Contains(normalizedWeld)) continue;
                if (result.ContainsKey(normalizedWeld)) continue; // already found

                var afterType = trimmed[(weldMatch.Index + weldMatch.Length)..].Trim();
                var tokens = NonWhitespaceTokenRegex().Matches(afterType)
                    .Cast<Match>()
                    .Select(m => m.Value)
                    .ToList();

                foreach (var token in tokens)
                {
                    if (token.Length == 2 && char.IsDigit(token[0]) && char.IsDigit(token[1])
                        && knownSpoolNumbers.Contains(token))
                    {
                        result.TryAdd(normalizedWeld, token);
                        break;
                    }
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Infers MaterialA and MaterialB for a weld record using the BOM components
    /// available in the drawing. Uses weld type to determine the most likely
    /// component pairing from the available BOM descriptions.
    /// </summary>
    private static void InferWeldMaterialsFromBom(WeldRecord weld, HashSet<string> bomComponents)
    {
        var type = (weld.Type ?? "").ToUpperInvariant();

        // FJ welds are always STUD BOLT / FLANGE
        if (type == "FJ")
        {
            weld.MaterialA ??= "STUD BOLT";
            weld.MaterialB ??= "FLANGE";
            return;
        }

        // SP (Support Pad) welds are always PIPE / PAD
        if (type == "SP")
        {
            if (string.IsNullOrEmpty(weld.MaterialA))
                weld.MaterialA = BomContainsComponent(bomComponents, "PIPE") ? "PIPE" : "";
            if (string.IsNullOrEmpty(weld.MaterialB))
                weld.MaterialB = BomContainsComponent(bomComponents, "PAD") ? "PAD" : "";
            return;
        }

        // BR/LET (Branch/Lethole) welds: primary is PIPE, secondary is OLET
        if (type is "BR" or "LET")
        {
            if (string.IsNullOrEmpty(weld.MaterialA))
                weld.MaterialA = BomContainsComponent(bomComponents, "PIPE") ? "PIPE" : "";
            if (string.IsNullOrEmpty(weld.MaterialB))
                weld.MaterialB = BomContainsComponent(bomComponents, "OLET") ? "OLET" : "";
            return;
        }

        // SOF (Socket-on-Flange): PIPE / FLANGE
        if (type == "SOF")
        {
            if (string.IsNullOrEmpty(weld.MaterialA))
                weld.MaterialA = BomContainsComponent(bomComponents, "PIPE") ? "PIPE" : "";
            if (string.IsNullOrEmpty(weld.MaterialB))
                weld.MaterialB = BomContainsComponent(bomComponents, "FLANGE") ? "FLANGE" : "";
            return;
        }

        // For BW, FW, SW, TH: MaterialA is typically PIPE.
        if (string.IsNullOrEmpty(weld.MaterialA))
        {
            weld.MaterialA = BomContainsComponent(bomComponents, "PIPE") ? "PIPE" : "";
        }

        // For SW/TH, the secondary side is typically a coupling
        if (type is "SW" or "TH" && string.IsNullOrEmpty(weld.MaterialB))
        {
            weld.MaterialB = BomContainsComponent(bomComponents, "COUPLING") ? "COUPLING" : "";
        }

        // For BW/FW, keep PIPE as the Material B default.
        // The downstream quantity-aware Strategy 1 in AssociatePtNosWithWelds
        // will replace PIPE with the correct fitting type (ELBOW, FLANGE, TEE,
        // etc.) using BOM quantities and weld-connection budgets. Setting a
        // specific fitting here would prevent Strategy 1's guard condition
        // (MaterialB == defaultMatB) from triggering, blocking the refinement.
        if (type is "BW" or "FW" && string.IsNullOrEmpty(weld.MaterialB))
        {
            weld.MaterialB = BomContainsComponent(bomComponents, "PIPE") ? "PIPE" : "";
        }
    }

    /// <summary>
    /// Checks if the BOM component set contains a component or any of its aliases.
    /// For example, BomContainsComponent(set, "OLET") returns true if the set
    /// contains "OLET", "BRANCH", "WELDOLET", etc.
    /// </summary>
    private static bool BomContainsComponent(HashSet<string> bomComponents, string componentType)
    {
        if (bomComponents.Contains(componentType)) return true;
        foreach (var alias in GetComponentAliases(componentType))
        {
            if (bomComponents.Contains(alias)) return true;
        }
        return false;
    }

    /// <summary>
    /// Infers the grade from a piping component type (MATD_Description).
    /// Falls back to common CS (carbon steel) defaults when no BOM data is available.
    /// Values must match Material_tbl.MAT_GRADE exactly.
    /// Handles standard MATD_Description values and their aliases.
    /// </summary>
    private static string InferGradeFromComponent(string componentType)
    {
        if (string.IsNullOrEmpty(componentType)) return "";
        var normalized = NormalizeBomComponentName(componentType);
        return normalized switch
        {
            "OLET" or "PAD" => "A105",
            _ => ""
        };
    }

    private static bool HasCompleteMaterialAndGrade(WeldRecord weld)
        => !string.IsNullOrEmpty(weld.MaterialA)
           && !string.IsNullOrEmpty(weld.MaterialB)
           && !string.IsNullOrEmpty(weld.GradeA)
           && !string.IsNullOrEmpty(weld.GradeB);

    private static bool NeedsMaterialGradeRefinement(WeldRecord weld)
        => string.IsNullOrEmpty(weld.MaterialA)
           || string.IsNullOrEmpty(weld.MaterialB)
           || string.IsNullOrEmpty(weld.GradeA)
           || string.IsNullOrEmpty(weld.GradeB);

    private static bool HasAuthoritativeTableData(WeldRecord weld)
        => weld.ExplicitTableData && HasCompleteMaterialAndGrade(weld);

    /// <summary>
    /// Extracts up to two ASTM/API grades from trailing text after a weld table match.
    /// Uses the BOM-grade regex for broader patterns (API 5L PSL2 GR.B, A420-WPL6, etc.).
    /// Existing grades are preserved; only empty slots are populated.
    /// </summary>
    private static (string? gradeA, string? gradeB) ExtractGradesFromTail(string tail, string? existingA, string? existingB)
    {
        string? gradeA = string.IsNullOrWhiteSpace(existingA) ? null : existingA;
        string? gradeB = string.IsNullOrWhiteSpace(existingB) ? null : existingB;

        if (string.IsNullOrWhiteSpace(tail))
            return (gradeA, gradeB);

        var matches = PdfBomGradeRx().Matches(tail);
        if (matches.Count == 0)
            return (gradeA, gradeB);

        // First match → Grade A, second match → Grade B when missing
        foreach (Match m in matches)
        {
            var normalized = ExtractBomGradeValue(m);
            if (string.IsNullOrEmpty(normalized))
                continue;

            if (gradeA == null)
            {
                gradeA = normalized;
                continue;
            }

            if (gradeB == null)
            {
                gradeB = normalized;
                break;
            }
        }

        return (gradeA, gradeB);
    }

    /// <summary>
    /// Infers GradeA and GradeB for a weld record using the BOM-extracted
    /// component-to-grade mappings. Prefers BOM-sourced grades over the
    /// hardcoded <see cref="InferGradeFromComponent"/> defaults, because the
    /// BOM table in the drawing carries the actual project/spec grades
    /// (e.g. A106-B, A234-WPB, A105, A193, API 5L PSL2, A420-WPL6, etc.).
    ///
    /// Also overrides hardcoded defaults when BOM provides different grades
    /// (e.g. BOM has A333-6 for PIPE instead of default A106-B for low-temp CS).
    /// </summary>
    private static void InferWeldGradesFromBom(WeldRecord weld, Dictionary<string, string> bomGradeMap)
    {
        if (bomGradeMap.Count == 0) return;

        // Preserve fully populated explicit table data, but allow BOM to fill
        // any missing/placeholder grades when OCR skipped them.
        if (HasAuthoritativeTableData(weld)) return;

        // FJ welds: GradeA is empty (stud bolt), GradeB from BOM if available
        if ((weld.Type ?? "").Equals("FJ", StringComparison.OrdinalIgnoreCase))
        {
            // GradeB for FJ = stud bolt grade
            var fjGrade = LookupBomGrade("STUD BOLT", bomGradeMap)
                       ?? LookupBomGrade("BOLT", bomGradeMap);
            if (!string.IsNullOrEmpty(fjGrade))
                weld.GradeB = fjGrade;
            else if (string.IsNullOrEmpty(weld.GradeB))
                weld.GradeB = InferGradeFromComponent("STUD BOLT");
            return;
        }

        // Grade A: look up from MaterialA component type in BOM.
        // Override existing hardcoded defaults if BOM provides a different grade
        // (e.g. BOM has A333-6 for PIPE instead of default A106-B).
        if (!string.IsNullOrEmpty(weld.MaterialA))
        {
            var bomGradeA = LookupBomGrade(weld.MaterialA, bomGradeMap);
            if (!string.IsNullOrEmpty(bomGradeA))
            {
                // BOM is authoritative — override hardcoded defaults
                if (string.IsNullOrEmpty(weld.GradeA) || IsHardcodedDefault(weld.GradeA, weld.MaterialA))
                    weld.GradeA = bomGradeA;
            }
            else if (string.IsNullOrEmpty(weld.GradeA))
            {
                weld.GradeA = InferGradeFromComponent(weld.MaterialA);
            }
        }

        // Grade B: look up from MaterialB component type in BOM
        if (!string.IsNullOrEmpty(weld.MaterialB))
        {
            var bomGradeB = LookupBomGrade(weld.MaterialB, bomGradeMap);
            if (!string.IsNullOrEmpty(bomGradeB))
            {
                if (string.IsNullOrEmpty(weld.GradeB) || IsHardcodedDefault(weld.GradeB, weld.MaterialB))
                    weld.GradeB = bomGradeB;
            }
            else if (string.IsNullOrEmpty(weld.GradeB))
            {
                weld.GradeB = InferGradeFromComponent(weld.MaterialB);
            }
        }
    }

    /// <summary>
    /// Returns true if the given grade matches the hardcoded default for the
    /// component type. Used to determine if a BOM-sourced grade should
    /// override an existing grade value.
    /// </summary>
    private static bool IsHardcodedDefault(string grade, string componentType)
    {
        var defaultGrade = InferGradeFromComponent(componentType);
        return !string.IsNullOrEmpty(defaultGrade)
            && string.Equals(grade, defaultGrade, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns true if the given grade matches the BOM-sourced grade for the
    /// component type. When Material changes (e.g. PIPE → TEE), the grade set
    /// for the OLD material (e.g. A358-316/316L for PIPE) is no longer valid and
    /// must be overridden with the grade for the NEW material.
    /// <see cref="IsHardcodedDefault"/> only knows carbon-steel defaults;
    /// this method handles stainless-steel and other alloy grades from the BOM.
    /// </summary>
    private static bool IsBomSourcedGrade(string grade, string componentType, Dictionary<string, string> bomGradeMap)
    {
        if (string.IsNullOrEmpty(grade) || string.IsNullOrEmpty(componentType)) return false;
        var bomGrade = LookupBomGrade(componentType, bomGradeMap);
        return !string.IsNullOrEmpty(bomGrade)
            && string.Equals(grade, bomGrade, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns true if the grade can be overridden — it is either empty,
    /// a hardcoded CS default, or a BOM-sourced grade for the current component.
    /// </summary>
    private static bool IsOverridableGrade(string? grade, string componentType, Dictionary<string, string> bomGradeMap)
    {
        if (string.IsNullOrEmpty(grade)) return true;
        return IsHardcodedDefault(grade, componentType)
            || IsBomSourcedGrade(grade, componentType, bomGradeMap);
    }

    /// <summary>
    /// Looks up the grade for a component type in the BOM grade map.
    /// Tries the normalized name first, then all known aliases.
    /// Returns null if not found (caller falls back to hardcoded inference).
    /// </summary>
    private static string? LookupBomGrade(string componentType, Dictionary<string, string> bomGradeMap)
    {
        if (string.IsNullOrEmpty(componentType)) return null;
        var normalized = NormalizeBomComponentName(componentType);

        // Direct lookup
        if (bomGradeMap.TryGetValue(normalized, out var grade) && !string.IsNullOrEmpty(grade))
            return grade;

        // Try all aliases for this component type
        foreach (var alias in GetComponentAliases(normalized))
        {
            if (bomGradeMap.TryGetValue(alias, out var aliasGrade) && !string.IsNullOrEmpty(aliasGrade))
                return aliasGrade;
        }

        return null;
    }

    /// <summary>
    /// Refines Material A/B and Grade A/B for each weld using BOM item data.
    /// Uses two complementary strategies:
    ///
    /// <b>Strategy 1: Size-based BOM matching (most reliable)</b>
    /// Matches each weld's NPS to BOM items at the same size. The weld type
    /// determines the expected component pairing (BW→PIPE+fitting, SW→PIPE+COUPLING,
    /// SOF→PIPE+FLANGE, etc.). When only one fitting type exists in the BOM at the
    /// weld's NPS, Material B can be set unambiguously.
    ///
    /// <b>Strategy 2: PT NO text proximity (supplementary)</b>
    /// Scans the drawing body text for standalone PT NO annotations (small squared
    /// numbers next to components) and associates the 2 nearest PT NOs with each weld
    /// by OCR text line proximity. Less reliable due to 2D→1D text flattening.
    ///
    /// Priority: explicit weld table data &gt; size-based BOM &gt; PT NO proximity &gt; defaults.
    /// </summary>
    private static void AssociatePtNosWithWelds(
        List<WeldRecord> welds,
        List<MaterialInfo> bomItems,
        Dictionary<string, int> weldLineIndices,
        string[] lines,
        Dictionary<string, string> bomGradeMap,
        List<string> warnings)
    {
        if (welds.Count == 0) return;

        // ═══════════════════════════════════════════════════════════════════
        // Strategy 1: Size-based BOM matching (requires bomItems)
        // ═══════════════════════════════════════════════════════════════════

        // Build NPS → List<MaterialInfo> lookup (only items with known size and type)
        var bomBySize = new Dictionary<string, List<MaterialInfo>>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in bomItems)
        {
            if (string.IsNullOrEmpty(item.Size) || string.IsNullOrEmpty(item.MaterialType))
                continue;
            var normalizedSize = NormalizeFractionalSize(item.Size);
            if (string.IsNullOrEmpty(normalizedSize)) continue;

            if (!bomBySize.TryGetValue(normalizedSize, out var list))
            {
                list = [];
                bomBySize[normalizedSize] = list;
            }
            list.Add(item);
        }

        // Build PT NO → MaterialInfo lookup for Strategy 2
        var ptNoLookup = new Dictionary<string, MaterialInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in bomItems)
        {
            if (!string.IsNullOrEmpty(item.PartNo) && !string.IsNullOrEmpty(item.MaterialType))
                ptNoLookup.TryAdd(item.PartNo, item);
        }

        // Build quantity-aware weld-connection budget per NPS per fitting type.
        // Each fitting consumes N BW weld connections based on its geometry
        // (ELBOW=2, TEE=3, FLANGE=1, REDUCER=2, etc.). The budget tracks how
        // many weld connections are still available for each fitting type at
        // each NPS, enabling distribution across BW welds instead of assigning
        // the same fitting to every weld at the same pipe size.
        var remainingWeldConnections = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        // Component types that are "fittings" (not pipe) — used to find Material B.
        // PAD is excluded: it's a support pad that only connects via SP welds,
        // which are handled by InferMaterialFromWeldType. Including PAD here
        // causes SUPPORTS BOM items to be selected as Material B for BW/FW welds.
        var fittingTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ELBOW", "TEE", "REDUCER", "FLANGE", "VALVE", "OLET", "NIPPLE",
            "CAP", "PLUG", "COUPLING", "UNION", "SWAGE", "HOSE", "PLATE"
        };

        // Populate the weld-connection budget from BOM items
        foreach (var item in bomItems)
        {
            var nps = NormalizeFractionalSize(item.Size ?? "");
            var type = NormalizeBomComponentName(item.MaterialType ?? "");
            if (string.IsNullOrEmpty(nps) || string.IsNullOrEmpty(type) || !fittingTypes.Contains(type))
                continue;

            if (!remainingWeldConnections.TryGetValue(nps, out var typeMap))
            {
                typeMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                remainingWeldConnections[nps] = typeMap;
            }

            int multiplier = GetFittingWeldMultiplier(type);
            typeMap.TryGetValue(type, out var existing);
            typeMap[type] = existing + (item.Quantity * multiplier);
        }

        if (remainingWeldConnections.Count > 0)
        {
            var budgetSummary = remainingWeldConnections
                .SelectMany(kv => kv.Value.Select(tv => $"{tv.Key}@{kv.Key}\"={tv.Value}"))
                .Take(15);
            warnings.Add($"BOM weld-connection budget: {string.Join(", ", budgetSummary)}");
        }

        int sizeMatchCount = 0;
        foreach (var weld in welds)
        {
            if (string.IsNullOrEmpty(weld.Number)) continue;
            var needsRefinement = NeedsMaterialGradeRefinement(weld);
            // FJ welds always have fixed materials — skip
            if ((weld.Type ?? "").Equals("FJ", StringComparison.OrdinalIgnoreCase)) continue;
            // Skip only when explicit table data is already complete
            if (weld.ExplicitTableData && !needsRefinement) continue;

            var weldSize = weld.Size;
            if (string.IsNullOrEmpty(weldSize)) continue;

            var normalizedWeldSize = NormalizeFractionalSize(weldSize);
            if (string.IsNullOrEmpty(normalizedWeldSize) ||
                !bomBySize.TryGetValue(normalizedWeldSize, out var sizeItems))
                continue;

            // Get default materials for this weld type
            var (defaultMatA, defaultMatB) = InferMaterialFromWeldType(weld.Type);

            // Separate into PIPE and fitting items at this NPS
            var pipeItems = sizeItems
                .Where(i => NormalizeBomComponentName(i.MaterialType ?? "")
                    .Equals("PIPE", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var fittingItems = sizeItems
                .Where(i => fittingTypes.Contains(NormalizeBomComponentName(i.MaterialType ?? "")))
                .ToList();

            // Material A: use PIPE item at this NPS if available
            if (pipeItems.Count > 0)
            {
                var pipeItem = pipeItems[0];
                if (string.IsNullOrEmpty(weld.MaterialA) ||
                    string.Equals(weld.MaterialA, defaultMatA, StringComparison.OrdinalIgnoreCase))
                {
                    weld.MaterialA = NormalizeBomComponentName(pipeItem.MaterialType ?? "");
                }
                if (!string.IsNullOrEmpty(pipeItem.Grade) &&
                    (string.IsNullOrEmpty(weld.GradeA) || IsHardcodedDefault(weld.GradeA, weld.MaterialA ?? "")))
                {
                    weld.GradeA = pipeItem.Grade;
                }
            }

            // Material B: find the best fitting at this NPS
            if (fittingItems.Count > 0 &&
                (string.IsNullOrEmpty(weld.MaterialB) ||
                 string.Equals(weld.MaterialB, defaultMatB, StringComparison.OrdinalIgnoreCase)))
            {
                MaterialInfo? bestFitting = null;

                // If only one fitting type at this NPS → unambiguous
                var distinctFittingTypes = fittingItems
                    .Select(i => NormalizeBomComponentName(i.MaterialType ?? ""))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (distinctFittingTypes.Count == 1)
                {
                    // Single fitting type at this NPS — use it if the weld-connection
                    // budget still has capacity. This prevents over-assigning when the
                    // BOM says e.g. 3 elbows (= 6 connections) but there are 10 BW welds.
                    var singleType = distinctFittingTypes[0];
                    bool hasBudget = true;
                    if (remainingWeldConnections.TryGetValue(normalizedWeldSize, out var singleBudget)
                        && singleBudget.TryGetValue(singleType, out var singleRemaining))
                    {
                        hasBudget = singleRemaining > 0;
                    }

                    if (hasBudget)
                        bestFitting = fittingItems[0];
                }
                else
                {
                    // Multiple fitting types at this NPS → use weld type to prefer one
                    string? preferredType = weld.Type?.ToUpperInvariant() switch
                    {
                        "SOF" => "FLANGE",
                        "BR" or "LET" => "OLET",
                        "SW" => null,  // SW can be COUPLING, ELBOW, TEE — ambiguous
                        "BW" => null,  // BW can be any fitting
                        _ => null
                    };

                    if (preferredType != null)
                    {
                        bestFitting = fittingItems.FirstOrDefault(i =>
                            NormalizeBomComponentName(i.MaterialType ?? "")
                                .Equals(preferredType, StringComparison.OrdinalIgnoreCase));
                    }

                    // For BW/FW welds, distribute fittings using the weld-connection
                    // budget. Pick the fitting type with the most remaining connections
                    // rather than the most frequent BOM item. This ensures different
                    // welds get different fitting types when the BOM has e.g. 3 elbows
                    // and 2 flanges at the same NPS.
                    if (bestFitting == null && weld.Type?.ToUpperInvariant() is "BW" or "FW")
                    {
                        if (remainingWeldConnections.TryGetValue(normalizedWeldSize, out var connMap))
                        {
                            var bestType = connMap
                                .Where(kv => kv.Value > 0 && fittingTypes.Contains(kv.Key))
                                .OrderByDescending(kv => kv.Value)
                                .Select(kv => kv.Key)
                                .FirstOrDefault();

                            if (bestType != null)
                            {
                                bestFitting = fittingItems.FirstOrDefault(i =>
                                    NormalizeBomComponentName(i.MaterialType ?? "")
                                        .Equals(bestType, StringComparison.OrdinalIgnoreCase));
                            }
                        }

                        // Fallback: most frequent fitting type if no budget data
                        bestFitting ??= fittingItems
                            .GroupBy(i => NormalizeBomComponentName(i.MaterialType ?? ""), StringComparer.OrdinalIgnoreCase)
                            .OrderByDescending(g => g.Count())
                            .First()
                            .First();
                    }
                }

                if (bestFitting != null)
                {
                    // Decrement the weld-connection budget for the assigned fitting
                    var assignedType = NormalizeBomComponentName(bestFitting.MaterialType ?? "");
                    if (remainingWeldConnections.TryGetValue(normalizedWeldSize, out var budget)
                        && budget.TryGetValue(assignedType, out var budgetRemaining)
                        && budgetRemaining > 0)
                    {
                        budget[assignedType] = budgetRemaining - 1;
                    }

                    var oldMatB = weld.MaterialB ?? "";
                    weld.MaterialB = assignedType;
                    // Check grade overridability against the OLD material type.
                    // The current grade may have been set for the previous component
                    // (e.g. A333-6 for PIPE) and must be updated when the component
                    // changes (e.g. to ELBOW with A420-WPL6).
                    if (!string.IsNullOrEmpty(bestFitting.Grade) &&
                        (string.IsNullOrEmpty(weld.GradeB) || IsOverridableGrade(weld.GradeB, oldMatB, bomGradeMap)))
                    {
                        weld.GradeB = bestFitting.Grade;
                    }
                }
            }

            sizeMatchCount++;
        }

        if (sizeMatchCount > 0)
        {
            warnings.Add($"BOM size matching: refined Material/Grade for {sizeMatchCount} weld(s) using NPS-matched BOM items.");
        }

        // ═══════════════════════════════════════════════════════════════════
        // Strategy 2: PT NO text proximity (supplementary, requires bomItems)
        // ═══════════════════════════════════════════════════════════════════
        if (ptNoLookup.Count > 0 && weldLineIndices.Count > 0)
        {

        var knownPtNos = new HashSet<string>(ptNoLookup.Keys, StringComparer.OrdinalIgnoreCase);

        // Scan drawing body lines for PT NO annotations (standalone small numbers)
        var ptNoPositions = new List<(int LineIndex, string PtNo)>();
        bool inBom = false, inCuttingList = false, inWeldList = false;
        int sectionBlank = 0;

        for (int li = 0; li < lines.Length; li++)
        {
            var trimmedLine = lines[li].Trim();

            // Track section boundaries
            if (PdfBomSectionHeaderRegex().IsMatch(trimmedLine))
            { inBom = true; sectionBlank = 0; continue; }
            if (PdfCuttingListHeaderRegex().IsMatch(trimmedLine))
            { inCuttingList = true; sectionBlank = 0; continue; }
            if (PdfWeldListHeaderRegex().IsMatch(trimmedLine))
            { inWeldList = true; sectionBlank = 0; continue; }

            if (string.IsNullOrWhiteSpace(trimmedLine))
            {
                sectionBlank++;
                if (sectionBlank > 8) { inBom = false; inCuttingList = false; inWeldList = false; }
                continue;
            }
            sectionBlank = 0;

            // Skip lines inside structured table sections
            if (inBom || inCuttingList || inWeldList) continue;

            // Skip elevation, continuation, NPS lines
            if (trimmedLine.Contains("EL+", StringComparison.OrdinalIgnoreCase) ||
                trimmedLine.Contains("EL-", StringComparison.OrdinalIgnoreCase) ||
                trimmedLine.Contains("EL=", StringComparison.OrdinalIgnoreCase) ||
                trimmedLine.Contains("CONT.", StringComparison.OrdinalIgnoreCase) ||
                trimmedLine.Contains("NPS", StringComparison.OrdinalIgnoreCase))
                continue;

            // PT NO annotations on drawing body appear on short-to-medium lines
            if (trimmedLine.Length > 120) continue;

            foreach (Match ptMatch in PdfPtNoAnnotationRegex().Matches(trimmedLine))
            {
                var candidate = ptMatch.Groups["ptno"].Value;
                if (!knownPtNos.Contains(candidate)) continue;

                int startPos = ptMatch.Index;
                int endPos = startPos + ptMatch.Length;

                // Filter out false positives
                if (startPos >= 2)
                {
                    var prefix = trimmedLine[(startPos - 2)..startPos].ToUpperInvariant();
                    if (prefix.EndsWith("EL", StringComparison.Ordinal) ||
                        prefix.EndsWith("SP", StringComparison.Ordinal))
                        continue;
                }
                if (startPos >= 1)
                {
                    var prevChar = trimmedLine[startPos - 1];
                    if (char.IsDigit(prevChar) || prevChar == '.') continue;
                }
                if (endPos < trimmedLine.Length)
                {
                    var nextChar = trimmedLine[endPos];
                    if (char.IsDigit(nextChar) || nextChar is '"' or '\'' or '/' or '.') continue;
                }

                ptNoPositions.Add((li, candidate));
            }
        }

        if (ptNoPositions.Count == 0) return;

        // For welds that STILL have default materials after Strategy 1,
        // try PT NO proximity as a supplementary refinement
        int proximityCount = 0;
        foreach (var weld in welds)
        {
            if (string.IsNullOrEmpty(weld.Number)) continue;
            var needsRefinement = NeedsMaterialGradeRefinement(weld);
            if ((weld.Type ?? "").Equals("FJ", StringComparison.OrdinalIgnoreCase)) continue;
            // When Material/Grade came from the structured weld list table, skip proximity override only if complete
            if (weld.ExplicitTableData && !needsRefinement) continue;

            var (defaultMatA, defaultMatB) = InferMaterialFromWeldType(weld.Type);

            // Only apply proximity if material is still at defaults
            bool matAIsDefault = string.IsNullOrEmpty(weld.MaterialA) ||
                string.Equals(weld.MaterialA, defaultMatA, StringComparison.OrdinalIgnoreCase);
            bool matBIsDefault = string.IsNullOrEmpty(weld.MaterialB) ||
                string.Equals(weld.MaterialB, defaultMatB, StringComparison.OrdinalIgnoreCase);

            bool fillGradeA = IsOverridableGrade(weld.GradeA, weld.MaterialA ?? "", bomGradeMap);
            bool fillGradeB = IsOverridableGrade(weld.GradeB, weld.MaterialB ?? "", bomGradeMap);

            if (!needsRefinement && !matAIsDefault && !matBIsDefault && !fillGradeA && !fillGradeB)
                continue;

            if (!weldLineIndices.TryGetValue(weld.Number, out var weldLineIdx))
            {
                if (weld.SourceLineIndex >= 0) weldLineIdx = weld.SourceLineIndex;
                else continue;
            }

            var nearestPtNos = ptNoPositions
                .Select(p => new { p.PtNo, p.LineIndex, Distance = Math.Abs(p.LineIndex - weldLineIdx) })
                .OrderBy(p => p.Distance)
                .Take(6)
                .ToList();

            if (nearestPtNos.Count == 0) continue;

            MaterialInfo? itemA = null;
            MaterialInfo? itemB = null;

            foreach (var nearest in nearestPtNos)
            {
                if (!ptNoLookup.TryGetValue(nearest.PtNo, out var item)) continue;
                var normalizedType = NormalizeBomComponentName(item.MaterialType ?? "");

                // Skip support items (PAD) in proximity matching — SUPPORTS BOM
                // items have PT NOs on the drawing near pipe supports, and their
                // proximity to weld callouts causes PAD to be incorrectly assigned
                // as Material B for BW/FW welds. SP welds already get PAD from
                // InferMaterialFromWeldType.
                if (string.Equals(normalizedType, "PAD", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (itemA == null) { itemA = item; continue; }

                var typeA = NormalizeBomComponentName(itemA.MaterialType ?? "");
                if (!string.Equals(normalizedType, typeA, StringComparison.OrdinalIgnoreCase))
                {
                    itemB = item;
                    break;
                }
                itemB ??= item;
            }

            if (itemA == null) continue;

            bool changed = false;
            if (matAIsDefault)
            {
                var ptCompA = NormalizeBomComponentName(itemA.MaterialType ?? "");
                if (!string.IsNullOrEmpty(ptCompA))
                {
                    weld.MaterialA = ptCompA;
                    if (fillGradeA)
                    {
                        if (!string.IsNullOrEmpty(itemA.Grade))
                            weld.GradeA = itemA.Grade;
                        else
                        {
                            // Material changed but BOM item has no grade — re-infer
                            var bg = LookupBomGrade(ptCompA, bomGradeMap);
                            weld.GradeA = !string.IsNullOrEmpty(bg) ? bg : InferGradeFromComponent(ptCompA);
                        }
                    }
                    changed = true;
                }
            }
            else if (fillGradeA && !string.IsNullOrEmpty(itemA.Grade))
            {
                weld.GradeA = itemA.Grade;
                changed = true;
            }

            if (matBIsDefault && itemB != null)
            {
                var ptCompB = NormalizeBomComponentName(itemB.MaterialType ?? "");
                if (!string.IsNullOrEmpty(ptCompB))
                {
                    weld.MaterialB = ptCompB;
                    if (fillGradeB)
                    {
                        if (!string.IsNullOrEmpty(itemB.Grade))
                            weld.GradeB = itemB.Grade;
                        else
                        {
                            // Material changed but BOM item has no grade — re-infer
                            var bg = LookupBomGrade(ptCompB, bomGradeMap);
                            weld.GradeB = !string.IsNullOrEmpty(bg) ? bg : InferGradeFromComponent(ptCompB);
                        }
                    }
                    changed = true;
                }
            }
            else if (itemB != null && fillGradeB && !string.IsNullOrEmpty(itemB.Grade))
            {
                weld.GradeB = itemB.Grade;
                changed = true;
            }

            if (changed) proximityCount++;
        }

        if (proximityCount > 0)
        {
            warnings.Add($"PT NO proximity: refined {proximityCount} weld(s) with drawing body annotations.");
        }

        } // end Strategy 2 if-block

        // ═══════════════════════════════════════════════════════════════════
        // Strategy 3: Connectivity-based chain propagation
        // ═══════════════════════════════════════════════════════════════════
        // In a piping isometric, consecutive welds share a component:
        //   PIPE ─[6]─ TEE ─[7]─ FLANGE       TEE ─[8]─ ELBOW ─[9]─ PIPE
        // Weld 6's Material B (TEE) becomes Weld 7's Material A.
        // Weld 8's Material B (ELBOW) becomes Weld 9's Material A.
        //
        // Uses three phases:
        //   Phase 1: FJ adjacency — BW weld next to FJ must have FLANGE
        //   Phase 2: Forward chain with TEE branch tracking
        //   Phase 3: Reverse chain (skips TEE branch boundaries)
        int chainCount = 0;
        var chainableTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "BW", "FW"
        };
        // Terminal fittings end a pipe run — they have only one BW
        // connection, so they must NOT propagate forward across a branch.
        var terminalFittings = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "FLANGE", "CAP", "VALVE"
        };

        // ── Phase 1: FJ adjacency — identify FLANGE connections ──
        for (int i = 0; i < welds.Count; i++)
        {
            if (!(welds[i].Type ?? "").Equals("FJ", StringComparison.OrdinalIgnoreCase))
                continue;

            // BW weld immediately before the FJ → its B-side is FLANGE
            if (i > 0)
            {
                var prev = welds[i - 1];
                if (chainableTypes.Contains(prev.Type ?? "")
                    && !(prev.ExplicitTableData && !NeedsMaterialGradeRefinement(prev)))
                {
                    var (_, defMatB) = InferMaterialFromWeldType(prev.Type);
                    if (string.IsNullOrEmpty(prev.MaterialB) ||
                        string.Equals(prev.MaterialB, defMatB, StringComparison.OrdinalIgnoreCase))
                    {
                        prev.MaterialB = "FLANGE";
                        if (string.IsNullOrEmpty(prev.GradeB) || IsHardcodedDefault(prev.GradeB, defMatB))
                        {
                            var bg = LookupBomGrade("FLANGE", bomGradeMap);
                            if (!string.IsNullOrEmpty(bg)) prev.GradeB = bg;
                        }
                        chainCount++;
                    }
                }
            }

            // BW weld immediately after the FJ → its A-side is FLANGE
            if (i + 1 < welds.Count)
            {
                var next = welds[i + 1];
                if (chainableTypes.Contains(next.Type ?? "")
                    && !(next.ExplicitTableData && !NeedsMaterialGradeRefinement(next)))
                {
                    var (defMatA, _) = InferMaterialFromWeldType(next.Type);
                    if (string.IsNullOrEmpty(next.MaterialA) ||
                        string.Equals(next.MaterialA, defMatA, StringComparison.OrdinalIgnoreCase))
                    {
                        next.MaterialA = "FLANGE";
                        if (string.IsNullOrEmpty(next.GradeA) || IsHardcodedDefault(next.GradeA, defMatA))
                        {
                            var bg = LookupBomGrade("FLANGE", bomGradeMap);
                            if (!string.IsNullOrEmpty(bg)) next.GradeA = bg;
                        }
                        chainCount++;
                    }
                }
            }
        }

        // ── Phase 2: Forward chain with TEE branch tracking ──
        // Track the most recently encountered TEE and how many of its
        // 3 connections have been assigned. When a chain break occurs
        // (terminal fitting or PIPE) but the active TEE still has free
        // connections, the next BW weld starts from the TEE branch.
        string? activeTeeNps = null;
        int activeTeeUsed = 0;
        var teeBranchStarts = new HashSet<int>(); // indices assigned via TEE branch

        for (int i = 0; i < welds.Count; i++)
        {
            var curr = welds[i];
            if (!chainableTypes.Contains(curr.Type ?? "")) continue;
            if (curr.ExplicitTableData && !NeedsMaterialGradeRefinement(curr)) continue;

            var (defaultMatA, _) = InferMaterialFromWeldType(curr.Type);
            var currNps = NormalizeFractionalSize(curr.Size ?? "");
            bool matAIsDefault = string.IsNullOrEmpty(curr.MaterialA) ||
                string.Equals(curr.MaterialA, defaultMatA, StringComparison.OrdinalIgnoreCase);

            // ── Try forward propagation from nearest BW/FW weld ──
            if (matAIsDefault && i > 0)
            {
                // Search backward for the nearest BW/FW weld — skip non-chainable
                // welds (BR, LET, FJ, SOF, etc.) that sit between BW welds and
                // would otherwise contaminate the chain with their fixed pairings.
                WeldRecord? prevBw = null;
                for (int j = i - 1; j >= Math.Max(0, i - 5); j--)
                {
                    if (chainableTypes.Contains(welds[j].Type ?? ""))
                    {
                        prevBw = welds[j];
                        break;
                    }
                }

                var prevMatB = (prevBw?.MaterialB ?? "").ToUpperInvariant();

                if (prevBw != null &&
                    !string.IsNullOrEmpty(prevMatB) &&
                    !string.Equals(prevMatB, "PIPE", StringComparison.OrdinalIgnoreCase) &&
                    !terminalFittings.Contains(prevMatB))
                {
                    // Linear chain: prev.MaterialB → curr.MaterialA
                    curr.MaterialA = prevMatB;
                    if (IsOverridableGrade(curr.GradeA, defaultMatA, bomGradeMap))
                    {
                        var bg = LookupBomGrade(prevMatB, bomGradeMap);
                        curr.GradeA = !string.IsNullOrEmpty(bg) ? bg
                            : !string.IsNullOrEmpty(prevBw.GradeB) ? prevBw.GradeB
                            : curr.GradeA;
                    }
                    chainCount++;
                    matAIsDefault = false;
                }
                else if (activeTeeNps != null && activeTeeUsed < 3 &&
                         (string.IsNullOrEmpty(currNps) ||
                          string.Equals(currNps, activeTeeNps, StringComparison.OrdinalIgnoreCase)))
                {
                    // TEE branch: the active TEE still has unused connections
                    curr.MaterialA = "TEE";
                    if (IsOverridableGrade(curr.GradeA, defaultMatA, bomGradeMap))
                    {
                        var bg = LookupBomGrade("TEE", bomGradeMap);
                        if (!string.IsNullOrEmpty(bg)) curr.GradeA = bg;
                    }
                    activeTeeUsed++;
                    teeBranchStarts.Add(i);
                    chainCount++;
                    matAIsDefault = false;
                }
            }

            // ── Track TEE connections ──
            var matA = (curr.MaterialA ?? "").ToUpperInvariant();
            var matB = (curr.MaterialB ?? "").ToUpperInvariant();

            if (string.Equals(matA, "TEE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(matB, "TEE", StringComparison.OrdinalIgnoreCase))
            {
                // Start or continue tracking a TEE at this NPS
                if (activeTeeNps == null ||
                    (!string.IsNullOrEmpty(currNps) &&
                     !string.Equals(currNps, activeTeeNps, StringComparison.OrdinalIgnoreCase)))
                {
                    activeTeeNps = !string.IsNullOrEmpty(currNps) ? currNps : activeTeeNps;
                    activeTeeUsed = 0;
                }
                if (string.Equals(matA, "TEE", StringComparison.OrdinalIgnoreCase)) activeTeeUsed++;
                if (string.Equals(matB, "TEE", StringComparison.OrdinalIgnoreCase)) activeTeeUsed++;
            }

            // TEE fully consumed → reset
            if (activeTeeUsed >= 3)
            {
                activeTeeNps = null;
                activeTeeUsed = 0;
            }
        }

        // ── Phase 3: Reverse chain propagation ──
        // Walk welds right-to-left. If the nearest downstream BW/FW weld's
        // MaterialA is a non-PIPE fitting and curr.MaterialB is at default,
        // propagate backward. Skip non-chainable welds and TEE branch boundaries.
        for (int i = welds.Count - 2; i >= 0; i--)
        {
            var curr = welds[i];

            if (!chainableTypes.Contains(curr.Type ?? "")) continue;
            if (curr.ExplicitTableData && !NeedsMaterialGradeRefinement(curr)) continue;

            // Search forward for the nearest BW/FW weld
            WeldRecord? nextBw = null;
            int nextBwIdx = -1;
            for (int j = i + 1; j < Math.Min(welds.Count, i + 6); j++)
            {
                if (chainableTypes.Contains(welds[j].Type ?? ""))
                {
                    nextBw = welds[j];
                    nextBwIdx = j;
                    break;
                }
            }
            if (nextBw == null) continue;

            // Don't reverse-propagate across a TEE branch boundary.
            if (teeBranchStarts.Contains(nextBwIdx)) continue;

            var nextMatA = (nextBw.MaterialA ?? "").ToUpperInvariant();
            if (string.IsNullOrEmpty(nextMatA)) continue;
            if (string.Equals(nextMatA, "PIPE", StringComparison.OrdinalIgnoreCase)) continue;

            var (_, defaultMatB) = InferMaterialFromWeldType(curr.Type);
            bool matBIsDefault = string.IsNullOrEmpty(curr.MaterialB) ||
                string.Equals(curr.MaterialB, defaultMatB, StringComparison.OrdinalIgnoreCase);

            if (!matBIsDefault) continue;

            curr.MaterialB = nextMatA;

            if (IsOverridableGrade(curr.GradeB, defaultMatB, bomGradeMap))
            {
                var bg = LookupBomGrade(nextMatA, bomGradeMap);
                curr.GradeB = !string.IsNullOrEmpty(bg) ? bg
                    : !string.IsNullOrEmpty(nextBw.GradeA) ? nextBw.GradeA
                    : curr.GradeB;
            }

            chainCount++;
        }

        if (chainCount > 0)
        {
            warnings.Add($"Chain propagation: refined Material/Grade for {chainCount} weld(s) using component connectivity.");
        }

        // ── Phase 4: Normalize Material A/B convention ──
        // After all strategies, ensure Material A = PIPE (pipe side) and
        // Material B = fitting for BW/FW/SW/TH welds. Chain propagation
        // may have set Material A to a fitting (topologically correct but
        // violates the display convention). Swap when needed.
        NormalizeMaterialABConvention(welds, bomGradeMap);
    }

    /// <summary>
    /// Refines Material B assignments using cutting list pipe-piece topology.
    /// Within a spool, each pipe piece (from the cutting list) is connected to
    /// the next via a fitting. Shop BW welds that still show PIPE/PIPE likely
    /// connect to a fitting, not pipe-to-pipe, because true pipe-to-pipe butt
    /// welds within a spool are rare (they occur only when a straight run is
    /// longer than a single pipe length).
    ///
    /// Algorithm:
    /// 1. Count pipe pieces per spool from the cutting list.
    /// 2. Count WS BW welds per spool with PIPE/PIPE assignment.
    /// 3. The number of internal fittings ≈ (pipe pieces - 1).
    /// 4. If there are more PIPE/PIPE WS BW welds than expected pipe-to-pipe
    ///    joints, some of them should be pipe-to-fitting instead.
    /// 5. Uses cutting list END codes (TH→COUPLING, SW→COUPLING) to
    ///    validate weld type-specific Material B assignments.
    /// </summary>
    private static void RefineWeldMaterialsFromCuttingList(
        List<WeldRecord> welds,
        List<CuttingListEntry> cuttingListEntries,
        Dictionary<string, string> spoolSizes,
        Dictionary<string, string> bomGradeMap,
        List<string> warnings)
    {
        if (welds.Count == 0 || cuttingListEntries.Count == 0) return;

        // Group cutting list entries by normalized spool number
        var spoolPieces = new Dictionary<string, List<CuttingListEntry>>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in cuttingListEntries)
        {
            if (string.IsNullOrEmpty(entry.SpoolNo)) continue;
            var spoolMatch = SpoolTagRegex().Match(entry.SpoolNo);
            var normalizedSpool = spoolMatch.Success
                ? spoolMatch.Groups[1].Value.PadLeft(2, '0')
                : entry.SpoolNo;

            if (!spoolPieces.TryGetValue(normalizedSpool, out var pieces))
            {
                pieces = [];
                spoolPieces[normalizedSpool] = pieces;
            }
            pieces.Add(entry);
        }

        int refinedCount = 0;

        foreach (var (spool, pieces) in spoolPieces)
        {
            // Count END codes that imply specific fitting types
            int thEndCount = 0;  // Threaded → COUPLING
            int swEndCount = 0;  // Socket weld → COUPLING

            foreach (var piece in pieces)
            {
                var end1 = (piece.End1 ?? "").ToUpperInvariant();
                var end2 = (piece.End2 ?? "").ToUpperInvariant();
                if (end1 == "TH") thEndCount++;
                if (end2 == "TH") thEndCount++;
                if (end1 is "SW" or "SOW") swEndCount++;
                if (end2 is "SW" or "SOW") swEndCount++;
            }

            // Get welds in this spool
            var spoolWelds = welds
                .Where(w => string.Equals(w.SpoolNo, spool, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Refine TH welds: ensure they have COUPLING (not PIPE)
            foreach (var w in spoolWelds)
            {
                if (!(w.Type ?? "").Equals("TH", StringComparison.OrdinalIgnoreCase)) continue;
                if (HasAuthoritativeTableData(w)) continue;

                var matB = (w.MaterialB ?? "").ToUpperInvariant();
                if (matB is "" or "PIPE")
                {
                    w.MaterialB = "COUPLING";
                    if (string.IsNullOrEmpty(w.GradeB) || IsHardcodedDefault(w.GradeB, "PIPE"))
                    {
                        w.GradeB = LookupBomGrade("COUPLING", bomGradeMap)
                            ?? InferGradeFromComponent("COUPLING");
                    }
                    refinedCount++;
                }
            }

            // Refine SW welds: ensure they have COUPLING
            foreach (var w in spoolWelds)
            {
                if (!(w.Type ?? "").Equals("SW", StringComparison.OrdinalIgnoreCase)) continue;
                if (HasAuthoritativeTableData(w)) continue;

                var matB = (w.MaterialB ?? "").ToUpperInvariant();
                if (matB is "" or "PIPE")
                {
                    w.MaterialB = "COUPLING";
                    if (string.IsNullOrEmpty(w.GradeB) || IsHardcodedDefault(w.GradeB, "PIPE"))
                    {
                        w.GradeB = LookupBomGrade("COUPLING", bomGradeMap)
                            ?? InferGradeFromComponent("COUPLING");
                    }
                    refinedCount++;
                }
            }

            // Pipe-piece count heuristic for BW shop welds:
            // Within a spool of N pipe pieces, there are approximately (N-1) internal
            // fittings connecting consecutive pieces. If there are M BW WS welds with
            // PIPE/PIPE and M > 0, at least some of them should be PIPE/FITTING.
            // The number of true pipe-to-pipe joints ≈ max(0, M - 2*(N-1))
            // because each fitting consumes ~2 BW connections.
            int pieceCount = pieces.Count;
            if (pieceCount > 1)
            {
                var bwPipePipeWelds = spoolWelds
                    .Where(w => (w.Type ?? "").Equals("BW", StringComparison.OrdinalIgnoreCase)
                        && (w.Location ?? "").Equals("WS", StringComparison.OrdinalIgnoreCase)
                        && (w.MaterialB ?? "").Equals("PIPE", StringComparison.OrdinalIgnoreCase)
                        && !HasAuthoritativeTableData(w))
                    .ToList();

                if (bwPipePipeWelds.Count > 0)
                {
                    // Estimate minimum fittings: a spool with N pieces has ≥ N-1
                    // connections, of which most go through fittings. Mark those
                    // PIPE/PIPE WS BW welds as needing fitting refinement — the
                    // chain propagation and downstream NormalizeMaterialABConvention
                    // will determine the specific fitting type.
                    int maxPipeToPipe = Math.Max(0, bwPipePipeWelds.Count - (pieceCount - 1));
                    int toRefine = bwPipePipeWelds.Count - maxPipeToPipe;

                    // Clear Material B on welds that should have fittings,
                    // allowing downstream BOM/chain strategies to fill them.
                    for (int i = 0; i < toRefine && i < bwPipePipeWelds.Count; i++)
                    {
                        // Leave MaterialB empty so InferMaterialFromWeldType and
                        // the quantity-aware strategy can provide a better value.
                        // Don't clear if the grade suggests a specific component.
                        var w = bwPipePipeWelds[i];
                        if (!string.IsNullOrEmpty(w.GradeB))
                        {
                            var inferred = InferComponentFromGrade(w.GradeB);
                            if (!string.IsNullOrEmpty(inferred) && !inferred.Equals("PIPE", StringComparison.OrdinalIgnoreCase))
                            {
                                w.MaterialB = inferred;
                                refinedCount++;
                            }
                        }
                    }
                }
            }
        }

        if (refinedCount > 0)
        {
            warnings.Add($"Cutting list topology: refined Material B for {refinedCount} weld(s) using pipe-piece structure.");
        }
    }

    /// <summary>
    /// Normalizes the Material A/B display convention for BW/FW/SW/TH welds.
    /// Material A should always be the "pipe" side (PIPE, NIPPLE, SWAGE) and
    /// Material B should be the fitting side (ELBOW, FLANGE, TEE, etc.).
    /// When Material A is a fitting and Material B is pipe-like (or empty),
    /// swaps both Material and Grade values, then re-applies BOM grade lookup
    /// to ensure grades match the corrected materials.
    /// </summary>
    private static void NormalizeMaterialABConvention(List<WeldRecord> welds, Dictionary<string, string> bomGradeMap)
    {
        var pipeLikeTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "PIPE", "NIPPLE", "SWAGE"
        };
        var fittingTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ELBOW", "TEE", "FLANGE", "REDUCER", "VALVE", "OLET", "CAP",
            "PLUG", "COUPLING", "UNION", "HOSE", "PLATE", "PAD"
        };

        foreach (var weld in welds)
        {
            var type = (weld.Type ?? "").ToUpperInvariant();
            // Only normalize for weld types where Material A = PIPE is the convention
            if (type is "FJ") continue;
            if (type is not ("BW" or "FW" or "SW" or "TH" or "SOF" or "SP" or "BR" or "LET")) continue;
            // Don't touch fully populated explicit table data
            if (weld.ExplicitTableData && !NeedsMaterialGradeRefinement(weld)) continue;

            var matA = (weld.MaterialA ?? "").ToUpperInvariant();
            var matB = (weld.MaterialB ?? "").ToUpperInvariant();

            bool needsSwap = false;

            // Case 1: Material A is a fitting and Material B is pipe-like → swap
            if (fittingTypes.Contains(matA) && pipeLikeTypes.Contains(matB))
            {
                needsSwap = true;
            }
            // Case 2: Both are fittings that differ → force Material A to PIPE
            else if (fittingTypes.Contains(matA) && fittingTypes.Contains(matB)
                     && !string.Equals(matA, matB, StringComparison.OrdinalIgnoreCase))
            {
                weld.MaterialA = "PIPE";
                weld.GradeA = LookupBomGrade("PIPE", bomGradeMap)
                    ?? InferGradeFromComponent("PIPE");
                continue;
            }
            // Case 3: Material A is a fitting and Material B is empty → move to B side
            else if (fittingTypes.Contains(matA) && string.IsNullOrEmpty(matB))
            {
                weld.MaterialB = weld.MaterialA;
                weld.GradeB = weld.GradeA;
                weld.MaterialA = "PIPE";
                weld.GradeA = LookupBomGrade("PIPE", bomGradeMap)
                    ?? InferGradeFromComponent("PIPE");
                continue;
            }

            if (needsSwap)
            {
                (weld.MaterialA, weld.MaterialB) = (weld.MaterialB, weld.MaterialA);
                (weld.GradeA, weld.GradeB) = (weld.GradeB, weld.GradeA);

                // After swap, re-apply BOM grades to ensure consistency
                if (bomGradeMap.Count > 0)
                {
                    var bomGA = LookupBomGrade(weld.MaterialA ?? "", bomGradeMap);
                    if (!string.IsNullOrEmpty(bomGA))
                        weld.GradeA = bomGA;
                    var bomGB = LookupBomGrade(weld.MaterialB ?? "", bomGradeMap);
                    if (!string.IsNullOrEmpty(bomGB))
                        weld.GradeB = bomGB;
                }
            }
        }
    }

    /// <summary>
    /// Returns the number of BW weld connections a single instance of a fitting
    /// type consumes. Used to compute the total weld-connection budget from BOM
    /// quantities so that fittings are distributed across BW welds rather than
    /// assigning the same fitting type to every weld at the same NPS.
    ///
    /// Geometry rules:
    /// - ELBOW (2 ends):   2 BW welds (inlet + outlet)
    /// - TEE   (3 ends):   3 BW welds (run-in + run-out + branch)
    /// - REDUCER (2 ends): 2 BW welds (large end + small end)
    /// - FLANGE (1 weld):  1 BW weld  (weld neck side; other side is FJ)
    /// - CAP (1 end):      1 BW weld
    /// - NIPPLE (2 ends):  2 BW welds
    /// - COUPLING:         1 SW/TH weld per coupling end
    /// - OLET:             1 BR/LET weld
    /// - VALVE:            0 direct BW (connected via flanges → FJ)
    /// - PAD:              1 SP weld
    /// </summary>
    private static int GetFittingWeldMultiplier(string fittingType)
    {
        return (fittingType ?? "").ToUpperInvariant() switch
        {
            "ELBOW" => 2,
            "TEE" => 3,
            "REDUCER" => 2,
            "NIPPLE" => 2,
            "FLANGE" => 1,
            "CAP" => 1,
            "COUPLING" => 1,
            "OLET" => 1,
            "PAD" => 1,
            "PLUG" => 1,
            "UNION" => 2,
            "SWAGE" => 2,
            "VALVE" => 0,
            "HOSE" => 0,
            "PLATE" => 1,
            _ => 1
        };
    }

    /// <summary>
    /// Normalizes a BOM component name to match MATD_Description conventions.
    /// Maps piping component aliases (BRANCH, WELDOLET, SOCKOLET, etc.) to
    /// the standard MATD_Description values used in Material_Des_tbl.
    /// </summary>
    private static string NormalizeBomComponentName(string component)
    {
        var upper = component.Trim().ToUpperInvariant();
        // Normalize "STUD BOLT" variants
        if (upper.Contains("STUD") && upper.Contains("BOLT")) return "STUD BOLT";
        if (upper == "BOLT" || upper == "NUT") return "STUD BOLT";
        // Map piping component aliases to standard MATD_Description values
        return upper switch
        {
            "BRANCH" or "WELDOLET" or "NIPOLET" or "LATROLET" => "OLET",
            "SOCKOLET" or "THREADOLET" => "COUPLING",
            "HALF COUPLING" or "HALFCOUPLING" => "COUPLING",
            "WLD NECK" or "WELDNECK" or "WELD NECK" => "FLANGE",
            "CON REDUCER" or "ECC REDUCER" or "CONRED" or "ECCRED" => "REDUCER",
            "90 ELBOW" or "45 ELBOW" or "LR ELBOW" or "SR ELBOW" => "ELBOW",
            "EQUAL TEE" or "RED TEE" or "REDUCING TEE" => "TEE",
            "BLIND" or "BLIND FLANGE" => "FLANGE",
            "GATE" or "GLOBE" or "CHECK" or "BALL" or "BUTTERFLY" => "VALVE",
            _ => upper
        };
    }

    /// <summary>
    /// Returns all possible aliases for a standard MATD_Description component type.
    /// Used for BOM grade map lookups when the exact key doesn't match.
    /// </summary>
    private static IEnumerable<string> GetComponentAliases(string componentType)
    {
        var upper = (componentType ?? "").Trim().ToUpperInvariant();
        yield return upper;
        switch (upper)
        {
            case "OLET":
                yield return "BRANCH";
                yield return "WELDOLET";
                yield return "NIPOLET";
                yield return "LATROLET";
                break;
            case "COUPLING":
                yield return "SOCKOLET";
                yield return "THREADOLET";
                yield return "HALF COUPLING";
                break;
            case "FLANGE":
                yield return "WLD NECK";
                yield return "WELDNECK";
                yield return "WELD NECK";
                yield return "BLIND";
                yield return "BLIND FLANGE";
                break;
            case "REDUCER":
                yield return "CON REDUCER";
                yield return "ECC REDUCER";
                break;
            case "ELBOW":
                yield return "90 ELBOW";
                yield return "45 ELBOW";
                break;
            case "TEE":
                yield return "EQUAL TEE";
                yield return "RED TEE";
                yield return "REDUCING TEE";
                break;
            case "STUD BOLT":
                yield return "BOLT";
                yield return "NUT";
                break;
            case "VALVE":
                yield return "GATE";
                yield return "GLOBE";
                yield return "CHECK";
                yield return "BALL";
                break;
        }
    }

    /// <summary>
    /// Extracts a normalized ASTM grade value from a <see cref="PdfBomComponentGradeRegex"/> match.
    /// Handles patterns like "A106 Gr.B" → "A106-B", "A234-WPB" → "A234-WPB",
    /// "API 5L PSL2 GR.B" → "A106-B" (API 5L Grade B ≡ A106-B for CS).
    /// </summary>
    private static string ExtractBomGradeValue(Match match)
    {
        // Try named groups in priority order
        string raw;
        if (match.Groups["grade"].Success && !string.IsNullOrWhiteSpace(match.Groups["grade"].Value))
        {
            // API 5L pattern → map to MAT_GRADE values
            var apiGrade = match.Groups["grade"].Value.Trim().ToUpperInvariant();
            return apiGrade switch
            {
                "B" or "B1" => "API 5L B",
                "X42" or "X46" or "X52" or "X56" or "X60" or "X65" or "X70" => $"API 5L-{apiGrade}",
                _ => $"API 5L {apiGrade}"
            };
        }

        // EEMUA specification (Cu-Ni piping): "EEMUA 234" → "Cu-Ni UNS 7060X"
        // The MAT_GRADE value for EEMUA/Cu-Ni piping is "Cu-Ni UNS 7060X".
        if (match.Groups["grade4"].Success && !string.IsNullOrWhiteSpace(match.Groups["grade4"].Value))
        {
            return "Cu-Ni UNS 7060X";
        }

        // UNS alloy designation: "UNS C70600", "UNS 7060X" → "Cu-Ni UNS 7060X"
        // Cu-Ni UNS designations (C706xx / 7060X) map to the same MAT_GRADE.
        if (match.Groups["grade5"].Success && !string.IsNullOrWhiteSpace(match.Groups["grade5"].Value))
        {
            var uns = match.Groups["grade5"].Value.Trim().ToUpperInvariant();
            // Cu-Ni alloy family: UNS C70600, UNS 7060X, etc.
            if (uns.Contains("706", StringComparison.Ordinal) ||
                uns.Contains("C706", StringComparison.Ordinal))
            {
                return "Cu-Ni UNS 7060X";
            }
            // Return as-is for other UNS designations
            return WhitespaceRegex().Replace(uns, " ");
        }

        // BS EN standard: "BS EN 12451" → keep as-is with spaces
        if (match.Groups["grade6"].Success && !string.IsNullOrWhiteSpace(match.Groups["grade6"].Value))
        {
            var bsen = match.Groups["grade6"].Value.Trim().ToUpperInvariant();
            return WhitespaceRegex().Replace(bsen, " ");
        }

        raw = match.Groups["grade2"].Success ? match.Groups["grade2"].Value
            : match.Groups["grade3"].Success ? match.Groups["grade3"].Value
            : "";

        if (string.IsNullOrWhiteSpace(raw)) return "";

        return NormalizeAstmGrade(raw.Trim());
    }

    /// <summary>
    /// Normalizes an ASTM grade string to match Material_tbl.MAT_GRADE format.
    /// The DB uses mixed conventions:
    ///   - Hyphen between spec and grade letter: "A106-B", "A234-WPB", "A333-6"
    ///   - Space between spec and grade code:    "A193 B7", "A320 L7", "A193 B8/B8M"
    ///   - Hyphen with class suffix:             "A350-LF2 Cl.1", "A671-CC60 Cl.22"
    /// This method maps common OCR-extracted patterns to the correct DB format.
    /// </summary>
    private static string NormalizeAstmGrade(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";

        var s = raw.Trim().ToUpperInvariant();

        // Remove "GR." / "GR " / "GRADE" prefixes within the string
        s = AstmGradeCleanRegex().Replace(s, "-");

        // Collapse multiple spaces to single space
        s = WhitespaceRegex().Replace(s, " ");

        // Remove double hyphens
        s = MultiHyphenRegex().Replace(s, "-");

        // Trim leading/trailing hyphens and spaces
        s = s.Trim('-', ' ');

        // Map to known MAT_GRADE DB formats.
        // The DB uses specific conventions per ASTM spec:
        //   A193 B7, A193 B8/B8M, A320 L7, A320 B8/B8M → space separator
        //   A106-B, A234-WPB, A333-6, A350-LF2, A420-WPL6 → hyphen separator
        //   A182-F316/316L, A312-TP316/316L → hyphen separator
        //   A671-CC60 Cl.22, A350-LF2 Cl.1 → hyphen then space+Cl.
        // Apply known mappings:
        var mapped = MapToKnownGrade(s);
        return mapped ?? s;
    }

    /// <summary>
    /// Maps a normalized grade string to the exact MAT_GRADE value in the database.
    /// Returns null if no mapping is found (caller uses the input as-is).
    /// </summary>
    private static string? MapToKnownGrade(string normalized)
    {
        if (string.IsNullOrEmpty(normalized)) return null;

        // Direct common mappings from OCR patterns to DB values
        // The key is the normalized extracted value; the value is the exact DB MAT_GRADE.
        return normalized switch
        {
            // Carbon Steel pipe
            "A106-B" or "A106 B" => "A106-B",
            // CS fittings
            "A234-WPB" or "A234 WPB" => "A234-WPB",
            "A234-WPBW" or "A234 WPBW" => "A234-WPBW",
            // CS forged
            "A105" => "A105",
            // Stud bolt CS
            "A193" or "A193-B7" or "A193 B7" => "A193 B7",
            // Stud bolt SS
            "A193-B8" or "A193 B8" or "A193-B8/B8M" or "A193 B8/B8M"
                or "A193-B8M" or "A193 B8M" => "A193 B8/B8M",
            // Low-temp CS
            "A333-6" or "A333 6" => "A333-6",
            "A350-LF2" or "A350 LF2" => "A350-LF2",
            "A350-LF2-CL.1" or "A350-LF2 CL.1" or "A350 LF2 CL.1"
                or "A350-LF2-CL1" or "A350 LF2 CL1" => "A350-LF2 Cl.1",
            "A352-LCB" or "A352 LCB" => "A352-LCB",
            "A420-WPL6" or "A420 WPL6" => "A420-WPL6",
            // Low-temp bolting
            "A320-L7" or "A320 L7" => "A320 L7",
            "A320-B8" or "A320 B8" or "A320-B8/B8M" or "A320 B8/B8M"
                or "A320-B8M" or "A320 B8M" => "A320 B8/B8M",
            // High-temp
            "A516-70" or "A516 70" => "A516-70",
            "A671-CC60" or "A671 CC60" => "A671-CC60",
            "A671-CC60-CL.22" or "A671-CC60 CL.22" or "A671 CC60 CL.22"
                or "A671-CC60-CL22" or "A671 CC60 CL22" => "A671-CC60 Cl.22",
            // Cast iron / ductile
            "A395" => "A395",
            "A216-WCB" or "A216 WCB" => "A216-WCB",
            // API line pipe
            "API 5L B" or "API-5L-B" or "API 5L-B" => "API 5L B",
            "API 5L-B PSL 2" or "API-5L-B-PSL-2" or "API 5L B PSL 2"
                or "API-5L-B-PSL2" or "API 5L-B PSL2" => "API 5L-B PSL 2",
            // Stainless steel
            "A182-F316/316L" or "A182 F316/316L" or "A182-F316"
                or "A182 F316" => "A182-F316/316L",
            "A182-F51" or "A182 F51" => "A182-F51",
            "A240-316/316L" or "A240 316/316L" => "A240-316/316L",
            "A312-TP316/316L" or "A312 TP316/316L" or "A312-TP316"
                or "A312 TP316" => "A312-TP316/316L",
            "A351-CF8M" or "A351 CF8M" => "A351-CF8M",
            "A358-316/316L" or "A358 316/316L" => "A358-316/316L",
            "A403-WP316/316L" or "A403 WP316/316L" or "A403-WP316"
                or "A403 WP316" => "A403-WP316/316L",
            "A307" => "A307",
            "A815" => "A815",
            // Duplex
            "A790-S31803" or "A790 S31803" => "A790-S31803",
            "A815-S31803" or "A815 S31803" => "A815-S31803",
            // Cu-Ni
            "CU-NI UNS 7060X" or "CU-NI-UNS-7060X" or "CUNI UNS 7060X"
                or "CU NI UNS 7060X" => "Cu-Ni UNS 7060X",
            "EEMUA-234" or "EEMUA 234" => "Cu-Ni UNS 7060X",
            "UNS-C70600" or "UNS C70600" or "UNS-7060X"
                or "UNS 7060X" => "Cu-Ni UNS 7060X",
            _ => null
        };
    }

    private static bool IsJointRecordRowEmpty(JointRecordAnalysisRow row)
    {
        return string.IsNullOrWhiteSpace(row.LayoutNo)
            && string.IsNullOrWhiteSpace(row.LRev)
            && string.IsNullOrWhiteSpace(row.LSRev)
            && string.IsNullOrWhiteSpace(row.UnitNumber)
            && string.IsNullOrWhiteSpace(row.ISO)
            && string.IsNullOrWhiteSpace(row.LineClass)
            && string.IsNullOrWhiteSpace(row.Material)
            && string.IsNullOrWhiteSpace(row.LineNo)
            && string.IsNullOrWhiteSpace(row.Fluid)
            && string.IsNullOrWhiteSpace(row.LSSheet)
            && string.IsNullOrWhiteSpace(row.LSDiameter)
            && string.IsNullOrWhiteSpace(row.LSScope)
            && string.IsNullOrWhiteSpace(row.Location)
            && string.IsNullOrWhiteSpace(row.WeldNumber)
            && string.IsNullOrWhiteSpace(row.JAdd)
            && string.IsNullOrWhiteSpace(row.WeldType)
            && string.IsNullOrWhiteSpace(row.SpoolNumber)
            && string.IsNullOrWhiteSpace(row.Diameter)
            && string.IsNullOrWhiteSpace(row.SpDia)
            && string.IsNullOrWhiteSpace(row.Schedule)
            && string.IsNullOrWhiteSpace(row.OLDiameter)
            && string.IsNullOrWhiteSpace(row.OLSchedule)
            && string.IsNullOrWhiteSpace(row.MaterialA)
            && string.IsNullOrWhiteSpace(row.MaterialB)
            && string.IsNullOrWhiteSpace(row.GradeA)
            && string.IsNullOrWhiteSpace(row.GradeB)
            && string.IsNullOrWhiteSpace(row.Remarks);
    }

    private static JointRecordAnalysisRow BuildJointRecordFromFileName(string fileName)
    {
        var row = new JointRecordAnalysisRow { SourceFile = fileName };
        var name = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;

        var pattern = @"^(?<unit>\d+)[-_](?<fluid>[A-Z0-9]+)[-_](?<line>\d+)[-_](?<sheet>\d+)_Rev(?<rev>[A-Za-z0-9]+(?:_[A-Za-z0-9]+)*)";
        var match = Regex.Match(name, pattern, RegexOptions.IgnoreCase);
        if (match.Success)
        {
            row.UnitNumber = match.Groups["unit"].Value;
            row.Fluid = match.Groups["fluid"].Value;
            row.LineNo = match.Groups["line"].Value;
            row.LSSheet = match.Groups["sheet"].Value;
            var revTag = NormalizeRevisionTagForEquivalence(match.Groups["rev"].Value);
            row.LRev = revTag;
            row.LSRev = revTag;
            row.LayoutNo = $"{row.Fluid}-{row.LineNo}";
            return row;
        }

        var patterns = new[]
        {
            @"^(?<unit>\d+)[-_](?<fluid>[A-Z0-9]+)[-_](?<line>\d+)",
            @"(?<fluid>[A-Z0-9]+)[-_](?<line>\d+)",
            @"(?<line>\d+)"
        };

        foreach (var pat in patterns)
        {
            match = Regex.Match(name, pat, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                if (match.Groups["unit"].Success)
                    row.UnitNumber = match.Groups["unit"].Value;
                if (match.Groups["fluid"].Success)
                    row.Fluid = match.Groups["fluid"].Value;
                if (match.Groups["line"].Success)
                    row.LineNo = match.Groups["line"].Value;

                if (!string.IsNullOrEmpty(row.Fluid) && !string.IsNullOrEmpty(row.LineNo))
                    row.LayoutNo = $"{row.Fluid}-{row.LineNo}";

                break;
            }
        }

        return row;
    }

    private static byte[] BuildJointRecordExcel(List<JointRecordAnalysisRow> rows)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("JointRecord");

        var headers = rows.Count > 0 ? JointRecordExportHeaders : JointRecordTemplateHeaders;

        // Mandatory column indexes (0-based) for Joints scope: Layout No(3), Sheet No.(7), Weld No(11), J-Add(12)
        var requiredCols = new HashSet<int> { 3, 7, 11, 12 };

        for (var i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.LightBlue;

            if (rows.Count == 0 && requiredCols.Contains(i))
            {
                // Rich text: header name + red " *"
                cell.GetRichText().AddText(headers[i]).SetBold(true);
                cell.GetRichText().AddText(" *").SetBold(true).SetFontColor(XLColor.Red);
            }
            else
            {
                cell.Value = headers[i];
            }
        }

        int r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.UnitNumber;
            ws.Cell(r, 2).Value = row.Fluid;
            ws.Cell(r, 3).Value = row.LineNo;
            ws.Cell(r, 4).Value = row.LayoutNo;
            ws.Cell(r, 5).Value = row.LineClass;
            ws.Cell(r, 6).Value = row.ISO;
            // Export compound revision tag matching Drawings controller format
            ws.Cell(r, 7).Value = NormalizeRevisionTagForEquivalence(row.LSRev ?? row.LRev);
            ws.Cell(r, 8).Value = row.LSSheet;
            ws.Cell(r, 9).Value = row.Material;
            ws.Cell(r, 10).Value = row.LSDiameter;
            ws.Cell(r, 11).Value = row.Location;
            ws.Cell(r, 12).Value = row.WeldNumber;
            ws.Cell(r, 13).Value = row.JAdd;
            ws.Cell(r, 14).Value = row.WeldType;
            ws.Cell(r, 15).Value = row.SpoolNumber;
            ws.Cell(r, 16).Value = row.Diameter;
            ws.Cell(r, 17).Value = row.Schedule;
            ws.Cell(r, 18).Value = row.SpDia;
            ws.Cell(r, 19).Value = row.OLDiameter;
            ws.Cell(r, 20).Value = row.OLSchedule;
            ws.Cell(r, 21).Value = row.MaterialA;
            ws.Cell(r, 22).Value = row.MaterialB;
            ws.Cell(r, 23).Value = row.GradeA;
            ws.Cell(r, 24).Value = row.GradeB;
            ws.Cell(r, 25).Value = row.Deleted || row.Cancelled ? "TRUE" : "";
            ws.Cell(r, 26).Value = row.LSScope ?? "AIC";
            ws.Cell(r, 27).Value = row.Remarks;
            r++;
        }

        // When generating a blank template, add an example row so users know the expected format
        if (rows.Count == 0)
        {
            var ex = 2;
            ws.Cell(ex, 1).Value = "09";              // Unit No
            ws.Cell(ex, 2).Value = "P";               // Fluid
            ws.Cell(ex, 3).Value = "11645";            // Line No
            ws.Cell(ex, 4).Value = "P-11645";          // Layout No
            ws.Cell(ex, 5).Value = "1CS2U01";          // Line Class
            ws.Cell(ex, 6).Value = "P-11645-1CS2U01";  // ISO
            ws.Cell(ex, 7).Value = "00_00";            // Revision
            ws.Cell(ex, 8).Value = "001";              // Sheet No.
            ws.Cell(ex, 9).Value = "CS";               // Material
            ws.Cell(ex, 10).Value = "6";               // ISO Dia. In.
            ws.Cell(ex, 11).Value = "WS";              // Location
            ws.Cell(ex, 12).Value = "01";              // Weld No
            ws.Cell(ex, 13).Value = "NEW";             // J-Add
            ws.Cell(ex, 14).Value = "BW";              // Weld Type
            ws.Cell(ex, 15).Value = "01";              // Spool
            ws.Cell(ex, 16).Value = "6";               // Dia. In.
            ws.Cell(ex, 17).Value = "40";              // Schedule
            ws.Cell(ex, 18).Value = "6";               // Spool Dia. In.
            ws.Cell(ex, 19).Value = "";                // OL Dia
            ws.Cell(ex, 20).Value = "";                // OL Schedule
            ws.Cell(ex, 21).Value = "PIPE";            // Material A
            ws.Cell(ex, 22).Value = "ELBOW";           // Material B
            ws.Cell(ex, 23).Value = "A106-B";          // Grade A
            ws.Cell(ex, 24).Value = "A234-WPB";        // Grade B
            ws.Cell(ex, 25).Value = "TRUE or FALSE";   // Delete / Cancel
            ws.Cell(ex, 26).Value = "AIC";             // Scope

            // Style the example row in italic light grey so it's clearly sample data
            var exRange = ws.Range(ex, 1, ex, headers.Length);
            exRange.Style.Font.Italic = true;
            exRange.Style.Font.FontColor = XLColor.Gray;
        }

        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private async Task PopulateActionStatusRemarksAsync(int projectId, string scope, List<JointRecordAnalysisRow> rows)
    {
        if (projectId <= 0 || rows.Count == 0) return;

        // Local helper: normalize a key the same way CommitJointRecord does
        static string Norm(string? s, int maxLen)
        {
            s = (s ?? "").Trim();
            if (s.Length > maxLen) s = s[..maxLen];
            return s.ToUpperInvariant();
        }

        try
        {
            if (scope.Equals(ScopeJoints, StringComparison.OrdinalIgnoreCase))
            {
                var existingDfr = await _context.DFR_tbl
                    .Where(d => d.Project_No == projectId)
                    .Select(d => new { d.LAYOUT_NUMBER, d.WELD_NUMBER, d.J_Add, d.SHEET, d.FITUP_DATE })
                    .ToListAsync();

                // Build a dictionary for O(1) lookups instead of O(n) per row
                var dfrLookup = new Dictionary<(string, string, string, string), DateTime?>(
                    existingDfr.Count);
                foreach (var d in existingDfr)
                {
                    var key = (
                        (d.LAYOUT_NUMBER ?? "").Trim().ToUpperInvariant(),
                        (d.WELD_NUMBER ?? "").Trim().ToUpperInvariant(),
                        (d.J_Add ?? "").Trim().ToUpperInvariant(),
                        (d.SHEET ?? "").Trim().ToUpperInvariant());
                    dfrLookup.TryAdd(key, d.FITUP_DATE);
                }

                foreach (var row in rows)
                {
                    var layout = Norm(row.LayoutNo, 10);
                    var weld = Norm(row.WeldNumber, 6);
                    var jAdd = Norm(row.JAdd, 8);
                    if (jAdd.Length == 0) jAdd = "NEW";
                    var sheet = Norm(row.LSSheet, 5);

                    var lookupKey = (layout, weld, jAdd, sheet);
                    bool deleteOrCancelRequested = row.Deleted || row.Cancelled;
                    if (!dfrLookup.TryGetValue(lookupKey, out var fitupDate))
                    {
                        row.Remarks = "Upload";
                    }
                    else if (fitupDate != null)
                    {
                        row.FitupDate = fitupDate.Value.ToString("dd-MMM-yyyy");
                        if (deleteOrCancelRequested)
                        {
                            // FITUP_DATE present → this will be a Delete operation
                            row.Deleted = true;
                            row.Cancelled = false;
                            row.Remarks = "Delete";
                        }
                        else
                        {
                            row.Remarks = $"Cannot update \u2014 Fitup already completed ({fitupDate:dd-MMM-yyyy})";
                        }
                    }
                    else
                    {
                        if (deleteOrCancelRequested)
                        {
                            // FITUP_DATE null → this will be a Cancel operation
                            row.Cancelled = true;
                            row.Deleted = false;
                            row.Remarks = "Cancel";
                        }
                        else
                        {
                            row.Remarks = "Update";
                        }
                    }
                }
            }
            else if (scope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
            {
                var existingSheets = await _context.Set<LineSheet>().AsNoTracking()
                    .Select(ls => new { ls.LS_LAYOUT_NO, ls.LS_SHEET })
                    .ToListAsync();

                var sheetLookup = new HashSet<(string, string)>(
                    existingSheets.Select(ls => (
                        (ls.LS_LAYOUT_NO ?? "").Trim().ToUpperInvariant(),
                        (ls.LS_SHEET ?? "").Trim().ToUpperInvariant())));

                foreach (var row in rows)
                {
                    var layout = Norm(row.LayoutNo, 10);
                    var sheet = Norm(row.LSSheet, 5);

                    row.Remarks = sheetLookup.Contains((layout, sheet)) ? "Update" : "Upload";
                }
            }
            else // LineList
            {
                var existingLayouts = await _context.Set<LineList>().AsNoTracking()
                    .Select(l => l.LAYOUT_NO)
                    .ToListAsync();

                var layoutLookup = new HashSet<string>(
                    existingLayouts
                        .Where(l => l != null)
                        .Select(l => l!.Trim().ToUpperInvariant()),
                    StringComparer.OrdinalIgnoreCase);

                foreach (var row in rows)
                {
                    var layout = Norm(row.LayoutNo, 10);

                    row.Remarks = layoutLookup.Contains(layout) ? "Update" : "Upload";
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to populate action status remarks");
        }
    }

    /// <summary>
    /// Validates extracted row values against project-scoped lookup tables.
    /// Corrects values to the closest permitted match or flags them with a warning.
    /// </summary>
    private async Task ValidateAgainstLookupTablesAsync(
        int projectId, List<JointRecordAnalysisRow> rows, List<string> warnings)
    {
        if (projectId <= 0 || rows.Count == 0) return;

        try
        {
            var weldersProjectId = await ResolveWeldersProjectIdAsync(projectId);

            // Load permitted values from lookup tables
            var permittedLocations = await _context.PMS_Location_tbl.AsNoTracking()
                .Where(l => l.LO_Project_No == weldersProjectId && l.P_Location != null && l.P_Location != "")
                .Select(l => l.P_Location!)
                .Distinct()
                .ToListAsync();

            var permittedWeldTypes = await _context.PMS_Weld_Type_tbl.AsNoTracking()
                .Where(w => w.W_Project_No == weldersProjectId && w.P_Weld_Type != null && w.P_Weld_Type != "")
                .Select(w => w.P_Weld_Type!)
                .Distinct()
                .ToListAsync();

            var permittedMaterialDescriptions = await _context.Material_Des_tbl.AsNoTracking()
                .Where(m => m.MATD_Description != null && m.MATD_Description != "")
                .Select(m => m.MATD_Description)
                .Distinct()
                .ToListAsync();

            var permittedGrades = await _context.Material_tbl.AsNoTracking()
                .Where(m => m.MAT_GRADE != null && m.MAT_GRADE != "")
                .Select(m => m.MAT_GRADE)
                .Distinct()
                .ToListAsync();

            // Build case-insensitive lookup sets for fast matching
            var locationSet = new HashSet<string>(permittedLocations, StringComparer.OrdinalIgnoreCase);
            var weldTypeSet = new HashSet<string>(permittedWeldTypes, StringComparer.OrdinalIgnoreCase);
            var materialDesSet = new HashSet<string>(permittedMaterialDescriptions, StringComparer.OrdinalIgnoreCase);
            var gradeSet = new HashSet<string>(permittedGrades, StringComparer.OrdinalIgnoreCase);

            // Find the default weld type for fallback
            var defaultWeldType = await _context.PMS_Weld_Type_tbl.AsNoTracking()
                .Where(w => w.W_Project_No == weldersProjectId && w.Default_Value)
                .Select(w => w.P_Weld_Type)
                .FirstOrDefaultAsync();

            foreach (var row in rows)
            {
                // 1. Validate Location
                if (!string.IsNullOrEmpty(row.Location))
                {
                    if (locationSet.Count > 0 && !locationSet.Contains(row.Location))
                    {
                        // Try case-insensitive exact match from the original list
                        var corrected = permittedLocations.FirstOrDefault(l =>
                            string.Equals(l, row.Location, StringComparison.OrdinalIgnoreCase));
                        if (corrected != null)
                        {
                            row.Location = corrected;
                        }
                        else
                        {
                            AppendRowRemark(row, warnings,
                                $"Location '{row.Location}' not found in project lookup. Permitted: {string.Join(", ", permittedLocations.Take(8))}");
                        }
                    }
                }

                // 2. Validate Weld Type
                if (!string.IsNullOrEmpty(row.WeldType))
                {
                    if (weldTypeSet.Count > 0 && !weldTypeSet.Contains(row.WeldType))
                    {
                        var corrected = permittedWeldTypes.FirstOrDefault(w =>
                            string.Equals(w, row.WeldType, StringComparison.OrdinalIgnoreCase));
                        if (corrected != null)
                        {
                            row.WeldType = corrected;
                        }
                        else
                        {
                            var fallback = defaultWeldType ?? permittedWeldTypes.FirstOrDefault() ?? "BW";
                            AppendRowRemark(row, warnings,
                                $"Weld type '{row.WeldType}' not found in project lookup. Defaulting to '{fallback}'. Permitted: {string.Join(", ", permittedWeldTypes.Take(8))}");
                            row.WeldType = fallback;
                        }
                    }
                }

                // 3. Auto-populate Material A/B from weld type if still empty, then validate
                if (string.IsNullOrEmpty(row.MaterialA) || string.IsNullOrEmpty(row.MaterialB))
                {
                    var (inferredA, inferredB) = InferMaterialFromWeldType(row.WeldType);
                    if (string.IsNullOrEmpty(row.MaterialA) && !string.IsNullOrEmpty(inferredA))
                    {
                        // Only apply if the inferred value exists in Material_Des_tbl
                        if (materialDesSet.Count == 0 || materialDesSet.Contains(inferredA))
                            row.MaterialA = inferredA;
                    }
                    if (string.IsNullOrEmpty(row.MaterialB) && !string.IsNullOrEmpty(inferredB))
                    {
                        if (materialDesSet.Count == 0 || materialDesSet.Contains(inferredB))
                            row.MaterialB = inferredB;
                    }
                }

                // 4. Validate Material A against Material_Des_tbl
                if (!string.IsNullOrEmpty(row.MaterialA))
                {
                    if (materialDesSet.Count > 0 && !materialDesSet.Contains(row.MaterialA))
                    {
                        var corrected = permittedMaterialDescriptions.FirstOrDefault(m =>
                            string.Equals(m, row.MaterialA, StringComparison.OrdinalIgnoreCase))
                            ?? MapToDropdownValue(row.MaterialA, permittedMaterialDescriptions);
                        if (corrected != null)
                        {
                            row.MaterialA = corrected;
                        }
                        else
                        {
                            AppendRowRemark(row, warnings,
                                $"Material A '{row.MaterialA}' not found in Material_Des_tbl");
                        }
                    }
                }

                // 5. Validate Material B against Material_Des_tbl
                if (!string.IsNullOrEmpty(row.MaterialB))
                {
                    if (materialDesSet.Count > 0 && !materialDesSet.Contains(row.MaterialB))
                    {
                        var corrected = permittedMaterialDescriptions.FirstOrDefault(m =>
                            string.Equals(m, row.MaterialB, StringComparison.OrdinalIgnoreCase))
                            ?? MapToDropdownValue(row.MaterialB, permittedMaterialDescriptions);
                        if (corrected != null)
                        {
                            row.MaterialB = corrected;
                        }
                        else
                        {
                            AppendRowRemark(row, warnings,
                                $"Material B '{row.MaterialB}' not found in Material_Des_tbl");
                        }
                    }
                }

                // 6. Validate Grade A
                if (string.IsNullOrEmpty(row.GradeA) && !string.IsNullOrEmpty(row.MaterialA))
                {
                    var inferred = InferGradeFromComponent(row.MaterialA);
                    if (!string.IsNullOrEmpty(inferred))
                    {
                        var mapped = gradeSet.Count == 0 || gradeSet.Contains(inferred)
                            ? inferred
                            : MapToDropdownValue(inferred, permittedGrades);
                        if (!string.IsNullOrEmpty(mapped))
                            row.GradeA = mapped;
                    }
                }

                if (!string.IsNullOrEmpty(row.GradeA))
                {
                    if (gradeSet.Count > 0 && !gradeSet.Contains(row.GradeA))
                    {
                        var corrected = permittedGrades.FirstOrDefault(g =>
                            string.Equals(g, row.GradeA, StringComparison.OrdinalIgnoreCase))
                            ?? MapToDropdownValue(row.GradeA, permittedGrades);
                        if (corrected != null)
                        {
                            row.GradeA = corrected;
                        }
                        else
                        {
                            AppendRowRemark(row, warnings,
                                $"Grade A '{row.GradeA}' not found in Material_tbl.MAT_GRADE");
                        }
                    }
                }

                // 7. Validate Grade B
                if (string.IsNullOrEmpty(row.GradeB) && !string.IsNullOrEmpty(row.MaterialB))
                {
                    var inferred = InferGradeFromComponent(row.MaterialB);
                    if (!string.IsNullOrEmpty(inferred))
                    {
                        var mapped = gradeSet.Count == 0 || gradeSet.Contains(inferred)
                            ? inferred
                            : MapToDropdownValue(inferred, permittedGrades);
                        if (!string.IsNullOrEmpty(mapped))
                            row.GradeB = mapped;
                    }
                }

                if (!string.IsNullOrEmpty(row.GradeB))
                {
                    if (gradeSet.Count > 0 && !gradeSet.Contains(row.GradeB))
                    {
                        var corrected = permittedGrades.FirstOrDefault(g =>
                            string.Equals(g, row.GradeB, StringComparison.OrdinalIgnoreCase))
                            ?? MapToDropdownValue(row.GradeB, permittedGrades);
                        if (corrected != null)
                        {
                            row.GradeB = corrected;
                        }
                        else
                        {
                            AppendRowRemark(row, warnings,
                                $"Grade B '{row.GradeB}' not found in Material_tbl.MAT_GRADE");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to validate against lookup tables for project {ProjectId}", projectId);
        }
    }

    private static string NormalizeDropdownToken(string value)
    {
        var upper = value.Trim().ToUpperInvariant();
        var sb = new StringBuilder(upper.Length);
        foreach (var ch in upper)
        {
            if (char.IsLetterOrDigit(ch)) sb.Append(ch);
        }
        return sb.ToString();
    }

    private static string? MapToDropdownValue(string? value, List<string> options)
    {
        if (string.IsNullOrWhiteSpace(value) || options.Count == 0) return null;

        // Try known grade mapping first (handles ASTM/API/EEMUA/UNS → DB format)
        var knownGrade = MapToKnownGrade(value.Trim().ToUpperInvariant());
        if (knownGrade != null)
        {
            // Verify it exists in the options list
            var dbMatch = options.FirstOrDefault(o =>
                string.Equals(o, knownGrade, StringComparison.OrdinalIgnoreCase));
            if (dbMatch != null) return dbMatch;
        }

        var normalized = NormalizeDropdownToken(value);
        if (normalized.Length == 0) return null;

        // Exact normalized match
        var exact = options.FirstOrDefault(o =>
            NormalizeDropdownToken(o) == normalized);
        if (exact != null) return exact;

        // Prefix/contains match
        var prefix = options.FirstOrDefault(o =>
            NormalizeDropdownToken(o).StartsWith(normalized, StringComparison.OrdinalIgnoreCase) ||
            normalized.StartsWith(NormalizeDropdownToken(o), StringComparison.OrdinalIgnoreCase));
        if (prefix != null) return prefix;

        return null;
    }

    private static bool HasIsoSeparators(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return value.Contains('-') || value.Contains('_') ||
               value.Contains('\u2013') || value.Contains('\u2014') ||
               value.Contains(' ');
    }

    private static bool IsPlausibleLineClass(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        var v = value.Trim().ToUpperInvariant();

        if (v.Length < 3 || v.Length > 12) return false;

        string[] banned = {
            "SUPPORT", "SUPPORTS", "INSULATION", "FABRICATION", "FABRIC",
            "TOTAL", "MATERIAL", "WELD", "SCHEDULE", "DIAMETER", "DRAWING",
            "ISOMETRIC", "ISO", "PIPING", "PIPE", "VALVE", "FLANGE",
            "THICK", "THICKNESS", "SPECIFICATION", "SPEC", "DESCRIPTION",
            "GASKET", "GASKETS", "SPI", "SPOOL", "SIZE"
        };
        if (banned.Any(b => v.Contains(b, StringComparison.Ordinal))) return false;

        var pattern = @"^(?=.*[A-Z])(?=(?:.*\d){2,})\d[A-Z0-9]{2,11}$";
        if (!Regex.IsMatch(v, pattern, RegexOptions.CultureInvariant)) return false;

        var letters = v.Count(char.IsLetter);
        var digits = v.Count(char.IsDigit);
        if (digits > letters * 3) return false;
        return true;
    }

    private static bool ShouldReplaceLineClass(string? current, string candidate)
    {
        if (!IsPlausibleLineClass(candidate)) return false;
        if (string.IsNullOrWhiteSpace(current)) return true;
        if (!IsPlausibleLineClass(current)) return true;
        return candidate.Length > current.Length;
    }

    #endregion

    #region Template Download / Save / Export

    [SessionAuthorization]
    [HttpGet]
    public IActionResult DownloadJointRecordTemplate()
    {
        var bytes = BuildJointRecordExcel(new List<JointRecordAnalysisRow>());
        var fileName = $"JointRecord_Template_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public IActionResult DownloadJointRecordAnalysis([FromBody] JointRecordAnalysisResult payload)
    {
        if (payload?.Rows == null || payload.Rows.Count == 0)
        {
            return BadRequest(new { error = "No rows to download." });
        }

        var bytes = BuildJointRecordExcel(payload.Rows);
        var fileName = $"JointRecord_Analysis_{AppClock.Now:yyyyMMdd_HHmmss}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    #endregion

    #region Joint Record Intake – Load / Commit / Layouts / Sheets

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetJointRecordLayouts(int projectId, string? scope)
    {
        var normalizedScope = (scope ?? ScopeJoints).Trim();
        List<string> layouts;

        if (normalizedScope.Equals(ScopeLineList, StringComparison.OrdinalIgnoreCase))
        {
            layouts = await _context.Set<LineList>().AsNoTracking()
                .Where(l => l.LAYOUT_NO != null && l.LAYOUT_NO != "")
                .Select(l => l.LAYOUT_NO!)
                .Distinct().OrderBy(x => x).ToListAsync();
        }
        else if (normalizedScope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
        {
            layouts = await _context.Set<LineSheet>().AsNoTracking()
                .Where(ls => ls.LS_LAYOUT_NO != null && ls.LS_LAYOUT_NO != "")
                .Select(ls => ls.LS_LAYOUT_NO!)
                .Distinct().OrderBy(x => x).ToListAsync();
        }
        else
        {
            if (projectId <= 0) return Json(Array.Empty<string>());
            layouts = await _context.DFR_tbl.AsNoTracking()
                .Where(d => d.Project_No == projectId && d.LAYOUT_NUMBER != null && d.LAYOUT_NUMBER != "")
                .Select(d => d.LAYOUT_NUMBER!)
                .Distinct().OrderBy(x => x).ToListAsync();
        }

        return Json(layouts);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetJointRecordSheets(int projectId, string? scope, string? layout)
    {
        if (string.IsNullOrWhiteSpace(layout)) return Json(Array.Empty<string>());
        var normalizedScope = (scope ?? ScopeJoints).Trim();
        List<string> sheets;

        if (normalizedScope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
        {
            sheets = await _context.Set<LineSheet>().AsNoTracking()
                .Where(ls => ls.LS_LAYOUT_NO == layout && ls.LS_SHEET != null && ls.LS_SHEET != "")
                .Select(ls => ls.LS_SHEET!)
                .Distinct().OrderBy(x => x).ToListAsync();
        }
        else
        {
            if (projectId <= 0) return Json(Array.Empty<string>());
            sheets = await _context.DFR_tbl.AsNoTracking()
                .Where(d => d.Project_No == projectId && d.LAYOUT_NUMBER == layout && d.SHEET != null && d.SHEET != "")
                .Select(d => d.SHEET!)
                .Distinct().OrderBy(x => x).ToListAsync();
        }

        return Json(sheets);
    }

    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> LoadJointRecordData(int projectId, string? scope, string? layout, string? sheet)
    {
        if (string.IsNullOrWhiteSpace(layout))
            return BadRequest(new { error = "Layout No is required" });

        var result = new JointRecordAnalysisResult { ProjectId = projectId };
        var normalizedScope = (scope ?? ScopeJoints).Trim();

        try
        {
            if (normalizedScope.Equals(ScopeLineList, StringComparison.OrdinalIgnoreCase))
            {
                var lineListRows = await _context.Set<LineList>().AsNoTracking()
                    .Where(l => l.LAYOUT_NO == layout)
                    .ToListAsync();

                result.Rows = lineListRows.Select(l => new JointRecordAnalysisRow
                {
                    LayoutNo = l.LAYOUT_NO,
                    UnitNumber = l.Unit_Number,
                    ISO = l.ISO,
                    LineClass = l.Line_Class,
                    Material = l.Material,
                    LineNo = l.LINE_NO,
                    Fluid = l.Fluid,
                    LRev = l.L_REV,
                }).ToList();
            }
            else if (normalizedScope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
            {
                var query = from ls in _context.Set<LineSheet>().AsNoTracking()
                            join ll in _context.Set<LineList>().AsNoTracking()
                                on ls.Line_ID_LS equals ll.Line_ID into llGroup
                            from ll in llGroup.DefaultIfEmpty()
                            where ls.LS_LAYOUT_NO == layout
                            select new { ls, ll };

                if (!string.IsNullOrWhiteSpace(sheet))
                    query = query.Where(x => x.ls.LS_SHEET == sheet);

                var data = await query.OrderBy(x => x.ls.LS_SHEET).ToListAsync();

                result.Rows = data.Select(x => new JointRecordAnalysisRow
                {
                    LayoutNo = x.ls.LS_LAYOUT_NO ?? x.ll?.LAYOUT_NO,
                    UnitNumber = x.ll?.Unit_Number,
                    ISO = x.ll?.ISO,
                    LineClass = x.ll?.Line_Class,
                    Material = x.ll?.Material,
                    LineNo = x.ll?.LINE_NO,
                    Fluid = x.ll?.Fluid,
                    LRev = x.ll?.L_REV,
                    LSSheet = x.ls.LS_SHEET,
                    LSRev = x.ls.LS_REV,
                    LSDiameter = x.ls.LS_DIAMETER?.ToString(CultureInfo.InvariantCulture),
                    LSScope = x.ls.LS_Scope,
                }).ToList();
            }
            else // Joints
            {
                if (projectId <= 0)
                    return BadRequest(new { error = "Project is required for Joints scope" });

                var dfrQuery = _context.DFR_tbl.AsNoTracking()
                    .Where(d => d.Project_No == projectId && d.LAYOUT_NUMBER == layout);

                if (!string.IsNullOrWhiteSpace(sheet))
                    dfrQuery = dfrQuery.Where(d => d.SHEET == sheet);

                var dfrData = await dfrQuery
                    .OrderBy(d => d.SHEET).ThenBy(d => d.WELD_NUMBER)
                    .ToListAsync();

                var lineSheetIds = dfrData.Where(d => d.Line_Sheet_ID_DFR.HasValue)
                    .Select(d => d.Line_Sheet_ID_DFR!.Value).Distinct().ToList();

                var lineSheets = lineSheetIds.Count == 0
                    ? new Dictionary<int, LineSheet>()
                    : await _context.Set<LineSheet>().AsNoTracking()
                        .Where(ls => lineSheetIds.Contains(ls.Line_Sheet_ID))
                        .ToDictionaryAsync(ls => ls.Line_Sheet_ID);

                var lineIds = lineSheets.Values
                    .Where(ls => ls.Line_ID_LS.HasValue)
                    .Select(ls => ls.Line_ID_LS!.Value).Distinct().ToList();

                var lineLists = lineIds.Count == 0
                    ? new Dictionary<int, LineList>()
                    : await _context.Set<LineList>().AsNoTracking()
                        .Where(l => lineIds.Contains(l.Line_ID))
                        .ToDictionaryAsync(l => l.Line_ID);

                // Also try fallback lookup by LAYOUT_NO
                var fallbackLine = lineLists.Count == 0
                    ? await _context.Set<LineList>().AsNoTracking()
                        .FirstOrDefaultAsync(l => l.LAYOUT_NO == layout)
                    : null;

                // Resolve SP_DIA from SP_Release_tbl via Spool_ID_DFR for each loaded row
                var spoolIds = dfrData.Where(d => d.Spool_ID_DFR.HasValue && d.Spool_ID_DFR.Value > 0)
                    .Select(d => d.Spool_ID_DFR!.Value).Distinct().ToList();

                var spoolDiaLookup = spoolIds.Count == 0
                    ? new Dictionary<int, double?>()
                    : await _context.SP_Release_tbl.AsNoTracking()
                        .Where(sp => spoolIds.Contains(sp.Spool_ID))
                        .ToDictionaryAsync(sp => sp.Spool_ID, sp => sp.SP_DIA);

                result.Rows = dfrData.Select(d =>
                {
                    LineSheet? ls = d.Line_Sheet_ID_DFR.HasValue
                        && lineSheets.TryGetValue(d.Line_Sheet_ID_DFR.Value, out var lsVal) ? lsVal : null;
                    LineList? ll = ls?.Line_ID_LS.HasValue == true
                        && lineLists.TryGetValue(ls.Line_ID_LS!.Value, out var llVal) ? llVal : null;
                    ll ??= fallbackLine;

                    // Reconstruct SpDia from SP_Release_tbl.SP_DIA
                    string? spDia = null;
                    if (d.Spool_ID_DFR.HasValue && d.Spool_ID_DFR.Value > 0
                        && spoolDiaLookup.TryGetValue(d.Spool_ID_DFR.Value, out var spDiaVal)
                        && spDiaVal.HasValue)
                    {
                        spDia = spDiaVal.Value.ToString(CultureInfo.InvariantCulture);
                    }

                    return new JointRecordAnalysisRow
                    {
                        LayoutNo = d.LAYOUT_NUMBER ?? ll?.LAYOUT_NO,
                        UnitNumber = ll?.Unit_Number,
                        ISO = ll?.ISO,
                        LineClass = ll?.Line_Class,
                        Material = ll?.Material,
                        LineNo = ll?.LINE_NO,
                        Fluid = ll?.Fluid,
                        LRev = ll?.L_REV,
                        LSSheet = d.SHEET ?? ls?.LS_SHEET,
                        LSRev = ls?.LS_REV ?? d.DFR_REV,
                        LSDiameter = ls?.LS_DIAMETER?.ToString(CultureInfo.InvariantCulture),
                        LSScope = ls?.LS_Scope,
                        Location = d.LOCATION,
                        WeldNumber = d.WELD_NUMBER,
                        JAdd = d.J_Add,
                        WeldType = d.WELD_TYPE,
                        SpoolNumber = d.SPOOL_NUMBER,
                        Diameter = d.DIAMETER?.ToString(CultureInfo.InvariantCulture),
                        SpDia = spDia,
                        Schedule = d.SCHEDULE,
                        OLDiameter = d.OL_DIAMETER?.ToString(CultureInfo.InvariantCulture),
                        OLSchedule = d.OL_SCHEDULE,
                        OLThick = d.OL_Thick?.ToString(CultureInfo.InvariantCulture),
                        MaterialA = d.MATERIAL_A,
                        MaterialB = d.MATERIAL_B,
                        GradeA = d.GRADE_A,
                        GradeB = d.GRADE_B,
                        Deleted = d.Deleted,
                        Cancelled = d.Cancelled,
                        FitupDate = d.FITUP_DATE?.ToString("dd-MMM-yyyy"),
                    };
                }).ToList();
            }

            // Normalize revisions loaded from DB (DFR_REV may be compound "00_00")
            NormalizeRevisionFields(result.Rows);

            result.Warnings.Add($"Loaded {result.Rows.Count} records for editing ({normalizedScope} scope).");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "LoadJointRecordData failed for project={ProjectId} scope={Scope}", projectId, scope);
            result.Errors.Add($"Failed to load data: {ex.Message}");
        }

        return new JsonResult(result, JointRecordJsonOptions);
    }

    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CommitJointRecord([FromBody] JointRecordCommitRequest request)
    {
        if (request?.Rows == null || request.Rows.Count == 0)
            return BadRequest(new { error = "No rows to commit." });

        var result = new JointRecordCommitResult();
        var userId = HttpContext.Session.GetInt32("UserID");
        var now = AppClock.Now;
        var normalizedScope = (request.Scope ?? ScopeJoints).Trim();
        var isUploadMode = (request.Mode ?? ModeUpload).Equals(ModeUpload, StringComparison.OrdinalIgnoreCase);
        var updateOnlyNonNull = request.UpdateOnlyNonNull;
        var cancellationToken = HttpContext.RequestAborted;
        var newDfrEntities = new List<Dfr>();

        // Determine where the revision tag should be stored based on Projects_tbl.Line_Sheet
        var lineSheetMode = await _context.Projects_tbl.AsNoTracking()
            .Where(p => p.Project_ID == request.ProjectId)
            .Select(p => p.Line_Sheet)
            .FirstOrDefaultAsync(cancellationToken);
        var revToLine = string.Equals(lineSheetMode, "line", StringComparison.OrdinalIgnoreCase);
        var revToSheet = string.Equals(lineSheetMode, "sheet", StringComparison.OrdinalIgnoreCase);

        await using var transaction = await _context.Database.BeginTransactionAsync(cancellationToken);
        try
        {
            for (int i = 0; i < request.Rows.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var row = request.Rows[i];
                try
                {
                    // 1. Find or create LINE_LIST record
                    var layoutKey = Clean(row.LayoutNo, 10);
                    if (string.IsNullOrWhiteSpace(layoutKey))
                    {
                        result.Skipped++;
                        result.Errors.Add($"Row skipped: Layout No is required.");
                        result.RowRemarks[i] = "Layout No is required.";
                        continue;
                    }

                    var lineList = await _context.Set<LineList>()
                        .FirstOrDefaultAsync(l => l.LAYOUT_NO == layoutKey);

                    bool lineListCreated = false;
                    if (lineList == null)
                    {
                        lineList = new LineList { LAYOUT_NO = layoutKey, PMI = false, PWHT_Y_N = false, PWHT_20mm = false };
                        _context.Set<LineList>().Add(lineList);
                        lineListCreated = true;
                    }

                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.UnitNumber))
                        lineList.Unit_Number = Clean(row.UnitNumber, 6);
                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.ISO))
                        lineList.ISO = Clean(row.ISO, 24);
                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.LineClass))
                        lineList.Line_Class = Clean(row.LineClass, 10);
                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.Material))
                        lineList.Material = Clean(row.Material, 10);
                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.LineNo))
                        lineList.LINE_NO = Clean(row.LineNo, 20);
                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.Fluid))
                        lineList.Fluid = Clean(row.Fluid, 4);
                    // Store revision in L_REV only when Projects_tbl.Line_Sheet = 'Line'
                    if (revToLine)
                    {
                        var lRevTag = NormalizeRevisionTagForEquivalence(row.LRev ?? row.LSRev);
                        lineList.L_REV = Clean(lRevTag, 8);
                    }
                    lineList.Line_List_Updated_Date = now;
                    lineList.Line_List_Updated_By = userId;

                    if (normalizedScope.Equals(ScopeLineList, StringComparison.OrdinalIgnoreCase))
                    {
                        await _context.SaveChangesAsync(cancellationToken);
                        if (lineListCreated) result.Created++; else result.Updated++;
                        continue;
                    }

                    // Save LINE_LIST first to get Line_ID
                    await _context.SaveChangesAsync(cancellationToken);

                    // 2. Find or create Line_Sheet record
                    var sheetKey = Clean(row.LSSheet, 5);
                    if (string.IsNullOrWhiteSpace(sheetKey))
                    {
                        if (normalizedScope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
                        {
                            if (lineListCreated) result.Created++; else result.Updated++;
                            continue;
                        }
                        result.Skipped++;
                        result.Errors.Add($"Row {layoutKey}: Sheet No is required for Joints scope.");
                        result.RowRemarks[i] = "Sheet No is required for Joints scope.";
                        continue;
                    }

                    var lineSheet = await _context.Set<LineSheet>()
                        .FirstOrDefaultAsync(ls => ls.LS_LAYOUT_NO == layoutKey && ls.LS_SHEET == sheetKey);

                    bool lineSheetCreated = false;
                    if (lineSheet == null)
                    {
                        lineSheet = new LineSheet
                        {
                            LS_LAYOUT_NO = layoutKey,
                            LS_SHEET = sheetKey,
                            Line_ID_LS = lineList.Line_ID,
                            Project_No = request.ProjectId,
                        };
                        _context.Set<LineSheet>().Add(lineSheet);
                        lineSheetCreated = true;
                    }

                    // Store revision in LS_REV only when Projects_tbl.Line_Sheet = 'Sheet'
                    if (revToSheet)
                    {
                        var lsRevTag = NormalizeRevisionTagForEquivalence(row.LSRev ?? row.LRev);
                        if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(lsRevTag))
                            lineSheet.LS_REV = Clean(lsRevTag, 8);
                    }
                    if (double.TryParse(row.LSDiameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var lsDia))
                        lineSheet.LS_DIAMETER = lsDia;
                    if (!updateOnlyNonNull || !string.IsNullOrWhiteSpace(row.LSScope))
                        lineSheet.LS_Scope = Clean(row.LSScope, 10);
                    lineSheet.Line_ID_LS = lineList.Line_ID;
                    lineSheet.Project_No = request.ProjectId;
                    lineSheet.LS_Updated_Date = now;
                    lineSheet.LS_Updated_By = userId;

                    if (normalizedScope.Equals(ScopeLineSheet, StringComparison.OrdinalIgnoreCase))
                    {
                        await _context.SaveChangesAsync(cancellationToken);
                        if (lineListCreated || lineSheetCreated) result.Created++; else result.Updated++;
                        continue;
                    }

                    // Save Line_Sheet first to get Line_Sheet_ID
                    await _context.SaveChangesAsync(cancellationToken);

                    // 3. Find or create DFR record (Joints scope)
                    var weldKey = Clean(row.WeldNumber, 6);
                    var jAddKey = Clean(row.JAdd, 8)?.ToUpperInvariant() ?? "NEW";
                    if (string.IsNullOrWhiteSpace(weldKey))
                    {
                        result.Skipped++;
                        result.Errors.Add($"Row {layoutKey}/{sheetKey}: Weld No is required.");
                        result.RowRemarks[i] = "Weld No is required.";
                        continue;
                    }

                    var dfr = await _context.DFR_tbl
                        .FirstOrDefaultAsync(d => d.Project_No == request.ProjectId
                            && d.LAYOUT_NUMBER == layoutKey
                            && d.SHEET == sheetKey
                            && d.WELD_NUMBER == weldKey
                            && d.J_Add == jAddKey);

                    bool dfrCreated = false;

                    bool deleteOrCancelRequested = row.Deleted || row.Cancelled;

                    if (isUploadMode)
                    {
                        // Upload mode: create new records, or update existing when Fitup not yet completed
                        if (dfr != null)
                        {
                            if (dfr.FITUP_DATE != null && !deleteOrCancelRequested)
                            {
                                result.Skipped++;
                                result.Errors.Add($"Row {layoutKey}/{sheetKey}/{weldKey}: Cannot update — Fitup already completed ({dfr.FITUP_DATE:dd-MMM-yyyy}).");
                                result.RowRemarks[i] = $"Cannot update — Fitup already completed ({dfr.FITUP_DATE:dd-MMM-yyyy}).";
                                continue;
                            }
                            // Fitup not completed or delete/cancel requested — allow update
                        }
                        else
                        {
                            dfr = new Dfr
                            {
                                Project_No = request.ProjectId,
                                LAYOUT_NUMBER = layoutKey,
                                SHEET = sheetKey,
                                WELD_NUMBER = weldKey,
                                J_Add = jAddKey,
                                Deleted = false,
                                Cancelled = false,
                                Fitup_Confirmed = false,
                            };
                            _context.DFR_tbl.Add(dfr);
                            dfrCreated = true;
                        }
                    }
                    else
                    {
                        // Edit/Update mode: only update existing records where FITUP_DATE is null
                        if (dfr == null)
                        {
                            result.Skipped++;
                            result.Errors.Add($"Row {layoutKey}/{sheetKey}/{weldKey}: Record not found.");
                            result.RowRemarks[i] = "Record not found.";
                            continue;
                        }

                        if (dfr.FITUP_DATE != null && !deleteOrCancelRequested)
                        {
                            result.Skipped++;
                            result.Errors.Add($"Row {layoutKey}/{sheetKey}/{weldKey}: Cannot update — Fitup already completed ({dfr.FITUP_DATE:dd-MMM-yyyy}).");
                            result.RowRemarks[i] = $"Cannot update — Fitup already completed ({dfr.FITUP_DATE:dd-MMM-yyyy}).";
                            continue;
                        }
                    }

                    if (updateOnlyNonNull && !dfrCreated)
                    {
                        // Update only non-null cells from the Excel template
                        if (!string.IsNullOrWhiteSpace(row.Location))
                            dfr.LOCATION = Clean(row.Location, 4);
                        if (!string.IsNullOrWhiteSpace(row.WeldType))
                            dfr.WELD_TYPE = NormalizeWeldType(row.WeldType);
                        if (!string.IsNullOrWhiteSpace(row.SpoolNumber))
                            dfr.SPOOL_NUMBER = Clean(row.SpoolNumber, 9);
                        if (!string.IsNullOrWhiteSpace(row.Diameter))
                        {
                            if (double.TryParse(row.Diameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var diaNn))
                                dfr.DIAMETER = diaNn;
                        }
                        if (!string.IsNullOrWhiteSpace(row.Schedule))
                            dfr.SCHEDULE = Clean(row.Schedule, 8);
                        if (!string.IsNullOrWhiteSpace(row.MaterialA))
                            dfr.MATERIAL_A = Clean(row.MaterialA, 20);
                        if (!string.IsNullOrWhiteSpace(row.MaterialB))
                            dfr.MATERIAL_B = Clean(row.MaterialB, 20);
                        if (!string.IsNullOrWhiteSpace(row.GradeA))
                            dfr.GRADE_A = Clean(row.GradeA, 30);
                        if (!string.IsNullOrWhiteSpace(row.GradeB))
                            dfr.GRADE_B = Clean(row.GradeB, 30);
                        if (!string.IsNullOrWhiteSpace(row.OLDiameter))
                        {
                            if (double.TryParse(row.OLDiameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var olDiaNn))
                                dfr.OL_DIAMETER = olDiaNn;
                        }
                        if (!string.IsNullOrWhiteSpace(row.OLSchedule))
                            dfr.OL_SCHEDULE = Clean(row.OLSchedule, 8);
                        if (!string.IsNullOrWhiteSpace(row.OLThick))
                        {
                            if (double.TryParse(row.OLThick, NumberStyles.Any, CultureInfo.InvariantCulture, out var olThNn))
                                dfr.OL_Thick = olThNn;
                        }
                    }
                    else
                    {
                        dfr.LOCATION = Clean(row.Location, 4);
                        dfr.WELD_TYPE = NormalizeWeldType(row.WeldType);
                        dfr.SPOOL_NUMBER = Clean(row.SpoolNumber, 9);
                        if (double.TryParse(row.Diameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var dia))
                            dfr.DIAMETER = dia;
                        else dfr.DIAMETER = null;
                        dfr.SCHEDULE = Clean(row.Schedule, 8);
                        dfr.MATERIAL_A = Clean(row.MaterialA, 20);
                        dfr.MATERIAL_B = Clean(row.MaterialB, 20);
                        dfr.GRADE_A = Clean(row.GradeA, 30);
                        dfr.GRADE_B = Clean(row.GradeB, 30);
                        if (double.TryParse(row.OLDiameter, NumberStyles.Any, CultureInfo.InvariantCulture, out var olDia))
                            dfr.OL_DIAMETER = olDia;
                        else dfr.OL_DIAMETER = null;
                        dfr.OL_SCHEDULE = Clean(row.OLSchedule, 8);
                        if (double.TryParse(row.OLThick, NumberStyles.Any, CultureInfo.InvariantCulture, out var olTh))
                            dfr.OL_Thick = olTh;
                    }
                    // Apply Delete / Cancel status with business-rule enforcement:
                    // FITUP_DATE present  + delete/cancel requested → Deleted = true
                    // FITUP_DATE null     + delete/cancel requested → Cancelled = true
                    if (deleteOrCancelRequested)
                    {
                        if (dfr.FITUP_DATE != null)
                        {
                            dfr.Deleted = true;
                            dfr.Cancelled = false;
                        }
                        else
                        {
                            dfr.Cancelled = true;
                            dfr.Deleted = false;
                        }
                    }
                    else
                    {
                        dfr.Deleted = false;
                        dfr.Cancelled = false;
                    }

                    dfr.Line_Sheet_ID_DFR = lineSheet.Line_Sheet_ID;
                    dfr.DFR_Updated_By = userId;
                    dfr.DFR_Updated_Date = now;

                    await _context.SaveChangesAsync(cancellationToken);

                    if (dfrCreated)
                    {
                        newDfrEntities.Add(dfr);
                    }

                    if (dfrCreated) result.Created++; else result.Updated++;
                }
                catch (Exception ex)
                {
                    result.Skipped++;
                    var rowLabel = $"{row.LayoutNo}/{row.LSSheet}/{row.WeldNumber}";
                    // Dig into the innermost exception for the actual DB error
                    // (DbUpdateException wraps the real cause in InnerException).
                    var root = ex;
                    while (root.InnerException != null) root = root.InnerException;
                    var reason = root.Message.Length < 200 ? root.Message : "Save error.";
                    result.Errors.Add($"Row {rowLabel}: {reason}");
                    result.RowRemarks[i] = reason;
                    _logger.LogWarning(ex, "CommitJointRecord row failed");

                    // Clear change tracker immediately so failed entities don't
                    // cascade into subsequent SaveChangesAsync calls.
                    _context.ChangeTracker.Clear();
                }

                // Periodically clear the change tracker to prevent memory bloat
                if ((i + 1) % CommitBatchSize == 0)
                {
                    _context.ChangeTracker.Clear();
                }
            }

            await transaction.CommitAsync(cancellationToken);

            result.Success = result.Errors.Count == 0;
            result.Warnings.Add($"Committed ({normalizedScope}): {result.Created} created, {result.Updated} updated, {result.Skipped} skipped.");

            // Deferred link updates: run after the transaction commits successfully
            foreach (var dfr in newDfrEntities)
            {
                try { await UpdateLineSheetAndSpoolRefsAsync(dfr, userId, now); }
                catch (Exception linkEx) { _logger.LogWarning(linkEx, "CommitJointRecord link refs failed Joint_ID={JointId}", dfr.Joint_ID); }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CommitJointRecord failed");
            var root = ex;
            while (root.InnerException != null) root = root.InnerException;
            result.Errors.Add($"Commit failed: {root.Message}");
        }

        return new JsonResult(result, JointRecordJsonOptions);
    }

    #endregion

    #region Bulk Analyze (Latest Drawings)

    /// <summary>
    /// Returns the list of Layout/Sheet pairs that have a latest drawing file (PDF)
    /// in the Drawings system for the given project.
    /// </summary>
    [SessionAuthorization]
    [HttpGet]
    public async Task<IActionResult> GetBulkAnalyzeDrawings(int projectId)
    {
        if (projectId <= 0) return Json(Array.Empty<BulkAnalyzeDrawingItem>());

        try
        {
            // Get all latest drawings per Line_Sheet (Sheet mode)
            var sheetDrawings = await (
                from d in _context.DWG_File_tbl.AsNoTracking()
                join ls in _context.Line_Sheet_tbl.AsNoTracking()
                    on d.Line_Sheet_ID equals ls.Line_Sheet_ID
                where d.Project_ID == projectId
                      && d.Mode == "Sheet"
                      && d.Line_Sheet_ID.HasValue
                select new
                {
                    d.Id,
                    d.RevisionOrder,
                    d.RevisionTag,
                    d.FileName,
                    d.Line_Sheet_ID,
                    ls.LS_LAYOUT_NO,
                    ls.LS_SHEET
                }).ToListAsync();

            // Group by Line_Sheet_ID and pick highest RevisionOrder (latest)
            var latestSheetDrawings = sheetDrawings
                .GroupBy(x => x.Line_Sheet_ID)
                .Select(g => g.OrderByDescending(x => x.RevisionOrder).ThenByDescending(x => x.Id).First())
                .Where(x => !string.IsNullOrWhiteSpace(x.LS_LAYOUT_NO))
                .Select(x => new BulkAnalyzeDrawingItem
                {
                    ProjectId = projectId,
                    Layout = x.LS_LAYOUT_NO!,
                    Sheet = x.LS_SHEET,
                    RevisionTag = x.RevisionTag,
                    FileName = x.FileName,
                    DrawingId = x.Id
                })
                .OrderBy(x => x.Layout).ThenBy(x => x.Sheet)
                .ToList();

            // Pre-fetch line IDs already covered by Sheet-mode drawings
            var coveredLineIds = await (
                from sd in _context.DWG_File_tbl.AsNoTracking()
                join ls2 in _context.Line_Sheet_tbl.AsNoTracking()
                    on sd.Line_Sheet_ID equals ls2.Line_Sheet_ID
                where sd.Project_ID == projectId
                      && sd.Mode == "Sheet"
                      && sd.Line_Sheet_ID.HasValue
                      && ls2.Line_ID_LS.HasValue
                select ls2.Line_ID_LS!.Value
            ).Distinct().ToListAsync();

            var coveredLineIdSet = new HashSet<int>(coveredLineIds);

            // Also get latest Line-mode drawings (no sheet)
            var lineDrawings = await (
                from d in _context.DWG_File_tbl.AsNoTracking()
                join ll in _context.LINE_LIST_tbl.AsNoTracking()
                    on d.Line_ID equals ll.Line_ID
                where d.Project_ID == projectId
                      && d.Mode == "Line"
                      && d.Line_ID.HasValue
                select new
                {
                    d.Id,
                    d.RevisionOrder,
                    d.RevisionTag,
                    d.FileName,
                    d.Line_ID,
                    ll.LAYOUT_NO
                }).ToListAsync();

            var latestLineDrawings = lineDrawings
                .Where(x => !coveredLineIdSet.Contains(x.Line_ID!.Value))
                .GroupBy(x => x.Line_ID)
                .Select(g => g.OrderByDescending(x => x.RevisionOrder).ThenByDescending(x => x.Id).First())
                .Where(x => !string.IsNullOrWhiteSpace(x.LAYOUT_NO))
                .Select(x => new BulkAnalyzeDrawingItem
                {
                    ProjectId = projectId,
                    Layout = x.LAYOUT_NO!,
                    Sheet = null,
                    RevisionTag = x.RevisionTag,
                    FileName = x.FileName,
                    DrawingId = x.Id
                })
                .OrderBy(x => x.Layout)
                .ToList();

            var all = latestSheetDrawings.Concat(latestLineDrawings).ToList();
            return Json(all);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetBulkAnalyzeDrawings failed for project={ProjectId}", projectId);
            return StatusCode(500, new { error = "Failed to retrieve drawings." });
        }
    }

    /// <summary>
    /// Bulk analyze all latest drawing PDFs for the given project.
    /// Retrieves each latest drawing from the Drawings/Revisions system, runs OCR analysis,
    /// and returns combined results.
    /// </summary>
    [SessionAuthorization]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> BulkAnalyzeJointRecord([FromBody] BulkAnalyzeRequest request)
    {
        if (request == null || request.ProjectId <= 0)
            return BadRequest(new { error = "Invalid request." });

        var bulkResult = new BulkAnalyzeResult { ProjectId = request.ProjectId };
        var analyzeScope = (request.Scope ?? ScopeJoints).Trim();
        var weldersProjectId = await ResolveWeldersProjectIdAsync(request.ProjectId);

        try
        {
            // Step 1: Collect all latest drawing rows from DWG_File_tbl
            var latestDrawings = await GetLatestDrawingRowsForProjectAsync(request.ProjectId);

            // Filter to selected drawings when specific IDs are provided
            if (request.DrawingIds is { Count: > 0 })
            {
                var selectedIds = new HashSet<int>(request.DrawingIds);
                latestDrawings = latestDrawings
                    .Where(d => selectedIds.Contains(d.Row.Id))
                    .ToList();
            }

            bulkResult.TotalDrawings = latestDrawings.Count;

            if (latestDrawings.Count == 0)
            {
                bulkResult.Warnings.Add("No latest drawings found for this project in the Drawings system.");
                return new JsonResult(bulkResult, JointRecordJsonOptions);
            }

            // Step 2: J-Add options (shared across all drawings)
            var jAddOpts = await _context.PMS_J_Add_tbl.AsNoTracking()
                .Where(x => x.Add_Project_No == weldersProjectId)
                .Where(x => x.Add_J_Add != null && x.Add_J_Add != "")
                .Select(x => x.Add_J_Add!)
                .Distinct()
                .ToListAsync();
            if (jAddOpts.Count == 0) jAddOpts = ["NEW", "R1", "R2"];

            // Step 3: Process each drawing
            var allRows = new List<JointRecordAnalysisRow>();
            foreach (var drawingInfo in latestDrawings)
            {
                try
                {
                    var filePath = TryResolveDrawingFilePath(drawingInfo.Row);
                    if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
                    {
                        bulkResult.Warnings.Add($"Drawing file not found: {drawingInfo.Layout}/{drawingInfo.Sheet} ({drawingInfo.Row.FileName})");
                        bulkResult.FailedDrawings++;
                        continue;
                    }

                    // Only process PDFs
                    var ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".pdf")
                    {
                        bulkResult.Warnings.Add($"Skipped non-PDF: {drawingInfo.Layout}/{drawingInfo.Sheet} ({ext})");
                        bulkResult.FailedDrawings++;
                        continue;
                    }

                    var drawingWarnings = new List<string>();
                    var fileName = drawingInfo.Row.FileName ?? Path.GetFileName(filePath);

                    await using var fileStream = System.IO.File.OpenRead(filePath);
                    var drawingRows = await AnalyzeSingleDrawingStreamAsync(
                        fileStream, fileName, request.ProjectId, analyzeScope, jAddOpts, drawingWarnings);

                    if (drawingRows.Count > 0)
                    {
                        allRows.AddRange(drawingRows);
                        bulkResult.SuccessfulDrawings++;
                        bulkResult.ProcessedDrawings.Add(new BulkAnalyzeDrawingItem
                        {
                            ProjectId = request.ProjectId,
                            Layout = drawingInfo.Layout,
                            Sheet = drawingInfo.Sheet,
                            RevisionTag = drawingInfo.Row.RevisionTag,
                            FileName = fileName,
                            DrawingId = drawingInfo.Row.Id
                        });
                    }
                    else
                    {
                        bulkResult.Warnings.Add($"No welds extracted from: {drawingInfo.Layout}/{drawingInfo.Sheet} ({fileName})");
                        bulkResult.FailedDrawings++;
                    }

                    // Prefix per-drawing warnings
                    foreach (var w in drawingWarnings)
                    {
                        bulkResult.Warnings.Add($"[{drawingInfo.Layout}/{drawingInfo.Sheet}] {w}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "BulkAnalyze: failed to process drawing {Layout}/{Sheet}",
                        drawingInfo.Layout, drawingInfo.Sheet);
                    bulkResult.Warnings.Add($"Error processing {drawingInfo.Layout}/{drawingInfo.Sheet}: {ex.Message}");
                    bulkResult.FailedDrawings++;
                }
            }

            // Step 4: Deduplicate and clean combined rows
            if (allRows.Count > 0)
            {
                bulkResult.Rows = CleanAndDeduplicateRows(allRows);
            }

            // Step 5: Populate action status remarks
            if (request.ProjectId > 0 && bulkResult.Rows.Count > 0)
            {
                await PopulateActionStatusRemarksAsync(request.ProjectId, analyzeScope, bulkResult.Rows);
            }

            bulkResult.Warnings.Add(
                $"Bulk analysis complete: {bulkResult.SuccessfulDrawings}/{bulkResult.TotalDrawings} drawings processed, {bulkResult.Rows.Count} total joint records.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "BulkAnalyzeJointRecord failed for project={ProjectId}", request.ProjectId);
            bulkResult.Errors.Add($"Bulk analysis error: {ex.Message}");
        }

        return new JsonResult(bulkResult, JointRecordJsonOptions);
    }

    /// <summary>
    /// Retrieves all latest drawing rows (highest RevisionOrder per Layout/Sheet)
    /// for the given project, including both Sheet-mode and Line-mode drawings.
    /// </summary>
    private async Task<List<(string Layout, string? Sheet, DwgFile Row)>> GetLatestDrawingRowsForProjectAsync(int projectId)
    {
        var result = new List<(string Layout, string? Sheet, DwgFile Row)>();

        // Sheet-mode drawings
        var sheetData = await (
            from d in _context.DWG_File_tbl.AsNoTracking()
            join ls in _context.Line_Sheet_tbl.AsNoTracking()
                on d.Line_Sheet_ID equals ls.Line_Sheet_ID
            where d.Project_ID == projectId && d.Mode == "Sheet" && d.Line_Sheet_ID.HasValue
            select new { Drawing = d, ls.LS_LAYOUT_NO, ls.LS_SHEET }
        ).ToListAsync();

        var latestSheet = sheetData
            .GroupBy(x => x.Drawing.Line_Sheet_ID)
            .Select(g => g.OrderByDescending(x => x.Drawing.RevisionOrder).ThenByDescending(x => x.Drawing.Id).First())
            .Where(x => !string.IsNullOrWhiteSpace(x.LS_LAYOUT_NO));

        foreach (var item in latestSheet)
        {
            result.Add((item.LS_LAYOUT_NO!, item.LS_SHEET, item.Drawing));
        }

        // Line-mode drawings (only if no Sheet-mode drawings exist for that line)
        // Batch-fetch all Line_ID_LS values in a single query instead of N+1
        var sheetLineSheetIds = sheetData
            .Where(x => x.Drawing.Line_Sheet_ID.HasValue)
            .Select(x => x.Drawing.Line_Sheet_ID!.Value)
            .Distinct()
            .ToList();

        var coveredLineIds = sheetLineSheetIds.Count > 0
            ? new HashSet<int>(
                await _context.Line_Sheet_tbl.AsNoTracking()
                    .Where(ls => sheetLineSheetIds.Contains(ls.Line_Sheet_ID) && ls.Line_ID_LS.HasValue)
                    .Select(ls => ls.Line_ID_LS!.Value)
                    .Distinct()
                    .ToListAsync())
            : new HashSet<int>();

        var lineData = await (
            from d in _context.DWG_File_tbl.AsNoTracking()
            join ll in _context.LINE_LIST_tbl.AsNoTracking()
                on d.Line_ID equals ll.Line_ID
            where d.Project_ID == projectId && d.Mode == "Line" && d.Line_ID.HasValue
            select new { Drawing = d, ll.Line_ID, ll.LAYOUT_NO }
        ).ToListAsync();

        var latestLine = lineData
            .Where(x => !coveredLineIds.Contains(x.Line_ID))
            .GroupBy(x => x.Line_ID)
            .Select(g => g.OrderByDescending(x => x.Drawing.RevisionOrder).ThenByDescending(x => x.Drawing.Id).First())
            .Where(x => !string.IsNullOrWhiteSpace(x.LAYOUT_NO));

        foreach (var item in latestLine)
        {
            result.Add((item.LAYOUT_NO!, null, item.Drawing));
        }

        return result.OrderBy(x => x.Layout).ThenBy(x => x.Sheet).ToList();
    }

    /// <summary>
    /// Core analysis logic for a single drawing stream.
    /// Extracted from AnalyzeJointRecord to enable reuse in bulk analysis.
    /// </summary>
    private async Task<List<JointRecordAnalysisRow>> AnalyzeSingleDrawingStreamAsync(
        Stream drawingStream,
        string fileName,
        int projectId,
        string scope,
        List<string> jAddOpts,
        List<string> warnings)
    {
        DrawingMetadata? metadata = null;
        var fileExtension = Path.GetExtension(fileName).ToLowerInvariant();

        switch (fileExtension)
        {
            case ".pcf":
                metadata = ParsePcfFile(drawingStream, fileName, warnings);
                break;
            case ".sha":
                metadata = ParseShaFile(drawingStream, fileName, warnings);
                break;
            case ".pdf":
                metadata = await ParsePdfDrawingAsync(drawingStream, fileName, warnings);
                break;
            default:
                warnings.Add($"Unsupported file format: {fileExtension}");
                return [];
        }

        if (metadata == null || metadata.Welds.Count == 0)
        {
            metadata ??= new DrawingMetadata();
            ExtractMetadataFromFileName(fileName, metadata);
        }

        var generatedRows = GenerateRowsFromTemplateMatch(metadata, [], fileName, warnings);
        if (generatedRows.Count == 0) return [];

        ApplyBusinessRules(generatedRows, warnings);
        SplitCompoundWeldNumbers(generatedRows, warnings);
        ExtractJAddFromWeldNumbers(generatedRows, jAddOpts, warnings);

        var orderedSpools = metadata?.OrderedSpools ?? [];
        var cuttingListSpoolSizes = metadata?.CuttingListSpoolSizes
            ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ReassignSpoolsBySuccessiveWsGroup(generatedRows, orderedSpools, warnings);
        PopulateSpoolDiameterFromGroup(generatedRows, cuttingListSpoolSizes, warnings);
        PopulateOLFieldsFromGroupMax(generatedRows, warnings);
        PopulateIsoDiameterFromGroupMax(generatedRows);
        await PopulateMaterialFromGroupMaxGradeAsync(generatedRows, warnings);

        NormalizeRevisionFields(generatedRows);
        NormalizeLineNumberFromLayout(generatedRows);
        PopulateIsoFromLayoutAndClass(generatedRows);
        BackfillFromIso(generatedRows);

        foreach (var row in generatedRows)
        {
            ValidateAndCorrectJointRecordRow(row, warnings);
            AutoPopulateMissingFields(row);
        }

        if (projectId > 0)
        {
            await ValidateAgainstLookupTablesAsync(projectId, generatedRows, warnings);
        }

        return CleanAndDeduplicateRows(generatedRows);
    }

    #endregion
}
