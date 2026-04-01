using System.Collections.Generic;
using System.Threading.Tasks;
using PMS.Models;

namespace PMS.Services;

public interface IWpsSelectorService
{
    /// <summary>
    /// Resolve (or derive) the effective thickness that should be used for WPS selection.
    /// Returns (thickness, reason/null).
    /// </summary>
    Task<(double? thickness, string? reason)> ResolveThicknessAsync(
        int projectId,
        string? explicitLineClass,
        double? explicitThickness,
        string? sch,
        double? dia,
        string? olSch,
        double? olDia,
        double? olThick);

    /// <summary>
    /// Resolve OLET actual thickness based on schedule join and line material rules.
    /// When material matches 'SS' => OLET_THICK_SS, 'SSS' => OLET_THICK_SSS, contains 'CS' => OLET_THICK_CS.
    /// Falls back to THICK if specific column is null. Returns (thickness, reason/null)
    /// </summary>
    Task<(double? thickness, string? reason)> ResolveOletThicknessAsync(string? material, string? olSchedule, double? olDiameter);

    /// <summary>
    /// Get WPS candidates filtered by project (or global), line class tokens and thickness.
    /// Returns ordered list (PWHT, Thickness_To, Code).
    /// </summary>
    Task<IReadOnlyList<WpsOption>> GetCandidatesAsync(int projectId, string? lineClass, double? effectiveThickness);
}
