using System;

namespace PMS.Models;

public class DwrRowUpdateDto
{
    public int JointId { get; set; }
    public string? FitupReport { get; set; }
    public int? RfiId { get; set; }
    public string? RfiDisplay { get; set; }
    public DateTime? FitupDate { get; set; } // DATE_WELDED
    public DateTime? ActualDate { get; set; } // ACTUAL_DATE_WELDED
    public int? WpsId { get; set; }
    public string? Wps { get; set; }

    // DWR crew fields
    public string? TackWelder { get; set; } // ROOT_A
    public string? TackWelderB { get; set; } // ROOT_B
    public string? TackWelderFillA { get; set; } // FILL_A
    public string? TackWelderFillB { get; set; } // FILL_B
    public string? TackWelderCapA { get; set; } // CAP_A
    public string? TackWelderCapB { get; set; } // CAP_B

    public double? PreheatTempC { get; set; }
    public string? IP_or_T { get; set; }
    public string? Open_Closed { get; set; }

    // Row state and remarks
    public string? DwrRemarks { get; set; }
    public string? Remarks { get; set; }
    public bool? Deleted { get; set; }
    public bool? Cancelled { get; set; }
    public bool? FitupConfirmed { get; set; }
}
