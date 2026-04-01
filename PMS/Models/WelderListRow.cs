namespace PMS.Models;

public sealed class WelderListRow
{
    public int Welder_ID { get; set; }
    public string? Welder_Symbol { get; set; }
    public string? Name { get; set; }
    public string? Welder_Location { get; set; }
    // Aggregated (comma separated distinct) welding processes for this welder
    public string? Welding_Process { get; set; }
    // Keep original dates (not shown now, but Status uses Demobilization_Date)
    public DateTime? Mobilization_Date { get; set; }
    public DateTime? Demobilization_Date { get; set; }
    // Display status (auto "Demobilized" when Demobilization_Date has value)
    public string? Status { get; set; }
    public string? Project_Welder { get; set; }
    // First batch kept for legacy (if needed elsewhere)
    public int? Batch_No { get; set; }
    // Aggregated (comma separated distinct) batch numbers for display
    public string? Batch_Nos { get; set; }
    // Computed next continuity date (DATE_OF_LAST_CONTINUITY or Test_Date) + 6 months
    public DateTime? Next_Continuity { get; set; }
}