namespace PMS.Services;

/// <summary>
/// Thread-safe singleton that tracks whether the site is in maintenance mode.
/// Admin users can toggle this at runtime from the dashboard header and
/// optionally set an estimated completion time shown to end users.
/// </summary>
public sealed class MaintenanceService
{
    private volatile bool _isEnabled;
    private string _message = "The system is currently undergoing scheduled maintenance. Please try again later.";
    private DateTime? _estimatedEndUtc;

    public bool IsEnabled => _isEnabled;

    public string Message => _message;

    /// <summary>UTC time when maintenance is expected to finish (null = unknown).</summary>
    public DateTime? EstimatedEndUtc => _estimatedEndUtc;

    public void Enable(string? message = null, DateTime? estimatedEndUtc = null)
    {
        if (!string.IsNullOrWhiteSpace(message))
            _message = message;
        _estimatedEndUtc = estimatedEndUtc;
        _isEnabled = true;
    }

    public void Disable()
    {
        _isEnabled = false;
        _estimatedEndUtc = null;
    }
}
