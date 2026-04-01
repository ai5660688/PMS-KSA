using System;

namespace PMS.Options;

public class AzureDocumentIntelligenceOptions
{
    public string? Endpoint { get; set; }
    public string? ApiKey { get; set; }
    public string ModelId { get; set; } = "prebuilt-read";
    public int PollIntervalMs { get; set; } = 1000;
    public int MaxWaitSeconds { get; set; } = 60;
}
