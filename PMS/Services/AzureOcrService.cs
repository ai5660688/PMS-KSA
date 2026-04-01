using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PMS.Options;

namespace PMS.Services;

public interface IOcrService
{
    Task<string?> ExtractTextAsync(Stream pdfStream, string fileName, CancellationToken cancellationToken = default);
}

public class AzureOcrService(HttpClient httpClient, IOptions<AzureDocumentIntelligenceOptions> options, ILogger<AzureOcrService> logger) : IOcrService
{
    private readonly HttpClient _httpClient = httpClient;
    private readonly AzureDocumentIntelligenceOptions _options = options.Value;
    private readonly ILogger<AzureOcrService> _logger = logger;

    public async Task<string?> ExtractTextAsync(Stream pdfStream, string fileName, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(_options.Endpoint) || string.IsNullOrWhiteSpace(_options.ApiKey))
        {
            _logger.LogDebug("Azure Document Intelligence not configured; skipping OCR for {File}.", fileName);
            return null;
        }

        // Copy to a buffer so we don't dispose the caller's stream
        using var ms = new MemoryStream();
        pdfStream.Position = 0;
        await pdfStream.CopyToAsync(ms, cancellationToken).ConfigureAwait(false);
        ms.Position = 0;

        using var content = new ByteArrayContent(ms.ToArray());
        content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

        var endpoint = _options.Endpoint.TrimEnd('/');
        var modelId = string.IsNullOrWhiteSpace(_options.ModelId) ? "prebuilt-read" : _options.ModelId;
        var requestUri = $"{endpoint}/documentintelligence/documentModels/{modelId}:analyze?api-version=2024-02-29-preview";

        using var request = new HttpRequestMessage(HttpMethod.Post, requestUri)
        {
            Content = content
        };
        request.Headers.Add("Ocp-Apim-Subscription-Key", _options.ApiKey);

        var response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            _logger.LogWarning("OCR analyze call failed for {File} with status {Status}.", fileName, response.StatusCode);
            return null;
        }

        if (!response.Headers.TryGetValues("operation-location", out var locations))
        {
            _logger.LogWarning("OCR analyze call missing operation-location for {File}.", fileName);
            return null;
        }

        var operationLocation = locations.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(operationLocation))
        {
            _logger.LogWarning("OCR analyze call returned empty operation-location for {File}.", fileName);
            return null;
        }

        var deadline = DateTimeOffset.UtcNow.AddSeconds(Math.Max(5, _options.MaxWaitSeconds));
        while (DateTimeOffset.UtcNow < deadline)
        {
            await Task.Delay(Math.Max(200, _options.PollIntervalMs), cancellationToken).ConfigureAwait(false);

            using var statusRequest = new HttpRequestMessage(HttpMethod.Get, operationLocation);
            statusRequest.Headers.Add("Ocp-Apim-Subscription-Key", _options.ApiKey);
            var statusResponse = await _httpClient.SendAsync(statusRequest, cancellationToken).ConfigureAwait(false);
            if (!statusResponse.IsSuccessStatusCode)
            {
                _logger.LogWarning("OCR status call failed for {File} with status {Status}.", fileName, statusResponse.StatusCode);
                return null;
            }

            var json = await statusResponse.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var status = root.GetProperty("status").GetString();
            if (string.Equals(status, "succeeded", StringComparison.OrdinalIgnoreCase))
            {
                if (root.TryGetProperty("analyzeResult", out var analyzeResult))
                {
                    if (analyzeResult.TryGetProperty("content", out var contentElement))
                    {
                        return contentElement.GetString();
                    }

                    // Fallback: concatenate lines if content is missing
                    if (analyzeResult.TryGetProperty("pages", out var pages) && pages.ValueKind == JsonValueKind.Array)
                    {
                        var text = string.Join("\n", pages.EnumerateArray()
                            .SelectMany(p => p.GetProperty("lines").EnumerateArray())
                            .Select(l => l.GetProperty("content").GetString() ?? string.Empty));
                        return string.IsNullOrWhiteSpace(text) ? null : text;
                    }
                }

                return null;
            }

            if (string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogWarning("OCR analyze failed for {File}. Response: {Json}", fileName, json);
                return null;
            }
        }

        _logger.LogWarning("OCR analyze timed out for {File}.", fileName);
        return null;
    }
}
