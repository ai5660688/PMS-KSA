using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MimeKit;
using System;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic; // added for inline resources
using System.Linq; // added for address splitting and list handling

namespace PMS.Services
{
    public class EmailService(IOptions<EmailSettings> settings, ILogger<EmailService> logger)
    {
        private readonly EmailSettings _settings = settings.Value;
        private readonly ILogger<EmailService> _logger = logger;

        public async Task SendEmailAsync(string toEmail, string subject, string htmlMessage)
        {
            await SendEmailAsync(new[] { toEmail }, null, subject, htmlMessage, false, inlineResources: null);
        }

        // New overload: support inline (CID) resources like images
        public async Task SendEmailAsync(
            string toEmail,
            string subject,
            string htmlMessage,
            Dictionary<string, (byte[] Content, string ContentType)>? inlineResources)
        {
            await SendEmailAsync(new[] { toEmail }, null, subject, htmlMessage, false, inlineResources);
        }

        // NEW: multi-recipient + cc + importance overload
        public async Task SendEmailAsync(IEnumerable<string> toEmails, IEnumerable<string>? ccEmails, string subject, string htmlMessage, bool highImportance = false)
        {
            await SendEmailAsync(toEmails, ccEmails, subject, htmlMessage, highImportance, inlineResources: null);
        }

        // Core implementation
        public async Task SendEmailAsync(
            IEnumerable<string> toEmails,
            IEnumerable<string>? ccEmails,
            string subject,
            string htmlMessage,
            bool highImportance,
            Dictionary<string, (byte[] Content, string ContentType)>? inlineResources)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_settings.Host) ||
                    _settings.Port <= 0 ||
                    string.IsNullOrWhiteSpace(_settings.User) ||
                    string.IsNullOrWhiteSpace(_settings.Password))
                {
                    throw new Exception("SMTP settings are missing or invalid.");
                }

                var userAddress = _settings.User.Trim();
                // Prefer the configured From if provided; otherwise fall back to the authenticated user
                var fromAddress = string.IsNullOrWhiteSpace(_settings.From) ? userAddress : _settings.From.Trim();

                // Use configurable display name
                var displayName = string.IsNullOrWhiteSpace(_settings.DisplayName) ? "PMS System" : _settings.DisplayName.Trim();

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(displayName, fromAddress));
                // Only set Sender when it differs from From (avoids some clients showing "on behalf of")
                if (!fromAddress.Equals(userAddress, StringComparison.OrdinalIgnoreCase))
                {
                    message.Sender = new MailboxAddress(displayName, userAddress);
                }

                // Add recipients
                var toList = (toEmails ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).SelectMany(SplitAddresses).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (toList.Count == 0) throw new Exception("No recipients provided.");
                foreach (var addr in toList)
                {
                    message.To.Add(MailboxAddress.Parse(addr));
                }
                var ccList = (ccEmails ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).SelectMany(SplitAddresses).Distinct(StringComparer.OrdinalIgnoreCase).Except(toList, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var addr in ccList)
                {
                    message.Cc.Add(MailboxAddress.Parse(addr));
                }

                message.Subject = subject;

                // High importance headers
                if (highImportance)
                {
                    message.Headers.Add("X-Priority", "1"); // High
                    message.Headers.Add("Importance", "high");
                    message.Headers.Add("Priority", "urgent");
                }

                var builder = new BodyBuilder { HtmlBody = htmlMessage };
                if (inlineResources != null)
                {
                    foreach (var kvp in inlineResources)
                    {
                        var id = kvp.Key; // also used as Content-Id
                        var (bytes, contentType) = kvp.Value;
                        var lr = builder.LinkedResources.Add(id, bytes);
                        lr.ContentId = id;
                        lr.ContentDisposition = new ContentDisposition(ContentDisposition.Inline);
                        // Set media type by updating existing ContentType (property is read-only but mutable)
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(contentType) && contentType.Contains('/'))
                            {
                                var parts = contentType.Split('/', 2);
                                lr.ContentType.MediaType = parts[0];
                                lr.ContentType.MediaSubtype = parts[1];
                            }
                        }
                        catch { /* ignore parse/mutation errors */ }
                    }
                }
                message.Body = builder.ToMessageBody();

                using var client = new SmtpClient
                {
                    // Reduce chance of hanging forever when egress is blocked
                    Timeout = 15000 // 15 seconds
                };

                // Prefer explicit options when provided; otherwise infer from common ports
                var preferredOptions = _settings.Port == 465 ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTls;

                _logger.LogInformation("Connecting to SMTP server: {Server}:{Port}", _settings.Host, _settings.Port);

                // Provide a cancellation budget for the whole connect+auth sequence
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(20));

                await ConnectAuthenticateWithFallbackAsync(client, _settings.Host, _settings.Port, preferredOptions, userAddress, _settings.Password, cts.Token);

                _logger.LogInformation("Sending email to: {Recipient}", string.Join(", ", toList));
                await client.SendAsync(message, cts.Token);
                await client.DisconnectAsync(true, cts.Token);
                _logger.LogInformation("Email sent successfully to: {Recipient}", string.Join(", ", toList));
            }
            catch (AuthenticationException authEx)
            {
                _logger.LogError(authEx, "SMTP authentication failed for user: {User}", _settings.User);
                throw new Exception("Email service authentication failed. Please check your credentials.", authEx);
            }
            catch (MailKit.Net.Smtp.SmtpCommandException smtpEx)
            {
                _logger.LogError(smtpEx, "SMTP command failed: {StatusCode} {Message}", smtpEx.StatusCode, smtpEx.Message);
                throw new Exception("Failed to send email due to SMTP command error.", smtpEx);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to send email");
                throw new Exception("Failed to send email. Please try again later.", ex);
            }
        }

        private static IEnumerable<string> SplitAddresses(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) yield break;
            foreach (var part in s.Split(new[] { ';', ',', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var v = part.Trim();
                if (!string.IsNullOrWhiteSpace(v)) yield return v;
            }
        }

        private async Task ConnectAuthenticateWithFallbackAsync(SmtpClient client, string host, int configuredPort, SecureSocketOptions preferredOptions, string user, string password, CancellationToken ct)
        {
            // Build attempts: configured, then the common alternative port with appropriate TLS mode
            var attempts = new (int Port, SecureSocketOptions Options, string Reason)[]
            {
                (configuredPort, preferredOptions, "configured"),
                (configuredPort == 465 ? 587 : 465, configuredPort == 465 ? SecureSocketOptions.StartTls : SecureSocketOptions.SslOnConnect, "fallback")
            };

            Exception? lastError = null;
            foreach (var attempt in attempts)
            {
                try
                {
                    _logger.LogInformation("SMTP connect attempt ({Reason}): {Host}:{Port} with {Mode}", attempt.Reason, host, attempt.Port, attempt.Options);
                    await client.ConnectAsync(host, attempt.Port, attempt.Options, ct);

                    // Ensure we use basic auth (LOGIN/PLAIN) when XOAUTH2 is not configured
                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                    await client.AuthenticateAsync(new NetworkCredential(user, password), ct);
                    return; // success
                }
                catch (SocketException se)
                {
                    lastError = se;
                    _logger.LogWarning(se, "SMTP socket error on {Host}:{Port} ({Reason}). Will{Next} retry with alternative.", host, attempt.Port, attempt.Reason, attempt.Equals(attempts[^1]) ? " not" : "");
                    try { if (client.IsConnected) await client.DisconnectAsync(true, ct); } catch { /* ignore */ }
                }
                catch (ServiceNotConnectedException snc)
                {
                    lastError = snc;
                    _logger.LogWarning(snc, "SMTP not connected on {Host}:{Port} ({Reason}). Will{Next} retry with alternative.", host, attempt.Port, attempt.Reason, attempt.Equals(attempts[^1]) ? " not" : "");
                    try { if (client.IsConnected) await client.DisconnectAsync(true, ct); } catch { /* ignore */ }
                }
                catch (SmtpProtocolException spe)
                {
                    lastError = spe;
                    _logger.LogWarning(spe, "SMTP protocol error on {Host}:{Port} ({Reason}). Will{Next} retry with alternative.", host, attempt.Port, attempt.Reason, attempt.Equals(attempts[^1]) ? " not" : "");
                    try { if (client.IsConnected) await client.DisconnectAsync(true, ct); } catch { /* ignore */ }
                }
            }

            throw new Exception($"Unable to connect/authenticate to SMTP server '{host}' using ports {configuredPort}/" + (configuredPort == 465 ? 587 : 465) + ". See inner error for details.", lastError);
        }
    }
}