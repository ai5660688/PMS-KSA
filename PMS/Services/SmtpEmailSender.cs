using System.Net;
using System.Net.Mail;
using Microsoft.Extensions.Options;
using PMS.Options;

namespace PMS.Services;

public sealed class SmtpEmailSender(IOptions<SmtpOptions> opts) : IEmailSender
{
    private readonly SmtpOptions _o = opts.Value;

    public async Task SendHtmlAsync(string subject, string htmlBody, string? fromOverride = null, string? toOverride = null)
    {
        using var msg = new MailMessage
        {
            From = new MailAddress(fromOverride ?? _o.From),
            Subject = subject,
            Body = htmlBody,
            IsBodyHtml = true
        };
        msg.To.Add(new MailAddress(toOverride ?? _o.To));

        using var client = new SmtpClient(_o.Host, _o.Port)
        {
            EnableSsl = _o.EnableSsl,
            Credentials = new NetworkCredential(_o.User, _o.Password)
        };
        await client.SendMailAsync(msg);
    }
}