using System.Threading.Tasks;

namespace PMS.Services
{
    public interface IEmailSender
    {
        // Revert to the original signature to avoid Hot Reload ENC0023
        Task SendHtmlAsync(string subject, string htmlBody, string? fromOverride = null, string? toOverride = null);
    }
}