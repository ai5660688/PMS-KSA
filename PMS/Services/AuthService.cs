using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PMS.Data;
using PMS.Infrastructure; // added for AppClock
using PMS.Models;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace PMS.Services
{
    public class AuthService(AppDbContext context, ILogger<AuthService> logger, IPasswordHasher<PMSLogin> passwordHasher, IConfiguration configuration)
    {
        private readonly AppDbContext _context = context;
        private readonly ILogger<AuthService> _logger = logger;
        private readonly IPasswordHasher<PMSLogin> _passwordHasher = passwordHasher;

        private readonly byte[] _resetTokenKey = GetResetTokenKey(configuration);

        internal static string HashToken(string token)
            => Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(token)));

        private static byte[] GetResetTokenKey(IConfiguration configuration)
        {
            var key = configuration["Security:ResetTokenKey"];
            if (string.IsNullOrWhiteSpace(key))
                throw new InvalidOperationException("Security:ResetTokenKey is not configured. Set it in appsettings.json, User Secrets, or Key Vault.");
            return Encoding.UTF8.GetBytes(key);
        }

        private string HashResetToken(string token)
        {
            ArgumentException.ThrowIfNullOrEmpty(token);
            using var hmac = new HMACSHA256(_resetTokenKey);
            return Convert.ToBase64String(hmac.ComputeHash(Encoding.UTF8.GetBytes(token)));
        }

        public string HashPasswordForStorage(PMSLogin user, string password)
            => _passwordHasher.HashPassword(user, password);

        public string GenerateResetToken()
        {
            Span<byte> buffer = stackalloc byte[32];
            RandomNumberGenerator.Fill(buffer);
            return Convert.ToBase64String(buffer); // transient token returned to caller (never stored directly)
        }

        public async Task<PMSLogin?> Authenticate(string username, string password)
        {
            try
            {
                _logger.LogInformation("Authenticating user: {UserName}", username);

                var user = await _context.PMS_Login_tbl
                    .FirstOrDefaultAsync(u => u.UserName != null && u.Password != null && u.UserName.Trim() == username);

                if (user == null)
                {
                    _logger.LogWarning("Authentication failed for user: {UserName}. Invalid credentials.", username);
                    return default;
                }

                var storedPassword = (user.Password ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(storedPassword))
                {
                    _logger.LogWarning("Authentication failed for user: {UserName}. Empty stored password.", username);
                    return default;
                }

                // Detect legacy/plain values (short length or non-Identity format)
                bool looksLikeHash = storedPassword.Length >= 40 && storedPassword.StartsWith("AQAAAA", StringComparison.Ordinal);

                var verification = looksLikeHash
                    ? _passwordHasher.VerifyHashedPassword(user, storedPassword, password)
                    : PasswordVerificationResult.Failed;

                if (verification == PasswordVerificationResult.SuccessRehashNeeded)
                {
                    user.Password = _passwordHasher.HashPassword(user, password);
                    await _context.SaveChangesAsync();
                }
                else if (verification == PasswordVerificationResult.Failed)
                {
                    // Legacy plain-text (or truncated) support: migrate to hashed if it matches after trimming
                    if (string.Equals(storedPassword, password, StringComparison.Ordinal))
                    {
                        user.Password = _passwordHasher.HashPassword(user, password);
                        await _context.SaveChangesAsync();
                    }
                    else
                    {
                        _logger.LogWarning("Authentication failed for user: {UserName}. Invalid credentials.", username);
                        return default;
                    }
                }

                _logger.LogInformation("Credentials valid for user: {UserName}", username);
                return user;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error authenticating user: {UserName}", username);
                return default;
            }
        }

        public async Task<PMSLogin?> GetUserByEmail(string email)
        {
            try
            {
                var input = (email ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(input)) return null;
                _logger.LogInformation("Looking up user by email: {Email}", input);

                // Case-insensitive match relies on database collation (default SQL Server is case-insensitive).
                // Removed ToUpper() normalization to satisfy CA1862 and avoid unnecessary allocations.
                var user = await _context.PMS_Login_tbl
                    .AsNoTracking()
                    .FirstOrDefaultAsync(u =>
                        u.Email != null && u.Email.Trim() == input);

                if (user == null)
                {
                    _logger.LogWarning("No user found for email (normalized match failed): {Email}", input);
                }
                return user;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error looking up user by email: {Email}", email);
                return default;
            }
        }

        public async Task CreateUser(PMSLogin user)
        {
            try
            {
                _logger.LogInformation("Creating new user: {UserName}", user.UserName);

                // Keep reset fields null on sign-up (set only during password reset)
                user.ResetToken = null;
                user.ResetExpiry = null;

                // Optional: stamp created date on sign-up (project local time)
                user.CreatedDate ??= AppClock.Now;

                _context.PMS_Login_tbl.Add(user);
                await _context.SaveChangesAsync();
                _logger.LogInformation("Successfully created user: {UserName}", user.UserName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating user: {UserName}", user.UserName);
                throw;
            }
        }

        public async Task UpdateResetToken(string email, string token)
        {
            var input = (email ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(input))
            {
                _logger.LogWarning("Reset token update skipped: empty email input");
                return;
            }
            // Case-insensitive match relies on database collation; removed ToUpper() per CA1862.
            var user = await _context.PMS_Login_tbl.FirstOrDefaultAsync(u => u.Email != null && u.Email.Trim() == input);
            if (user == null)
            {
                _logger.LogWarning("Password reset requested for unknown email: {Email}", input);
                return; // keep response generic
            }

            user.ResetToken = HashResetToken(token);
            // Expiry still best expressed in absolute UTC for validation logic
            user.ResetExpiry = AppClock.UtcNow.AddMinutes(30); // token valid for 30 minutes
            await _context.SaveChangesAsync();

            _logger.LogInformation("Set reset token for UserID {UserID} expiring at {ExpiryUtc}", user.UserID, user.ResetExpiry);
        }

        public string HashResetTokenForLookup(string token) => HashResetToken(token);

    }
}