using System;

namespace PMS.Infrastructure
{
    /// <summary>
    /// Centralized application clock. Project default time zone: UTC+03:00 (e.g. Arabia Standard Time).
    /// Use AppClock.Now for all persisted timestamps that should reflect the project local time.
    /// Keep UtcNow wrapper where true UTC is explicitly required.
    /// </summary>
    public static class AppClock
    {
        private static readonly TimeSpan ProjectOffset = TimeSpan.FromHours(3);

        /// <summary>
        /// Project local time (UTC+03:00).
        /// </summary>
        public static DateTime Now => DateTime.UtcNow.Add(ProjectOffset);

        /// <summary>
        /// True UTC time (pass-through) when needed for external integrations.
        /// </summary>
        public static DateTime UtcNow => DateTime.UtcNow;

        /// <summary>
        /// Normalize any incoming DateTime to project local (UTC+03:00) without carrying a Kind/offset.
        /// Rules:
        /// - If dt.Kind == Utc, shift by +03:00.
        /// - If dt.Kind == Local, shift by (targetOffset - systemLocalOffset at that time).
        /// - If dt.Kind == Unspecified, assume it's already project local and keep as-is.
        /// The returned DateTime has Kind=Unspecified for storage in SQL datetime columns.
        /// </summary>
        public static DateTime ToProjectLocal(DateTime dt)
        {
            var targetOffset = ProjectOffset;
            if (dt.Kind == DateTimeKind.Utc)
            {
                return DateTime.SpecifyKind(dt, DateTimeKind.Unspecified).Add(targetOffset);
            }
            if (dt.Kind == DateTimeKind.Local)
            {
                var sysOffset = TimeZoneInfo.Local.GetUtcOffset(dt);
                return DateTime.SpecifyKind(dt, DateTimeKind.Unspecified).Add(targetOffset - sysOffset);
            }
            // Unspecified -> treat as already in project local
            return dt;
        }
    }
}
