namespace PMS.Services
{
    public class EmailSettings
    {
        public string Host { get; set; } = "";
        public int Port { get; set; }
        public bool EnableSsl { get; set; }
        public string User { get; set; } = "";
        public string Password { get; set; } = "";
        public string From { get; set; } = "";
        public string AdminEmail { get; set; } = "";
        public string InternalContactEmail { get; set; } = "";
        public string OwnerEmail { get; set; } = "";
        public string OwnerPassword { get; set; } = "";
        public string DisplayName { get; set; } = "PMS System";
    }
}