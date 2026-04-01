using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PMS.Models;

namespace PMS.Data
{
    public class ApplicationDbContext(DbContextOptions<ApplicationDbContext> options, ILogger<ApplicationDbContext> logger) : DbContext(options)
    {
        private readonly ILogger<ApplicationDbContext> _logger = logger;

        public DbSet<PMSLogin> PMS_Login_tbl { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PMSLogin>().ToTable("PMS_Login_tbl");
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.LogTo(message => _logger.LogInformation("{Message}", message));
        }
    }
}