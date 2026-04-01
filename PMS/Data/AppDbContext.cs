using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PMS.Models;

namespace PMS.Data
{
    public class AppDbContext(DbContextOptions<AppDbContext> options, ILogger<AppDbContext> logger) : DbContext(options)
    {
        private readonly ILogger<AppDbContext> _logger = logger;

        public DbSet<PMSLogin> PMS_Login_tbl { get; set; }
        public DbSet<Project> Projects_tbl { get; set; }
        public DbSet<Wps> WPS_tbl { get; set; } // NEW
        public DbSet<Welder> Welders_tbl { get; set; }
        public DbSet<WelderQualification> Welder_List_tbl { get; set; }
        public DbSet<Dfr> DFR_tbl { get; set; }
        public DbSet<Rfi> RFI_tbl { get; set; }
        public DbSet<LineSheet> Line_Sheet_tbl { get; set; }
        public DbSet<LineList> LINE_LIST_tbl { get; set; }
        public DbSet<Schedule> Schedule_tbl { get; set; }
        // Lookup tables for DFR dropdowns
        public DbSet<PmsLocation> PMS_Location_tbl { get; set; }
        public DbSet<PmsJAdd> PMS_J_Add_tbl { get; set; }
        public DbSet<PmsWeldType> PMS_Weld_Type_tbl { get; set; }
        public DbSet<SpRelease> SP_Release_tbl { get; set; }
        public DbSet<UpdatedConfirmed> PMS_Updated_Confirmed_tbl { get; set; }
        public DbSet<MaterialTrace> MATERIAL_TRACE_tbl { get; set; }
        public DbSet<MaterialDes> Material_Des_tbl { get; set; }
        public DbSet<Material> Material_tbl { get; set; }
        public DbSet<PmsIpT> PMS_IP_T_tbl { get; set; }
        // Added: DWR and Other NDE tables
        public DbSet<Dwr> DWR_tbl { get; set; }
        public DbSet<OtherNde> Other_NDE_tbl { get; set; }
        // Added: RT and PWHT/HT tables
        public DbSet<Rt> RT_tbl { get; set; }
        public DbSet<PwhtHt> PWHT_HT_tbl { get; set; }
        // Added: Coating and Dispatch/Receiving tables
        public DbSet<Coating> Coating_tbl { get; set; }
        public DbSet<DispatchReceiving> Dispatch_Receiving_tbl { get; set; }
        // Added: DWG files
        public DbSet<DwgFile> DWG_File_tbl { get; set; }
        // Added: Receivers (keyless)
        public DbSet<Receiver> Receivers_tbl { get; set; }
        public DbSet<ReportNoForm> Report_No_Form_tbl { get; set; }
        public DbSet<Pln> PLN_tbl { get; set; }
        public DbSet<LotNo> Lot_No_tbl { get; set; }
        public DbSet<TransmittalLog> Transmittal_Log_tbl { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PMSLogin>().ToTable("PMS_Login_tbl");
            modelBuilder.Entity<Project>().ToTable("Projects_tbl");
            modelBuilder.Entity<Wps>().ToTable("WPS_tbl");
            modelBuilder.Entity<Welder>().ToTable("Welders_tbl");
            modelBuilder.Entity<WelderQualification>().ToTable("Welder_List_tbl");
            modelBuilder.Entity<Dfr>().ToTable("DFR_tbl");
            modelBuilder.Entity<Rfi>().ToTable("RFI_tbl");
            modelBuilder.Entity<LineSheet>().ToTable("Line_Sheet_tbl");
            modelBuilder.Entity<LineList>().ToTable("LINE_LIST_tbl");
            modelBuilder.Entity<Schedule>().ToTable("Schedule_tbl");
            modelBuilder.Entity<PmsLocation>().ToTable("PMS_Location_tbl");
            modelBuilder.Entity<PmsJAdd>().ToTable("PMS_J_Add_tbl");
            modelBuilder.Entity<PmsWeldType>().ToTable("PMS_Weld_Type_tbl");
            modelBuilder.Entity<SpRelease>().ToTable("SP_Release_tbl");
            modelBuilder.Entity<UpdatedConfirmed>().ToTable("PMS_Updated_Confirmed_tbl");
            modelBuilder.Entity<MaterialTrace>().ToTable("MATERIAL_TRACE_tbl");
            modelBuilder.Entity<MaterialDes>().ToTable("Material_Des_tbl");
            modelBuilder.Entity<Material>().ToTable("Material_tbl");
            modelBuilder.Entity<PmsIpT>().ToTable("PMS_IP_T_tbl");
            modelBuilder.Entity<Dwr>().ToTable("DWR_tbl");
            modelBuilder.Entity<OtherNde>().ToTable("Other_NDE_tbl");
            // Explicitly map RT and PWHT/HT tables (also annotated on models)
            modelBuilder.Entity<Rt>().ToTable("RT_tbl");
            modelBuilder.Entity<PwhtHt>().ToTable("PWHT_HT_tbl");
            modelBuilder.Entity<Coating>().ToTable("Coating_tbl");
            modelBuilder.Entity<DispatchReceiving>().ToTable("Dispatch_Receiving_tbl");
            modelBuilder.Entity<DwgFile>().ToTable("DWG_File_tbl");
            // Keyless entity mapping
            modelBuilder.Entity<Receiver>().HasNoKey().ToTable("Receivers_tbl");
            modelBuilder.Entity<ReportNoForm>().HasNoKey().ToTable("Report_No_Form_tbl");
            modelBuilder.Entity<Pln>().ToTable("PLN_tbl");
            modelBuilder.Entity<LotNo>().ToTable("Lot_No_tbl");
            modelBuilder.Entity<TransmittalLog>().ToTable("Transmittal_Log_tbl");

             // WelderQualification: primary key is JCC_No (set via data annotations)
            modelBuilder.Entity<WelderQualification>()
                .HasIndex(q => q.Welder_ID_WL); // index to speed up lookups by welder

            // One-to-many relationship: Welder has many Qualifications
            modelBuilder.Entity<Welder>()
                .HasMany(w => w.Qualifications)
                .WithOne(q => q.Welder)
                .HasForeignKey(q => q.Welder_ID_WL)
                .OnDelete(DeleteBehavior.Cascade);
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.LogTo(message => _logger.LogInformation("{Message}", message));
        }
    }
}