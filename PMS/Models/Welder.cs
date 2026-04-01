using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models
{
    [Table("Welders_tbl")]
    public class Welder
    {
        [Key]
        public int Welder_ID { get; set; }

        [Required, StringLength(12)]
        public string Welder_Symbol { get; set; } = string.Empty;

        [StringLength(24)]
        public string? Iqama_No { get; set; }

        [StringLength(24)]
        public string? Passport { get; set; }

        [Required, StringLength(100)]
        public string Name { get; set; } = string.Empty;

        [StringLength(5)]
        public string? Welder_Location { get; set; }

        [StringLength(12)]
        public string? Mobile_No { get; set; }

        [StringLength(50), EmailAddress]
        public string? Email { get; set; }

        public DateTime? Mobilization_Date { get; set; }
        public DateTime? Demobilization_Date { get; set; }

        [StringLength(35)]
        public string? Status { get; set; }

        // FK to Project
        public int? Project_Welder { get; set; }

        // Audit
        public DateTime? Welders_Updated_Date { get; set; }
        public int? Welders_Updated_By { get; set; }

        // One-to-many: a welder can have multiple qualifications
        public ICollection<WelderQualification> Qualifications { get; set; } = new List<WelderQualification>();
    }
}