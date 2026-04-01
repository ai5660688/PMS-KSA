using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.EntityFrameworkCore;

namespace PMS.Models
{
    public class PMSLogin
    {
        [Key]
        public int UserID { get; set; } // This is your primary key

        [MaxLength(20)]
        public string? FirstName { get; set; }

        [MaxLength(20)]
        public string? LastName { get; set; }

        [MaxLength(40)]
        public string? UserName { get; set; }

        [Required]
        [MaxLength(200)]
        [Unicode(false)] // varchar(200)
        public string Password { get; set; } = string.Empty;

        [MaxLength(50)]
        public string? Email { get; set; }

        [MaxLength(50)]
        public string? Position { get; set; }

        [Required]
        public bool Approved { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime? CreatedDate { get; set; }

        [MaxLength(20)]
        public string? Access { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime? DemobilisedDate { get; set; }

        [MaxLength(20)]
        public string? Company { get; set; }

        [MaxLength(200)]
        [Unicode(false)] // varchar(200)
        public string? ResetToken { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime? ResetExpiry { get; set; }

        [MaxLength(200)]
        [Unicode(false)] // varchar(200)
        public string? SessionToken { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime? SessionIssuedUtc { get; set; }

        [Column(TypeName = "datetime")]
        public DateTime? Login_Approval_Date { get; set; }

        public int? Login_Approval_By { get; set; }
    }
}