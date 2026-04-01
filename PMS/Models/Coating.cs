using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("Coating_tbl")]
public class Coating
{
 [Key]
 public int Spool_ID_PA { get; set; }

 [MaxLength(15)]
 public string? Actual_Paint_Code { get; set; }

 public int? Surface_Preparation_3_1 { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Surface_Preparation_3_1_Date { get; set; }

 public int? Primer_Application_3_2 { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Primer_Application_3_2_Date { get; set; }

 public int? Primer_Inspection_3_3 { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Primer_Inspection_3_3_Date { get; set; }

 public int? Top_Coat_3_4 { get; set; }

 // Note: column name ends with lowercase 'date' per provided schema
 [Column(TypeName = "datetime")]
 public DateTime? Top_Coat_3_4_date { get; set; }

 public int? Intermediate_Coat_3_4 { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Intermediate_Coat_3_4_date { get; set; }

 public int? Final_Coating_3_5 { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Final_Coating_3_5_date { get; set; }

 public int? S_P_3_1_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? S_P_3_1_Date_SUPERSEDED { get; set; }

 public int? P_A_3_2_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? P_A_3_2_Date_SUPERSEDED { get; set; }

 public int? P_I_3_3_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? P_I_3_3_Date_SUPERSEDED { get; set; }

 public int? T_C_3_4_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? T_C_3_4_date_SUPERSEDED { get; set; }

 public int? I_C_3_4_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? I_C_3_4_date_SUPERSEDED { get; set; }

 public int? F_C_3_5_SUPERSEDED { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? F_C_3_5_date_SUPERSEDED { get; set; }

 [MaxLength(35)]
 public string? Coating_REMARKS { get; set; }

 [Column(TypeName = "datetime")]
 public DateTime? Coating_Updated_Date { get; set; }

 public int? Coating_Updated_By { get; set; }
}
