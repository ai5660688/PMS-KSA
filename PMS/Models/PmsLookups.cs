using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("PMS_Location_tbl")]
public class PmsLocation
{
    [Key]
    public int LO_ID { get; set; }
    [MaxLength(8)] public string? LO_Location { get; set; }
    [MaxLength(8)] public string? P_Location { get; set; }
    [MaxLength(100)] public string? LO_Notes { get; set; }
    public int LO_Project_No { get; set; }
}

[Table("PMS_J_Add_tbl")]
public class PmsJAdd
{
    [Key]
    public int Add_ID { get; set; }
    [MaxLength(8)] public string? Add_J_Add { get; set; }
    [MaxLength(8)] public string? P_J_Add { get; set; }
    [MaxLength(100)] public string? Add_Notes { get; set; }
    public int Add_Project_No { get; set; }
}

[Table("PMS_Weld_Type_tbl")]
public class PmsWeldType
{
    [Key]
    public int W_Type_ID { get; set; }
    [MaxLength(8)] public string? W_Weld_Type { get; set; }
    [MaxLength(8)] public string? P_Weld_Type { get; set; }
    public bool Default_Value { get; set; } // bit NOT NULL
    public bool PROG_Default_Value { get; set; } // bit NOT NULL
    [MaxLength(100)] public string? W_Notes { get; set; }
    public int W_Project_No { get; set; }
}

[Table("PMS_IP_T_tbl")]
public class PmsIpT
{
    [Key]
    public int IP_T_ID { get; set; }
    [MaxLength(8)] public string? IP_T_List { get; set; }
    [MaxLength(8)] public string? P_IP_T_List { get; set; }
    [MaxLength(100)] public string? IP_T_Notes { get; set; }
    public int IP_T_Project_No { get; set; }
}
