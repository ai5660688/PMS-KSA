using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("LINE_LIST_tbl")]
public class LineList
{
    [Key]
    public int Line_ID { get; set; }

    [MaxLength(10)] public string? LAYOUT_NO { get; set; }
    [MaxLength(6)] public string? Unit_Number { get; set; }
    [MaxLength(15)] public string? PDS_Model { get; set; }
    [MaxLength(24)] public string? ISO { get; set; }
    [MaxLength(10)] public string? Line_Class { get; set; }
    [MaxLength(30)] public string? Material_Des { get; set; }
    [MaxLength(10)] public string? Material { get; set; }
    [MaxLength(20)] public string? LINE_NO { get; set; }
    [MaxLength(4)] public string? Fluid { get; set; }
    [MaxLength(10)] public string? Design_Temperature_F { get; set; }
    [MaxLength(40)] public string? Category { get; set; }

    public double? RT_Shop { get; set; }
    public double? RT_Field { get; set; }
    public double? RT_Field_Shop_SW { get; set; }
    public double? MT_Shop { get; set; }
    public double? MT_Field { get; set; }
    public double? PT_Shop { get; set; }
    public double? PT_Field { get; set; }

    public bool? PWHT_Y_N { get; set; }
    public bool? PWHT_20mm { get; set; }

    public double? HT { get; set; }
    public double? HT_After_PWHT { get; set; }

    public bool? PMI { get; set; }

    [MaxLength(50)] public string? DWG_Remarks { get; set; }
    [MaxLength(24)] public string? P_ID { get; set; }
    [MaxLength(8)] public string? System { get; set; }
    [MaxLength(11)] public string? Sub_System { get; set; }
    [MaxLength(35)] public string? Test_Package_No { get; set; }
    [MaxLength(10)] public string? Test_Type { get; set; }
    [MaxLength(7)] public string? Spool_No { get; set; }
    [MaxLength(35)] public string? Test_Package_No_WS { get; set; }

    [Column(TypeName = "datetime")] public DateTime? Line_List_Updated_Date { get; set; }
    public int? Line_List_Updated_By { get; set; }

    [MaxLength(100)] public string? DWG_L_FileName { get; set; }
    public int? DWG_L_FileSize { get; set; }
    [Column(TypeName = "datetime2")] public DateTime? DWG_L_UploadDate { get; set; }
    [MaxLength(105)] public string? DWG_L_BlobName { get; set; }

    [MaxLength(8)] public string? L_REV { get; set; }
}
