using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PMS.Models;

[Table("RT_tbl")]
public class Rt
{
    [Key]
    public int Joint_ID_RT { get; set; }

    [MaxLength(8)] public string? BSR_NDE_REQUEST { get; set; }
    [Column(TypeName = "datetime")] public DateTime? BSR_DATE_NDE_WAS_REQ { get; set; }
    [Column(TypeName = "datetime")] public DateTime? BSR_Assessment_Date { get; set; }
    [MaxLength(11)] public string? BSR_RT_REPORT_NO { get; set; }
    [Column(TypeName = "datetime")] public DateTime? BSR_RT_DATE { get; set; }
    [MaxLength(4)] public string? BSR_RT_RESULT { get; set; }
    [MaxLength(150)] public string? LOCATION_MARKERS { get; set; }
    [MaxLength(30)] public string? BSR_NDE_Third_P { get; set; }
    public double? BSR_No_of_filmes { get; set; }
    [MaxLength(8)] public string? NDE_REQUEST { get; set; }
    [Column(TypeName = "datetime")] public DateTime? DATE_NDE_WAS_REQUESTED { get; set; }
    [Column(TypeName = "datetime")] public DateTime? Assessment_Date { get; set; }
    [MaxLength(9)] public string? NDT_LOG_NUMBER { get; set; }
    [MaxLength(11)] public string? RT_REPORT_NUMBER { get; set; }
    [Column(TypeName = "datetime")] public DateTime? Final_RT_DATE { get; set; }
    [MaxLength(6)] public string? Final_RT_RESULT { get; set; }
    [MaxLength(30)] public string? NDE_Third_P { get; set; }
    public double? No_of_filmes { get; set; }
    [MaxLength(30)] public string? REPAIR_TYPE { get; set; }
    [Column("LINEAR_REJECTED_LENGTH_(mm)")] public double? LINEAR_REJECTED_LENGTH_mm { get; set; }
    public double? REVIEWED { get; set; }
    [MaxLength(10)] public string? RS_Number_1 { get; set; }
    [Column(TypeName = "datetime")] public DateTime? RS_Date_1 { get; set; }
    [MaxLength(10)] public string? RS_Number_2 { get; set; }
    public DateTime? RS_Date_2 { get; set; }
    [MaxLength(10)] public string? RS_Number_3 { get; set; }
    [Column(TypeName = "datetime")] public DateTime? RS_Date_3 { get; set; }
    [MaxLength(10)] public string? RS_Number_4 { get; set; }
    [Column(TypeName = "datetime")] public DateTime? RS_Date_4 { get; set; }
    [Column(TypeName = "datetime")] public DateTime? THIRD_PARTY_DATE { get; set; }
    [MaxLength(4)] public string? THIRD_PARTY_RESULT { get; set; }
    [Column(TypeName = "datetime")] public DateTime? PRIME_CONTRACTOR_DATE { get; set; }
    [MaxLength(4)] public string? PRIME_CONTRACTOR_RESULT { get; set; }
    [MaxLength(10)] public string? RS_Request_1 { get; set; }
    [MaxLength(10)] public string? RS_Request_2 { get; set; }
    [MaxLength(10)] public string? RS_Request_3 { get; set; }
    [MaxLength(10)] public string? RS_Request_4 { get; set; }
    public double? OID_Films_Reviewed { get; set; }
    public bool? OID_Disagreement { get; set; }
    [MaxLength(50)] public string? NDE_REMARKS { get; set; }
    [MaxLength(9)] public string? Repair_Welder { get; set; }
    [MaxLength(6)] public string? NDE_Category { get; set; }
    [Column(TypeName = "datetime")] public DateTime? RT_Updated_Date { get; set; }
    public int? RT_Updated_By { get; set; }
}
