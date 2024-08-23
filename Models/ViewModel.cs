using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace CORSYS_API.Models
{
    public partial class AuthModel
    {
        public string DBName { get; set; }
        public string DestPath { get; set; }
        public string Token { get; set; }
        public string ProjectID { get; set; }
        public string CladLineNo { get; set; }
        public string json { get; set; }
        public bool isLive { get; set; }
        public bool isActive { get; set; }
    }
    public partial class UserModel
    {
        public string UserName { get; set; }
        public string Name { get; set; }
        public string Badge { get; set; }
        public string Title { get; set; }
        public string Code { get; set; }
        public string GroupCode { get; set; }
        public string JMessage { get; set; }
        public string JStatus { get; set; }
    }

    public class CRModel
    {
        public string Param { get; set; }
        public string Address { get; set; }
        public string Service { get; set; }
        public DataSet DS { get; set; }
        public List<SetSubReport> SubDS { get; set; }
        public System.Data.DataTable[] DT { get; set; }
        public System.Data.DataTable[] SubDT { get; set; }
        public string ReportID { get; set; }
        public int ID { get; set; }
        public string DestPath { get; set; }
    }
    public class SetSubReport
    {
        public int Sub { get; set; }
        public int DataSet { get; set; }
    }
    public class ExlHeader
    {
        public bool isMerge { get; set; }
        public int start_a { get; set; }
        public int end_a { get; set; }
        public int start_b { get; set; }
        public int end_b { get; set; }
        public string lbl { get; set; }
        public string rowno { get; set; }
    }
    public class DoubleString
    {
        public string str_a { get; set; }
        public string str_b { get; set; }
    }
    public class IntString
    {
        public int int_a { get; set; }
        public string str_b { get; set; }
    }

    public class GenericModel
    {
        public int int_a { get; set; }
        public int int_b { get; set; }
        public int int_c { get; set; }
        public int int_d { get; set; }
        public int int_e { get; set; }
        public int int_f { get; set; }
        public int int_g { get; set; }
        public int int_h { get; set; }
        public int int_i { get; set; }
        public int int_j { get; set; }
        public string str_a { get; set; }
        public string str_b { get; set; }
        public string str_c { get; set; }
        public string str_d { get; set; }
        public string str_e { get; set; }
        public string str_f { get; set; }
        public string str_g { get; set; }
        public string str_h { get; set; }
        public string str_i { get; set; }
        public string str_j { get; set; }
        public double double_a { get; set; }
        public double double_b { get; set; }
        public double double_c { get; set; }
        public double double_d { get; set; }
        public double double_e { get; set; }
        public double double_f { get; set; }
        public double double_g { get; set; }
        public double double_h { get; set; }
        public double double_i { get; set; }
        public double double_j { get; set; }
    }

    public class CladLineNoModel
    {
        public int ID { get; set; }
        public string CladUniqID { get; set; }
        public string CladItemDesc { get; set; }
        public string CladLineNo { get; set; }
        public string CustPOItemNo { get; set; }
    }

    public class DetailedCladUniqModel
    {
        public string LinerCladUniqNo { get; set; }
        public double LinerMaterialLength { get; set; }
        public string LinerPackageNo { get; set; }
        public string LinerHeatNo { get; set; }
        public string LinerMaterialCert { get; set; }
        public string PipeCladUniqNo { get; set; }
        public string PipeHeatNo { get; set; }
        public string PipeOriNo { get; set; }
        public string PipeMaterialCert { get; set; }
        public double PipeMaterialLength { get; set; }
        public string PipeCustPONo { get; set; }
        public string PipeCustTagNo { get; set; }
        public string PartType { get; set; }
        public string EndP { get; set; }
        public string WOLPre { get; set; }
        public string WelderR { get; set; }
        public string WelderY { get; set; }
        public string WOLDone { get; set; }
        public int ReportStatus { get; set; }
        public int ReprocessFlag { get; set; }
        public string Remark { get; set; }
    }

    public class ClsEmailModel
    {
        public DateTime date { get; set; }
    }

    public class WireModel
    {
        public string id { get; set; }
        public string WBNo { get; set; }
        public string WireBrand { get; set; }
    }
    public partial class MachineOEE
    {
        public string rig { get; set; }
        public decimal mB { get; set; }
        public decimal pB { get; set; }
        public decimal p { get; set; }
        public decimal rP { get; set; }
        public decimal rW { get; set; }
        public decimal d { get; set; }
        public decimal aV { get; set; }
    }
}