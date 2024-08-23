using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Mvc;
using CrystalDecisions.CrystalReports.Engine;
using HttpPostAttribute = System.Web.Http.HttpPostAttribute;
using CrystalDecisions.Shared;
using CORSYS_API.Models;
using iTextSharp.text.pdf;
using System.Web.Razor.Text;
using System.Web.Script.Serialization;
using System.Text;
using System.Drawing;
using System.Web.Helpers;

namespace CORSYS_API.Controllers
{
    public class ClsMLPController : ApiController
    {
        public ClsSql db = new ClsSql();

        [HttpPostAttribute]
        public HttpResponseMessage GetPendingLiner([FromBody] CRModel text)
        {
            try
            {
                string query = "select * from matahari.dbo.WelderRegister --where dlt=1";
                var Data = db.GetDataTableSQL(query);
                text.DS.Tables.Add(Data);
                query = "select * from matahari.dbo.WelderRegister --where dlt=1";
                Data = db.GetDataTableSQL(query);
                text.DS.Tables.Add(Data);

                return ClsCR.TOCR(text);
            }
            catch (Exception x)
            {
                return ClsPdf.ErrPdf(x, text.ReportID, text.DestPath);
            }

        }
    }
}
