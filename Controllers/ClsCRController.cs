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

namespace CORSYS_API.Controllers
{
    public class ClsCRController : ApiController
    {
        [HttpPostAttribute]
        public HttpResponseMessage PostCRWithSub([FromBody] CRModel text)
        {
            try
            {
                ReportDocument rd = new ReportDocument();
                var pathCR = Path.Combine(text.DestPath, text.ReportID + ".rpt");
                rd.Load(pathCR);
                var no = 0;
                foreach (DataTable table in text.DS.Tables)
                {
                    rd.Database.Tables[table.TableName].SetDataSource(table);
                    text.DT[no].TableName = table.TableName;
                    no++;
                }
                if (text.SubDS!=null)
                {
                    foreach (var i in text.SubDS)
                    {
                        rd.Subreports[i.Sub].SetDataSource(text.DT[i.DataSet]);
                        //rd.Subreports[1].SetDataSource(i);
                    }
                }
                Stream stream = rd.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                rd.Close();
                rd.Dispose();
                stream.Seek(0, SeekOrigin.Begin);
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StreamContent(stream); 
                result.Content.Headers.ContentLength = stream.Length;
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                result.Content.Headers.ContentDisposition.FileName = text.ReportID;
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                return result;
            }
            catch (Exception x)
            {
                return ClsPdf.ErrPdf(x,text.ReportID,text.DestPath);
            }

        }
        [HttpPostAttribute]
        public HttpResponseMessage PostCR([FromBody] CRModel text)
        {
            try
            {
                ReportDocument rd = new ReportDocument();
                var pathCR = Path.Combine(text.DestPath, text.ReportID + ".rpt");
                rd.Load(pathCR);
                foreach (DataTable table in text.DS.Tables)
                {
                    rd.Database.Tables[table.TableName].SetDataSource(table);
                }
                Stream stream = rd.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                rd.Close();
                rd.Dispose();
                stream.Seek(0, SeekOrigin.Begin);
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StreamContent(stream); 
                result.Content.Headers.ContentLength = stream.Length;
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                result.Content.Headers.ContentDisposition.FileName = text.ReportID;
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                return result;
            }
            catch (Exception x)
            {
                return ClsPdf.ErrPdf(x,text.ReportID,text.DestPath);
            }

        }
        [HttpPostAttribute]
        public HttpResponseMessage PostCRToPDF([FromBody] CRModel text)
        {
            try
            {
                ReportDocument rd = new ReportDocument();
                var pathPdf = Path.Combine(text.DestPath, "PDF");
                System.IO.Directory.CreateDirectory(pathPdf);
                var pathCR = Path.Combine(text.DestPath, text.ReportID + ".rpt");
                rd.Load(pathCR);
                List<string> fileList = new List<string>();
                // alamat pdf report duluan baru di masukkan alamat graph nya
                var Finalpath = Path.Combine(text.DestPath, "PDF", text.ReportID + "_" + text.ID.ToString() + ".pdf");
                fileList.Add(Finalpath);

                ReportDocument rdgraph = new ReportDocument();
                var pathgraph = Path.Combine(text.DestPath, "Graph" + text.ReportID + ".rpt");
                rdgraph.Load(pathgraph);

                foreach (DataTable table in text.DS.Tables)
                {
                    if (table.TableName.ToLower() == "graph")
                    {
                        rdgraph.Database.Tables[table.TableName].SetDataSource(table);
                    }
                    else if (table.TableName.ToLower() == "dt")
                    {
                        rd.Database.Tables[table.TableName].SetDataSource(table);
                        rdgraph.Database.Tables[table.TableName].SetDataSource(table);
                    }
                    else
                    {
                        rd.Database.Tables[table.TableName].SetDataSource(table);
                    }
                }

                var rptgraph = rdgraph.FileName.Replace("Reports", "Reports/PDF");
                rdgraph.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, rptgraph.Replace(".rpt", "_" + text.ID.ToString() + ".pdf").Replace("rassdk://", ""));
                rdgraph.Close();
                rdgraph.Dispose();
                fileList.Add(Path.Combine(text.DestPath, "PDF", "Graph" + text.ReportID + "_" + text.ID.ToString() + ".pdf"));

                var rpt = rd.FileName.Replace("Reports", "Reports/PDF");
                rd.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, rpt.Replace(".rpt", "_" + text.ID.ToString() + ".pdf").Replace("rassdk://", ""));
                rd.Close();
                rd.Dispose();

                string strPath = Path.Combine(text.DestPath, "PDF", text.ReportID + "_" + text.ID.ToString() + "_Combine.pdf");
                ClsPdf.MergePDFs(fileList, strPath);
                FileStream stream = new FileStream(strPath, FileMode.Open);

                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StreamContent(stream);
                result.Content.Headers.ContentLength = stream.Length;
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                result.Content.Headers.ContentDisposition.FileName = text.ReportID;
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                //HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                //result.Content = new StringContent(Finalpath);
                return result;
            }
            catch (Exception x)
            {
                return ClsPdf.ErrPdf(x, text.ReportID, text.DestPath);
            }
        }
        [HttpPostAttribute]
        public HttpResponseMessage PostPDF([FromBody] CRModel text)
        {
            try
            {
                ReportDocument rd = new ReportDocument();
                var pathPdf = Path.Combine(text.DestPath, "PDF");
                System.IO.Directory.CreateDirectory(pathPdf);
                var pathCR = Path.Combine(text.DestPath, text.ReportID + ".rpt");
                rd.Load(pathCR);
                foreach (DataTable table in text.DS.Tables)
                {
                    rd.Database.Tables[table.TableName].SetDataSource(table);
                }

                var rpt = rd.FileName.Replace("Reports", "Reports/PDF");
                rd.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, rpt.Replace(".rpt", "_" + text.ID.ToString() + ".pdf").Replace("rassdk://", ""));
                rd.Close();
                rd.Dispose();
                var Finalpath  = Path.Combine(text.DestPath, "PDF", text.ReportID + "_" + text.ID.ToString() + ".pdf");
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StringContent(Finalpath); 
                return result;
            }
            catch (Exception x)
            {
                return ClsPdf.ErrPdf(x,text.ReportID,text.DestPath);
            }

        }
        [HttpPostAttribute]
        public HttpResponseMessage PostHTML([FromBody] CRModel text)
        {
            try
            {
                ReportDocument rd = new ReportDocument();
                var pathPdf = Path.Combine(text.DestPath, "HTML");
                System.IO.Directory.CreateDirectory(pathPdf);
                var pathCR = Path.Combine(text.DestPath, text.ReportID + ".rpt");
                rd.Load(pathCR);
                foreach (DataTable table in text.DS.Tables)
                {
                    rd.Database.Tables[table.TableName].SetDataSource(table);
                }
                rd.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.HTML40, rd.FileName.Replace("Reports", "Reports/HTML").Replace(".rpt", ".html").Replace("rassdk://", ""));
                String str = System.IO.File.ReadAllText(pathCR.Replace(".rpt", ".html"));
                var rpt = rd.FileName;
                rd.Close();
                rd.Dispose();
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StringContent(str); 
                return result;
            }
            catch (Exception x)
            {
                return ClsPdf.ErrPdf(x,text.ReportID,text.DestPath);
            }

        }

        public ClsSql db = new ClsSql();

        [HttpPostAttribute]
        public HttpResponseMessage GetWelder([FromBody] CRModel text)
        {
            try
            {
                db.OpenConn();
                string query = "select * from matahari.dbo.WelderRegister";
                var Data = db.GetDataTableSQL(query);
                var stringContent = new StringContent(DataTableToJSON(Data).ToString(), UnicodeEncoding.UTF8, "application/json");
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = stringContent; 
                return result;
            }
            catch (Exception x)
            {
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                return result;
            }

        }
        public string toJSON(DataTable dt)
        {
            JavaScriptSerializer oSerializer = new JavaScriptSerializer();
            oSerializer.MaxJsonLength = Int32.MaxValue;
            string result = oSerializer.Serialize(dt);
            return result;
        }
        public string toJSON(DataTable[] dt)
        {
            JavaScriptSerializer oSerializer = new JavaScriptSerializer();
            oSerializer.MaxJsonLength = Int32.MaxValue;
            string result = oSerializer.Serialize(dt);
            return result;
        }
        public static object DataTableToJSON(DataTable table)
        {
            var lst = new List<Dictionary<string, object>>();
            foreach (DataRow row in table.Rows)
            {
                var dict = new Dictionary<string, object>();
                foreach (DataColumn col in table.Columns)
                {
                    dict[col.ColumnName] = (Convert.ToString(row[col]));
                }
                lst.Add(dict);
            }
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Serialize(lst);
        }
    }
}
