using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using CrystalDecisions.CrystalReports.Engine;
using System.Data;
using adminlte.Modul;

namespace CORSYS_API.Models
{
    public static partial class ClsPdf
    {
        public static HttpResponseMessage ErrPdf(Exception err, string reportID, string path)
        {
            iTextSharp.text.Document document = new iTextSharp.text.Document();
            MemoryStream stream = new MemoryStream();
            try
            {
                PdfWriter pdfWriter = PdfWriter.GetInstance(document, stream);
                pdfWriter.CloseStream = false;
                document.Open();
                document.Add(new iTextSharp.text.Paragraph("Report on " + reportID));
                document.Add(new iTextSharp.text.Paragraph("Error Occured : " + err));
            }
            catch (iTextSharp.text.DocumentException de)
            {
                Console.Error.WriteLine(de.Message);
            }
            catch (IOException ioe)
            {
                Console.Error.WriteLine(ioe.Message);
            }

            document.Close();
            stream.Flush(); //Always catches me out
            stream.Position = 0; 
            stream.Seek(0, SeekOrigin.Begin);
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentLength = stream.Length;
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = reportID;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            return result;
        }
        public static bool MergePDFs(IEnumerable<string> fileNames, string targetPdf)
        {
            bool merged = true;
            using (FileStream stream = new FileStream(targetPdf, FileMode.Create))
            {
                Document document = new Document();
                PdfCopy pdf = new PdfCopy(document, stream);
                PdfReader reader = null;
                try
                {
                    document.Open();
                    foreach (string file in fileNames)
                    {
                        reader = new PdfReader(file);
                        pdf.AddDocument(reader);
                        reader.Close();
                    }
                }
                catch (Exception)
                {
                    merged = false;
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
                finally
                {
                    if (document != null)
                    {
                        document.Close();
                    }
                }
            }
            return merged;
        }
    }

    public static partial class ClsCR
    {
        public static HttpResponseMessage TOCR(CRModel text)
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
                return ClsPdf.ErrPdf(x, text.ReportID, text.DestPath);
            }

        }
        public static string CheckUserNameCOR(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetMethodUserValidation(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string CheckAllUserNameCOR(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetMethodUserValidation(data);
                var stream = response.ReadAsStringAsync().Result;
                if (stream == "User Not Found")
                {
                    response = api.GetMethodScadaUserValidation(data);
                    stream = response.ReadAsStringAsync().Result;
                }
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetProject(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetProject(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
     
        public static string GetCladLineNo(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetCladLineNo(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetItemForWOLScada(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetItemForWOLScada(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetItemForMachiningScada(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetItemForMachiningScada(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetItemForPAWIIScada(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetItemForPAWIIScada(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetWelderStamp(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetWelderStamp(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetWPS(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetWPS(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetWPSPAWII(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetWPSPAWII(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string PostNewReportWOL(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.PostNewReportWOL(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string PostNewReportPAW(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.PostNewReportPAW(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string PostReportWOLFitting(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.PostReportWOLFitting(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        public static string GetWire(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetWire(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetGas(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetGas(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        public static string GetPipeDetail(string data)
        {
            try
            {
                ClsAPI api = new ClsAPI();
                var response = api.GetPipeDetail(data);
                var stream = response.ReadAsStringAsync().Result;
                return stream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}