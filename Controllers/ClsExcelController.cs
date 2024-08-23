using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using HttpPostAttribute = System.Web.Http.HttpPostAttribute;
using Microsoft.Office.Interop.Excel;
using CORSYS_API.Models;
using System.Web.Script.Serialization;
using DataTable = System.Data.DataTable;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

namespace CORSYS_API.Controllers
{

    public class ClsExcelController : ApiController
    {

        public ClsSql db = new ClsSql();

        Application excel;
        Workbook excelworkBook;
        Workbooks excelworkBooks;
        Worksheet excelSheet;
        Range excelCellrange;
        public void ExcelDispose()
        {
            if (excelworkBook != null && excel != null)
            {
                // Cleanup
                excelworkBook.Close(false);
                excel.Quit();
            }

            // Manual disposal because of COM
            if (Marshal.ReleaseComObject(excel) != 0)
            {
                while (Marshal.ReleaseComObject(excel) != 0) { }
            }
            if (Marshal.ReleaseComObject(excelworkBook) != 0)
            {
                while (Marshal.ReleaseComObject(excelworkBook) != 0) { }
            }
            if (Marshal.ReleaseComObject(excelSheet) != 0)
            {
                while (Marshal.ReleaseComObject(excelSheet) != 0) { }
            }
            if (Marshal.ReleaseComObject(excelCellrange) != 0)
            {
                while (Marshal.ReleaseComObject(excelCellrange) != 0) { }
            }
            if (Marshal.ReleaseComObject(excelworkBooks) != 0)
            {
                while (Marshal.ReleaseComObject(excelworkBooks) != 0) { }
            }

            excel = null;
            excelworkBooks = null;
            excelworkBook = null;
            excelSheet = null;
            excelCellrange = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        [HttpPostAttribute]
        public HttpResponseMessage GetListAvailableFinalLiner([FromBody] AuthModel AM)
        {
            try
            {
                db.OpenConn();
                AM = db.CheckDB(AM);
                List<string> colHeader = new List<string> { "No", "Cladatek Unique ID No", "Skid No/Package No", "Heat No", "Material Certificates", "Final Material Length", "Remark" };
                string query = String.Format(@"
                    select ROW_NUMBER() OVER ( ORDER BY a.CladUniqNo ) No, a.CladUniqNo as LinerCladUniqNo, b.PackageNo as LinerPackageNo, b.HeatNo as LinerHeatNo, b.MatCert as LinerMaterialCert, c.Length as LinerMaterialLength, b.Remark from {0}.dbo.FVIHD a 
                    left join (
	                    select distinct x.MatCert, x.PackageNo, x.HeatNo, x.CladUniqNo, x.ReportStatus,  x.Remark from {0}.dbo.CoilCutDT x where x.Dlt!=1 and x.ReportStatus = (select MAX(reportstatus) from {0}.dbo.CoilCutDT where CladUniqNo = x.CladUniqNo)
                    ) b on a.CladUniqNo = b.CladUniqNo
                    left join (
	                    select distinct x.Length, x.CladUniqNo, x.ReportStatus from {0}.dbo.LnDCHD x where x.Dlt!=1 and x.ReportStatus = (select MAX(reportstatus) from {0}.dbo.LnDCHD where CladUniqNo = x.CladUniqNo)
                    ) c on a.CladUniqNo = c.CladUniqNo
                    where a.Dlt != 1 and a.ProjectID = '{1}' and a.ItemNo = '{2}' and a.CladUniqNo not in (
                    select top 1 p.LinerCladUniqNo from {0}.dbo.PreTelAssyDT p where p.LinerCladUniqNo = a.CladUniqNo and p.Dlt!=1 order by p.ReportStatus DESC)

                    ", AM.DBName, AM.ProjectID, AM.CladLineNo);
                var Data = db.GetDataTableSQL(query);
                string targetpath = AM.DestPath;
                string name = "List Available Release Liner.xlsx";
                string filename = Path.Combine(targetpath, name);
                var a = 2;
                var b = a + 1;
                var c = b + 1;
                try
                {
                    //Get Excel using  Microsoft.Office.Interop.Excel;  
                    object misValue = System.Reflection.Missing.Value;
                    excel = new Application();
                    excel.ODBCTimeout = 0;
                    excel.Visible = false;
                    excel.ScreenUpdating = false;
                    excelworkBooks = excel.Workbooks;
                    excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
                    excelSheet = (Worksheet)excelworkBook.ActiveSheet;

                    var rowcount = Data.Rows.Count;
                    var columncount = Data.Columns.Count;
                    int colIndex = 0;
                    int rowIndex = 2;
                    foreach (var i in colHeader)
                    {
                        colIndex++;
                        excelSheet.Cells[rowIndex, colIndex] = i;
                    }
                    excelSheet.Cells[1, 1] = "Last Updated";
                    excelSheet.Cells[1, 2] = DateTime.Now;
                    var obj = db.ToObject(Data);
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + rowIndex, columncount]];
                    excelCellrange.Value = obj;

                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + rowIndex, columncount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    Borders border = excelCellrange.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
                    excelCellrange.Font.Bold = true;
                    excelCellrange.WrapText = true;
                    excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
                    excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    excel.DisplayAlerts = false;
                    excelworkBook.SaveAs(filename);
                    ExcelDispose();

                    byte[] fileBook = File.ReadAllBytes(filename);
                    MemoryStream stream = new MemoryStream();
                    string excelBase64String = Convert.ToBase64String(fileBook);
                    StreamWriter excelWriter = new StreamWriter(stream);
                    excelWriter.Write(excelBase64String);
                    excelWriter.Flush();
                    stream.Position = 0;

                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                    result.Content = new StringContent(name);
                    return result;

                }
                catch (COMException e)
                {
                    ExcelDispose();
                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                    result.Content = new StringContent(e.Message);
                    return result;
                }

            }
            catch (Exception x)
            {
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                result.Content = new StringContent(x.Message);
                return result;
            }

        }

        [HttpPostAttribute]
        public HttpResponseMessage GetPendingLiner([FromBody] AuthModel AM)
        {
            try
            {
                db.OpenConn();
                AM = db.CheckDB(AM);

                JavaScriptSerializer oSerializer = new JavaScriptSerializer();
                oSerializer.MaxJsonLength = Int32.MaxValue;
                var aaa = oSerializer.DeserializeObject(AM.json);
                var aaa1 = oSerializer.Deserialize<dynamic>(AM.json);
                DataTable Data = (DataTable)JsonConvert.DeserializeObject(AM.json, (typeof(DataTable)));
                //DataTable Data2 = (DataTable)JsonConvert.DeserializeObject(aaa, (typeof(DataTable)));
                //DataTable Data = (DataTable)JsonConvert.DeserializeObject(AM.json, (typeof(DataTable)));
                List<string> colHeader = new List<string> { "No", "Process", "Cladatek Unique ID No", "Skid No/Package No", "Heat No", "Material Certificates", "Time Delay",  "Remark" };
                List<string> colHeaderName = new List<string> { "No", "Station", "CladUniqNo", "PackageNo", "HeatNo", "MatCert", "TimeDelay",  "Remark" };

                string targetpath = AM.DestPath;
                string name = "List Pending CRA Liner.xlsx";
                string filename = Path.Combine(targetpath, name);
                try
                {
                    object misValue = System.Reflection.Missing.Value;
                    excel = new Application();
                    excel.ODBCTimeout = 0;
                    excel.Visible = false;
                    excel.ScreenUpdating = false;
                    excelworkBooks = excel.Workbooks;
                    excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
                    excelSheet = (Worksheet)excelworkBook.ActiveSheet;

                    var rowcount = Data.Rows.Count;
                    var columncount = colHeader.Count;
                    int rowIndex = 2;
                    int FirstRow = 1;
                    excelSheet.Cells[1, 1] = "Last Updated";
                    excelSheet.Cells[1, 2] = DateTime.Now;
                    var obj = db.ToObjectWithHeader(Data, colHeader, colHeaderName, true);
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
                    excelCellrange.Value = obj;
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    Borders border = excelCellrange.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
                    excelCellrange.Font.Bold = true;
                    excelCellrange.WrapText = true;
                    excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
                    excel.DisplayAlerts = false;
                    excelworkBook.SaveAs(filename);
                    ExcelDispose();

                    byte[] fileBook = File.ReadAllBytes(filename);
                    MemoryStream stream = new MemoryStream();
                    string excelBase64String = Convert.ToBase64String(fileBook);
                    StreamWriter excelWriter = new StreamWriter(stream);
                    excelWriter.Write(excelBase64String);
                    excelWriter.Flush();
                    stream.Position = 0;

                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                    result.Content = new StringContent(name);
                    return result;

                }
                catch (COMException e)
                {
                    ExcelDispose();
                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                    result.Content = new StringContent(e.Message);
                    return result;
                }

            }
            catch (Exception x)
            {
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                result.Content = new StringContent(x.Message);
                return result;
            }

        }
        [HttpPostAttribute]
        public HttpResponseMessage GetPendingPipe([FromBody] AuthModel AM)
        {
            try
            {
                db.OpenConn();
                AM = db.CheckDB(AM);

                JavaScriptSerializer oSerializer = new JavaScriptSerializer();
                oSerializer.MaxJsonLength = Int32.MaxValue;
                var aaa = oSerializer.DeserializeObject(AM.json);
                var aaa1 = oSerializer.Deserialize<dynamic>(AM.json);
                DataTable Data = (DataTable)JsonConvert.DeserializeObject(AM.json, (typeof(DataTable)));
                List<string> colHeader = new List<string> { "No", "Process", " Lined Pipe Cladatek Unique ID No", "Original Material No", "Heat No", "Material Certificates", "Time Delay", "Remark" };
                List<string> colHeaderName = new List<string> { "No", "Station", "PipeCladUniqNo", "PipeOriNo", "PipeHeatNo", "PipeMaterialCert", "TimeDelay",  "Remark" };

                string targetpath = AM.DestPath;
                string name = "List Pending Lined Pipe.xlsx";
                string filename = Path.Combine(targetpath, name);
                try
                {
                    object misValue = System.Reflection.Missing.Value;
                    excel = new Application();
                    excel.ODBCTimeout = 0;
                    excel.Visible = false;
                    excel.ScreenUpdating = false;
                    excelworkBooks = excel.Workbooks;
                    excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
                    excelSheet = (Worksheet)excelworkBook.ActiveSheet;

                    var rowcount = Data.Rows.Count;
                    var columncount = colHeader.Count;
                    int rowIndex = 2;
                    int FirstRow = 1;
                    excelSheet.Cells[1, 1] = "Last Updated";
                    excelSheet.Cells[1, 2] = DateTime.Now;
                    var obj = db.ToObjectWithHeader(Data, colHeader, colHeaderName, true);
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount+ FirstRow, columncount]];
                    excelCellrange.Value = obj;
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    Borders border = excelCellrange.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
                    excelCellrange.Font.Bold = true;
                    excelCellrange.WrapText = true;
                    excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
                    excel.DisplayAlerts = false;
                    excelworkBook.SaveAs(filename);
                    ExcelDispose();

                    byte[] fileBook = File.ReadAllBytes(filename);
                    MemoryStream stream = new MemoryStream();
                    string excelBase64String = Convert.ToBase64String(fileBook);
                    StreamWriter excelWriter = new StreamWriter(stream);
                    excelWriter.Write(excelBase64String);
                    excelWriter.Flush();
                    stream.Position = 0;

                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                    result.Content = new StringContent(name);
                    return result;

                }
                catch (COMException e)
                {
                    ExcelDispose();
                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                    result.Content = new StringContent(e.Message);
                    return result;
                }

            }
            catch (Exception x)
            {
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                result.Content = new StringContent(x.Message);
                return result;
            }

        }


        [HttpPostAttribute]
        public HttpResponseMessage GetMTSLiner([FromBody] AuthModel AM)
        {
            try
            {
                db.OpenConn();
                AM = db.CheckDB(AM);
                string query = String.Format(@"
                    select  ROW_NUMBER() OVER ( ORDER BY CladUniqNo ) No, CladUniqNo, PackageNo, HeatNo, MaterialCert,  
                    MrtsDate	,	Mrts	,	MrtsResult	,	
                    DimenCheckDate	,	DimenCheck	,	DimenCheckResult	,	
                    PmiDate	,	Pmi	,	PmiResult	,	
                    CoilCutDate	,	CoilCut	,	CoilCutResult	,	CoilLength,
                    DeCoilDcDate	,	DeCoilDc	,	DeCoilDcResult	,	
                    DeCoilUTDate	,	DeCoilUT	,	DeCoilUTResult	,	
                    CrimpingDate	,	Crimping	,	CrimpingResult	,	
                    FormingDate	,	Forming	,	FormingResult	,
                    convert(varchar,TimeStart,8)	TimeStart,convert(varchar,TimeFinish,8) TimeFinish	, WelderID	, WeldStation	,
                    WBNo	,  GBNoShielding	, GBNoPlasma	, GBNoBacking	, GBNoTrailing	, GBNoGTAW	, WpsNo	,
                    a.Remarks	,  WeldDate	,	Weld	,	WeldResult	,	
                    CalibDate	,	Calib	,	CalibResult	,	
                    BHTDate	,	BHT	,	BHTResult	,	
                    ECTDate	,	ECT	,	ECTResult	,	
                    LPTDate	,	LPT	,	LPTResult	,	
                    FnCutDate	,	FnCut	,	FnCutResult	,	FnLength,
                    AnnealDate	,	Anneal	,	AnnealResult	,	
                    BlastDate	,	Blast	,	BlastResult	,	
                    LnDCDate	,	LnDC	,	LnDCResult	,	LnDCLength,
                    AHTDate	,	AHT	,	AHTResult	,	
                    DRTDate	,	DRT	,	DRTResult	,	
                    FVIDate	,	FVI	,	FVIResult	,	
                    IRNDate	,	IRN	
                    from {0}.dbo.MtsLmpDT a
					join {0}.dbo.MtsLmp b on b.ID = ParentID
					where b.ItemNo = '{1}' ", AM.DBName, AM.CladLineNo);
                var Data = db.GetDataTableSQL(query);
                string targetpath = AM.DestPath;
                string name = "Mts Liner.xlsx";
                string filename = Path.Combine(targetpath, name);
                var row1 = 2;
                var row2 = row1 + 1;
                var row3 = row2 + 1;
                try
                {
                    object misValue = System.Reflection.Missing.Value;
                    excel = new Application();
                    excel.ODBCTimeout = 0;
                    excel.Visible = false;
                    excel.ScreenUpdating = false;
                    excelworkBooks = excel.Workbooks;
                    excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
                    excelSheet = (Worksheet)excelworkBook.ActiveSheet;

                    var rowcount = Data.Rows.Count;
                    var columncount = Data.Columns.Count;
                    int colIndex = row3 + 1;
                    List<ExlHeader> headers = new List<ExlHeader>();
                    //row 1
                    headers.Add(new ExlHeader { start_a = row1, end_a = 1, start_b = row2, end_b = 1, isMerge = true, lbl = "", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 1, start_b = row3, end_b = 1, isMerge = false, lbl = "No", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 2, start_b = row2, end_b = 5, isMerge = true, lbl = "Identification", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 6, start_b = row2, end_b = 8, isMerge = true, lbl = "Material Receival", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 9, start_b = row2, end_b = 11, isMerge = true, lbl = "10% DC Original", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 12, start_b = row2, end_b = 14, isMerge = true, lbl = "PMI Testing", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 15, start_b = row2, end_b = 18, isMerge = true, lbl = "Cutting", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 19, start_b = row2, end_b = 21, isMerge = true, lbl = "DC Coil", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 22, start_b = row2, end_b = 24, isMerge = true, lbl = "Thickness Measurement", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 25, start_b = row2, end_b = 27, isMerge = true, lbl = "Crimping", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 28, start_b = row2, end_b = 30, isMerge = true, lbl = "Forming", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 31, start_b = row1, end_b = 45, isMerge = true, lbl = "Welding Tally Sheet", rowno = "1" });
                    // row 2
                    headers.Add(new ExlHeader { start_a = row2, end_a = 31, start_b = row2, end_b = 39, isMerge = false, lbl = "", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row2, end_a = 41, start_b = row2, end_b = 45, isMerge = false, lbl = "", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row2, end_a = 36, start_b = row2, end_b = 40, isMerge = true, lbl = "Gas Consumable Batch No", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 46, start_b = row2, end_b = 48, isMerge = true, lbl = "Liner Dimensional Certification Tally Sheet", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 49, start_b = row2, end_b = 51, isMerge = true, lbl = "Visual - After Welding", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 52, start_b = row2, end_b = 54, isMerge = true, lbl = "Eddy Current Testing", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 55, start_b = row2, end_b = 57, isMerge = true, lbl = "Liquid Penetrant Testing", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 58, start_b = row2, end_b = 61, isMerge = true, lbl = "Final Cutting", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 62, start_b = row2, end_b = 64, isMerge = true, lbl = "Annealing", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 65, start_b = row2, end_b = 67, isMerge = true, lbl = "External Surface Blasting", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 68, start_b = row2, end_b = 71, isMerge = true, lbl = "Final Dimensional Inspection", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 72, start_b = row2, end_b = 74, isMerge = true, lbl = "Visual Inspection - After Heatreatment", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 75, start_b = row2, end_b = 77, isMerge = true, lbl = "Digital Radiography Testing", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 78, start_b = row2, end_b = 80, isMerge = true, lbl = "Final Visual Inspection", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 81, start_b = row2, end_b = 82, isMerge = true, lbl = "Inspection Release Note", rowno = "2" });
                    // row 3
                    headers.Add(new ExlHeader { start_a = row3, end_a = 2, start_b = row3, end_b = 2, isMerge = false, lbl = "Cladtek Unique ID No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 3, start_b = row3, end_b = 2, isMerge = false, lbl = "Skid No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 4, start_b = row3, end_b = 2, isMerge = false, lbl = "Heat No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 5, start_b = row3, end_b = 2, isMerge = false, lbl = "Certificate", rowno = "3" });
                    // looping
                    headers.Add(new ExlHeader { start_a = row3, end_a = 6, start_b = row3, end_b = 2, isMerge = false, lbl = "Date", rowno = "L" });
                    //headers.Add(new ExlHeader { start_a = row3, end_a = 7, start_b = row3, end_b = 2, isMerge = false, lbl = "Report No" , rowno = "L" });
                    //headers.Add(new ExlHeader { start_a = row3, end_a = 8, start_b = row3, end_b = 2, isMerge = false, lbl = "Result" , rowno = "L" });
                    // sampai 30
                    headers.Add(new ExlHeader { start_a = row3, end_a = 31, start_b = row3, end_b = 31, isMerge = false, lbl = "Time Start", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 32, start_b = row3, end_b = 32, isMerge = false, lbl = "Time Finish", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 33, start_b = row3, end_b = 33, isMerge = false, lbl = "Welder ID", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 34, start_b = row3, end_b = 34, isMerge = false, lbl = "Welder Station", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 35, start_b = row3, end_b = 35, isMerge = false, lbl = "Weld Wire Batch No", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 36, start_b = row3, end_b = 2, isMerge = false, lbl = "Shielding", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 37, start_b = row3, end_b = 2, isMerge = false, lbl = "Plasma", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 38, start_b = row3, end_b = 2, isMerge = false, lbl = "Backing", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 39, start_b = row3, end_b = 2, isMerge = false, lbl = "Trailing", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 40, start_b = row3, end_b = 2, isMerge = false, lbl = "Shielding GTAW", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 41, start_b = row3, end_b = 41, isMerge = false, lbl = "Wps No", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 42, start_b = row3, end_b = 45, isMerge = false, lbl = "Remarks", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 43, start_b = row3, end_b = 42, isMerge = false, lbl = "Report Date", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 44, start_b = row3, end_b = 43, isMerge = false, lbl = "Report No", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 45, start_b = row3, end_b = 44, isMerge = false, lbl = "Result", rowno = "2" });
                    // looping
                    headers.Add(new ExlHeader { start_a = row3, end_a = 46, start_b = row3, end_b = 2, isMerge = false, lbl = "Date", rowno = "L" });
                    //headers.Add(new ExlHeader { start_a = row3, end_a = 47, start_b = row3, end_b = 2, isMerge = false, lbl = "Report No" , rowno = "L" });
                    //headers.Add(new ExlHeader { start_a = row3, end_a = 48, start_b = row3, end_b = 2, isMerge = false, lbl = "Result" , rowno = "L" });
                    // sampai 82
                    // masukkan ke dalam excel
                    for (int i = 0; i < headers.Count(); i++)
                    {
                        try
                        {
                            if (headers[i].rowno == "L")
                            {
                                int maxLoop = 0;
                                if (headers[i].end_a == 6)
                                {
                                    maxLoop = headers[i + 1].end_a;
                                }
                                else
                                {
                                    maxLoop = columncount;
                                }
                                for (int j = headers[i].end_a; j < maxLoop;)
                                {
                                    int noa = j;
                                    int nob = noa + 1;
                                    int noc = nob + 1;
                                    excelSheet.Cells[headers[i].start_a, noa] = "Date";
                                    excelSheet.Cells[headers[i].start_a, nob] = "Report No";
                                    //excelSheet.Cells[headers[i].start_a, noc] = "Result";
                                    if (j < maxLoop-1)
                                    {
                                        excelSheet.Cells[headers[i].start_a, noc] = "Result";
                                    }

                                    if (noc + 1 == 18 || noc + 1 == 61 || noc + 1 == 71)
                                    {
                                        noc += 1;
                                        excelSheet.Cells[headers[i].start_a, noc] = "Length";
                                    }
                                    j = noc + 1;
                                }
                            }
                            else
                            {
                                if (headers[i].isMerge)
                                {
                                    excelCellrange = excelSheet.Range[excelSheet.Cells[headers[i].start_a, headers[i].end_a], excelSheet.Cells[headers[i].start_b, headers[i].end_b]];
                                    excelCellrange.Merge(misValue);
                                    excelCellrange.Value = headers[i].lbl;
                                }
                                else
                                {
                                    excelSheet.Cells[headers[i].start_a, headers[i].end_a] = headers[i].lbl;
                                }
                            }
                        }
                        catch (Exception x)
                        {
                            var temp = i;
                            throw x;
                        }
                    }
                    excelSheet.Cells[1, 1] = "Last Updated";
                    excelSheet.Cells[1, 2] = DateTime.Now;
                    // Data Detail
                    var obj = db.ToObject(Data);
                    excelCellrange = excelSheet.Range[excelSheet.Cells[row3 + 1, 1], excelSheet.Cells[rowcount + row3, columncount]];
                    excelCellrange.Value = obj;
                    // Format Style
                    excelCellrange = excelSheet.Range[excelSheet.Cells[row1, 1], excelSheet.Cells[rowcount + row3, columncount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    Borders border = excelCellrange.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    excelCellrange = excelSheet.Range[excelSheet.Cells[row1, 1], excelSheet.Cells[row3, columncount]];
                    excelCellrange.Font.Bold = true;
                    excelCellrange.WrapText = true;
                    excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
                    excel.DisplayAlerts = false;
                    excelworkBook.SaveAs(filename);
                    ExcelDispose();

                    byte[] fileBook = File.ReadAllBytes(filename);
                    MemoryStream stream = new MemoryStream();
                    string excelBase64String = Convert.ToBase64String(fileBook);
                    StreamWriter excelWriter = new StreamWriter(stream);
                    excelWriter.Write(excelBase64String);
                    excelWriter.Flush();
                    stream.Position = 0;

                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                    result.Content = new StringContent(name);
                    return result;

                }
                catch (COMException e)
                {
                    ExcelDispose();
                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                    result.Content = new StringContent(e.Message);
                    return result;
                }

            }
            catch (Exception x)
            {
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                result.Content = new StringContent(x.Message);
                return result;
            }

        }
        [HttpPostAttribute]
        public HttpResponseMessage GetMTSPipe([FromBody] AuthModel AM)
        {
            try
            {
                db.OpenConn();
                AM = db.CheckDB(AM);
                string query = String.Format(@"
                    select  ROW_NUMBER() OVER ( ORDER BY PipeCladUniqNo ) No, 
					a.PipeCladUniqNo, a.PipeOriNo, a.PipeHeatNo, a.PipeMaterialCert,
					a.PackageNo, a.LinerCladUniqNo, a.LinerHeatNo, a.LinerMaterialCert,
					a.MrtsPipeDate, a.MrtsPipe, a.MrtsPipeResult,
					a.MrtsPipeDCDate, a.MrtsPipeDC, a.MrtsPipeDCResult,
					a.ISBlastingDate, a.ISBlasting, a.ISBlastingResult,
					a.PreTelAssyDate, a.PreTelAssy, a.PreTelAssyResult,
					a.VLWeldDate, a.VLWeld, 'CT'+VLWelderID as VLWelderID, a.VLWeldResult,

					a.WOLEndTimeStart, a.WOLEndTimeFinish, 
					'CT'+a.WOLEndWelderR+' CT'+WOLEndWelderY as WOLEndWelderID, --a.WOLEndWelderY, 
					a.WOLEndWeldStation,
					a.WOLEndWire1HeatNo+'  '+a.WOLEndWire2HeatNo as WOLEndWireHeatNo, 
					a.WOLEndSG1BatchNo+'  '+a.WOLEndSG2BatchNo as WOLEndSGBatchNo, 
					a.WOLEndWpsNo, a.WOLEndResult, a.WOLEndStation, a.WOLEndRemarks,

					a.UTBODate, a.UTBOResult,
					a.UTAODate, a.UTAO, a.UTAOResult,
					a.FinalMachDate, a.FinalMach, a.FinalMachResult,
					a.LPTDate, a.LPT, a.LPTResult,
					a.UTLamdisDate, a.UTLamdis, a.UTLamdisResult,
					a.HydroExpanScadaDate, a.HydroExpanScada, a.HydroExpanScadaResult,
					a.PWJetCleaningDate, a.PWJetCleaningResult, a.PWJetCleaningRemarks,
					a.PMIPipeTestingDate, a.PMIPipeTesting, a.PMIPipeTestingResult,
					a.FinalDimensionDate, a.FinalDimension, a.FinalDimensionResult, a.FinalDimensionLength,
					a.FinalVisualDate, a.FinalVisual, a.FinalVisualResult,
                    a.IRNDate	, a.IRN	
                    from {0}.dbo.MtsPipeDT a
					join {0}.dbo.MtsPipeHD b on b.ID = ParentID
					where b.PipeCladLineNo = '{1}' 
                    ", AM.DBName, AM.CladLineNo);
                var Data = db.GetDataTableSQL(query);
                string targetpath = AM.DestPath;
                string name = "MTS Pipe.xlsx";
                string filename = Path.Combine(targetpath, name);
                var row1 = 2;
                var row2 = row1 + 1;
                var row3 = row2 + 1;
                try
                {
                    object misValue = System.Reflection.Missing.Value;
                    excel = new Application();
                    excel.ODBCTimeout = 0;
                    excel.Visible = false;
                    excel.ScreenUpdating = false;
                    excelworkBooks = excel.Workbooks;
                    excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
                    excelSheet = (Worksheet)excelworkBook.ActiveSheet;

                    var rowcount = Data.Rows.Count;
                    var columncount = Data.Columns.Count;
                    int colIndex = row3 + 1;
                    List<ExlHeader> headers = new List<ExlHeader>();
                    //row 1
                    headers.Add(new ExlHeader { start_a = row1, end_a = 1, start_b = row3, end_b = 1, isMerge = true, lbl = "No", rowno = "1" });
                    //headers.Add(new ExlHeader { start_a = row3, end_a = 1, start_b = row3, end_b = 1, isMerge = false, lbl = "No", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 2, start_b = row2, end_b = 9, isMerge = true, lbl = "Identification", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 10, start_b = row2, end_b = 12, isMerge = true, lbl = "Material Receival - CS Pipes", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 13, start_b = row2, end_b = 15, isMerge = true, lbl = "10% Dimensional Check - CS Pipes", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 16, start_b = row2, end_b = 18, isMerge = true, lbl = "Internal Blasting of CS Pipes", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 19, start_b = row2, end_b = 21, isMerge = true, lbl = "Pre Assembly Check and Telescopic Assembly", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 22, start_b = row2, end_b = 25, isMerge = true, lbl = "Vacuum and Laser Transition Weld", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 26, start_b = row2, end_b = 35, isMerge = true, lbl = "Welding", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 36, start_b = row1, end_b = 40, isMerge = true, lbl = "UT Thickness Gauging", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 41, start_b = row2, end_b = 43, isMerge = true, lbl = "Final Machining Ends", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 44, start_b = row2, end_b = 46, isMerge = true, lbl = "Liquid Penetrant Testing - Machined Ends", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 47, start_b = row2, end_b = 49, isMerge = true, lbl = "UT LaminationDisbonding", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 50, start_b = row2, end_b = 52, isMerge = true, lbl = "High Pressure Expansion - Hydro Static Test ", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 53, start_b = row2, end_b = 55, isMerge = true, lbl = "Surface Treatment & Cleaning", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 56, start_b = row2, end_b = 58, isMerge = true, lbl = "PMI Testing", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 59, start_b = row2, end_b = 62, isMerge = true, lbl = "Final Dimensional Inspection", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 63, start_b = row2, end_b = 65, isMerge = true, lbl = "Final Visual Inspection", rowno = "1" });
                    headers.Add(new ExlHeader { start_a = row1, end_a = 66, start_b = row2, end_b = 67, isMerge = true, lbl = "Inspection Release Note - IRN", rowno = "1" });
                    // row 2
                    headers.Add(new ExlHeader { start_a = row2, end_a = 36, start_b = row2, end_b = 37, isMerge = true, lbl = "Prior Weld Overlay", rowno = "2" });
                    headers.Add(new ExlHeader { start_a = row2, end_a = 37, start_b = row2, end_b = 39, isMerge = true, lbl = "After Machined", rowno = "2" });
                    // row 3
                    headers.Add(new ExlHeader { start_a = row3, end_a = 2, start_b = row3, end_b = 2, isMerge = false, lbl = "Lined Pipe Cladtek Unique ID No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 3, start_b = row3, end_b = 3, isMerge = false, lbl = "Original Material No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 4, start_b = row3, end_b = 4, isMerge = false, lbl = "Heat No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 5, start_b = row3, end_b = 5, isMerge = false, lbl = "Material Certificate No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 6, start_b = row3, end_b = 6, isMerge = false, lbl = "Skid No / Packages No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 7, start_b = row3, end_b = 7, isMerge = false, lbl = "CRA Liner Cladtek Unique ID No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 8, start_b = row3, end_b = 8, isMerge = false, lbl = "Heat No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 9, start_b = row3, end_b = 9, isMerge = false, lbl = "Material Certificate No", rowno = "3" });
                    // looping
                    headers.Add(new ExlHeader { start_a = row3, end_a = 10, start_b = row3, end_b = 10, isMerge = false, lbl = "Date", rowno = "L" });
                    // sampai 24
                    headers.Add(new ExlHeader { start_a = row3, end_a = 22, start_b = row3, end_b = 22, isMerge = false, lbl = "Date", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 23, start_b = row3, end_b = 23, isMerge = false, lbl = "Report No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 24, start_b = row3, end_b = 24, isMerge = false, lbl = "Welder ID", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 25, start_b = row3, end_b = 25, isMerge = false, lbl = "Result", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 26, start_b = row3, end_b = 26, isMerge = false, lbl = "Date Start", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 27, start_b = row3, end_b = 27, isMerge = false, lbl = "Date Finish", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 28, start_b = row3, end_b = 28, isMerge = false, lbl = "Welder ID", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 29, start_b = row3, end_b = 29, isMerge = false, lbl = "Weld Station", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 30, start_b = row3, end_b = 30, isMerge = false, lbl = "Wire Heat No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 31, start_b = row3, end_b = 31, isMerge = false, lbl = "Gas Consumable Batch No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 32, start_b = row3, end_b = 32, isMerge = false, lbl = "WPS No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 33, start_b = row3, end_b = 33, isMerge = false, lbl = "Result", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 34, start_b = row3, end_b = 34, isMerge = false, lbl = "Station", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 35, start_b = row3, end_b = 35, isMerge = false, lbl = "Remarks", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 36, start_b = row3, end_b = 36, isMerge = false, lbl = "Date", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 37, start_b = row3, end_b = 37, isMerge = false, lbl = "Result", rowno = "3" });
                    // looping
                    headers.Add(new ExlHeader { start_a = row3, end_a = 38, start_b = row3, end_b = 38, isMerge = false, lbl = "Date", rowno = "L" });
                    // sampai 51
                    headers.Add(new ExlHeader { start_a = row3, end_a = 53, start_b = row3, end_b = 53, isMerge = false, lbl = "Date", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 54, start_b = row3, end_b = 54, isMerge = false, lbl = "Result", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 55, start_b = row3, end_b = 55, isMerge = false, lbl = "Remarks", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 56, start_b = row3, end_b = 56, isMerge = false, lbl = "Date", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 57, start_b = row3, end_b = 57, isMerge = false, lbl = "Report No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 58, start_b = row3, end_b = 58, isMerge = false, lbl = "Result", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 59, start_b = row3, end_b = 59, isMerge = false, lbl = "Date", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 60, start_b = row3, end_b = 60, isMerge = false, lbl = "Report No", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 61, start_b = row3, end_b = 61, isMerge = false, lbl = "Result", rowno = "3" });
                    headers.Add(new ExlHeader { start_a = row3, end_a = 62, start_b = row3, end_b = 62, isMerge = false, lbl = "Length", rowno = "3" });
                    //looping
                    headers.Add(new ExlHeader { start_a = row3, end_a = 63, start_b = row3, end_b = 63, isMerge = false, lbl = "Date", rowno = "L" });
                    // sampai 65

                    // masukkan ke dalam excel
                    for (int i = 0; i < headers.Count(); i++)
                    {
                        try
                        {
                            if (headers[i].rowno == "L")
                            {
                                int maxLoop = 0;
                                //maxLoop = columncount;
                                if (headers[i].end_a == 10 || headers[i].end_a == 37)
                                {
                                    maxLoop = headers[i + 1].end_a;
                                }
                                else
                                {
                                    maxLoop = columncount;
                                }
                                for (int j = headers[i].end_a; j < maxLoop;)
                                {
                                    int noa = j;
                                    int nob = noa + 1;
                                    int noc = nob + 1;
                                    excelSheet.Cells[headers[i].start_a, noa] = "Date";
                                    excelSheet.Cells[headers[i].start_a, nob] = "Report No";
                                    if (j < maxLoop - 1)
                                    {
                                        excelSheet.Cells[headers[i].start_a, noc] = "Result";
                                    }
                                    //if (noc + 1 == 53)
                                    //{
                                    //    excelSheet.Cells[headers[i].start_a, nob] = "Result";
                                    //    noc += 1;
                                    //    excelSheet.Cells[headers[i].start_a, noc] = "Remarks";
                                    //}
                                    j = noc + 1;
                                }
                            }
                            else
                            {
                                if (headers[i].isMerge)
                                {
                                    excelCellrange = excelSheet.Range[excelSheet.Cells[headers[i].start_a, headers[i].end_a], excelSheet.Cells[headers[i].start_b, headers[i].end_b]];
                                    excelCellrange.Merge(misValue);
                                    excelCellrange.Value = headers[i].lbl;
                                }
                                else
                                {
                                    excelSheet.Cells[headers[i].start_a, headers[i].end_a] = headers[i].lbl;
                                }
                            }
                        }
                        catch (Exception x)
                        {
                            var temp = i;
                            throw x;
                        }
                    }
                    excelSheet.Cells[1, 2] = "Last Updated";
                    excelSheet.Cells[1, 3] = DateTime.Now;
                    // Data Detail
                    var obj = db.ToObject(Data);
                    excelCellrange = excelSheet.Range[excelSheet.Cells[row3 + 1, 1], excelSheet.Cells[rowcount + row3, columncount]];
                    excelCellrange.Value = obj;
                    // Format Style
                    excelCellrange = excelSheet.Range[excelSheet.Cells[row1, 1], excelSheet.Cells[rowcount + row3, columncount]];
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    Borders border = excelCellrange.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    excelCellrange = excelSheet.Range[excelSheet.Cells[row1, 1], excelSheet.Cells[row3, columncount]];
                    excelCellrange.Font.Bold = true;
                    excelCellrange.WrapText = true;
                    excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
                    excel.DisplayAlerts = false;
                    excelworkBook.SaveAs(filename);
                    ExcelDispose();

                    byte[] fileBook = File.ReadAllBytes(filename);
                    MemoryStream stream = new MemoryStream();
                    string excelBase64String = Convert.ToBase64String(fileBook);
                    StreamWriter excelWriter = new StreamWriter(stream);
                    excelWriter.Write(excelBase64String);
                    excelWriter.Flush();
                    stream.Position = 0;

                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                    result.Content = new StringContent(name);
                    return result;

                }
                catch (COMException e)
                {
                    ExcelDispose();
                    HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                    result.Content = new StringContent(e.Message);
                    return result;
                }

            }
            catch (Exception x)
            {
                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
                result.Content = new StringContent(x.Message);
                return result;
            }

        }
    }

}
