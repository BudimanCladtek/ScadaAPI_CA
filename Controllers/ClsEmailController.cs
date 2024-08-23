using CORSYS_API.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using HttpPostAttribute = System.Web.Http.HttpPostAttribute;
using HttpGetAttribute = System.Web.Http.HttpGetAttribute;
using System.Web.Script.Serialization;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;
using MimeKit;
using Microsoft.Office.Interop.Excel;

namespace SCADA_API.Controllers
{
    public class ClsEmailController : ApiController
    {
        public ClsScadaSql db = new ClsScadaSql();

		[HttpPostAttribute]
		public HttpResponseMessage EndClad([FromBody] CRModel text)
		{
			try
			{
				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				string fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));

				GenerateEndCladExcel(filename, fileWeeklySummary);
				//				db.OpenConn("SCADA SHV");

				try
                {

                    // Set email subject
                    String subject = "  Target VS Actual and Daily Rig Report (End Clad) #";
                    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                        subject = "  Target VS Actual and Daily and Weekly Report (End Clad) #";
                    // Set email body
                    string emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily Rig and Target VS Actual (End Clad) \n\n\nThanks & Regards,\n\n\n\nProduction Automation";
                    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                        emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily and Weekly Target VS Actual (End CLad) \n\n\nThanks & Regards,\n\n\n\nProduction Automation";
                    //Attaching File to Mail  
                    var builder = new BodyBuilder();
                    builder.Attachments.Add(filename);

                    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                        builder.Attachments.Add(fileWeeklySummary);

                    builder.HtmlBody = emailBody;
                    var message = new MimeMessage()
                    {
                        Body = builder.ToMessageBody(),
                        Subject = subject

                    };
					//message.To.Add(new MailboxAddress("tomoyuki.ueno@cladtek.com", "tomoyuki.ueno@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("aang.junaidi@cladtek.com", "aang.junaidi@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("iswantika.putra@cladtek.com", "iswantika.putra@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("dayanandan@cladtek.com", "dayanandan@cladtek.com"));
					//message.Bcc.Add(new MailboxAddress("muhammad.shofiq@cladtek.com", "muhammad.shofiq@cladtek.com"));
					//message.Bcc.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));

					//message.To.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));
					//					message.To.Add(new MailboxAddress("muhammad.shofiq@cladtek.com", "muhammad.shofiq@cladtek.com"));
					//					message.Cc.Add(new MailboxAddress("aang.junaidi@cladtek.com", "aang.junaidi@cladtek.com"));
					//					message.Cc.Add(new MailboxAddress("iswantika.putra@cladtek.com", "iswantika.putra@cladtek.com"));
					//					message.Bcc.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));
					//message.To.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));

					String recepientType;
					String recepientName;
					String recepientAddress;
					var recepient = db.GetDataTableSQL("SCADA SHV", "Select * from tblEmailDaily where category='endclad' and deleted=0");
					foreach (var data in recepient.Rows)
					{
						recepientType = ((DataRow)data).ItemArray[1].ToString();
						recepientName = ((DataRow)data).ItemArray[2].ToString();
						recepientAddress = ((DataRow)data).ItemArray[3].ToString();
						if (recepientType.ToLower() == "to")
							message.To.Add(new MailboxAddress(recepientAddress, recepientAddress));
						if (recepientType.ToLower() == "cc")
							message.Cc.Add(new MailboxAddress(recepientAddress, recepientAddress));
						if (recepientType.ToLower() == "bcc")
							message.Bcc.Add(new MailboxAddress(recepientAddress, recepientAddress));
					}
					
					message.From.Add(new MailboxAddress("Scada System", "noreply.cor@cladtek.com"));
                    using (var emailClient = new MailKit.Net.Smtp.SmtpClient())
                    {

                        emailClient.Connect("smtp.gmail.com", 587, false);
                        emailClient.Authenticate("noreply.cor@cladtek.com", "tetanggaberisik!");
                        emailClient.Send(message);
                        emailClient.Disconnect(true);
                    }

                }

                catch (Exception ex)
                {
                    var stringContent1 = new StringContent(ex.Message.ToString(), UnicodeEncoding.UTF8, "application/json");
                    HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
                    result1.Content = stringContent1;
                    return result1;
                    //					return new HttpResponseMessage(HttpStatusCode.BadRequest);
                }

				var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
				var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
				result.Content = stringContent;
				return result;
			}
			catch (Exception ex)
			{
				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				var stringContent1 = new StringContent(String.Format(@"{0}. {1}", filename, ex.Message.ToString()), UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
				result1.Content = stringContent1;
				return result1;
//				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
//				return result;
			}

		}
		private void GenerateEndCladExcel(string filename, string fileWeeklySummary)
        {
			var ds = new DataSet();

			//string query = String.Format(@"Select 
			//			b.machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
			//			ISNULL(c.total_time/60/60,0) as Actual_Arc_Time,
			//			DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) as Actual_Time, 
			//			Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//			else Convert(Decimal(24,2),Round(
			//				100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
			//			, 2)) 
			//			end as ArcOnTime, 
			//			Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//			else Convert(Decimal(24,2),Round(
			//				100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
			//			, 2)) 
			//			end as ArcEff, ActualfinishDate
			//		from
			//		(
			//			Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
			//			Convert(decimal(24,2),Target_Hours) as Target_Total, DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
			//			b.finishdate as ActualfinishDate
			//			from TargetArcOn a
			//			outer apply 
			//			(
			//					Select * from sfnPipeOnMachineSatus(a.Machine, a.pipeno, '{0} 08:00') b
			//			) b
			//		) a
			//		inner join tblMachine b on a.Machine_no=b.Machine_no and b.classification='endclad'
			//		left outer join sfnGetWeldingProcessByDate('{0} 08:00') c on b.Machine_no=c.rig_no and a.Id_pipe=c.id_pipe
			//		where a.ActualfinishDate is null and a.startTarget<='{0} 08:00'
			//		order by b.machine_no
			//		", DateTime.Now.ToString("yyyy-MM-dd"));
			string query = String.Format(@"Select 
				a.rig_no as machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
				ISNULL(c.total_time/60,0) as Actual_Arc_Time,
				Convert(Decimal(24,2), Convert(Decimal(24,6),DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60) as Actual_Time, 
				Case DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
				else Convert(Decimal(24,2),Round(
					(100 * Convert(Decimal(24,2), ISNULL(c.total_time,0))/
						Convert(Decimal(24,2), Convert(Decimal(24,6), DateDiff(second, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60)
					)
				, 2)) 
				end as ArcOnTime
			from
			(
				Select a.id_pipe, startdatetime as startTarget, 
				MAX(Convert(decimal(24,2),Target_Hours)) as Target_Total, 
				MAX(DateAdd(hour, Convert(decimal(24,2),Target_Hours), startdatetime)) as finishTarget,
				MAX(c.finishdate) as ActualfinishDate,
				MAX(a.rig_no) as rig_no
				from (
					Select a.id_pipe, min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time))) as startdatetime, TRIM(master.dbo.Group_concat(distinct concat(' ', a.rig_no))) as rig_no
					from TblArcON a
					inner join tblMachine b on a.rig_no=b.Machine_no 
					and b.classification='endclad'
					where a.status not in('Breakdown')
					and a.rig_no not in('SHV50')
					and len(a.id_pipe)>1 and id_pipe not like '%maintenance%'
					group by a.id_pipe
					having min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time)))>DateAdd(month, -1, getDate())
				) a
				left outer join (
					Select a.PipeNo, MIN(a.Actual_Start) as Actual_Start, MAX(Target_Hours) as Target_Hours, MAX(DATEADD(hour, COnvert(decimal(24,2),Target_Hours), a.Actual_Start)) as Target_Finish
					from TargetArcOn a
					group by a.PipeNo
				) b on a.id_pipe=b.PipeNo
				outer apply 
				(
						Select * from sfnPipeStatus(a.id_pipe, '{0} 08:00') b
				) c
				group by a.id_pipe, startdatetime
				Having MAX(c.finishdate) is null
			) a
			left outer join 
			(
				Select c.id_pipe, Sum(c.total_time) as total_time 
				from sfnGetWeldingProcessByDate('{0} 08:00') c
				group by c.id_pipe
			) c on a.Id_pipe=c.id_pipe
			where a.ActualfinishDate is null and a.startTarget<='{0} 08:00'
			order by a.startTarget
			", DateTime.Now.AddDays(-25).ToString("yyyy-MM-dd"));
			//var Summary = db.GetDataTableSQL("SCADA SHV", query, 1200);
            //var Summary = db.GetDataTableSQL("SCADA_SHV_22Jul", query);
            query = String.Format(@"Declare @datDate DateTime='{0}'
						Select a.*,
							Case when Machine='Average' or ds=0 then '-' else dbo.sfnGetWelderByMachineShift(@datDate, Machine, 'd')end as welder_ds,
							Case when Machine='Average' or ns=0 then '-' else dbo.sfnGetWelderByMachineShift(@datDate, Machine, 'n') end as welder_ns
						from(
						Select rig_no as Machine, concat(convert(varchar(max), dsp), ' %') as dsp, concat(convert(varchar(max), nsp), ' %') as nsp, 
							convert(decimal(24,2),convert(decimal(24,6),ds)) as ds, convert(decimal(24,2),convert(decimal(24,6),ns)) as ns
						from dbo.sfnGetWeldingProcessByMachine(@datDate) a
						inner join TblMachine b on a.rig_no=b.machine_no and b.classification='endclad'
						where CharIndex(' ',a.rig_no)>0
						union all
/*						Select 'Average' as rig_no, concat(convert(varchar(max), Convert(Decimal(24,2),Round(Avg(dsp),2))), ' %') as dsp, 
							concat(convert(varchar(max), Convert(Decimal(24,2),Round(Avg(nsp),2))), ' %') as nsp, 
							convert(Decimal(24,2),Avg(convert(Decimal(24,12),ds)),2) as ds, convert(Decimal(24,2),Avg(convert(Decimal(24,12),ns)),2) as ns*/
						Select 'Average' as rig_no, 
							concat(convert(varchar(max), Case Sum(Case when dsp>0 then 1 else 0 end) when 0 then 0 else Convert(Decimal(24,2),Round(SUM(dsp)/Sum(Case when dsp>0 then 1 else 0 end),2)) end), ' %')as dsp, 
							concat(convert(varchar(max), Case Sum(Case when nsp>0 then 1 else 0 end) when 0 then 0 else Convert(Decimal(24,2),Round(Sum(nsp)/Sum(Case when nsp>0 then 1 else 0 end),2)) end), ' %') as nsp, 
							Case Sum(Case when ds>0 then 1 else 0 end) when 0 then 0 else convert(Decimal(24,2),Sum(convert(Decimal(24,12),ds))/Sum(Case when ds>0 then 1 else 0 end),2) end as ds, 
							Case Sum(Case when ns>0 then 1 else 0 end) when 0 then 0 else convert(Decimal(24,2),Sum(convert(Decimal(24,12),ns))/Sum(Case when ns>0 then 1 else 0 end),2) end as ns
						from dbo.sfnGetWeldingProcessByMachine(@datDate) a
						inner join TblMachine b on a.rig_no=b.machine_no and b.classification='endclad'
						where CharIndex(' ',a.rig_no)>0
						) a
						order by Replace(Left(MAchine,1),'A','Z'), ISNULL(Try_Convert(bigint,SUBSTRING(Machine,5,len(Machine)-4)),100)
				", DateTime.Now.AddDays(-26).ToString("yyyy-MM-dd"));

            //var SummaryMachine = db.GetDataTableSQL("SCADA SHV", query, 1200);
			//var SummaryMachine = db.GetDataTableSQL("SCADA_SHV_22Jul", query);
			//        query = String.Format(@"Select 
			//	b.machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, a.ActualfinishDate,
			//	ISNULL(c.total_time/60,0) as Actual_Arc_Time,
			//	DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) as Actual_Time, 
			//	Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//	else Convert(Decimal(24,2),Round(
			//		100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
			//	, 2)) 
			//	end as ArcOnTime, 
			//	Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//	else Convert(Decimal(24,2),Round(
			//		Target_Total/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))*100
			//	, 2)) 
			//	end as ArcEff
			//from
			//(
			//	Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
			//	Convert(decimal(24,2),Target_Hours) as Target_Total, DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
			//	b.finishdate as ActualfinishDate
			//	from TargetArcOn a
			//	outer apply 
			//	(
			//			Select * from sfnPipeOnMachineSatus(a.Machine, a.pipeno, '{0} 08:00') b
			//	) b
			//) a
			//inner join tblMachine b on a.Machine_no=b.Machine_no and b.classification='endclad'
			//outer apply
			//(
			//	Select * from sfnGetWeldingProcessByDate(ISNULL(a.ActualfinishDate, '{0} 08:00')) c
			//	where b.Machine_no=c.rig_no and a.Id_pipe=c.id_pipe
			//) c 
			//where a.ActualfinishDate is not null and DATEDIFF(hour, a.ActualfinishDate, '{0} 08:00') between 0 and 24
			//order by LEFT(b.machine_no,1), Try_Convert(bigint,SUBSTRING(b.machine_no,5,len(b.machine_no)-4))
			//", DateTime.Now.ToString("yyyy-MM-dd"));
			query = String.Format(@"Select 
						a.rig_no as machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
						ISNULL(c.total_time/60,0) as Actual_Arc_Time,
						Convert(Decimal(24,2), Convert(Decimal(24,6),DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60) as Actual_Time, 
						Case DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
						else Convert(Decimal(24,2),Round(
							(100 * Convert(Decimal(24,2), ISNULL(c.total_time,0))/
								Convert(Decimal(24,2), Convert(Decimal(24,6), DateDiff(second, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60)
							)
						, 2)) 
						end as ArcOnTime, 
						Case ISNULL(Target_Total,0) when 0 then 100
						else Convert(Decimal(24,2),Round(
							Target_Total/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))*100
						, 2)) 
						end as ArcEff, ActualfinishDate
					from
					(
						Select a.id_pipe, startdatetime as startTarget, 
						MAX(Convert(decimal(24,2),Target_Hours)) as Target_Total, 
						MAX(DateAdd(hour, Convert(decimal(24,2),Target_Hours), startdatetime)) as finishTarget,
						MAX(c.finishdate) as ActualfinishDate,
						MAX(a.rig_no) as rig_no
						from (
							Select a.id_pipe, min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time))) as startdatetime, TRIM(master.dbo.Group_concat(distinct concat(' ', a.rig_no))) as rig_no
							from TblArcON a
							inner join tblMachine b on a.rig_no=b.Machine_no 
							and b.classification='endclad'
							where a.status not in('Breakdown')
							and a.rig_no not in('SHV50')
							and len(a.id_pipe)>1 and id_pipe not like '%maintenance%'
							group by a.id_pipe
							having min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time)))>DateAdd(month, -1, getDate())
						) a
						left outer join (
							Select a.PipeNo, MIN(a.Actual_Start) as Actual_Start, MAX(Target_Hours) as Target_Hours, MAX(DATEADD(hour, COnvert(decimal(24,2),Target_Hours), a.Actual_Start)) as Target_Finish
							from TargetArcOn a
							group by a.PipeNo
						) b on a.id_pipe=b.PipeNo
						outer apply 
						(
								Select * from sfnPipeStatus(a.id_pipe, '{0} 08:00') b
						) c
						group by a.id_pipe, startdatetime
						Having MAX(c.finishdate) is not null
					) a
					left outer join 
					(
						Select c.id_pipe, Sum(c.total_time) as total_time 
						from sfnGetWeldingProcessByDate('{0} 08:00') c
						group by c.id_pipe
					) c on a.Id_pipe=c.id_pipe
					where a.ActualfinishDate is not null and DATEDIFF(hour, a.ActualfinishDate, '{0} 08:00') between 0 and 24
					order by a.startTarget
				", DateTime.Now.AddDays(-25).ToString("yyyy-MM-dd"));
            //var FinishedItem = db.GetDataTableSQL("SCADA SHV", query, 1200);
            //var FinishedItem = db.GetDataTableSQL("SCADA_SHV_22Jul", query);

            //ds.Tables.Add(Summary);
            //ds.Tables[0].TableName = "Summary Items Target";
            //ds.Tables.Add(SummaryMachine);
            //ds.Tables[1].TableName = "Summary Machine Arc On";
            //ds.Tables.Add(FinishedItem);
            //ds.Tables[2].TableName = "Finished Items within 24 hours";

            //EC_ExportDataSetToExcelAppSummary(ds, filename);

            if (DateTime.Now.AddDays(-32).DayOfWeek.ToString() == "Monday")
            {
                query = String.Format(@"Declare @datDate DateTime='{0}'
					Select a1.machine_no as Machine,
						concat(convert(varchar(max), ISNULL(a.dsp,0)), ' %') as dsp1, concat(convert(varchar(max), ISNULL(a.nsp,0)), ' %') as nsp1, 
						convert(decimal(24,2), case when a.ds is null then 0 else convert(decimal(24,6),a.ds) end) as ds1, convert(decimal(24,2), case when a.ns is null then 0 else convert(decimal(24,6),a.ns) end) as ns1,
						concat(convert(varchar(max), ISNULL(b.dsp,0)), ' %') as dsp2, concat(convert(varchar(max), ISNULL(b.nsp,0)), ' %') as nsp2,
						convert(decimal(24,2),case when b.ds is null then 0 else convert(decimal(24,6),b.ds) end) as ds2, convert(decimal(24,2),case when b.ns is null then 0 else convert(decimal(24,6),b.ns) end) as ns2,
						concat(convert(varchar(max), ISNULL(c.dsp,0)), ' %') as dsp3, concat(convert(varchar(max), ISNULL(c.nsp,0)), ' %') as nsp3,
						convert(decimal(24,2),case when c.ds is null then 0 else convert(decimal(24,6),c.ds) end) as ds3, convert(decimal(24,2),case when c.ns is null then 0 else convert(decimal(24,6),c.ns) end) as ns3,
						concat(convert(varchar(max), ISNULL(d.dsp,0)), ' %') as dsp4, concat(convert(varchar(max), ISNULL(d.nsp,0)), ' %') as nsp4,
						convert(decimal(24,2),case when d.ds is null then 0 else convert(decimal(24,6),d.ds) end) as ds4, convert(decimal(24,2),case when d.ns is null then 0 else convert(decimal(24,6),d.ns) end) as ns4,
						concat(convert(varchar(max), ISNULL(e.dsp,0)), ' %') as dsp5, concat(convert(varchar(max), ISNULL(e.nsp,0)), ' %') as nsp5,
						convert(decimal(24,2),case when e.ds is null then 0 else convert(decimal(24,6),e.ds) end) as ds5, convert(decimal(24,2),case when e.ns is null then 0 else convert(decimal(24,6),e.ns) end) as ns5,
						concat(convert(varchar(max), ISNULL(f.dsp,0)), ' %') as dsp6, concat(convert(varchar(max), ISNULL(f.nsp,0)), ' %') as nsp6,
						convert(decimal(24,2),case when f.ds is null then 0 else convert(decimal(24,6),f.ds) end) as ds6, convert(decimal(24,2),case when f.ns is null then 0 else convert(decimal(24,6),f.ns) end) as ns6,
						concat(convert(varchar(max), ISNULL(g.dsp,0)), ' %') as dsp7, concat(convert(varchar(max), ISNULL(g.nsp,0)), ' %') as nsp7,
						convert(decimal(24,2),case when g.ds is null then 0 else convert(decimal(24,6),g.ds) end) as ds7, convert(decimal(24,2),case when g.ns is null then 0 else convert(decimal(24,6),g.ns) end) as ns7,
						convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0)))/7) as avgds, 
						convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0)))/7) as avgns,
						concat(convert(varchar(max), (ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0))/7*100/720), ' %') as avgdsp, 
						concat(convert(varchar(max), (ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0))/7*100/720), ' %') as avgnsp,
						isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget
					from tblMachine a1
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-7,@datDate)) a on a1.machine_no=a.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-6,@datDate)) b on a1.machine_no=b.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-5,@datDate)) c on a1.machine_no=c.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-4,@datDate)) d on a1.machine_no=d.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-3,@datDate)) e on a1.machine_no=e.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-2,@datDate)) f on a1.machine_no=f.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-1,@datDate)) g on a1.machine_no=g.rig_no
					left outer join
					(
						Select 
							b.machine_no, 
							sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
							sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
						from
						(
							Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
								DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
								b.finishdate as ActualfinishDate
							from TargetArcOn a
							outer apply 
							(
									Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@datDate)) b
							) b
							where a.Actual_Start between DateAdd(hour, 8, DateAdd(day, -7, @datDate)) and DateAdd(hour, 8, @datDate) or 
								b.finishdate between DateAdd(day, -7, @datDate) and @datDate
						) a
						inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no)) and b.classification='endclad'
						group by b.Machine_no
					) a2 on a1.machine_no=a2.Machine_No
					where a1.machine_no not like 'QH 50%' and a1.classification='endclad'
					order by LEFT(a1.machine_no,1), Try_Convert(bigint,SUBSTRING(a1.machine_no,5,len(a1.machine_no)-4))
						", DateTime.Now.AddDays(-33).ToString("yyyy-MM-dd"));
                var Summary = db.GetDataTableSQL("SCADA SHV", query, 1200);
                ds = new DataSet();
                ds.Tables.Add(Summary);
                ds.Tables[0].TableName = "Weekly Summary";
                EC_ExportDataSetToExcelAppWeeklySummary(ds, fileWeeklySummary);
            }
        }
		[HttpPostAttribute]
		public HttpResponseMessage SHVNSHH([FromBody] CRModel text)
		{
			try
			{
				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				string fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				//				db.OpenConn("SCADA SHV");
				GenerateShvNShhExcel(filename, fileWeeklySummary);

				try
                {

                    // Set email subject
                    String subject = "  Target VS Actual and Daily Rig Report (SHV and SHH) #";
                    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                        subject = "  Target VS Actual and Daily and Weekly Report (SHV and SHH) #";
                    // Set email body
                    string emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily Rig and Target VS Actual (SHV and SHH) \n\n\nThanks & Regards,\n\n\n\nProduction Automation";
                    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                        emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily and Weekly Target VS Actual (SHV and SHH) \n\n\nThanks & Regards,\n\n\n\nProduction Automation";
                    //Attaching File to Mail  
                    var builder = new BodyBuilder();
                    builder.Attachments.Add(filename);

                    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                        builder.Attachments.Add(fileWeeklySummary);

                    builder.HtmlBody = emailBody;
                    var message = new MimeMessage()
                    {
                        Body = builder.ToMessageBody(),
                        Subject = subject

                    };
					//message.To.Add(new MailboxAddress("tomoyuki.ueno@cladtek.com", "tomoyuki.ueno@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("aang.junaidi@cladtek.com", "aang.junaidi@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("iswantika.putra@cladtek.com", "iswantika.putra@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("dayanandan@cladtek.com", "dayanandan@cladtek.com"));
					//message.Bcc.Add(new MailboxAddress("muhammad.shofiq@cladtek.com", "muhammad.shofiq@cladtek.com"));
					//message.Bcc.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));

					//message.To.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));
					//					message.To.Add(new MailboxAddress("muhammad.shofiq@cladtek.com", "muhammad.shofiq@cladtek.com"));
					//					message.Cc.Add(new MailboxAddress("aang.junaidi@cladtek.com", "aang.junaidi@cladtek.com"));
					//					message.Cc.Add(new MailboxAddress("iswantika.putra@cladtek.com", "iswantika.putra@cladtek.com"));
					//					message.Bcc.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));
					//message.To.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));

					String recepientType;
					String recepientName;
					String recepientAddress;
					var recepient = db.GetDataTableSQL("SCADA SHV", "Select * from tblEmailDaily where category='shvnshh' and deleted=0");
					foreach(var data in recepient.Rows)
                    {
						recepientType = ((DataRow)data).ItemArray[1].ToString();
						recepientName = ((DataRow)data).ItemArray[2].ToString();
						recepientAddress = ((DataRow)data).ItemArray[3].ToString();
						if(recepientType.ToLower()=="to")
							message.To.Add(new MailboxAddress(recepientAddress, recepientAddress));
						if (recepientType.ToLower() == "cc")
							message.Cc.Add(new MailboxAddress(recepientAddress, recepientAddress));
						if (recepientType.ToLower() == "bcc")
							message.Bcc.Add(new MailboxAddress(recepientAddress, recepientAddress));
					}

					message.From.Add(new MailboxAddress("Scada System", "noreply.cor@cladtek.com"));
                    using (var emailClient = new MailKit.Net.Smtp.SmtpClient())
                    {

                        emailClient.Connect("smtp.gmail.com", 587, false);
                        emailClient.Authenticate("noreply.cor@cladtek.com", "tetanggaberisik!");
                        emailClient.Send(message);
                        emailClient.Disconnect(true);
                    }

                }

                catch (Exception ex)
                {
                    var stringContent1 = new StringContent(ex.Message.ToString(), UnicodeEncoding.UTF8, "application/json");
                    HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
                    result1.Content = stringContent1;
                    return result1;
                    //					return new HttpResponseMessage(HttpStatusCode.BadRequest);
                }

				var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
				var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
				result.Content = stringContent;
				return result;
			}
			catch (Exception ex)
			{
				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				var stringContent1 = new StringContent(String.Format(@"{0}. {1}", filename, ex.Message.ToString()), UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
				result1.Content = stringContent1;
				return result1;
				//				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
				//				return result;
			}

		}
		[HttpPostAttribute]
		public HttpResponseMessage AllMachine([FromBody] CRModel text)
		{
			try
			{

				// Set email subject
				String subject = "  Target VS Actual and Daily Rig Report#";
				if (DateTime.Now.DayOfWeek.ToString() == "Monday")
					subject = "  Target VS Actual and Daily and Weekly Report#";
				// Set email body
				string emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily Rig and Target VS Actual\n\n\nThanks & Regards,\n\n\n\nProduction Automation";
				if (DateTime.Now.DayOfWeek.ToString() == "Monday")
					emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily and Weekly Target VS Actual\n\n\nThanks & Regards,\n\n\n\nProduction Automation";
				//Attaching File to Mail  
				var builder = new BodyBuilder();

				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				string fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				GenerateQuadHeadExcel(filename, fileWeeklySummary);

				string filenameQH = filename;
				string fileWeeklySummaryQH = fileWeeklySummary;

				filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				GenerateShvNShhExcel(filename, fileWeeklySummary);

				string filenameShvNShh = filename;
				string fileWeeklySummaryShvNShh = fileWeeklySummary;

				filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));

				GenerateEndCladExcel(filename, fileWeeklySummary);

				string filenameEndClad = filename;
				string fileWeeklySummaryEndClad = fileWeeklySummary;

				builder.Attachments.Add(filenameQH);
				builder.Attachments.Add(filenameShvNShh);
				builder.Attachments.Add(filenameEndClad);

				if (DateTime.Now.DayOfWeek.ToString() == "Monday")
				{
					builder.Attachments.Add(fileWeeklySummaryQH);
					builder.Attachments.Add(fileWeeklySummaryShvNShh);
					builder.Attachments.Add(fileWeeklySummaryEndClad);
				}

				builder.HtmlBody = emailBody;
				var message = new MimeMessage()
				{
					Body = builder.ToMessageBody(),
					Subject = subject

				};

				String recepientType;
				String recepientName;
				String recepientAddress;
				var recepient = db.GetDataTableSQL("SCADA", "Select * from tblEmailDaily where deleted=0");
				foreach (var data in recepient.Rows)
				{
					recepientType = ((DataRow)data).ItemArray[1].ToString();
					recepientName = ((DataRow)data).ItemArray[2].ToString();
					recepientAddress = ((DataRow)data).ItemArray[3].ToString();
					if (recepientType.ToLower() == "to")
						message.To.Add(new MailboxAddress(recepientAddress, recepientAddress));
					if (recepientType.ToLower() == "cc")
						message.Cc.Add(new MailboxAddress(recepientAddress, recepientAddress));
					if (recepientType.ToLower() == "bcc")
						message.Bcc.Add(new MailboxAddress(recepientAddress, recepientAddress));
				}

				message.From.Add(new MailboxAddress("Scada System", "noreply.cor@cladtek.com"));
				using (var emailClient = new MailKit.Net.Smtp.SmtpClient())
				{

					emailClient.Connect("smtp.gmail.com", 587, false);
					emailClient.Authenticate("noreply.cor@cladtek.com", "tetanggaberisik!");
					emailClient.Send(message);
					emailClient.Disconnect(true);
				}

			}

			catch (Exception ex)
			{
				var JsonResult1 = String.Concat("{\"data\":\"[]\", \"error\":\"","\"", ex.Message.ToString() ,"\"}");
				var stringContent1 = new StringContent(JsonResult1, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
				result1.Content = stringContent1;
				return result1;
				//					return new HttpResponseMessage(HttpStatusCode.BadRequest);
			}

			var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
			var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
			HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
			result.Content = stringContent;
			return result;
		}

		[HttpPostAttribute]
		public HttpResponseMessage AllMachineSend([FromBody] CRModel text)
		{
			try
			{

				// Set email subject
				String subject = "  Target VS Actual and Daily Rig Report#";
				if (DateTime.Now.DayOfWeek.ToString() == "Monday")
					subject = "  Target VS Actual and Daily and Weekly Report#";
				// Set email body
				string emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily Rig and Target VS Actual\n\n\nThanks & Regards,\n\n\n\nProduction Automation";
				if (DateTime.Now.DayOfWeek.ToString() == "Monday")
					emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily and Weekly Target VS Actual\n\n\nThanks & Regards,\n\n\n\nProduction Automation";
				//Attaching File to Mail  
				var builder = new BodyBuilder();

				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				string fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));

				string filenameQH = filename;
				string fileWeeklySummaryQH = fileWeeklySummary;

				filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));

				string filenameShvNShh = filename;
				string fileWeeklySummaryShvNShh = fileWeeklySummary;

				filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));


				string filenameEndClad = filename;
				string fileWeeklySummaryEndClad = fileWeeklySummary;

				builder.Attachments.Add(filenameQH);
				builder.Attachments.Add(filenameShvNShh);
				builder.Attachments.Add(filenameEndClad);

				if (DateTime.Now.DayOfWeek.ToString() == "Monday")
				{
					builder.Attachments.Add(fileWeeklySummaryQH);
					//builder.Attachments.Add(fileWeeklySummaryShvNShh);
					//builder.Attachments.Add(fileWeeklySummaryEndClad);
				}

				builder.HtmlBody = emailBody;
				var message = new MimeMessage()
				{
					Body = builder.ToMessageBody(),
					Subject = subject

				};

				String recepientType;
				String recepientName;
				String recepientAddress;
				var recepient = db.GetDataTableSQL("SCADA", "Select * from tblEmailDaily where deleted=0");
				foreach (var data in recepient.Rows)
				{
					recepientType = ((DataRow)data).ItemArray[1].ToString();
					recepientName = ((DataRow)data).ItemArray[2].ToString();
					recepientAddress = ((DataRow)data).ItemArray[3].ToString();
					if (recepientType.ToLower() == "to")
						message.To.Add(new MailboxAddress(recepientAddress, recepientAddress));
					if (recepientType.ToLower() == "cc")
						message.Cc.Add(new MailboxAddress(recepientAddress, recepientAddress));
					if (recepientType.ToLower() == "bcc")
						message.Bcc.Add(new MailboxAddress(recepientAddress, recepientAddress));
				}

				message.From.Add(new MailboxAddress("Scada System", "noreply.cor@cladtek.com"));
				using (var emailClient = new MailKit.Net.Smtp.SmtpClient())
				{

					emailClient.Connect("smtp.gmail.com", 587, false);
					emailClient.Authenticate("noreply.cor@cladtek.com", "tetanggaberisik!");
					emailClient.Send(message);
					emailClient.Disconnect(true);
				}

			}

			catch (Exception ex)
			{
				var JsonResult1 = String.Concat("{\"data\":\"[]\", \"error\":\"", "\"", ex.Message.ToString(), "\"}");
				var stringContent1 = new StringContent(JsonResult1, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
				result1.Content = stringContent1;
				return result1;
				//					return new HttpResponseMessage(HttpStatusCode.BadRequest);
			}

			var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
			var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
			HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
			result.Content = stringContent;
			return result;
		}

		[HttpPostAttribute]
		public HttpResponseMessage AllMachineWithoutSend([FromBody] CRModel text)
		{
			try
			{

				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				string fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				//GenerateQuadHeadExcel(filename, fileWeeklySummary);

				string filenameQH = filename;
				string fileWeeklySummaryQH = fileWeeklySummary;

				filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-Summary-{0}.xlsx", DateTime.Now.AddDays(-18).ToString("dd-MM-yyyy")));
				fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(-18).ToString("dd-MM-yyyy")));
				GenerateShvNShhExcel(filename, fileWeeklySummary);

				string filenameShvNShh = filename;
				string fileWeeklySummaryShvNShh = fileWeeklySummary;

				filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-Summary-{0}.xlsx", DateTime.Now.AddDays(-18).ToString("dd-MM-yyyy")));
				fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(-18).ToString("dd-MM-yyyy")));

				GenerateEndCladExcel(filename, fileWeeklySummary);

				string filenameEndClad = filename;
				string fileWeeklySummaryEndClad = fileWeeklySummary;

			}

			catch (Exception ex)
			{
				var JsonResult1 = String.Concat("{\"data\":\"[]\", \"error\":\"", "\"", ex.Message.ToString(), "\"}");
				var stringContent1 = new StringContent(JsonResult1, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
				result1.Content = stringContent1;
				return result1;
				//					return new HttpResponseMessage(HttpStatusCode.BadRequest);
			}

			var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
			var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
			HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
			result.Content = stringContent;
			return result;
		}

		[HttpPostAttribute]
		public HttpResponseMessage AllMachineMonthlyWithoutSend([FromBody] ClsEmailModel text)
		{
			DateTime datDate = (text.date == new DateTime()) ? DateTime.Now : text.date;
			try
			{
				string fileMonthlySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("QH-MonthlySummary-{0}.xlsx", datDate.AddMonths(-1).ToString("MMM-yyyy")));
                GenerateMonthlyQuadHeadExcel(fileMonthlySummary, datDate);

                string fileMonthlySummaryQH = fileMonthlySummary;

                fileMonthlySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("SHV-MonthlySummary-{0}.xlsx", datDate.AddMonths(-1).ToString("MMM-yyyy")));
                GenerateMonthlyShvNShhExcel(fileMonthlySummary, datDate);

                string fileMonthlySummaryShvNShh = fileMonthlySummary;

                fileMonthlySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("EC-MonthlySummary-{0}.xlsx", datDate.AddMonths(-1).ToString("MMM-yyyy")));
                GenerateMonthlyEndCladExcel(fileMonthlySummary, datDate);

                string fileMonthlySummaryEndClad = fileMonthlySummary;

            }

			catch (Exception ex)
			{
				var JsonResult1 = String.Concat("{\"data\":\"[]\", \"error\":\"", "\"", ex.Message.ToString(), "\"}");
				var stringContent1 = new StringContent(JsonResult1, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result1 = new HttpResponseMessage(HttpStatusCode.OK);
				result1.Content = stringContent1;
				return result1;
				//					return new HttpResponseMessage(HttpStatusCode.BadRequest);
			}

			var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
			var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
			HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
			result.Content = stringContent;
			return result;
		}
		private void GenerateShvNShhExcel(string filename, string fileWeeklySummary)
		{
			var ds = new DataSet();

			//string query = String.Format(@"Select 
			//			b.machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
			//			ISNULL(c.total_time/60/60,0) as Actual_Arc_Time,
			//			DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) as Actual_Time, 
			//			Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//			else Convert(Decimal(24,2),Round(
			//				100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
			//			, 2)) 
			//			end as ArcOnTime, 
			//			Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//			else Convert(Decimal(24,2),Round(
			//				100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
			//			, 2)) 
			//			end as ArcEff, ActualfinishDate
			//		from
			//		(
			//			Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
			//			Convert(decimal(24,2),Target_Hours) as Target_Total, DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
			//			b.finishdate as ActualfinishDate
			//			from TargetArcOn a
			//			outer apply 
			//			(
			//					Select * from sfnPipeOnMachineSatus(a.Machine, a.pipeno, '{0} 08:00') b
			//			) b
			//		) a
			//		inner join tblMachine b on a.Machine_no=b.Machine_no and b.classification='shvnshh'
			//		left outer join sfnGetWeldingProcessByDate('{0} 08:00') c on b.Machine_no=c.rig_no and a.Id_pipe=c.id_pipe
			//		where a.ActualfinishDate is null and a.startTarget<='{0} 08:00'
			//		order by b.machine_no
			//		", DateTime.Now.ToString("yyyy-MM-dd"));
   //         var Summary = db.GetDataTableSQL("SCADA SHV", query, 1200);
			string query = String.Format(@"Select
						a.rig_no as machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
						ISNULL(c.total_time/60,0) as Actual_Arc_Time,
						Convert(Decimal(24,2), Convert(Decimal(24,6),DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60) as Actual_Time, 
						Case DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
						else Convert(Decimal(24,2),Round(
							(100 * Convert(Decimal(24,2), ISNULL(c.total_time,0))/
								Convert(Decimal(24,2), Convert(Decimal(24,6), DateDiff(second, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60)
							)
						, 2)) 
						end as ArcOnTime
					from
					(
						Select a.id_pipe, startdatetime as startTarget, 
						MAX(Convert(decimal(24,2),Target_Hours)) as Target_Total, 
						MAX(DateAdd(hour, Convert(decimal(24,2),Target_Hours), startdatetime)) as finishTarget,
						MAX(c.finishdate) as ActualfinishDate,
						MAX(a.rig_no) as rig_no
						from (
							Select a.id_pipe, min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time))) as startdatetime, TRIM(master.dbo.Group_concat(distinct concat(' ', a.rig_no))) as rig_no
							from TblArcON a
							inner join tblMachine b on a.rig_no=b.Machine_no 
							and b.classification='shvnshh'
							where a.status not in('Breakdown')
							and a.rig_no not in('SHV50')
							and len(a.id_pipe)>1 and id_pipe not like '%maintenance%'
							group by a.id_pipe
							having min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time)))>DateAdd(month, -1, getDate())
						) a
						left outer join (
							Select a.PipeNo, MIN(a.Actual_Start) as Actual_Start, MAX(Target_Hours) as Target_Hours, MAX(DATEADD(hour, COnvert(decimal(24,2),Target_Hours), a.Actual_Start)) as Target_Finish
							from TargetArcOn a
							group by a.PipeNo
						) b on a.id_pipe=b.PipeNo
						outer apply 
						(
								Select * from dbo.sfnPipeStatus(a.id_pipe, '{0} 08:00') b
						) c
						group by a.id_pipe, startdatetime
						Having MAX(c.finishdate) is null
					) a
					left outer join 
					(
						Select c.id_pipe, Sum(c.total_time) as total_time 
						from sfnGetWeldingProcessByDate('{0} 08:00') c
						group by c.id_pipe
					) c on a.Id_pipe=c.id_pipe
					where a.ActualfinishDate is null and a.startTarget<='{0} 08:00'
					order by a.startTarget
					", DateTime.Now.AddDays(-18).ToString("yyyy-MM-dd"));
			//var Summary = db.GetDataTableSQL("SCADA SHV", query, 1200);
			//var Summary = db.GetDataTableSQL("SCADA_SHV_22Jul", query);
//			query = String.Format(@"Declare @datDate DateTime='{0}'
//						Select a.*,
//							Case when Machine='Average' or ds=0 then '-' else dbo.sfnGetWelderByMachineShift(@datDate, Machine, 'd')end as welder_ds,
//							Case when Machine='Average' or ns=0 then '-' else dbo.sfnGetWelderByMachineShift(@datDate, Machine, 'n') end as welder_ns
//						from(
//						Select rig_no as Machine, concat(convert(varchar(max), dsp), ' %') as dsp, concat(convert(varchar(max), nsp), ' %') as nsp, 
//							convert(decimal(24,2),convert(decimal(24,6),ds)) as ds, convert(decimal(24,2),convert(decimal(24,6),ns)) as ns
//						from dbo.sfnGetWeldingProcessByMachine(@datDate) a
//						inner join TblMachine b on a.rig_no=b.machine_no and b.classification='shvnshh'
//						where CharIndex(' ',a.rig_no)>0
//and a.rig_no not in('SHV 12', 'SHV 47', 'SHV 58')
//						union all
///*						Select 'Average' as rig_no, concat(convert(varchar(max), Convert(Decimal(24,2),Round(Avg(dsp),2))), ' %') as dsp, 
//							concat(convert(varchar(max), Convert(Decimal(24,2),Round(Avg(nsp),2))), ' %') as nsp, 
//							convert(Decimal(24,2),Avg(convert(Decimal(24,12),ds)),2) as ds, convert(Decimal(24,2),Avg(convert(Decimal(24,12),ns)),2) as ns*/
//						Select 'Average' as rig_no, 
//							concat(convert(varchar(max), Case Sum(Case when dsp>0 then 1 else 0 end) when 0 then 0 else Convert(Decimal(24,2),Round(SUM(dsp)/Sum(Case when dsp>0 then 1 else 0 end),2)) end), ' %')as dsp, 
//							concat(convert(varchar(max), Case Sum(Case when nsp>0 then 1 else 0 end) when 0 then 0 else Convert(Decimal(24,2),Round(Sum(nsp)/Sum(Case when nsp>0 then 1 else 0 end),2)) end), ' %') as nsp, 
//							Case Sum(Case when ds>0 then 1 else 0 end) when 0 then 0 else convert(Decimal(24,2),Sum(convert(Decimal(24,12),ds))/Sum(Case when ds>0 then 1 else 0 end),2) end as ds, 
//							Case Sum(Case when ns>0 then 1 else 0 end) when 0 then 0 else convert(Decimal(24,2),Sum(convert(Decimal(24,12),ns))/Sum(Case when ns>0 then 1 else 0 end),2) end as ns
//						from dbo.sfnGetWeldingProcessByMachine(@datDate) a
//						inner join TblMachine b on a.rig_no=b.machine_no and b.classification='shvnshh'
//						where CharIndex(' ',a.rig_no)>0
//and a.rig_no not in('SHV 12', 'SHV 47', 'SHV 58')
//						) a
//						order by Replace(Left(MAchine,1),'A','Z'), ISNULL(Try_Convert(bigint,SUBSTRING(Machine,5,len(Machine)-4)),100)
//				", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
			query = String.Format(@"Declare @datDate DateTime='{0}'
						Select a.*,
							Case when Machine='Average' or ds=0 then '-' else dbo.sfnGetWelderByMachineShift(@datDate, Machine, 'd')end as welder_ds,
							Case when Machine='Average' or ns=0 then '-' else dbo.sfnGetWelderByMachineShift(@datDate, Machine, 'n') end as welder_ns
						from(
							Select a.rig_no as Machine,
							case when a.ds=0 then 'OFF' when c.ds>0 then 'OT' else 'NORMAL' end as dswh,
							case when c.ds>0 then 
								concat(convert(varchar(max), dsp), ' %')
							else
								concat(convert(varchar(max), convert(decimal(24,2), round(convert(decimal(24,6), a.ds)/8,2))*100), ' %')
							end as dsp, 
							case when a.ns=0 then 'OFF' when c.ns>0 then 'OT' else 'NORMAL' end as nswh,
							case when c.ns>0 then 
								concat(convert(varchar(max), nsp), ' %')
							else
								concat(convert(varchar(max), convert(decimal(24,2), round(convert(decimal(24,6), a.ns)/8,2))*100), ' %')
							end as nsp, 
							convert(decimal(24,2),convert(decimal(24,6),a.ds)) as ds, convert(decimal(24,2),convert(decimal(24,6),a.ns)) as ns
						from dbo.sfnGetWeldingProcessByMachine(@datDate) a
						inner join TblMachine b on a.rig_no=b.machine_no and b.classification='shvnshh'
						left join dbo.sfnGetOTWeldingProcessByMachine(@datDate) c on a.rig_no=c.rig_no
						where CharIndex(' ',a.rig_no)>0
						union all
						Select 'Average' as rig_no, 
							'-' as dswh,
							concat(convert(varchar(max), Case Sum(Case when dsp>0 then 1 else 0 end) when 0 then 0 else 
								Convert(Decimal(24,2),Round(SUM(
								case when c.ds>0 then 
									dsp
								else 
									convert(decimal(24,2), round(convert(decimal(24,6), a.ds)/8,2))*100
								end
								)/Sum(Case when dsp>0 then 1 else 0 end),2)) end), ' %')
							as dsp, 
							'-' as nswh,
							concat(convert(varchar(max), Case Sum(Case when nsp>0 then 1 else 0 end) when 0 then 0 else 
								Convert(Decimal(24,2),Round(Sum(
								case when c.ds>0 then 
									nsp
								else 
									convert(decimal(24,2), round(convert(decimal(24,6), a.ns)/8,2))*100
								end
								)/Sum(Case when nsp>0 then 1 else 0 end),2)) end), ' %') 
							as nsp, 
							Case Sum(Case when a.ds>0 then 1 else 0 end) when 0 then 0 else convert(Decimal(24,2),Sum(convert(Decimal(24,12),a.ds))/Sum(Case when a.ds>0 then 1 else 0 end),2) end as ds, 
							Case Sum(Case when a.ns>0 then 1 else 0 end) when 0 then 0 else convert(Decimal(24,2),Sum(convert(Decimal(24,12),a.ns))/Sum(Case when a.ns>0 then 1 else 0 end),2) end as ns
						from dbo.sfnGetWeldingProcessByMachine(@datDate) a
						inner join TblMachine b on a.rig_no=b.machine_no and b.classification='shvnshh'
						left join dbo.sfnGetOTWeldingProcessByMachine(@datDate) c on a.rig_no=c.rig_no
						where CharIndex(' ',a.rig_no)>0
						) a
						order by Replace(Left(MAchine,1),'A','Z'), ISNULL(Try_Convert(bigint,SUBSTRING(Machine,5,len(Machine)-4)),100)
				", DateTime.Now.AddDays(-19).ToString("yyyy-MM-dd"));

			//var SummaryMachine = db.GetDataTableSQL("SCADA SHV", query, 1200);
            //var SummaryMachine = db.GetDataTableSQL("SCADA_SHV_22Jul", query);
			//query = String.Format(@"Select 
			//		b.machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, a.ActualfinishDate,
			//		ISNULL(c.total_time/60,0) as Actual_Arc_Time,
			//		DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) as Actual_Time, 
			//		Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//		else Convert(Decimal(24,2),Round(
			//			100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
			//		, 2)) 
			//		end as ArcOnTime, 
			//		Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
			//		else Convert(Decimal(24,2),Round(
			//			Target_Total/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))*100
			//		, 2)) 
			//		end as ArcEff
			//	from
			//	(
			//		Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
			//		Convert(decimal(24,2),Target_Hours) as Target_Total, DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
			//		b.finishdate as ActualfinishDate
			//		from TargetArcOn a
			//		outer apply 
			//		(
			//				Select * from sfnPipeOnMachineSatus(a.Machine, a.pipeno, '{0} 08:00') b
			//		) b
			//	) a
			//	inner join tblMachine b on a.Machine_no=b.Machine_no and b.classification='shvnshh'
			//	outer apply
			//	(
			//		Select * from sfnGetWeldingProcessByDate(ISNULL(a.ActualfinishDate, '{0} 08:00')) c
			//		where b.Machine_no=c.rig_no and a.Id_pipe=c.id_pipe
			//	) c 
			//	where a.ActualfinishDate is not null and DATEDIFF(hour, a.ActualfinishDate, '{0} 08:00') between 0 and 24
			//	order by LEFT(b.machine_no,1), Try_Convert(bigint,SUBSTRING(b.machine_no,5,len(b.machine_no)-4))
			//	", DateTime.Now.ToString("yyyy-MM-dd"));
			query = String.Format(@"Select 
						a.rig_no as machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
						ISNULL(c.total_time/60,0) as Actual_Arc_Time,
						Convert(Decimal(24,2), Convert(Decimal(24,6),DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60) as Actual_Time, 
						Case DateDiff(minute, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
						else Convert(Decimal(24,2),Round(
							(100 * Convert(Decimal(24,2), ISNULL(c.total_time,0))/
								Convert(Decimal(24,2), Convert(Decimal(24,6), DateDiff(second, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')))/60)
							)
						, 2)) 
						end as ArcOnTime, 
						Case ISNULL(Target_Total,0) when 0 then 100
						else Convert(Decimal(24,2),Round(
							Target_Total/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))*100
						, 2)) 
						end as ArcEff, ActualfinishDate
					from
					(
						Select a.id_pipe, startdatetime as startTarget, 
						MAX(Convert(decimal(24,2),Target_Hours)) as Target_Total, 
						MAX(DateAdd(hour, Convert(decimal(24,2),Target_Hours), startdatetime)) as finishTarget,
						MAX(c.finishdate) as ActualfinishDate,
						MAX(a.rig_no) as rig_no
						from (
							Select a.id_pipe, min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time))) as startdatetime, TRIM(master.dbo.Group_concat(distinct concat(' ', a.rig_no))) as rig_no
							from TblArcON a
							inner join tblMachine b on a.rig_no=b.Machine_no 
							and b.classification='shvnshh'
							where a.status not in('Breakdown')
							and a.rig_no not in('SHV50')
							and len(a.id_pipe)>1 and id_pipe not like '%maintenance%'
							group by a.id_pipe
							having min(convert(datetime, dbo.sfnGetDateTimeFormat(a.date, a.start_time)))>DateAdd(month, -1, getDate())
						) a
						left outer join (
							Select a.PipeNo, MIN(a.Actual_Start) as Actual_Start, MAX(Target_Hours) as Target_Hours, MAX(DATEADD(hour, COnvert(decimal(24,2),Target_Hours), a.Actual_Start)) as Target_Finish
							from TargetArcOn a
							group by a.PipeNo
						) b on a.id_pipe=b.PipeNo
						outer apply 
						(
								Select * from dbo.sfnPipeStatus(a.id_pipe, '{0} 08:00') b
						) c
						group by a.id_pipe, startdatetime
						Having MAX(c.finishdate) is not null
					) a
					left outer join 
					(
						Select c.id_pipe, Sum(c.total_time) as total_time 
						from sfnGetWeldingProcessByDate('{0} 08:00') c
						group by c.id_pipe
					) c on a.Id_pipe=c.id_pipe
					where a.ActualfinishDate is not null and DATEDIFF(hour, a.ActualfinishDate, '{0} 08:00') between 0 and 24
					order by a.startTarget
				", DateTime.Now.AddDays(-18).ToString("yyyy-MM-dd"));
            //var FinishedItem = db.GetDataTableSQL("SCADA SHV", query, 1200);
            //var FinishedItem = db.GetDataTableSQL("SCADA_SHV_22Jul", query);

            //ds.Tables.Add(Summary);
            //ds.Tables[0].TableName = "Summary Items Target";
            //ds.Tables.Add(SummaryMachine);
            //ds.Tables[1].TableName = "Summary Machine Arc On";
            //ds.Tables.Add(FinishedItem);
            //ds.Tables[2].TableName = "Finished Items within 24 hours";

            //SHV_ExportDataSetToExcelAppSummary(ds, filename);

            if (DateTime.Now.AddDays(-32).DayOfWeek.ToString() == "Monday")
            {
                query = String.Format(@"Declare @datDate DateTime='{0}'
					Select a1.machine_no as Machine,
						concat(convert(varchar(max), ISNULL(a.dsp,0)), ' %') as dsp1, concat(convert(varchar(max), ISNULL(a.nsp,0)), ' %') as nsp1, 
						convert(decimal(24,2), case when a.ds is null then 0 else convert(decimal(24,6),a.ds) end) as ds1, convert(decimal(24,2), case when a.ns is null then 0 else convert(decimal(24,6),a.ns) end) as ns1,
						concat(convert(varchar(max), ISNULL(b.dsp,0)), ' %') as dsp2, concat(convert(varchar(max), ISNULL(b.nsp,0)), ' %') as nsp2,
						convert(decimal(24,2),case when b.ds is null then 0 else convert(decimal(24,6),b.ds) end) as ds2, convert(decimal(24,2),case when b.ns is null then 0 else convert(decimal(24,6),b.ns) end) as ns2,
						concat(convert(varchar(max), ISNULL(c.dsp,0)), ' %') as dsp3, concat(convert(varchar(max), ISNULL(c.nsp,0)), ' %') as nsp3,
						convert(decimal(24,2),case when c.ds is null then 0 else convert(decimal(24,6),c.ds) end) as ds3, convert(decimal(24,2),case when c.ns is null then 0 else convert(decimal(24,6),c.ns) end) as ns3,
						concat(convert(varchar(max), ISNULL(d.dsp,0)), ' %') as dsp4, concat(convert(varchar(max), ISNULL(d.nsp,0)), ' %') as nsp4,
						convert(decimal(24,2),case when d.ds is null then 0 else convert(decimal(24,6),d.ds) end) as ds4, convert(decimal(24,2),case when d.ns is null then 0 else convert(decimal(24,6),d.ns) end) as ns4,
						concat(convert(varchar(max), ISNULL(e.dsp,0)), ' %') as dsp5, concat(convert(varchar(max), ISNULL(e.nsp,0)), ' %') as nsp5,
						convert(decimal(24,2),case when e.ds is null then 0 else convert(decimal(24,6),e.ds) end) as ds5, convert(decimal(24,2),case when e.ns is null then 0 else convert(decimal(24,6),e.ns) end) as ns5,
						concat(convert(varchar(max), ISNULL(f.dsp,0)), ' %') as dsp6, concat(convert(varchar(max), ISNULL(f.nsp,0)), ' %') as nsp6,
						convert(decimal(24,2),case when f.ds is null then 0 else convert(decimal(24,6),f.ds) end) as ds6, convert(decimal(24,2),case when f.ns is null then 0 else convert(decimal(24,6),f.ns) end) as ns6,
						concat(convert(varchar(max), ISNULL(g.dsp,0)), ' %') as dsp7, concat(convert(varchar(max), ISNULL(g.nsp,0)), ' %') as nsp7,
						convert(decimal(24,2),case when g.ds is null then 0 else convert(decimal(24,6),g.ds) end) as ds7, convert(decimal(24,2),case when g.ns is null then 0 else convert(decimal(24,6),g.ns) end) as ns7,
						convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0)))/7) as avgds, 
						convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0)))/7) as avgns,
						concat(convert(varchar(max), (ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0))/7*100/720), ' %') as avgdsp, 
						concat(convert(varchar(max), (ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0))/7*100/720), ' %') as avgnsp,
						isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget
					from tblMachine a1
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-7,@datDate)) a on a1.machine_no=a.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-6,@datDate)) b on a1.machine_no=b.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-5,@datDate)) c on a1.machine_no=c.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-4,@datDate)) d on a1.machine_no=d.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-3,@datDate)) e on a1.machine_no=e.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-2,@datDate)) f on a1.machine_no=f.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-1,@datDate)) g on a1.machine_no=g.rig_no
					left outer join
					(
						Select 
							b.machine_no, 
							sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
							sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
						from
						(
							Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
								DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
								b.finishdate as ActualfinishDate
							from TargetArcOn a
							outer apply 
							(
									Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@datDate)) b
							) b
							where a.Actual_Start between DateAdd(hour, 8, DateAdd(day, -7, @datDate)) and DateAdd(hour, 8, @datDate) or 
								b.finishdate between DateAdd(day, -7, @datDate) and @datDate
						) a
						inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no)) and b.classification='shvnshh'
						group by b.Machine_no
					) a2 on a1.machine_no=a2.Machine_No
					where a1.machine_no not like 'QH 50%' and a1.classification='shvnshh'
					order by LEFT(a1.machine_no,1), Try_Convert(bigint,SUBSTRING(a1.machine_no,5,len(a1.machine_no)-4))
						", DateTime.Now.AddDays(-33).ToString("yyyy-MM-dd"));
                var Summary = db.GetDataTableSQL("SCADA SHV", query, 2400);
                ds = new DataSet();
                ds.Tables.Add(Summary);
                ds.Tables[0].TableName = "Weekly Summary";

                SHV_ExportDataSetToExcelAppWeeklySummary(ds, fileWeeklySummary);
            }
        }
		[HttpPostAttribute]
		public HttpResponseMessage QuadHead([FromBody] CRModel text)
		{
			try
			{
				string filename = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("Summary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				string fileWeeklySummary = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~/ExcelFiles"), String.Format("WeeklySummary-{0}.xlsx", DateTime.Now.AddDays(0).ToString("dd-MM-yyyy")));
				GenerateQuadHeadExcel(filename, fileWeeklySummary);
				try
				{

					// Set email subject
					String subject = "  Target VS Actual and Daily Rig Report#";
					if (DateTime.Now.DayOfWeek.ToString() == "Monday")
						subject = "  Target VS Actual and Daily and Weekly Report#";
					// Set email body
					string emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily Rig and Target VS Actual \n\n\nThanks & Regards,\n\n\n\nProduction Automation";
					if (DateTime.Now.DayOfWeek.ToString() == "Monday")
						emailBody = "Dear All Production (PIC Person),\n\n\nThis is Report for Daily and Weekly Target VS Actual \n\n\nThanks & Regards,\n\n\n\nProduction Automation";
					//Attaching File to Mail  
					var builder = new BodyBuilder();
					builder.Attachments.Add(filename);

					if (DateTime.Now.DayOfWeek.ToString() == "Monday")
						builder.Attachments.Add(fileWeeklySummary);

					builder.HtmlBody = emailBody;
					var message = new MimeMessage()
					{
						Body = builder.ToMessageBody(),
						Subject = subject

					};
					//message.To.Add(new MailboxAddress("tomoyuki.ueno@cladtek.com", "tomoyuki.ueno@cladtek.com"));
					//               message.Cc.Add(new MailboxAddress("aang.junaidi@cladtek.com", "aang.junaidi@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("iswantika.putra@cladtek.com", "iswantika.putra@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("craig.duncan@cladtek.com", "craig.duncan@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("dayanandan@cladtek.com", "dayanandan@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("suryanarayanan@cladtek.com", "suryanarayanan@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("wayne.williams@cladtek.com", "wayne.williams@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("ganesh.karri@cladtek.com", "ganesh.karri@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("patar.pangaribuan@cladtek.com", "patar.pangaribuan@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("dhani.purwadi@cladtek.com", "dhani.purwadi@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("vinod.upadhyay@cladtek.com", "vinod.upadhyay@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("nurfahmi@cladtek.com", "nurfahmi@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("jhonson.silitonga@cladtek.com", "jhonson.silitonga@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("suherimon@cladtek.com", "suherimon@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("zulhendri@cladtek.com", "zulhendri@cladtek.com"));
					//message.Cc.Add(new MailboxAddress("tari.rumantias@cladtek.com", "tari.rumantias@cladtek.com"));
					//message.Bcc.Add(new MailboxAddress("muhammad.shofiq@cladtek.com", "muhammad.shofiq@cladtek.com"));
					//message.Bcc.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));

					//message.To.Add(new MailboxAddress("putra.budiman@cladtek.com", "putra.budiman@cladtek.com"));

                    message.From.Add(new MailboxAddress("Scada System", "noreply.cor@cladtek.com"));
					using (var emailClient = new MailKit.Net.Smtp.SmtpClient())
					{

						emailClient.Connect("smtp.gmail.com", 587, false);
						emailClient.Authenticate("noreply.cor@cladtek.com", "tetanggaberisik!");
						emailClient.Send(message);
						emailClient.Disconnect(true);
					}

				}

				catch (Exception ex)
				{
					return new HttpResponseMessage(HttpStatusCode.BadRequest);
				}

				var JsonResult = String.Concat("{\"data\":\"[]\", \"success\":\"ok\"}");
				var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
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
		private void GenerateQuadHeadExcel(string filename, string fileWeeklySummary)
		{
			var ds = new DataSet();
			string query = String.Format(@"Select 
					b.machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
					ISNULL(c.total_time/60,0) as Actual_Arc_Time,
					DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) as Actual_Time, 
					Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
					else Convert(Decimal(24,2),Round(
						100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
					, 2)) 
					end as ArcOnTime, 
					Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
					else Convert(Decimal(24,2),Round(
						100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
					, 2)) 
					end as ArcEff
				from
				(
					Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
					Convert(decimal(24,2),Target_Hours) as Target_Total, DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
					b.finishdate as ActualfinishDate
					from (Select * from TargetArcOn where Actual_Start<='{0}' and finishDate is null) a
					outer apply 
					(
							Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, '{0} 08:00') b
					) b
				) a
				inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
				left outer join sfnGetWeldingProcessByDate('{0} 08:00') c on b.Machine_no=c.rig_no and a.Id_pipe=c.id_pipe
				where a.ActualfinishDate is null and a.startTarget<='{0} 08:00'
				order by Try_Convert(bigint,SUBSTRING(b.machine_no,4,len(b.machine_no)-4))
				", DateTime.Now.ToString("yyyy-MM-dd"));
			var Summary = db.GetDataTableSQL("SCADA QH", query, 1200);

			query = String.Format(@"Select * from(
							Select rig_no as Machine, concat(convert(varchar(max), dsp), ' %') as dsp, concat(convert(varchar(max), nsp), ' %') as nsp, 
								convert(decimal(24,2),convert(decimal(24,6),ds)/60) as ds, convert(decimal(24,2),convert(decimal(24,6),ns)/60) as ns
							from dbo.sfnGetWeldingProcessByMachine('{0}') a
							union all
							Select 'Average' as rig_no, concat(convert(varchar(max), Convert(Decimal(24,2),Round(Avg(dsp),2))), ' %') as dsp, 
								concat(convert(varchar(max), Convert(Decimal(24,2),Round(Avg(nsp),2))), ' %') as nsp, 
								convert(Decimal(24,2),Avg(convert(Decimal(24,12),ds)/60),2) as ds, convert(Decimal(24,2),Avg(convert(Decimal(24,12),ns)/60),2) as ns
							from dbo.sfnGetWeldingProcessByMachine('{0}') a
							) a
							order by ISNULL(Try_Convert(bigint,SUBSTRING(Machine,4,len(Machine)-4)),100)
				", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
			var SummaryMachine = db.GetDataTableSQL("SCADA QH", query);

			query = String.Format(@"Select 
					b.machine_no, a.id_pipe, a.startTarget, a.finishTarget, Round(Target_Total,2) as Target_Time, 
					ISNULL(c.total_time/60,0) as Actual_Arc_Time,
					DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) as Actual_Time, 
					Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
					else Convert(Decimal(24,2),Round(
						100 * ISNULL(c.total_time/60,0)/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))
					, 2)) 
					end as ArcOnTime, 
					Case DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00')) when 0 then 0
					else Convert(Decimal(24,2),Round(
						Target_Total/DateDiff(hour, a.startTarget, ISNULL(ActualFinishDate, '{0} 08:00'))*100
					, 2)) 
					end as ArcEff
				from
				(
					Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
					Convert(decimal(24,2),Target_Hours) as Target_Total, DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
					b.finishdate as ActualfinishDate
					from (Select * from TargetArcOn where Actual_Start<='{0}' and finishDate is null) a
					outer apply 
					(
							Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, '{0} 08:00') b
					) b
				) a
				inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
				outer apply
				(
					Select * from sfnGetWeldingProcessByDate(ISNULL(a.ActualfinishDate, '{0} 08:00')) c
					where b.Machine_no=c.rig_no and a.Id_pipe=c.id_pipe
				) c 
				where a.ActualfinishDate is not null and DATEDIFF(hour, a.ActualfinishDate, '{0} 08:00') between 0 and 24
				order by Try_Convert(bigint,SUBSTRING(b.machine_no,4,len(b.machine_no)-4))
					", DateTime.Now.ToString("yyyy-MM-dd"));
			var FinishedItem = db.GetDataTableSQL("SCADA QH", query, 1200);

			ds.Tables.Add(Summary);
			ds.Tables[0].TableName = "Summary Pipe Target";
			ds.Tables.Add(SummaryMachine);
			ds.Tables[1].TableName = "Summary Machine Arc On";
			ds.Tables.Add(FinishedItem);
			ds.Tables[2].TableName = "Finished Pipes within 24 hours";

			QH_ExportDataSetToExcelAppSummary(ds, filename);

			if (DateTime.Now.DayOfWeek.ToString() == "Monday")
			{
				query = String.Format(@"Declare @datDate DateTime='{0}'
												Select a1.machine_no as Machine,
						concat(convert(varchar(max), ISNULL(a.dsp,0)), ' %') as dsp1, concat(convert(varchar(max), ISNULL(a.nsp,0)), ' %') as nsp1, 
						convert(decimal(24,2), case when a.ds is null then 0 else convert(decimal(24,6),a.ds)/60 end) as ds1, convert(decimal(24,2), case when a.ns is null then 0 else convert(decimal(24,6),a.ns)/60 end) as ns1,
						concat(convert(varchar(max), ISNULL(b.dsp,0)), ' %') as dsp2, concat(convert(varchar(max), ISNULL(b.nsp,0)), ' %') as nsp2,
						convert(decimal(24,2),case when b.ds is null then 0 else convert(decimal(24,6),b.ds)/60 end) as ds2, convert(decimal(24,2),case when b.ns is null then 0 else convert(decimal(24,6),b.ns)/60 end) as ns2,
						concat(convert(varchar(max), ISNULL(c.dsp,0)), ' %') as dsp3, concat(convert(varchar(max), ISNULL(c.nsp,0)), ' %') as nsp3,
						convert(decimal(24,2),case when c.ds is null then 0 else convert(decimal(24,6),c.ds)/60 end) as ds3, convert(decimal(24,2),case when c.ns is null then 0 else convert(decimal(24,6),c.ns)/60 end) as ns3,
						concat(convert(varchar(max), ISNULL(d.dsp,0)), ' %') as dsp4, concat(convert(varchar(max), ISNULL(d.nsp,0)), ' %') as nsp4,
						convert(decimal(24,2),case when d.ds is null then 0 else convert(decimal(24,6),d.ds)/60 end) as ds4, convert(decimal(24,2),case when d.ns is null then 0 else convert(decimal(24,6),d.ns)/60 end) as ns4,
						concat(convert(varchar(max), ISNULL(e.dsp,0)), ' %') as dsp5, concat(convert(varchar(max), ISNULL(e.nsp,0)), ' %') as nsp5,
						convert(decimal(24,2),case when e.ds is null then 0 else convert(decimal(24,6),e.ds)/60 end) as ds5, convert(decimal(24,2),case when e.ns is null then 0 else convert(decimal(24,6),e.ns)/60 end) as ns5,
						concat(convert(varchar(max), ISNULL(f.dsp,0)), ' %') as dsp6, concat(convert(varchar(max), ISNULL(f.nsp,0)), ' %') as nsp6,
						convert(decimal(24,2),case when f.ds is null then 0 else convert(decimal(24,6),f.ds)/60 end) as ds6, convert(decimal(24,2),case when f.ns is null then 0 else convert(decimal(24,6),f.ns)/60 end) as ns6,
						concat(convert(varchar(max), ISNULL(g.dsp,0)), ' %') as dsp7, concat(convert(varchar(max), ISNULL(g.nsp,0)), ' %') as nsp7,
						convert(decimal(24,2),case when g.ds is null then 0 else convert(decimal(24,6),g.ds)/60 end) as ds7, convert(decimal(24,2),case when g.ns is null then 0 else convert(decimal(24,6),g.ns)/60 end) as ns7,
						convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0)))/7/60) as avgds, 
						convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0)))/7/60) as avgns,
						concat(convert(varchar(max), (ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0))/7*100/720), ' %') as avgdsp, 
						concat(convert(varchar(max), (ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0))/7*100/720), ' %') as avgnsp,
						isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget
					from tblMachine a1
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-7,@datDate)) a on a1.machine_no=a.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-6,@datDate)) b on a1.machine_no=b.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-5,@datDate)) c on a1.machine_no=c.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-4,@datDate)) d on a1.machine_no=d.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-3,@datDate)) e on a1.machine_no=e.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-2,@datDate)) f on a1.machine_no=f.rig_no
					left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-1,@datDate)) g on a1.machine_no=g.rig_no
					left outer join
					(
					Select 
						b.machine_no, 
						sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
						sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
					from
					(
						Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
							DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
							b.finishdate as ActualfinishDate
						from TargetArcOn a
						outer apply 
						(
								Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@datDate)) b
						) b
						where a.Actual_Start between DateAdd(hour, 8, DateAdd(day, -7, @datDate)) and DateAdd(hour, 8, @datDate) or 
							b.finishdate between DateAdd(day, -7, @datDate) and @datDate
					) a
					inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
					group by b.Machine_no
					) a2 on a1.machine_no=a2.Machine_No
					where a1.machine_no not like 'QH 50%'
					order by Try_Convert(bigint,SUBSTRING(a1.machine_no,4,len(a1.machine_no)-4))
						", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
				Summary = db.GetDataTableSQL("SCADA QH", query, 1200);
				ds = new DataSet();
				ds.Tables.Add(Summary);
				ds.Tables[0].TableName = "Weekly Summary";

				QH_ExportDataSetToExcelAppWeeklySummary(ds, fileWeeklySummary);
			}
		}

		private void GenerateMonthlyQuadHeadExcel(string fileMonthlySummary, DateTime datDate)
		{
			var ds = new DataSet();
			//string query = String.Format(@"Declare @datDate DateTime='{0}'
			//									Select a1.machine_no as Machine,
			//			concat(convert(varchar(max), ISNULL(a.dsp,0)), ' %') as dsp1, concat(convert(varchar(max), ISNULL(a.nsp,0)), ' %') as nsp1, 
			//			convert(decimal(24,2), case when a.ds is null then 0 else convert(decimal(24,6),a.ds)/60 end) as ds1, convert(decimal(24,2), case when a.ns is null then 0 else convert(decimal(24,6),a.ns)/60 end) as ns1,
			//			concat(convert(varchar(max), ISNULL(b.dsp,0)), ' %') as dsp2, concat(convert(varchar(max), ISNULL(b.nsp,0)), ' %') as nsp2,
			//			convert(decimal(24,2),case when b.ds is null then 0 else convert(decimal(24,6),b.ds)/60 end) as ds2, convert(decimal(24,2),case when b.ns is null then 0 else convert(decimal(24,6),b.ns)/60 end) as ns2,
			//			concat(convert(varchar(max), ISNULL(c.dsp,0)), ' %') as dsp3, concat(convert(varchar(max), ISNULL(c.nsp,0)), ' %') as nsp3,
			//			convert(decimal(24,2),case when c.ds is null then 0 else convert(decimal(24,6),c.ds)/60 end) as ds3, convert(decimal(24,2),case when c.ns is null then 0 else convert(decimal(24,6),c.ns)/60 end) as ns3,
			//			concat(convert(varchar(max), ISNULL(d.dsp,0)), ' %') as dsp4, concat(convert(varchar(max), ISNULL(d.nsp,0)), ' %') as nsp4,
			//			convert(decimal(24,2),case when d.ds is null then 0 else convert(decimal(24,6),d.ds)/60 end) as ds4, convert(decimal(24,2),case when d.ns is null then 0 else convert(decimal(24,6),d.ns)/60 end) as ns4,
			//			concat(convert(varchar(max), ISNULL(e.dsp,0)), ' %') as dsp5, concat(convert(varchar(max), ISNULL(e.nsp,0)), ' %') as nsp5,
			//			convert(decimal(24,2),case when e.ds is null then 0 else convert(decimal(24,6),e.ds)/60 end) as ds5, convert(decimal(24,2),case when e.ns is null then 0 else convert(decimal(24,6),e.ns)/60 end) as ns5,
			//			concat(convert(varchar(max), ISNULL(f.dsp,0)), ' %') as dsp6, concat(convert(varchar(max), ISNULL(f.nsp,0)), ' %') as nsp6,
			//			convert(decimal(24,2),case when f.ds is null then 0 else convert(decimal(24,6),f.ds)/60 end) as ds6, convert(decimal(24,2),case when f.ns is null then 0 else convert(decimal(24,6),f.ns)/60 end) as ns6,
			//			concat(convert(varchar(max), ISNULL(g.dsp,0)), ' %') as dsp7, concat(convert(varchar(max), ISNULL(g.nsp,0)), ' %') as nsp7,
			//			convert(decimal(24,2),case when g.ds is null then 0 else convert(decimal(24,6),g.ds)/60 end) as ds7, convert(decimal(24,2),case when g.ns is null then 0 else convert(decimal(24,6),g.ns)/60 end) as ns7,
			//			convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0)))/7/60) as avgds, 
			//			convert(decimal(24,2),COnvert(Decimal(24,6),(ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0)))/7/60) as avgns,
			//			concat(convert(varchar(max), (ISNULL(a.ds,0)+ISNULL(b.ds,0)+ISNULL(c.ds,0)+ISNULL(d.ds,0)+ISNULL(e.ds,0)+ISNULL(f.ds,0)+ISNULL(g.ds,0))/7*100/720), ' %') as avgdsp, 
			//			concat(convert(varchar(max), (ISNULL(a.ns,0)+ISNULL(b.ns,0)+ISNULL(c.ns,0)+ISNULL(d.ns,0)+ISNULL(e.ns,0)+ISNULL(f.ns,0)+ISNULL(g.ns,0))/7*100/720), ' %') as avgnsp,
			//			isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget
			//		from tblMachine a1
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-7,@datDate)) a on a1.machine_no=a.rig_no
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-6,@datDate)) b on a1.machine_no=b.rig_no
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-5,@datDate)) c on a1.machine_no=c.rig_no
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-4,@datDate)) d on a1.machine_no=d.rig_no
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-3,@datDate)) e on a1.machine_no=e.rig_no
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-2,@datDate)) f on a1.machine_no=f.rig_no
			//		left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,-1,@datDate)) g on a1.machine_no=g.rig_no
			//		left outer join
			//		(
			//		Select 
			//			b.machine_no, 
			//			sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
			//			sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
			//		from
			//		(
			//			Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
			//				DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
			//				b.finishdate as ActualfinishDate
			//			from TargetArcOn a
			//			outer apply 
			//			(
			//					Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@datDate)) b
			//			) b
			//			where a.Actual_Start between DateAdd(hour, 8, DateAdd(day, -7, @datDate)) and DateAdd(hour, 8, @datDate) or 
			//				b.finishdate between DateAdd(day, -7, @datDate) and @datDate
			//		) a
			//		inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
			//		group by b.Machine_no
			//		) a2 on a1.machine_no=a2.Machine_No
			//		where a1.machine_no not like 'QH 50%'
			//		order by Try_Convert(bigint,SUBSTRING(a1.machine_no,4,len(a1.machine_no)-4))
			//			", DateTime.Now.AddDays(0).ToString("yyyy-MM-dd"));
			Int64 daysinmonth = DateTime.DaysInMonth(datDate.AddMonths(-1).Year, datDate.AddMonths(-1).Month);
			string query = String.Format(@"Declare @datDate DateTime='{0}',
					@toDate DateTime,
					@bigDay bigint
			set @datDate = DATEADD(month, -1, @datDate)
			set @datDate = DATEFROMPARTS(year(@datDate), month(@datDate), 1)
			set @toDate = EOMONTH(@datDate)
			set @bigDay = day(EOMONTH(@datDate))
			Select a1.machine_no as Machine", datDate.AddDays(0).ToString("yyyy-MM-dd"));
			
			for (var i=0; i<daysinmonth; i++)
            {
				query += String.Format(@",
				concat(convert(varchar(max), ISNULL(b{0}.dsp, 0)), ' %') as dsp{0}, 
				concat(convert(varchar(max), ISNULL(b{0}.nsp, 0)), ' %') as nsp{0}, 
				convert(decimal(24, 2), case when b{0}.ds is null then 0 else convert(decimal(24, 6), b{0}.ds) / 60 end) as ds{0}, 
				convert(decimal(24, 2), case when b{0}.ns is null then 0 else convert(decimal(24, 6), b{0}.ns) / 60 end) as ns{0}", i+1);
            }

			query += @",
				convert(decimal(24,2),COnvert(Decimal(24,6),(0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ds, 0)", i + 1);
			}

			query += String.Format(@")/{0}/60)) as avgds", daysinmonth);

			query += @",
				convert(decimal(24,2),COnvert(Decimal(24,6),(0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ns, 0)", i + 1);
			}

			query += String.Format(@")/{0}/60)) as avgns", daysinmonth);

			query += @",
				concat(convert(varchar(max), (0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ds, 0)", i + 1);
			}

			query += String.Format(@")/{0}*100/720), ' %') as avgdsp", daysinmonth);

			query += @",
				concat(convert(varchar(max), (0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ns, 0)", i + 1);
			}

			query += String.Format(@")/{0}*100/720), ' %') as avgnsp", daysinmonth);

			query += @",
				isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget";

			query += @"
				from tblMachine a1";
			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"
				left outer join dbo.sfnGetWeldingProcessByMachine(DateAdd(day,{0},@datDate)) b{1} on a1.machine_no=b{1}.rig_no", i, i + 1);
			}
			query += @"
				left outer join
				(
					Select 
						b.machine_no, 
						sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
						sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
					from
					(
						Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
							DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
							b.finishdate as ActualfinishDate
						from TargetArcOn a
						outer apply 
						(
								Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@toDate)) b
						) b
						where a.Actual_Start between DateAdd(hour, 8, @datDate) and DateAdd(hour, 8, @toDate) or 
							b.finishdate between DateAdd(hour, 8, @datDate) and @toDate
					) a
					inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
					group by b.Machine_no
				) a2 on a1.machine_no=a2.Machine_No
				where a1.machine_no not like 'QH 50%'
				order by Try_Convert(bigint,SUBSTRING(a1.machine_no,4,len(a1.machine_no)-4))";
			var Summary = db.GetDataTableSQL("SCADA QH", query, 12000);
			ds = new DataSet();
			ds.Tables.Add(Summary);
			ds.Tables[0].TableName = "Monthly Summary";

			QH_ExportDataSetToExcelAppMonthlySummary(ds, fileMonthlySummary, daysinmonth, datDate);
		}
		private void GenerateMonthlyShvNShhExcel(string fileMonthlySummary, DateTime datDate)
		{
			var ds = new DataSet();
			Int64 daysinmonth = DateTime.DaysInMonth(datDate.AddMonths(-1).Year, datDate.AddMonths(-1).Month);
			string query = String.Format(@"Declare @datDate DateTime='{0}',
					@toDate DateTime,
					@bigDay bigint
			set @datDate = Convert(datetime, Convert(date, DATEADD(month, -1, @datDate)))
			set @datDate = DATEFROMPARTS(year(@datDate), month(@datDate), 1)
			set @toDate = EOMONTH(@datDate)
			set @bigDay = day(EOMONTH(@datDate))
			Select a1.machine_no as Machine", datDate.ToString("yyyy-MM-dd"));

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@",
				concat(convert(varchar(max), ISNULL(b{0}.dsp, 0)), ' %') as dsp{0}, 
				concat(convert(varchar(max), ISNULL(b{0}.nsp, 0)), ' %') as nsp{0}, 
				convert(decimal(24, 2), case when b{0}.ds is null then 0 else convert(decimal(24, 6), b{0}.ds) end) as ds{0}, 
				convert(decimal(24, 2), case when b{0}.ns is null then 0 else convert(decimal(24, 6), b{0}.ns)end) as ns{0}", i + 1);
			}

			query += @",
				convert(decimal(24,2),COnvert(Decimal(24,6),(0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ds, 0)", i + 1);
			}

			query += String.Format(@")/{0})) as avgds", daysinmonth);

			query += @",
				convert(decimal(24,2),COnvert(Decimal(24,6),(0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ns, 0)", i + 1);
			}

			query += String.Format(@")/{0})) as avgns", daysinmonth);

			query += @",
				concat(convert(varchar(max), (0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ds, 0)", i + 1);
			}

			query += String.Format(@")/{0}*100/12), ' %') as avgdsp", daysinmonth);

			query += @",
				concat(convert(varchar(max), (0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ns, 0)", i + 1);
			}

			query += String.Format(@")/{0}*100/12), ' %') as avgnsp", daysinmonth);

			query += @",
				--isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget
				0 as pipefinish, 0 as exceedtarget";

			query += @"
				from tblMachine a1";
			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"
				left outer join dbo.sfnRptWeldingProcessByMachine(DateAdd(day,{0},@datDate)) b{1} on a1.machine_no=b{1}.rig_no", i, i + 1);
			}
			query += @"
				/*left outer join
				(
					Select 
						b.machine_no, 
						sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
						sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
					from
					(
						Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
							DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
							b.finishdate as ActualfinishDate
						from TargetArcOn a
						outer apply 
						(
								Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@toDate)) b
						) b
						where a.Actual_Start between DateAdd(hour, 8, @datDate) and DateAdd(hour, 8, @toDate) or 
							b.finishdate between DateAdd(hour, 8, @datDate) and @toDate
					) a
					inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
					group by b.Machine_no
				) a2 on a1.machine_no=a2.Machine_No*/
				where a1.machine_no not like 'SHV50%' and a1.classification='shvnshh' and machine_no not in('SHV 00')
				order by Try_Convert(bigint,SUBSTRING(a1.machine_no,4,len(a1.machine_no)-4))";
			var Summary = db.GetDataTableSQL("SCADA SHV", query, 12000);
			ds = new DataSet();
			ds.Tables.Add(Summary);
			ds.Tables[0].TableName = "Monthly Summary";

			QH_ExportDataSetToExcelAppMonthlySummary(ds, fileMonthlySummary, daysinmonth, datDate);
		}
		private void GenerateMonthlyEndCladExcel(string fileMonthlySummary, DateTime datDate)
		{
			var ds = new DataSet();
			Int64 daysinmonth = DateTime.DaysInMonth(datDate.AddMonths(-1).Year, datDate.AddMonths(-1).Month);
			string query = String.Format(@"Declare @datDate DateTime='{0}',
					@toDate DateTime,
					@bigDay bigint
			set @datDate =  Convert(datetime, Convert(date, DATEADD(month, -1, @datDate)))
			set @datDate = DATEFROMPARTS(year(@datDate), month(@datDate), 1)
			set @toDate = EOMONTH(@datDate)
			set @bigDay = day(EOMONTH(@datDate))
			Select a1.machine_no as Machine", datDate.ToString("yyyy-MM-dd"));

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@",
				concat(convert(varchar(max), ISNULL(b{0}.dsp, 0)), ' %') as dsp{0}, 
				concat(convert(varchar(max), ISNULL(b{0}.nsp, 0)), ' %') as nsp{0}, 
				convert(decimal(24, 2), case when b{0}.ds is null then 0 else convert(decimal(24, 6), b{0}.ds) end) as ds{0}, 
				convert(decimal(24, 2), case when b{0}.ns is null then 0 else convert(decimal(24, 6), b{0}.ns)end) as ns{0}", i + 1);
			}

			query += @",
				convert(decimal(24,2),COnvert(Decimal(24,6),(0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ds, 0)", i + 1);
			}

			query += String.Format(@")/{0})) as avgds", daysinmonth);

			query += @",
				convert(decimal(24,2),COnvert(Decimal(24,6),(0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ns, 0)", i + 1);
			}

			query += String.Format(@")/{0})) as avgns", daysinmonth);

			query += @",
				concat(convert(varchar(max), (0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ds, 0)", i + 1);
			}

			query += String.Format(@")/{0}*100/12), ' %') as avgdsp", daysinmonth);

			query += @",
				concat(convert(varchar(max), (0";

			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"+ISNULL(b{0}.ns, 0)", i + 1);
			}

			query += String.Format(@")/{0}*100/12), ' %') as avgnsp", daysinmonth);

			query += @",
				--isnull(pipefinish,0) as pipefinish, ISNULL(exceedtarget,0) as exceedtarget
				0 as pipefinish, 0 as exceedtarget";

			query += @"
				from tblMachine a1";
			for (var i = 0; i < daysinmonth; i++)
			{
				query += String.Format(@"
				left outer join dbo.sfnRptWeldingProcessByMachine(DateAdd(day,{0},@datDate)) b{1} on a1.machine_no=b{1}.rig_no", i, i + 1);
			}
			query += @"
				/*left outer join
				(
					Select 
						b.machine_no, 
						sum(case when actualfinishdate is not null then 1 else 0 end) as pipefinish,
						sum(case when actualfinishdate is not null and finishTarget<ActualfinishDate then 1 else 0 end) as exceedtarget
					from
					(
						Select PipeNo as id_pipe, Machine as Machine_no, Actual_Start as startTarget, 
							DateAdd(hour, Convert(decimal(24,2),Target_Hours), Actual_Start) as finishTarget,
							b.finishdate as ActualfinishDate
						from TargetArcOn a
						outer apply 
						(
								Select * from sfnPipeOnMachineSatus(left(a.Machine,len(a.Machine)-1), a.pipeno, DateAdd(hour,8,@toDate)) b
						) b
						where a.Actual_Start between DateAdd(hour, 8, @datDate) and DateAdd(hour, 8, @toDate) or 
							b.finishdate between DateAdd(hour, 8, @datDate) and @toDate
					) a
					inner join tblMachine b on a.Machine_no=left(b.machine_no,len(a.Machine_no))
					group by b.Machine_no
				) a2 on a1.machine_no=a2.Machine_No*/
				where a1.machine_no not like 'SHV50%' and a1.classification='endclad'
				order by Try_Convert(bigint,SUBSTRING(a1.machine_no,4,len(a1.machine_no)-4))";
			var Summary = db.GetDataTableSQL("SCADA SHV", query, 12000);
			ds = new DataSet();
			ds.Tables.Add(Summary);
			ds.Tables[0].TableName = "Monthly Summary";

			QH_ExportDataSetToExcelAppMonthlySummary(ds, fileMonthlySummary, daysinmonth, datDate);
		}
		private void EC_ExportDataSetToExcelAppSummary(DataSet ds, string filename)
		{
			object misValue = System.Reflection.Missing.Value;
			string errDesc = "1";
			var app = new Application();
			errDesc = "2";
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				errDesc = "3";
				Workbooks excelworkBooks = app.Workbooks;
				errDesc = "4";
                Workbook excelworkBook = excelworkBooks.Add(misValue);
                errDesc = "5";
                Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;
                errDesc = "6";

                EC_SetSheetSummaryPipeTarget(excelSheet, ds);
                errDesc = "7";
				app.DisplayAlerts = false;

                excelSheet = (Worksheet)excelworkBook.Worksheets.Add(After: excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                EC_SetSheetSummaryMachine(excelSheet, ds.Tables[1]);


                excelSheet = (Worksheet)excelworkBook.Worksheets.Add(After: excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                EC_SetSheetSummaryFinish(excelSheet, ds.Tables[2]);

                //				excelworkBook.SaveAs(filename.Replace(@"\", @"\\"));
                //				excelworkBook.SaveAs(filename.Replace(@"\", @"/"), XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                excelworkBook.SaveCopyAs(filename.Replace(@"\",@"/"));
				errDesc = "8";
				app.Quit();
			}
			catch (Exception f)
			{
				app.Quit();
				throw new Exception(String.Format(@"Error on generate report {0} {1}. {2}", filename,  errDesc, f.Message.ToString()));
			}
		}
		private void SHV_ExportDataSetToExcelAppSummary(DataSet ds, string filename)
		{
			object misValue = System.Reflection.Missing.Value;
			var app = new Microsoft.Office.Interop.Excel.Application();
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				Workbooks excelworkBooks = app.Workbooks;
				Workbook excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
				Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;

				SHV_SetSheetSummaryPipeTarget(excelSheet, ds);
				app.DisplayAlerts = false;

				excelSheet = (Worksheet)excelworkBook.Worksheets.Add(After: excelworkBook.Sheets[excelworkBook.Sheets.Count]);
				SHV_SetSheetSummaryMachine(excelSheet, ds.Tables[1]);


				excelSheet = (Worksheet)excelworkBook.Worksheets.Add(After: excelworkBook.Sheets[excelworkBook.Sheets.Count]);
				SHV_SetSheetSummaryFinish(excelSheet, ds.Tables[2]);

				excelworkBook.SaveCopyAs(filename.Replace(@"\", @"/"));
				app.Quit();
			}
			catch (COMException e)
			{
				app.Quit();
			}
		}
		private void EC_ExportDataSetToExcelAppWeeklySummary(DataSet ds, string filename)
		{
			object misValue = System.Reflection.Missing.Value;
			var app = new Microsoft.Office.Interop.Excel.Application();
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				Workbooks excelworkBooks = app.Workbooks;
				Workbook excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
				Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;

				//excelworkBook.Worksheets.Add("ConditionalFormatting");
				EC_SetSheetWeeklySummary(excelSheet, ds);
				app.DisplayAlerts = false;

				//app.ActiveWorkbook.Sheets[1].Select(misValue);
				excelworkBook.SaveCopyAs(filename.Replace(@"\", @"/"));
				app.Quit();
			}
			catch (COMException e)
			{
				app.Quit();
			}
		}
		private void SHV_ExportDataSetToExcelAppWeeklySummary(DataSet ds, string filename)
		{
			object misValue = System.Reflection.Missing.Value;
			var app = new Microsoft.Office.Interop.Excel.Application();
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				Workbooks excelworkBooks = app.Workbooks;
				Workbook excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
				Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;

				//excelworkBook.Worksheets.Add("ConditionalFormatting");
				SHV_SetSheetWeeklySummary(excelSheet, ds);
				app.DisplayAlerts = false;

				//app.ActiveWorkbook.Sheets[1].Select(misValue);
				excelworkBook.SaveCopyAs(filename.Replace(@"\", @"/"));
				app.Quit();
			}
			catch (COMException e)
			{
				app.Quit();
			}
		}
		private void QH_ExportDataSetToExcelAppSummary(DataSet ds, string filename)
		{
			object misValue = System.Reflection.Missing.Value;
			var app = new Microsoft.Office.Interop.Excel.Application();
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				Workbooks excelworkBooks = app.Workbooks;
				Workbook excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
				Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;

				QH_SetSheetSummaryPipeTarget(excelSheet, ds);
				app.DisplayAlerts = false;

				excelSheet = (Worksheet)excelworkBook.Worksheets.Add(After: excelworkBook.Sheets[excelworkBook.Sheets.Count]);
				QH_SetSheetSummaryMachine(excelSheet, ds.Tables[1]);


				excelSheet = (Worksheet)excelworkBook.Worksheets.Add(After: excelworkBook.Sheets[excelworkBook.Sheets.Count]);
				QH_SetSheetSummaryFinish(excelSheet, ds.Tables[2]);

				excelworkBook.SaveCopyAs(filename.Replace(@"\", @"/"));
				app.Quit();
			}
			catch (COMException e)
			{
				app.Quit();
			}
		}
		private void QH_ExportDataSetToExcelAppWeeklySummary(DataSet ds, string filename)
		{
			object misValue = System.Reflection.Missing.Value;
			var app = new Microsoft.Office.Interop.Excel.Application();
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				Workbooks excelworkBooks = app.Workbooks;
				Workbook excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
				Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;

				//excelworkBook.Worksheets.Add("ConditionalFormatting");
				QH_SetSheetWeeklySummary(excelSheet, ds);
				app.DisplayAlerts = false;

				//app.ActiveWorkbook.Sheets[1].Select(misValue);
				excelworkBook.SaveCopyAs(filename.Replace(@"\", @"/"));
				app.Quit();
			}
			catch (COMException e)
			{
				app.Quit();
			}
		}
		private void QH_ExportDataSetToExcelAppMonthlySummary(DataSet ds, string filename, Int64 daysinmonth, DateTime datDate)
		{
			object misValue = System.Reflection.Missing.Value;
			var app = new Microsoft.Office.Interop.Excel.Application();
			//var ProcessID = excel.Hwnd;
			try
			{
				app.ODBCTimeout = 0;
				app.Visible = false;
				app.ScreenUpdating = false;
				Workbooks excelworkBooks = app.Workbooks;
				Workbook excelworkBook = (Workbook)(excelworkBooks.Add(misValue));
				Worksheet excelSheet = (Worksheet)excelworkBook.ActiveSheet;

				//excelworkBook.Worksheets.Add("ConditionalFormatting");
				QH_SetSheetMonthlySummary(excelSheet, ds, daysinmonth, datDate);
				app.DisplayAlerts = false;

				//app.ActiveWorkbook.Sheets[1].Select(misValue);
				excelworkBook.SaveCopyAs(filename.Replace(@"\", @"/"));
				app.Quit();
			}
			catch (COMException e)
			{
				app.Quit();
			}
		}
		private void EC_SetSheetSummaryPipeTarget(Worksheet excelSheet, DataSet ds)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8" };
			List<string> colHeader1 = new List<string> { "", "", "Actual Start", "Target Finish", "Target (Hours)", "Actual Arc On Time (Hours)", "Actual Completion (Hours)", "Arc On Time (%)" };
			List<string> colHeaderName = new List<string> { "id_pipe", "machine_no", "startTarget", "finishTarget", "Target_Time", "Actual_Arc_Time", "Actual_Time", "ArcOnTime" };
			List<string> colHeaderType = new List<string> { "text", "text", "date", "date", "number", "number", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 9;
			int FirstRow = 9;
			excelSheet.Cells[1, 1] = "Legend :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[2, 1].Interior.Color = Color.Red;
			excelSheet.Cells[3, 1].Interior.Color = Color.Yellow;
			excelSheet.Cells[5, 1] = "Report Date :";
			excelSheet.Cells[5, 1].Font.Bold = true;
			excelSheet.Cells[5, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[5, 2].NumberFormat = "dd MMM yyyy";
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production over the Target";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[3, 2], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production exceeds Target within 24 hours";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[6, colHeader1.Count], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 3], excelSheet.Cells[rowcount + FirstRow, 4]];
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 5], excelSheet.Cells[rowcount + FirstRow, 8]];
			excelCellrange.NumberFormat = "#,##0.00";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[FirstRow - 1, i], excelSheet.Cells[FirstRow - 1, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.NumberFormat = "@";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex - 1, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Item No";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 25;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 1, 2]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine";
			excelCellrange.ColumnWidth = 25;
			//			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 3], excelSheet.Cells[rowIndex - 2, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production";
			excelCellrange.ColumnWidth = 36;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 5], excelSheet.Cells[rowIndex - 2, 8]];
			excelCellrange.Merge();
			excelCellrange.Value = "Pipes Complete";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 28;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			if (rowcount > 0)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
				excelSheet.EnableFormatConditionsCalculation = true;
				string strFormula = String.Format(@"=(($D10-DATE({0}))*24)-8<0", DateTime.Now.ToString("yyyy,M,dd"));
				FormatCondition cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Red;
				cond.Font.Color = Color.White;
				strFormula = String.Format(@"=AND(24>(($D10-DATE({0}))*24)-8,(($D10-DATE({0}))*24)-8>0)", DateTime.Now.ToString("yyyy,M,dd"));
				cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Yellow;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = FirstRow;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void SHV_SetSheetSummaryPipeTarget(Worksheet excelSheet, DataSet ds)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8" };
			List<string> colHeader1 = new List<string> { "", "", "Actual Start", "Target Finish", "Target (Hours)", "Actual Arc On Time (Hours)", "Actual Completion (Hours)", "Arc On Time (%)" };
			List<string> colHeaderName = new List<string> { "id_pipe", "machine_no", "startTarget", "finishTarget", "Target_Time", "Actual_Arc_Time", "Actual_Time", "ArcOnTime" };
			List<string> colHeaderType = new List<string> { "text", "text", "date", "date", "number", "number", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 9;
			int FirstRow = 9;
			excelSheet.Cells[1, 1] = "Legend :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[2, 1].Interior.Color = Color.Red;
			excelSheet.Cells[3, 1].Interior.Color = Color.Yellow;
			excelSheet.Cells[5, 1] = "Report Date :";
			excelSheet.Cells[5, 1].Font.Bold = true;
			excelSheet.Cells[5, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[5, 2].NumberFormat = "dd MMM yyyy";
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production over the Target";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[3, 2], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production exceeds Target within 24 hours";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[6, colHeader1.Count], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 3], excelSheet.Cells[rowcount + FirstRow, 4]];
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 5], excelSheet.Cells[rowcount + FirstRow, 8]];
			excelCellrange.NumberFormat = "#,##0.00";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[FirstRow - 1, i], excelSheet.Cells[FirstRow - 1, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.NumberFormat = "@";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex - 1, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Item No";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 25;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 1, 2]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine";
			excelCellrange.ColumnWidth = 25;
			//			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 3], excelSheet.Cells[rowIndex - 2, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production";
			excelCellrange.ColumnWidth = 36;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 5], excelSheet.Cells[rowIndex - 2, 8]];
			excelCellrange.Merge();
			excelCellrange.Value = "Items Complete";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 28;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			if (rowcount > 0)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
				excelSheet.EnableFormatConditionsCalculation = true;
				string strFormula = String.Format(@"=(($D10-DATE({0}))*24)-8<0", DateTime.Now.ToString("yyyy,M,dd"));
				FormatCondition cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Red;
				cond.Font.Color = Color.White;
				strFormula = String.Format(@"=AND(24>(($D10-DATE({0}))*24)-8,(($D10-DATE({0}))*24)-8>0)", DateTime.Now.ToString("yyyy,M,dd"));
				cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Yellow;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = FirstRow;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void QH_SetSheetSummaryPipeTarget(Worksheet excelSheet, DataSet ds)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8" };
			List<string> colHeader1 = new List<string> { "", "", "Actual Start", "Target Finish", "Target (Hours)", "Actual Arc On Time (Hours)", "Actual Completion (Hours)", "Arc On Time (%)" };
			List<string> colHeaderName = new List<string> { "machine_no", "id_pipe", "startTarget", "finishTarget", "Target_Time", "Actual_Arc_Time", "Actual_Time", "ArcOnTime" };
			List<string> colHeaderType = new List<string> { "text", "text", "date", "date", "number", "number", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 9;
			int FirstRow = 9;
			excelSheet.Cells[1, 1] = "Legend :";
			excelSheet.Cells[1, 1].Font.Bold=true;
			excelSheet.Cells[2, 1].Interior.Color = Color.Red;
			excelSheet.Cells[3, 1].Interior.Color = Color.Yellow;
			excelSheet.Cells[5, 1] = "Report Date :";
			excelSheet.Cells[5, 1].Font.Bold = true;
			excelSheet.Cells[5, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[5, 2].NumberFormat = "dd MMM yyyy"; 
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production over the Target";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[3, 2], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production exceeds Target within 24 hours";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[6, colHeader1.Count], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 3], excelSheet.Cells[rowcount + FirstRow, 4]];
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 5], excelSheet.Cells[rowcount + FirstRow, 8]];
			excelCellrange.NumberFormat = "#,##0.00";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[FirstRow-1, i], excelSheet.Cells[FirstRow-1, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.NumberFormat = "@";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex - 1, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 15;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 1, 2]];
			excelCellrange.Merge();
			excelCellrange.Value = "Pipe No";
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 3], excelSheet.Cells[rowIndex - 2, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production";
			excelCellrange.ColumnWidth = 36;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 5], excelSheet.Cells[rowIndex - 2, 8]];
			excelCellrange.Merge();
			excelCellrange.Value = "Pipes Complete";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 28;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			if (rowcount > 0)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
				excelSheet.EnableFormatConditionsCalculation = true;
				string strFormula = String.Format(@"=(($D10-DATE({0}))*24)-8<0", DateTime.Now.ToString("yyyy,M,dd"));
				FormatCondition cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Red;
				cond.Font.Color = Color.White;
				strFormula = String.Format(@"=AND(24>(($D10-DATE({0}))*24)-8,(($D10-DATE({0}))*24)-8>0)", DateTime.Now.ToString("yyyy,M,dd"));
				cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Yellow;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = FirstRow;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void EC_SetSheetWeeklySummary(Worksheet excelSheet, DataSet ds)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35" };
			List<string> colHeader1 = new List<string> { "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Pipe Finish", "Exceed Target" };
			List<string> colHeaderName = new List<string> { "Machine", "ds1", "dsp1", "ns1", "nsp1", "ds2", "dsp2", "ns2", "nsp2", "ds3", "dsp3", "ns3", "nsp3", "ds4", "dsp4", "ns4", "nsp4", "ds5", "dsp5", "ns5", "nsp5", "ds6", "dsp6", "ns6", "nsp6", "ds7", "dsp7", "ns7", "nsp7", "avgds", "avgdsp", "avgns", "avgnsp", "pipefinish", "exceedtarget" };
			List<string> colHeaderType = new List<string> { "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 4;
			int FirstRow = 4;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			DateTime datDate = DateTime.Now.AddDays(-32);
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[3, 1]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 10;
			excelCellrange.Value = "Machine";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 2], excelSheet.Cells[1, 5]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Monday (", datDate.AddDays(-7).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 6], excelSheet.Cells[1, 9]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Tuesday (", datDate.AddDays(-6).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 10], excelSheet.Cells[1, 13]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Wednesday (", datDate.AddDays(-5).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 14], excelSheet.Cells[1, 17]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Thursday (", datDate.AddDays(-4).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 18], excelSheet.Cells[1, 21]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Friday (", datDate.AddDays(-3).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 22], excelSheet.Cells[1, 25]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Saturday (", datDate.AddDays(-2).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 26], excelSheet.Cells[1, 29]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Sunday (", datDate.AddDays(-1).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 30], excelSheet.Cells[1, 33]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Average";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 34], excelSheet.Cells[1, 35]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Pipe Production";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			for (var i = 2; i <= colHeader1.Count - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i + 1]];
				excelCellrange.Merge();
				i++;
			}
			for (var i = 2; i <= columncount - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i], excelSheet.Cells[3, i]];
				excelCellrange.Value = "Arc On Time (Hours)";
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i + 1], excelSheet.Cells[3, i + 1]];
				excelCellrange.Value = "Percentage";
				i++;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 34], excelSheet.Cells[3, 34]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 35], excelSheet.Cells[3, 35]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = 4;
			excelSheet.Application.ActiveWindow.SplitColumn = 1;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void SHV_SetSheetWeeklySummary(Worksheet excelSheet, DataSet ds)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35" };
			List<string> colHeader1 = new List<string> { "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Pipe Finish", "Exceed Target" };
			List<string> colHeaderName = new List<string> { "Machine", "ds1", "dsp1", "ns1", "nsp1", "ds2", "dsp2", "ns2", "nsp2", "ds3", "dsp3", "ns3", "nsp3", "ds4", "dsp4", "ns4", "nsp4", "ds5", "dsp5", "ns5", "nsp5", "ds6", "dsp6", "ns6", "nsp6", "ds7", "dsp7", "ns7", "nsp7", "avgds", "avgdsp", "avgns", "avgnsp", "pipefinish", "exceedtarget" };
			List<string> colHeaderType = new List<string> { "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 4;
			int FirstRow = 4;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			DateTime datDate = DateTime.Now.AddDays(-32);
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[3, 1]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 10;
			excelCellrange.Value = "Machine";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 2], excelSheet.Cells[1, 5]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Monday (", datDate.AddDays(-7).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 6], excelSheet.Cells[1, 9]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Tuesday (", datDate.AddDays(-6).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 10], excelSheet.Cells[1, 13]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Wednesday (", datDate.AddDays(-5).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 14], excelSheet.Cells[1, 17]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Thursday (", datDate.AddDays(-4).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 18], excelSheet.Cells[1, 21]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Friday (", datDate.AddDays(-3).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 22], excelSheet.Cells[1, 25]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Saturday (", datDate.AddDays(-2).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 26], excelSheet.Cells[1, 29]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Sunday (", datDate.AddDays(-1).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 30], excelSheet.Cells[1, 33]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Average";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 34], excelSheet.Cells[1, 35]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Pipe Production";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			for (var i = 2; i <= colHeader1.Count - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i + 1]];
				excelCellrange.Merge();
				i++;
			}
			for (var i = 2; i <= columncount - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i], excelSheet.Cells[3, i]];
				excelCellrange.Value = "Arc On Time (Hours)";
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i + 1], excelSheet.Cells[3, i + 1]];
				excelCellrange.Value = "Percentage";
				i++;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 34], excelSheet.Cells[3, 34]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 35], excelSheet.Cells[3, 35]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = 4;
			excelSheet.Application.ActiveWindow.SplitColumn = 1;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void QH_SetSheetWeeklySummary(Worksheet excelSheet, DataSet ds)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35" };
			List<string> colHeader1 = new List<string> { "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Pipe Finish", "Exceed Target" };
			List<string> colHeaderName = new List<string> { "Machine", "ds1", "dsp1", "ns1", "nsp1", "ds2", "dsp2", "ns2", "nsp2", "ds3", "dsp3", "ns3", "nsp3", "ds4", "dsp4", "ns4", "nsp4", "ds5", "dsp5", "ns5", "nsp5", "ds6", "dsp6", "ns6", "nsp6", "ds7", "dsp7", "ns7", "nsp7", "avgds", "avgdsp", "avgns", "avgnsp", "pipefinish", "exceedtarget" };
			List<string> colHeaderType = new List<string> { "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 4;
			int FirstRow = 4;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			DateTime datDate = DateTime.Now.AddDays(0);
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[3, 1]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 10;
			excelCellrange.Value = "Machine";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 2], excelSheet.Cells[1, 5]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Monday (", datDate.AddDays(-7).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 6], excelSheet.Cells[1, 9]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Tuesday (", datDate.AddDays(-6).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 10], excelSheet.Cells[1, 13]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Wednesday (", datDate.AddDays(-5).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 14], excelSheet.Cells[1, 17]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Thursday (", datDate.AddDays(-4).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 18], excelSheet.Cells[1, 21]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Friday (", datDate.AddDays(-3).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 22], excelSheet.Cells[1, 25]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Saturday (", datDate.AddDays(-2).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 26], excelSheet.Cells[1, 29]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = String.Concat("Sunday (", datDate.AddDays(-1).ToString("dd MMM yyyy"), ")");
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 30], excelSheet.Cells[1, 33]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Average";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 34], excelSheet.Cells[1, 35]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Pipe Production";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			for (var i = 2; i <= colHeader1.Count - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i + 1]];
				excelCellrange.Merge();
				i++;
			}
			for (var i = 2; i <= columncount - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i], excelSheet.Cells[3, i]];
				excelCellrange.Value = "Arc On Time (Hours)";
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i + 1], excelSheet.Cells[3, i + 1]];
				excelCellrange.Value = "Percentage";
				i++;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 34], excelSheet.Cells[3, 34]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 35], excelSheet.Cells[3, 35]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = 4;
			excelSheet.Application.ActiveWindow.SplitColumn = 1;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void QH_SetSheetMonthlySummary(Worksheet excelSheet, DataSet ds, Int64 daysinmonth, DateTime datDate)
		{
//			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35" };
			List<string> colHeader = new List<string>();
			for (var i=1; i<=(daysinmonth*4)+7; i++)
            {
				colHeader.Add(i.ToString());
            }
//			List<string> colHeader1 = new List<string> { "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Day Shift", "", "Night Shift", "", "Pipe Finish", "Exceed Target" };
			List<string> colHeader1 = new List<string> {""};
			for (var i = 1; i <= daysinmonth+1; i++)
			{
				colHeader1.AddRange(new List<string> { "Day Shift", "", "Night Shift", ""});
			}
			colHeader1.AddRange(new List<string> { "Pipe Finish", "Exceed Target" });
//			List<string> colHeaderName = new List<string> { "Machine", "ds1", "dsp1", "ns1", "nsp1", "ds2", "dsp2", "ns2", "nsp2", "ds3", "dsp3", "ns3", "nsp3", "ds4", "dsp4", "ns4", "nsp4", "ds5", "dsp5", "ns5", "nsp5", "ds6", "dsp6", "ns6", "nsp6", "ds7", "dsp7", "ns7", "nsp7", "avgds", "avgdsp", "avgns", "avgnsp", "pipefinish", "exceedtarget" };
			List<string> colHeaderName = new List<string> { "Machine"};
			for (var i = 1; i <= daysinmonth; i++)
			{
				colHeaderName.AddRange(new List<string> { "ds"+i, "dsp"+i, "ns"+i, "nsp"+i });
			}
			colHeaderName.AddRange(new List<string> { "avgds", "avgdsp", "avgns", "avgnsp", "pipefinish", "exceedtarget" });
//			List<string> colHeaderType = new List<string> { "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "text", "number", "number" };
			List<string> colHeaderType = new List<string> { "text" };
			for (var i = 1; i <= daysinmonth+1; i++)
			{
				colHeaderType.AddRange(new List<string> { "number", "text", "number", "text"});
			}
			colHeaderType.AddRange(new List<string> { "number", "number" });

			Range excelCellrange;
			excelSheet.Name = ds.Tables[0].TableName;
			var rowcount = ds.Tables[0].Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 4;
			int FirstRow = 4;
			var obj = ToObjectWithHeader(ds.Tables[0], colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			//DateTime datDate = DateTime.Now.AddDays(0);
			datDate = new DateTime((int)datDate.AddMonths(-1).Year, (int)datDate.AddMonths(-1).Month, 1); 
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[3, 1]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 10;
			excelCellrange.Value = "Machine";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			for (var i = 0; i < daysinmonth; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[1, (i*4)+2], excelSheet.Cells[1, (i * 4) +5]];
				excelCellrange.Merge();
				excelCellrange.ColumnWidth = 14;
				excelCellrange.Value = String.Concat(datDate.AddDays(i).ToString("dddd"), " (", datDate.AddDays(i).ToString("dd MMM yyyy"), ")");
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, (daysinmonth * 4) + 2], excelSheet.Cells[1, (daysinmonth * 4) + 5]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Average";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[1, (daysinmonth * 4) + 6], excelSheet.Cells[1, (daysinmonth * 4) + 7]];
			excelCellrange.Merge();
			excelCellrange.ColumnWidth = 14;
			excelCellrange.Value = "Pipe Production";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 3, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			for (var i = 2; i <= colHeader1.Count - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[2, i], excelSheet.Cells[2, i + 1]];
				excelCellrange.Merge();
				i++;
			}
			for (var i = 2; i <= columncount - 2; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i], excelSheet.Cells[3, i]];
				excelCellrange.Value = "Arc On Time (Hours)";
				excelCellrange = excelSheet.Range[excelSheet.Cells[3, i + 1], excelSheet.Cells[3, i + 1]];
				excelCellrange.Value = "Percentage";
				i++;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, columncount-1], excelSheet.Cells[3, columncount-1]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, columncount], excelSheet.Cells[3, columncount]];
			excelCellrange.Merge();
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = 4;
			excelSheet.Application.ActiveWindow.SplitColumn = 1;
			excelSheet.Application.ActiveWindow.FreezePanes = true;

		}
		private void EC_SetSheetSummaryMachine(Microsoft.Office.Interop.Excel.Worksheet excelSheet, System.Data.DataTable dt)
		{
			List<string> colHeader = new List<string> { "", "", "Day Shift", "", "", "Night Shift", "" };
			List<string> colHeaderName = new List<string> { "Machine", "welder_ds", "ds", "dsp", "welder_ns", "ns", "nsp" };
			List<string> colHeaderType = new List<string> { "text", "text", "number", "text", "text", "number", "text" };

			Range excelCellrange;
			excelSheet.Name = dt.TableName;
			var rowcount = dt.Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 5;
			int FirstRow = 5;
			excelSheet.Cells[1, 1] = "Report Date :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[1, 2].NumberFormat = "dd MMM yyyy";
			excelSheet.Cells[1, 2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			var obj = ToObjectWithHeader(dt, colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine"; excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 15;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 2, 7]];
			excelCellrange.Merge();
			excelCellrange.Value = "Arc On Time";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 1, 2], excelSheet.Cells[rowIndex - 1, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Day Shift";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 1, 5], excelSheet.Cells[rowIndex - 1, 7]];
			excelCellrange.Merge();
			excelCellrange.Value = "Night Shift";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 2], excelSheet.Cells[rowIndex, 2]];
			excelCellrange.Value = "Welder ID";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 3], excelSheet.Cells[rowIndex, 3]];
			excelCellrange.Value = "Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 4], excelSheet.Cells[rowIndex, 4]];
			excelCellrange.Value = "Percentage";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 5], excelSheet.Cells[rowIndex, 5]];
			excelCellrange.Value = "Welder ID";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 6], excelSheet.Cells[rowIndex, 6]];
			excelCellrange.Value = "Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 7], excelSheet.Cells[rowIndex, 7]];
			excelCellrange.Value = "Percentage";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, 1]];
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount + FirstRow, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Font.Bold = true;
		}
		private void SHV_SetSheetSummaryMachine(Microsoft.Office.Interop.Excel.Worksheet excelSheet, System.Data.DataTable dt)
		{
			List<string> colHeader = new List<string> { "", "", "", "Day Shift", "", "", "", "Night Shift", "" };
			List<string> colHeaderName = new List<string> { "Machine", "dswh", "welder_ds", "ds", "dsp", "nswh", "welder_ns", "ns", "nsp" };
			List<string> colHeaderType = new List<string> { "text", "text", "text", "number", "text", "text", "text", "number", "text" };

			Range excelCellrange;
			excelSheet.Name = dt.TableName;
			var rowcount = dt.Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 5;
			int FirstRow = 5;
			excelSheet.Cells[1, 1] = "Report Date :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[1, 2].NumberFormat = "dd MMM yyyy";
			excelSheet.Cells[1, 2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			var obj = ToObjectWithHeader(dt, colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine"; excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 15;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 2, 9]];
			excelCellrange.Merge();
			excelCellrange.Value = "Arc On Time";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 1, 2], excelSheet.Cells[rowIndex - 1, 5]];
			excelCellrange.Merge();
			excelCellrange.Value = "Day Shift";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 1, 6], excelSheet.Cells[rowIndex - 1, 9]];
			excelCellrange.Merge();
			excelCellrange.Value = "Night Shift";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 2], excelSheet.Cells[rowIndex, 2]];
			excelCellrange.Value = "Working Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 3], excelSheet.Cells[rowIndex, 3]];
			excelCellrange.Value = "Welder ID";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 4], excelSheet.Cells[rowIndex, 4]];
			excelCellrange.Value = "Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 5], excelSheet.Cells[rowIndex, 5]];
			excelCellrange.Value = "Percentage";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 6], excelSheet.Cells[rowIndex, 6]];
			excelCellrange.Value = "Working Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 7], excelSheet.Cells[rowIndex, 7]];
			excelCellrange.Value = "Welder ID";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 8], excelSheet.Cells[rowIndex, 8]];
			excelCellrange.Value = "Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 9], excelSheet.Cells[rowIndex, 9]];
			excelCellrange.Value = "Percentage";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, 1]];
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount + FirstRow, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Font.Bold = true;
			excelSheet.Application.ActiveWindow.SplitRow = FirstRow;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;
		}
		private void QH_SetSheetSummaryMachine(Microsoft.Office.Interop.Excel.Worksheet excelSheet, System.Data.DataTable dt)
		{
			List<string> colHeader = new List<string> { "", "Day Shift", "", "Night Shift", "" };
			List<string> colHeaderName = new List<string> { "Machine", "ds", "dsp", "ns", "nsp" };
			List<string> colHeaderType = new List<string> { "text", "number", "text", "number", "text" };

			Range excelCellrange;
			excelSheet.Name = dt.TableName;
			var rowcount = dt.Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 5;
			int FirstRow = 5;
			excelSheet.Cells[1, 1] = "Report Date :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[1, 2].NumberFormat = "dd MMM yyyy";
			excelSheet.Cells[1, 2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
			var obj = ToObjectWithHeader(dt, colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine"; excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 15;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 2, 5]];
			excelCellrange.Merge();
			excelCellrange.Value = "Arc On Time";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 1, 2], excelSheet.Cells[rowIndex - 1, 3]];
			excelCellrange.Merge();
			excelCellrange.Value = "Day Shift";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 1, 4], excelSheet.Cells[rowIndex - 1, 5]];
			excelCellrange.Merge();
			excelCellrange.Value = "Night Shift";
			excelCellrange.ColumnWidth = 18;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignRight;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 2], excelSheet.Cells[rowIndex, 2]];
			excelCellrange.Value = "Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 3], excelSheet.Cells[rowIndex, 3]];
			excelCellrange.Value = "Percentage";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 4], excelSheet.Cells[rowIndex, 4]];
			excelCellrange.Value = "Hours";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 5], excelSheet.Cells[rowIndex, 5]];
			excelCellrange.Value = "Percentage";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, 1]];
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount + FirstRow, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Font.Bold = true;
		}
		private void EC_SetSheetSummaryFinish(Microsoft.Office.Interop.Excel.Worksheet excelSheet, System.Data.DataTable dt)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
			List<string> colHeader1 = new List<string> { "", "", "Actual Start", "Target Finish", "Target (Hours)", "Actual Completion (Hours)", "Actual Arc On Time (Hours)", "Arc On Time (%)", "Actual Eff (%)" };
			List<string> colHeaderName = new List<string> { "id_pipe", "machine_no", "startTarget", "finishTarget", "Target_Time", "Actual_Arc_Time", "Actual_Time", "ArcOnTime", "ArcEff" };
			List<string> colHeaderType = new List<string> { "text", "text", "date", "date", "number", "number", "number", "number", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = dt.TableName;
			var rowcount = dt.Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 9;
			int FirstRow = 9;
			excelSheet.Cells[1, 1] = "Legend :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[2, 1].Interior.Color = Color.Red;
			excelSheet.Cells[3, 1].Interior.Color = Color.Yellow;
			excelSheet.Cells[5, 1] = "Report Date :";
			excelSheet.Cells[5, 1].Font.Bold = true;
			excelSheet.Cells[5, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[5, 2].NumberFormat = "dd MMM yyyy";
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Arc exceeds Target (Hours)";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[3, 2], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Actual Completion exceeds Target (Hours)";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[6, colHeader1.Count], excelSheet.Cells[7, colHeader1.Count]];
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			var obj = ToObjectWithHeader(dt, colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 3], excelSheet.Cells[rowcount + FirstRow, 4]];
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 5], excelSheet.Cells[rowcount + FirstRow, 8]];
			excelCellrange.NumberFormat = "#,##0.00";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[FirstRow - 1, i], excelSheet.Cells[FirstRow - 1, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.NumberFormat = "@";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex - 1, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Item No";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 25;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 1, 2]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine";
			excelCellrange.ColumnWidth = 25;
			//			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 3], excelSheet.Cells[rowIndex - 2, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production";
			excelCellrange.ColumnWidth = 36;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 5], excelSheet.Cells[rowIndex - 2, columncount]];
			excelCellrange.Merge();
			excelCellrange.Value = "Items Complete";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 28;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			if (rowcount > 0)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
				string strFormula = String.Format(@"=$G10>$E10");
				FormatCondition cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Red;
				excelSheet.EnableFormatConditionsCalculation = true;
				strFormula = String.Format(@"=$F10>$E10");
				cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Yellow;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = 7;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;
		}
		private void SHV_SetSheetSummaryFinish(Microsoft.Office.Interop.Excel.Worksheet excelSheet, System.Data.DataTable dt)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
			List<string> colHeader1 = new List<string> { "", "", "Actual Start", "Target Finish", "Target (Hours)", "Actual Completion (Hours)", "Actual Arc On Time (Hours)", "Arc On Time (%)", "Actual Eff (%)" };
			List<string> colHeaderName = new List<string> { "id_pipe", "machine_no", "startTarget", "finishTarget", "Target_Time", "Actual_Arc_Time", "Actual_Time", "ArcOnTime", "ArcEff" };
			List<string> colHeaderType = new List<string> { "text", "text", "date", "date", "number", "number", "number", "number", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = dt.TableName;
			var rowcount = dt.Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 9;
			int FirstRow = 9;
			excelSheet.Cells[1, 1] = "Legend :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[2, 1].Interior.Color = Color.Red;
			excelSheet.Cells[3, 1].Interior.Color = Color.Yellow;
			excelSheet.Cells[5, 1] = "Report Date :";
			excelSheet.Cells[5, 1].Font.Bold = true;
			excelSheet.Cells[5, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[5, 2].NumberFormat = "dd MMM yyyy";
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Arc exceeds Target (Hours)";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[3, 2], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Actual Completion exceeds Target (Hours)";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[6, colHeader1.Count], excelSheet.Cells[7, colHeader1.Count]];
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			var obj = ToObjectWithHeader(dt, colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 3], excelSheet.Cells[rowcount + FirstRow, 4]];
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 5], excelSheet.Cells[rowcount + FirstRow, 8]];
			excelCellrange.NumberFormat = "#,##0.00";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[FirstRow - 1, i], excelSheet.Cells[FirstRow - 1, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.NumberFormat = "@";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex - 1, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Item No";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 25;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 1, 2]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine";
			excelCellrange.ColumnWidth = 25;
			//			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 3], excelSheet.Cells[rowIndex - 2, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production";
			excelCellrange.ColumnWidth = 36;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 5], excelSheet.Cells[rowIndex - 2, columncount]];
			excelCellrange.Merge();
			excelCellrange.Value = "Items Complete";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 28;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			if (rowcount > 0)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
				string strFormula = String.Format(@"=$G10>$E10");
				FormatCondition cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Red;
				cond.Font.Color = Color.White;
				excelSheet.EnableFormatConditionsCalculation = true;
				strFormula = String.Format(@"=$F10>$E10");
				cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Yellow;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = FirstRow;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;
		}
		private void QH_SetSheetSummaryFinish(Microsoft.Office.Interop.Excel.Worksheet excelSheet, System.Data.DataTable dt)
		{
			List<string> colHeader = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
			List<string> colHeader1 = new List<string> { "", "", "Actual Start", "Target Finish", "Target (Hours)", "Actual Completion (Hours)", "Actual Arc On Time (Hours)", "Arc On Time (%)", "Actual Eff (%)" };
			List<string> colHeaderName = new List<string> { "machine_no", "id_pipe", "startTarget", "finishTarget", "Target_Time", "Actual_Time", "Actual_Arc_Time", "ArcOnTime", "ArcEff" };
			List<string> colHeaderType = new List<string> { "text", "text", "date", "date", "number", "number", "number", "number", "number" };

			Range excelCellrange;
			excelSheet.Name = dt.TableName;
			var rowcount = dt.Rows.Count;
			var columncount = colHeader.Count;
			int rowIndex = 9;
			int FirstRow = 9;
			excelSheet.Cells[1, 1] = "Legend :";
			excelSheet.Cells[1, 1].Font.Bold = true;
			excelSheet.Cells[2, 1].Interior.Color = Color.Red;
			excelSheet.Cells[3, 1].Interior.Color = Color.Yellow;
			excelSheet.Cells[5, 1] = "Report Date :";
			excelSheet.Cells[5, 1].Font.Bold = true;
			excelSheet.Cells[5, 2] = DateTime.Now.ToString("yyyy/M/dd");
			excelSheet.Cells[5, 2].NumberFormat = "dd MMM yyyy";
			excelCellrange = excelSheet.Range[excelSheet.Cells[2, 2], excelSheet.Cells[2, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Arc exceeds Target (Hours)";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[3, 2], excelSheet.Cells[3, colHeader1.Count]];
			excelCellrange.Merge();
			excelCellrange.Value = "Actual Completion exceeds Target (Hours)";
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
			excelCellrange = excelSheet.Range[excelSheet.Cells[6, colHeader1.Count], excelSheet.Cells[7, colHeader1.Count]];
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			var obj = ToObjectWithHeader(dt, colHeader, colHeaderName, colHeaderType, false);
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 3], excelSheet.Cells[rowcount + FirstRow, 4]];
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange.NumberFormat = "dd/MM/yyyy HH:mm";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 5], excelSheet.Cells[rowcount + FirstRow, 8]];
			excelCellrange.NumberFormat = "#,##0.00";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.Value = obj;
			for (var i = 1; i <= colHeader1.Count; i++)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[FirstRow - 1, i], excelSheet.Cells[FirstRow - 1, i]];
				excelCellrange.Value = colHeader1[i - 1];
				excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
				excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			}
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.NumberFormat = "@";
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex - 1, 1]];
			excelCellrange.Merge();
			excelCellrange.Value = "Machine";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 15;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 2], excelSheet.Cells[rowIndex - 1, 2]];
			excelCellrange.Merge();
			excelCellrange.Value = "Pipe No";
			excelCellrange.EntireColumn.AutoFit();
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 3], excelSheet.Cells[rowIndex - 2, 4]];
			excelCellrange.Merge();
			excelCellrange.Value = "Production";
			excelCellrange.ColumnWidth = 36;
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 5], excelSheet.Cells[rowIndex - 2, columncount]];
			excelCellrange.Merge();
			excelCellrange.Value = "Pipes Complete";
			excelCellrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange.ColumnWidth = 28;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowIndex, columncount]];
			excelCellrange.Font.Bold = true;
			excelCellrange.WrapText = true;
			excelCellrange.Interior.Color = XlRgbColor.rgbAqua;
			excelCellrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex - 2, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			Borders border = excelCellrange.Borders;
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2d;
			if (rowcount > 0)
			{
				excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex + 1, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
				string strFormula = String.Format(@"=$G10>$E10");
				FormatCondition cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Red;
				cond.Font.Color = Color.White;
				excelSheet.EnableFormatConditionsCalculation = true;
				strFormula = String.Format(@"=$F10>$E10");
				cond = excelCellrange.FormatConditions.Add(XlFormatConditionType.xlExpression, Type.Missing, strFormula, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				cond.StopIfTrue = false;
				cond.Interior.Color = Color.Yellow;
			}
			excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, 1], excelSheet.Cells[rowcount + FirstRow, columncount]];
			excelCellrange.AutoFilter(1);
			excelSheet.Application.ActiveWindow.SplitRow = FirstRow;
			excelSheet.Application.ActiveWindow.SplitColumn = 2;
			excelSheet.Application.ActiveWindow.FreezePanes = true;
		}
		public Object ToObjectWithHeader(System.Data.DataTable Data, List<string> colHeader, List<string> colHeaderName, List<string> colHeaderType, bool addNo)
		{
			try
			{
				var rowcount = Data.Rows.Count;
				var columncount = colHeader.Count();
				var obj = new object[rowcount + 1, columncount];
				int Temp = 0;
				int rw = 0;
				int cl = 0;

				foreach (var i in colHeader)
				{
					obj[rw, cl] = i;
					cl++;
				}

				foreach (DataRow dr in Data.Rows)
				{
					rw++;
					cl = 0;
					foreach (DataColumn dc in Data.Columns)
					{
						try
						{
							if (cl == 0 && addNo)
							{
								obj[rw, cl] = rw;
								cl++;
							}
							else if (colHeaderName.Contains(dc.ColumnName))
							{
								Temp = Enumerable.Range(0, colHeaderName.Count).Where(i => colHeaderName[i] == dc.ColumnName).FirstOrDefault();
								if (colHeaderType[Temp].ToString() == "number" || colHeaderType[Temp].ToString() == "date")
									obj[rw, Temp] = dr[dc.ColumnName];
								else
									obj[rw, Temp] = "'" + dr[dc.ColumnName];
								cl++;
							}
						}
						catch (Exception x)
						{
							x.ToString();
						}
					}
				}
				return obj;
			}
			catch (Exception x)
			{
				throw x;
			}
		}

		public string toJSON(System.Data.DataTable dt)
		{
			JavaScriptSerializer oSerializer = new JavaScriptSerializer();
			oSerializer.MaxJsonLength = Int32.MaxValue;
			string result = oSerializer.Serialize(dt);
			return result;
		}
		public string toJSON(System.Data.DataTable[] dt)
		{
			JavaScriptSerializer oSerializer = new JavaScriptSerializer();
			oSerializer.MaxJsonLength = Int32.MaxValue;
			string result = oSerializer.Serialize(dt);
			return result;
		}
		public static object DataTableToJSON(System.Data.DataTable table)
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