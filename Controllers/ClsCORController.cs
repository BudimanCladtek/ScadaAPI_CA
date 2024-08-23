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
using HttpGetAttribute = System.Web.Http.HttpGetAttribute;
using CrystalDecisions.Shared;
using CORSYS_API.Models;
using iTextSharp.text.pdf;
using System.Web.Razor.Text;
using System.Web.Script.Serialization;
using System.Text;
using System.Drawing;
using System.Web.Helpers;
using adminlte.Modul;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SCADA_API;
using System.Web.Http.Cors;

namespace CORSYS_API.Controllers
{
    [EnableCors(origins: "http://mainca.cor.sys", headers: "*", methods: "get,post")]
	//[EnableCors(origins: "http://localhost:44395", headers: "*", methods: "get,post")]
	//[Authentication]
	public class ClsCORController : ApiController
	{
		public ClsScadaSql sdb = new ClsScadaSql();

		[HttpPostAttribute]
		public HttpResponseMessage GetWCSFromScada([FromBody] AuthModel AM)
		{
			try
			{
				AM.json = string.IsNullOrEmpty(AM.json) ? "" : AM.json;
				var criteria = JValue.Parse(AM.json);
				string itemno = "";
				string status="";
				if (criteria["itemno"] != null)
					itemno = criteria["itemno"].Value<string>();
				if (criteria["status"] != null)
					status = criteria["status"].Value<string>();
				var wcs = sdb.GetDataTableSQL("SCADA", String.Format(@"
					Declare @item_no varchar(max)='{0}',
					@status varchar(50)='{1}'
				select * from (
				select 
					master.dbo.GROUP_CONCAT(distinct c.machine_no) as Machine, 
					master.dbo.GROUP_CONCAT(distinct convert(varchar,try_convert(int,y.strvalue))) as WelderID, 
					a.shift as Shift, a.pre_heat as Preheat, a.layer as Layer, a.weld_current as WeldCurrent,
					a.interpass_temp as InterpassTemp, a.hw_voltage as HWVoltage, a.hw_current as HWCurrent,
					a.datum as Datum, a.audited as Audited, a.status as Status, a.total_wire as TotalWire,
					a.voltage as Voltage, a.weld_speed as WeldSpeed, a.wire_speed1 as WireSpeed1, a.wire_speed2 as WireSpeed2,
					a.first_weight as FirstWeight,a.final_weight as FinalWeight, 
					a.wire1_length as Wire1Length, 
					a.wire2_length as Wire2Length, 
					a.wire3_length as Wire3Length, 
					a.wire4_length as Wire4Length, 
					a.setpoint_stepback as SetpointStepback,
					DateAdd(day, case left(convert(varchar,a.time_rec),2) when '24' then 1 else 0 end, 
					TRY_CONVERT(datetime,concat(date,' ', replace(replace(convert(varchar,a.time_rec),'24.','00.'), '.',':')),103)
					)DateRec,
					a.enclad_process
				from [SCADA QH].dbo.sfnGetWCS(@item_no, null) a 
				join [SCADA QH].dbo.TblPipe b on a.id_pipe = b.id_pipe
				cross apply [SCADA].dbo.sfnStringSplit(a.id_machine,',') x 
				cross apply [SCADA].dbo.sfnStringSplit(replace(replace(replace(replace(replace(a.welder_id,'CT',''),'-',''),'.',''),'=',''),' ',''),',') y 
				join [SCADA QH].dbo.TblMachine c on c.id_machine = x.strvalue
				where convert(varchar,b.item_no) like @item_no and convert(varchar,a.status) = @status
				--and (Datum like 'R%' or Datum like 'Y%' or Datum like 'X%' or Datum like 'Z%') 
				--and Replace(Datum, ' ', '') not like '%notwelding%' and c.id_machine != 0  
				--and a.enclad_process = @process
				group by a.shift, a.pre_heat, a.layer, a.weld_current, a.interpass_temp, a.hw_voltage, a.hw_current, a.datum, a.audited, a.status, 
				a.total_wire, a.voltage, a.weld_speed, a.wire_speed1, a.wire_speed2, a.first_weight, a.final_weight, a.wire1_length, a.wire2_length, 
				a.wire3_length, a.wire4_length, a.setpoint_stepback, a.time_rec, a.date,
				a.enclad_process
				) x 
				where isnull(LEN(x.WelderID), 0) > 0", itemno, status), 1200);
				var JsonResult = JsonConvert.SerializeObject(wcs);
				var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
				result.Content = stringContent;
				return result;

			}
			catch (Exception x)
			{
				var stringContent = new StringContent(x.Message, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
				result.Content = stringContent;
				return result;
			}
		}
		[HttpPostAttribute]
		public HttpResponseMessage GetWCSHeaderFromScada([FromBody] AuthModel AM)
		{
			try
			{
				AM.json = string.IsNullOrEmpty(AM.json) ? "" : AM.json;
				var criteria = JValue.Parse(AM.json);
				string itemno = "";
				string status = "";
				if (criteria["itemno"] != null)
					itemno = criteria["itemno"].Value<string>();
				if (criteria["status"] != null)
					status = criteria["status"].Value<string>();
				var wcs = sdb.GetDataTableSQL("SCADA", String.Format(@"
					Declare @item_no varchar(max)='{0}',
					@status varchar(50)='{1}'
					select distinct
                        b.item_no as PipeCladUniqNo,  b.wcs as WPSNo, b.wcs as WPSNoF, b.wcs2 as WPSNoB, b.wcs3 as WPSNoG,
                        b.wire_1 as Wire1, b.wire_2 as Wire2, b.wire_3 as Wire3, b.wire_4 as Wire4, 
                        b.HeatNoA1 as Wire1HeatNo, b.HeatNoA2 as Wire2HeatNo, b.HeatNoB1 as Wire3HeatNo, b.HeatNoB2 as Wire4HeatNo,
                        b.gas_batch as SG1BatchNo, b.gas_batch as SG2BatchNo, TotalWeldLength as WOLLength
                    from [SCADA QH].dbo.TblProcess a 
                    join [SCADA QH].dbo.TblPipe b on a.id_pipe = b.id_pipe
                    where convert(varchar,b.item_no) like @item_no and convert(varchar,a.status) = @status", itemno, status), 1200);
				var JsonResult = JsonConvert.SerializeObject(wcs);
				var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
				result.Content = stringContent;
				return result;

			}
			catch (Exception x)
			{
				var stringContent = new StringContent(x.Message, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
				result.Content = stringContent;
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
