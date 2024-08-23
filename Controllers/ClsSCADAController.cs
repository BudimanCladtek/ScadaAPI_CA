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
    [EnableCors(origins: "http://10.112.220.69", headers: "*", methods: "get,post")]
	//[EnableCors(origins: "http://localhost:44395", headers: "*", methods: "get,post")]
	//[Authentication]
	public class ClsSCADAController : ApiController
	{
//		public ClsSql db = new ClsSql();
		public ClsScadaSql sdb = new ClsScadaSql();

		//[HttpPostAttribute]
		//public HttpResponseMessage GetWelder([FromBody] CRModel text)
		//{
		//	try
		//	{
		//		db.OpenConn();
		//		string query = "select * from matahari.dbo.WelderRegister";
		//		var Data = db.GetDataTableSQL(query);
		//		var stringContent = new StringContent(DataTableToJSON(Data).ToString(), UnicodeEncoding.UTF8, "application/json");
		//		HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
		//		result.Content = stringContent;
		//		return result;
		//	}
		//	catch (Exception x)
		//	{
		//		HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
		//		return result;
		//	}

		//}
		//[HttpPostAttribute]
		//public IHttpActionResult GetWeldera([FromBody] CRModel text)
		//{
		//	try
		//	{
		//		db.OpenConn();
		//		string query = "select * from matahari.dbo.WelderRegister --where dlt=1";
		//		var Data = db.GetDataTableSQL(query);
		//		return Json(Data);
		//	}
		//	catch (Exception x)
		//	{
		//		return Json(x.Message);
		//	}

		//}
		[HttpPostAttribute]
		public HttpResponseMessage CheckUserSystem([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				string JsonResult = ClsCR.CheckUserNameCOR(AM.json);

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
		public HttpResponseMessage CheckAllUserSystem([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				string JsonResult = ClsCR.CheckAllUserNameCOR(AM.json);

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
		public HttpResponseMessage CheckUserSystemFromScada([FromBody] AuthModel AM)
		{
			try
			{
				var data = JValue.Parse(AM.json);
				string username = data[0]["UserName"].ToString();
				string password = data[0]["PassCode"].ToString();

				//var login = sdb.GetDataTableSQL("SCADA", String.Format(@"Select ISNULL(FullName, username) as Name, Username, 1 as IsActive, ISNULL(b.WelderStamp,'') as WelderID, '' as Title from userlogin a
				//	left join tblWelder b on a.username=b.EmployeeID where username='{0}' and password='{1}'", username, password), 1).Rows;
				var login = sdb.GetDataTableSQL("SCADA", String.Format(@"Select ISNULL(FullName, username) as Name, Username, Password, 1 as IsActive, ISNULL(b.WelderStamp,'') as WelderID, '' as Title, Role
					from userlogin a
					left join tblWelder b on a.username=b.EmployeeID where username='{0}'", username), 1).Rows;
				String JsonResult = "User Not Found";
				if (login.Count != 0)
				{
					JsonResult = JsonConvert.SerializeObject(login[0].Table);
					string passCode = JValue.Parse(JsonResult)[0]["Password"].ToString();
					if(!(password==CryptoAPI.EncryptHash(passCode) || password==passCode))
                    {
						JsonResult = "User Not Found or Wrong Password ";
					}
					else
                    {
						DataTable dataResult = login[0].Table;
						dataResult.Columns.Remove("Password");
						JsonResult = JsonConvert.SerializeObject(dataResult);
					}

				}
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
		public HttpResponseMessage GetProject([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				string JsonResult = ClsCR.GetProject(AM.json);

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
		public HttpResponseMessage GetProjectNew([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetProject(AM.json);
				int page = 0;
				int skip = 30;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				var data = JValue.Parse(JsonResult);
				JsonResult = JsonConvert.SerializeObject(data.Children().Select(x => new { ID = x["ID"], ProjectID = x["ProjectID"] }));
				data = JValue.Parse(JsonResult);
				Int32 total = data.Count();
				if (page != 0)
					JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}");

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
		public HttpResponseMessage GetItemForWOLScada([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				string JsonResult = ClsCR.GetItemForWOLScada(AM.json);

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
		public HttpResponseMessage GetItemForPAWIIScada([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				string JsonResult = ClsCR.GetItemForPAWIIScada(AM.json);

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
		public HttpResponseMessage GetItemForWOLScadaNew([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetItemForWOLScada(AM.json);
				int page = 0;
				int skip = 30;
				string filter = "";
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				if (criteria["filter"] != null)
					filter = criteria["filter"].Value<string>();
				var data = JValue.Parse(JsonResult);
				if (filter.Length > 0)
				{
					JsonResult = JsonConvert.SerializeObject(data.Children().Select(x => new { PipeCladUniqNo = x["PipeCladUniqNo"], ReportStatus = x["ReportStatus"] }).Where(x => x.PipeCladUniqNo.ToString().Contains(filter)));
					data = JValue.Parse(JsonResult);
				}
				if (criteria["compress"] != null)
				{
					JsonResult = JsonConvert.SerializeObject(data.Children().Select(x => new { UniqNo = x["PipeCladUniqNo"] }));
					data = JValue.Parse(JsonResult);
				}
				int total = data.Count();
				if (page != 0)
				{
					if (skip > (total - (page * skip)) && (total - ((page - 1) * skip)) > 0)
					{
						int take = total - ((page - 1) * skip);
						take = (take > skip) ? skip : take;
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(take));
					}
					else
					{
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
					}
				}
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"","'") , "\", \"total\":", total, "}");
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
		public HttpResponseMessage GetSearchItemForWOLScadaNew([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetItemForWOLScada(AM.json);
				int totData = 10;
				string filter = "";
				if (criteria["totData"] != null)
					totData = criteria["totData"].Value<int>();
				if (criteria["filter"] != null)
					filter = criteria["filter"].Value<string>();
				var data = JValue.Parse(JsonResult);
				if (criteria["compress"] != null)
				{
					JsonResult = JsonConvert.SerializeObject(data.Children().Select(x => new { UniqNo = x["PipeCladUniqNo"] }));
				}
                if (filter.Length > 0)
                {
					JsonResult = JsonConvert.SerializeObject(JArray.Parse(JsonResult).Where(i => i["UniqNo"].ToString().Contains(filter.ToString())));
				}
				data = JValue.Parse(JsonResult);
				int total = data.Count();
				JsonResult = JsonConvert.SerializeObject(data.Take(totData));
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}");
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
		public HttpResponseMessage GetItemForMachiningScadaNew([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetItemForMachiningScada(AM.json);
				int page = 0;
				int skip = 30;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				//				JsonResult = JsonResult.Replace("PipeCladUniqNo", "UniqNo").Replace("ReportStatus", "Status");
				var data = JValue.Parse(JsonResult);
				JsonResult = JsonConvert.SerializeObject(data.Children().Select(x => new { PipeCladUniqNo = x["PipeCladUniqNo"], ReportStatus = x["ReportStatus"] }));
				data = JValue.Parse(JsonResult);
				int total = data.Count();
				if (page != 0)
				{
					if (skip > (total - (page * skip)) && (total - ((page - 1) * skip)) > 0)
					{
						int take = total - ((page - 1) * skip);
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(take));
					}
					else
					{
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
					}
				}
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}");
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
		public HttpResponseMessage GetItemForWOLScadaNew1([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetItemForWOLScada(AM.json);
				int page = 0;
				int skip = 30;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				//				JsonResult = JsonResult.Replace("PipeCladUniqNo", "UniqNo").Replace("ReportStatus", "Status");
				var data = JValue.Parse(JsonResult);
				int total = data.Count();
				if (page != 0)
				{
					if (skip>(total - (page * skip)) && (total - ((page - 1) * skip))> 0)
					{
						int take = total - ((page - 1) * skip);
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(take));
						JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":\"", total, "\", \"test\":", (page - 1) * skip, "}");
					}
					else
					{
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
						JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":\"", total, "\", \"skip\":", (page * skip), "}");
//						JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}");
					}
				}
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
		public HttpResponseMessage GetItemForPAWIIScadaNew([FromBody] AuthModel AM)
		{
			try
			{
				//db.OpenConn();
				//ClsCR.CheckUserNameCOR(AM.json);
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetItemForPAWIIScada(AM.json);
				int page = 0;
				int skip = 30;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				//				JsonResult = JsonResult.Replace("PipeCladUniqNo", "UniqNo").Replace("ReportStatus", "Status");
				var data = JValue.Parse(JsonResult);
				Int32 total = data.Count();
				if (page != 0)
					JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}");
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
		public HttpResponseMessage GetCladLineNo([FromBody] AuthModel AM)
		{
			try
			{
				string JsonResult = ClsCR.GetCladLineNo(AM.json);

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
		public HttpResponseMessage GetCladLineNoNew([FromBody] AuthModel AM)
		{
			try
			{

				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetCladLineNo(AM.json);
				int page = 0;
				int skip = 30;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				var data = JValue.Parse(JsonResult);
				JsonResult = JsonConvert.SerializeObject(data.Children().Select(x => new { CladLineNo = x["CladLineNo"] }));
				data = JValue.Parse(JsonResult);
				Int32 total = data.Count();
				if (page != 0)
					JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}"); 

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
		public HttpResponseMessage GetWelderStamp([FromBody] AuthModel AM)
		{
			try
			{
				string JsonResult = ClsCR.GetWelderStamp(AM.json);

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
		public HttpResponseMessage GetWelderStampFromScada([FromBody] AuthModel AM)
		{
			try
			{
				var welderStamp = sdb.GetDataTableSQL("SCADA", String.Format(@"Select ISNULL(FullName, username) as Name, Username, 1 as IsActive, ISNULL(b.WelderStamp,'') as WelderID, '' as Title from userlogin a
					left join tblWelder b on a.username=b.EmployeeID"), 1200);
				var JsonResult = JsonConvert.SerializeObject(welderStamp);
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
		/// <summary>  
		/// Get all wps (by datum)  
		/// </summary>  
		/// <returns>List of Student</returns>  
		public HttpResponseMessage GetWPS([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				string weldpass = "";
				int wpsonly = 0;
				if (criteria["wpsonly"] != null)
					wpsonly = criteria["wpsonly"].Value<int>();
				if (criteria["weldpass"] != null)
					weldpass = criteria["weldpass"].Value<string>();
				string JsonResult = ClsCR.GetWPS(AM.json);
				if (JsonResult.Contains("Cannot find WPS"))
					JsonResult = "[]";
				dynamic listWPS = JsonConvert.DeserializeObject(JsonResult);
				if (weldpass != "")
                {
					try
					{
						listWPS = JsonConvert.DeserializeObject(JsonConvert.SerializeObject(JValue.Parse(JsonResult).Where(x => x["WeldPass"].ToString() == weldpass)));
					}
					catch
                    {
						listWPS =listWPS;
					}
				}

				for (int i = 0; i < ((JContainer)listWPS).Count; i++)
				{
					if (listWPS[i]["FillerDia"].Value == null)
						listWPS[i]["FillerDia"].Value = "";
					JsonResult = JsonConvert.SerializeObject(listWPS, Formatting.Indented);
				}

				if (wpsonly == 1)
				{
					var data = JValue.Parse(JsonResult);
					JsonResult = JsonConvert.SerializeObject(data.Children()
						.Select(x => new
						{
							ID = x["ID"],
							ProjectWPSNo = x["ProjectWPSNo"],
							Application = x["Application"],
							WeldPass = x["WeldPass"],
						}));

				}

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
		public HttpResponseMessage GetWPSPAWII([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				int wpsonly = 0;
				if (criteria["wpsonly"] != null)
					wpsonly = criteria["wpsonly"].Value<int>();
				string JsonResult = ClsCR.GetWPSPAWII(AM.json);
				if (wpsonly == 1)
				{
					var data = JValue.Parse(JsonResult);
					JsonResult = JsonConvert.SerializeObject(data.Children()
						.Select(x => new
						{
							ID = x["ID"],
							ProjectWPSNo = x["ProjectWPSNo"],
							Application = x["Application"],
							WeldPass = x["WeldPass"],
						}));

				}

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
		public HttpResponseMessage GetWPSNew([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				string JsonResult = ClsCR.GetWPS(AM.json); 
				int page = 0;
				int skip = 30;
				int wpsonly = 0;
				int wpsnoonly = 0;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				if (criteria["wpsonly"] != null)
					wpsonly = criteria["wpsonly"].Value<int>();
				if (criteria["wpsnoonly"] != null)
					wpsnoonly = criteria["wpsnoonly"].Value<int>();
				var data = JValue.Parse(JsonResult);
				if (wpsonly == 1)
				{
					if (wpsnoonly == 0)
					{
						JsonResult = JsonConvert.SerializeObject(data.Children()
							.Select(x => new
							{
								ID = x["ID"],
								ProjectWPSNo = x["ProjectWPSNo"],
								Application = x["Application"],
								WeldPass = x["WeldPass"],
							}));
                    }
                    else
                    {
						JsonResult = JsonConvert.SerializeObject(data.Children()
							.Select(x => new
							{
								ProjectWPSNo = x["ProjectWPSNo"]
							}));
					}

				}
				else
				{
					JsonResult = JsonConvert.SerializeObject(data.Children()
						.Select(x => new
						{
							ID = x["ID"],
							ProjectWPSNo = x["ProjectWPSNo"],
							Application = x["Application"],
							WeldPass = x["WeldPass"],
							MinWireSpeed = x["MinWireSpeed"],
							MaxWireSpeed = x["MaxWireSpeed"],
							MinAmperage = x["MinAmperage"],
							MaxAmperage = x["MaxAmperage"],
							MinVoltage = x["MinVoltage"],
							MaxVoltage = x["MaxVoltage"],
							MinTravSpeed = x["MinTravSpeed"],
							MaxTravSpeed = x["MaxTravSpeed"],
							MatGrade = x["MatGrade"],
							Speed = x["Speed"],
							Step = x["Step"],
							FillerDia = x["FillerDia"],
							WireFeed = x["WireFeed"]
						}));
				}
				data = JValue.Parse(JsonResult);
				Int32 total = data.Count();
				if (page != 0)
					JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}"); 

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
		public HttpResponseMessage GetWPSByNo([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				string wpsno = "";
				string JsonResult = ClsCR.GetWPS(AM.json);
				if (criteria["wpsno"] != null)
					wpsno = criteria["wpsno"].Value<string>();
				if (wpsno != "")
				{
					var data = JValue.Parse(JsonResult);
					JsonResult = JsonConvert.SerializeObject(data.Children()
						.Select(x => new
						{
							ID = x["ID"],
							ProjectWPSNo = x["ProjectWPSNo"],
							Application = x["Application"],
							WeldPass = x["WeldPass"],
							MinWireSpeed = x["MinWireSpeed"],
							MaxWireSpeed = x["MaxWireSpeed"],
							MinAmperage = x["MinAmperage"],
							MaxAmperage = x["MaxAmperage"],
							MinVoltage = x["MinVoltage"],
							MaxVoltage = x["MaxVoltage"],
							MinTravSpeed = x["MinTravSpeed"],
							MaxTravSpeed = x["MaxTravSpeed"],
							MatGrade = x["MatGrade"],
							Speed = x["Speed"],
							Step = x["Step"],
							FillerDia = x["FillerDia"],
							WireFeed = x["WireFeed"]
						})
						.Where(x=> x.ProjectWPSNo.ToString()==wpsno).ToList()
						);
					data = JValue.Parse(JsonResult);
					JsonResult = JsonConvert.SerializeObject(data);
				}
				else
				{
					JsonResult = "[]";
				}

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
		public HttpResponseMessage PostNewReportWOL([FromBody] AuthModel AM)
		{
			string JsonResult = "";
			var partType = JObject.Parse(AM.json)["PartType"];
			try
			{
				string pType = (partType != null) ? partType.Value<string>() : "";
				if (pType == "Liner")
				{
					JsonResult = ClsCR.PostNewReportPAW(AM.json);
				}
				else
				{
					JsonResult = ClsCR.PostNewReportWOL(AM.json);
				}

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
		public HttpResponseMessage PostReportWOLFitting([FromBody] AuthModel AM)
		{
			try
			{
				string JsonResult = ClsCR.PostReportWOLFitting(AM.json);

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
		public HttpResponseMessage PostInsertTargetArcOn()
		{
			try
			{
				string strJSON = @"{""ProjectID"":""',Word1,'"", ""CladLineNo"":""', Replace(Word3,'-',''), '""}";
				string strQuery = String.Format(@"Declare @item_no varchar(50),
						@mfg varchar(50),
						@BodyData nvarchar(max),
						@projectid varchar(10),
						@cladlineno varchar(50)
					Select @item_no=a.Pipe_No from MachineStatus a where a.Machine_id='SHV 84'
					IF(Select Count(*) from TblPipeLoad where PipeLoad= @item_no)=0
					BEGIN
						Select @BodyData=CONCAT('{0}'), @projectid=Word1, @cladlineno=Replace(Word3,'-', '') 
							from SCADA.dbo.sfnSplitPipe(@item_no)
						Exec APIGetManufacturing @BodyData = @BodyData, @mfg = @mfg output
						IF(@mfg = 'FLANGE')
						BEGIN TRY
							EXEC dbo.sprInsPipeLoadFlange @projectid, @cladlineno, @item_no
						END TRY
						BEGIN CATCH
							Print 'A'
						END CATCH
						IF(@mfg = 'FITTING')
						BEGIN
							EXEC dbo.sprInsPipeLoadFitting @projectid, @cladlineno, @item_no
						END
					END
					IF(Select Count(*) from TargetArcOn a where a.PipeNo = @item_no) = 0
					BEGIN
						Exec dbo.sprInsTargetArcOnByItemNo @item_no
					END", strJSON);

				sdb.RunQuery(strQuery, "Scada SHV");

				var stringContent = new StringContent("", UnicodeEncoding.UTF8, "application/json");
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
		public HttpResponseMessage PostMachineOEE([FromBody] AuthModel AM)
		{
			try
			{
				var data = JValue.Parse(AM.json);
				var JsonResult = JsonConvert.SerializeObject(data);
				var stringContent = new StringContent(AM.json.Length.ToString(), UnicodeEncoding.UTF8, "application/json");
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
		public HttpResponseMessage GetWire([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				var param = JObject.Parse(AM.json);
				if (param["WPSNo"].ToString() == "CT-00-WPS-1345")
					param["WPSNo"]="";
				if (param["WPSNo"].ToString() == "CT-00-WPS-1345 Rev 0")
					param["WPSNo"] = "";
				if (param["Qty"]!=null)
					param.Property("Qty").Remove();
				string brand = "";
				int Qty = 0;
				if (criteria["brand"] != null)
					brand = criteria["brand"].Value<string>();
				if (criteria["Qty"] != null)
					Qty = criteria["Qty"].Value<int>();

				string JsonResult = ClsCR.GetWire(param.ToString());

				List<WireModel> dataWire = JsonConvert.DeserializeObject<List<WireModel>>(JsonResult);

				if (JsonResult.Length != 0)
				{
					var data = JValue.Parse(JsonResult);
					if (brand != "")
						JsonResult = JsonConvert.SerializeObject(data.Where(x => x.Last.First.ToString().ToLower() == brand.ToLower()));
					if (Qty != 0)
					{
						var dataResult = JValue.Parse(JsonResult).Take(Qty);
						foreach(var a in dataWire.GroupBy(x=> x.WireBrand.ToUpper().Trim()).Select(y=> new WireModel { WireBrand = y.Key }).Take(7))
                        {
							if(dataResult.Where(x=> x.Last.First.ToString().ToUpper() == a.WireBrand.ToUpper()).Count() == 0)
							{
								dataResult=dataResult.Union(data.Where(x => x.Last.First.ToString().ToLower() == a.WireBrand.ToLower()).Take(1));
                            }
                        }
						JsonResult = JsonConvert.SerializeObject(dataResult);
//						JsonResult = JsonConvert.SerializeObject(JValue.Parse(JsonResult).Take(Qty));
					}
					//				JsonResult = JsonConvert.SerializeObject(data);
				}
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
		public HttpResponseMessage GetWireFromScada([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				var param = JObject.Parse(AM.json);
				if (param["Qty"] != null)
					param.Property("Qty").Remove();
				string brand = "";
				int Qty = 0;
				if (criteria["brand"] != null)
					brand = criteria["brand"].Value<string>();
				if (criteria["Qty"] != null)
					Qty = criteria["Qty"].Value<int>();

				var dataWire = sdb.GetDataTableSQL("SCADA", String.Format(@"Select WireID as ID, WireBrand, HeatNo as WBNo from tblWire"), 1200);
				var JsonResult = JsonConvert.SerializeObject(dataWire);

				if (JsonResult.Length != 0)
				{
					var data = JValue.Parse(JsonResult);
					if (brand != "")
						JsonResult = JsonConvert.SerializeObject(data.Where(x => x.Last.First.ToString().ToLower() == brand.ToLower()));
					if (Qty != 0)
					{
						JsonResult = JsonConvert.SerializeObject(JValue.Parse(JsonResult).Take(Qty));
					}
					//				JsonResult = JsonConvert.SerializeObject(data);
				}
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
		public HttpResponseMessage GetWireNew([FromBody] AuthModel AM)
		{
			try
			{
				var criteria = JValue.Parse(AM.json);
				var param = JObject.Parse(AM.json);
				if (param["Qty"] != null)
					param.Property("Qty").Remove();
				string brand = "";
				int Qty = 0;
				if (criteria["brand"] != null)
					brand = criteria["brand"].Value<string>();
				if (criteria["Qty"] != null)
					Qty = criteria["Qty"].Value<int>();
				string JsonResult = ClsCR.GetWire(param.ToString());
				if (JsonResult.Length != 0)
				{
					var data = JValue.Parse(JsonResult);
					if (brand != "")
                    {
						JsonResult = JsonConvert.SerializeObject(data.Where(x => x.Last.First.ToString().ToLower() == brand.ToLower()));
						JsonResult = JsonConvert.SerializeObject(JValue.Parse(JsonResult).Children()
											.Select(x => new {
												WBNo = x["WBNo"]
											}));
					}
					if (Qty != 0)
					{
						JsonResult = JsonConvert.SerializeObject(JValue.Parse(JsonResult).Take(Qty));
					}
					//				JsonResult = JsonConvert.SerializeObject(data);
				}
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
		public HttpResponseMessage GetWireAll([FromBody] AuthModel AM)
		{
			string strData = string.Empty;
			try
			{
				var criteria = JValue.Parse(AM.json);
				int page = 0;
				int skip = 50;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				var param = JObject.Parse(AM.json);
				if (param["Qty"] != null)
					param.Property("Qty").Remove();
				if (param["brand"] != null)
					param.Property("brand").Remove();
				string brand = "";
				int Qty = 0;
				if (criteria["brand"] != null)
					brand = criteria["brand"].Value<string>();
				if (criteria["Qty"] != null)
					Qty = criteria["Qty"].Value<int>();
				string JsonResult = ClsCR.GetWire(param.ToString());
				strData = JsonResult;
				Int32 total = JValue.Parse(JsonResult).Count();
				if (JsonResult.Length != 0)
				{
					var data = JValue.Parse(JsonResult);
					if (brand != "")
					{
						JsonResult = JsonConvert.SerializeObject(data.Where(x => x.Last.First.ToString().ToLower() == brand.ToLower()));
						JsonResult = JsonConvert.SerializeObject(JValue.Parse(JsonResult).Children()
											.Select(x => new {
												WBNo = x["WBNo"]
											}));
						data=JValue.Parse(JsonResult);
						total = JValue.Parse(JsonResult).Count();
					}
					if (Qty != 0)
					{
						JsonResult = JsonConvert.SerializeObject(JValue.Parse(JsonResult).Take(Qty));
					}
					if (page != 0)
						JsonResult = JsonConvert.SerializeObject(data.Skip((page - 1) * skip).Take(skip));
					//				JsonResult = JsonConvert.SerializeObject(data);
				}
				JsonResult = String.Concat("{\"data\":\"", JsonResult.Replace("\"", "'"), "\", \"total\":", total, "}");
				var stringContent = new StringContent(JsonResult, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
				result.Content = stringContent;
				return result;

			}
			catch (Exception x)
			{
				strData= String.Concat("{\"data\":\"[]\", \"error\":\"", x.Message.Replace("\"", "'"), "\", \"total\":0}");
				var stringContent = new StringContent(x.Message, UnicodeEncoding.UTF8, "application/json");
				HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
				result.Content = stringContent;
				return result;
			}

		}

		[HttpPostAttribute]
		public HttpResponseMessage GetGas([FromBody] AuthModel AM)
		{
			
			try
			{
				var criteria = JValue.Parse(AM.json);
				int page = 0;
				int skip = 90;
				if (criteria["page"] != null)
					page = criteria["page"].Value<int>();
				if (criteria["skip"] != null)
					skip = criteria["skip"].Value<int>();
				//				JsonResult = JsonResult.Replace("PipeCladUniqNo", "UniqNo").Replace("ReportStatus", "Status");

				string JsonResult = ClsCR.GetGas(AM.json);
				var data = JValue.Parse(JsonResult);
				if(page!=0)
					JsonResult = JsonConvert.SerializeObject(data.Skip((page-1) * skip).Take(skip));
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
		public HttpResponseMessage GetSearchGas([FromBody] AuthModel AM)
		{

			try
			{
				var criteria = JValue.Parse(AM.json);
				string filter = "";
				if (criteria["filter"] != null)
					filter = criteria["filter"].Value<string>();

				string JsonResult = ClsCR.GetGas(AM.json);
				var data = JValue.Parse(JsonResult);
				JsonResult = JsonConvert.SerializeObject(data.Where(x=> x["GBNo"].ToString().Contains(filter)).ToList().Take(10).ToList());
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
		public HttpResponseMessage GetGasFromScada([FromBody] AuthModel AM)
		{
			try
			{
				var gasBatch = sdb.GetDataTableSQL("SCADA", String.Format(@"Select GasBatch as GBNo from tblGasBatch"), 1200);
				var JsonResult = JsonConvert.SerializeObject(gasBatch);
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
		public HttpResponseMessage GetPipeDetail([FromBody] AuthModel AM)
		{
			try
			{
				string JsonResult = ClsCR.GetPipeDetail(AM.json);

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

		//[HttpPostAttribute]
		//public HttpResponseMessage GetProjectMLP([FromBody] AuthModel AM)
		//{
		//	try
		//	{
		//		db.OpenConn();
		//		AM = db.CheckDB(AM);
		//		string query = String.Format(@"Select * from {0}.dbo.MasterProject where dlt != 1 and Status = {1}", AM.DBName, (AM.isActive ? 1 : 0)); 
		//		var Data = db.GetDataTableSQL(query);
		//		var stringContent = new StringContent(DataTableToJSON(Data).ToString(), UnicodeEncoding.UTF8, "application/json");
		//		HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
		//		result.Content = stringContent;
		//		return result;
		//	}
		//	catch (Exception x)
		//	{
		//		HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
		//		return result;
		//	}

		//}
		//[HttpPostAttribute]
		//public HttpResponseMessage GetPipeCladLineNo([FromBody] AuthModel AM)
		//{
		//    try
		//    {
		//        db.OpenConn();
		//        AM = db.CheckDB(AM);
		//        string query = String.Format(@"Select distinct a.MatGrade, b.CustomerPONo, a.Tagno, b.LOINo, a.ProductDesc, c.Supplier, d.RawMat, d.CladLineNo,  a.CladLineNo as CladUniqID
		//                    from {1}.dbo.TblBomPipeDT a
		//                    left join {1}.dbo.TblProjectPO b on a.CustPoItemNo = b.CustomerPONo and b.dlt != 1
		//                    left join {1}.dbo.TblMatCertPipe c on c.CladLineNo  = a.CladLineNo and c.dlt != 1
		//                    left join {1}.dbo.TblPdlDT d on substring(d.CladUniqID, 1, (len(d.CladUniqID) - 1))= a.CladLineNo and d.dlt != 1
		//                    where  a.Dlt!=1 and a.ParentID = (
		//                    select top 1 aa.ID from {1}.dbo.TblBomPipeHD aa
		//                    WHERE aa.Dlt != 1
		//                    and aa.projectid = '{0}'
		//                    --and aa.PrepareBy is not null and aa.CheckBy  is not null
		//                    --and aa.ApproveBy is not null
		//                    order by aa.ID desc
		//                    ) and b.projectid = '{0}' and c.projectid = '{0}' and d.projectid = '{0}'", AM.ProjectID, AM.DBName); ;

		//        var Data = db.GetDataTableSQL(query);
		//        var stringContent = new StringContent(DataTableToJSON(Data).ToString(), UnicodeEncoding.UTF8, "application/json");
		//        HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
		//        result.Content = stringContent;
		//        return result;
		//    }
		//    catch (Exception x)
		//    {
		//        HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.BadRequest);
		//        return result;
		//    }

		//}

		//[HttpGet]
		////public HttpResponseMessage GetItemForHydroScada(AuthModel AM)
		//public HttpResponseMessage GetItemForHydro(string ProjectID, string CladLineNo, bool isLive = true)
		//{
		//	try
		//	{
		//		db.OpenConn();
		//		AuthModel AM = new AuthModel()
		//		{
		//			ProjectID = ProjectID,
		//			CladLineNo = CladLineNo,
		//			isLive = isLive
		//		};

		//		AM = db.CheckDB(AM);

		//		string qry = String.Format(@"
		//			Select PipeCladUniqNo From(
		//				Select a0.* FROM(
		//					Select 'PrevProcess' as ItemOrigin, a.PipeCladUniqNo, 0 as ReportStatus, 0 as ReprocessFlag, 'C' as Result
		//						FROM {1}.dbo.WOLEndDT a
		//					Join {1}.dbo.WOLEndHD b on a.ParentID = b.ID
		//					Join {1}.dbo.MasterItemPipe d on d.PipeCladUniqNo = a.PipeCladUniqNo And d.Dlt != 1
		//					full outer join(select x.PipeCladUniqNo, x.ReportStatus, x.ReprocessFlag, x.Dlt FROM {1}.dbo.HydroExpanScadaDT x WHERE x.Dlt != 1
		//				) g on a.PipeCladUniqNo = g.PipeCladUniqNo and isnull(g.ReportStatus,0) = 0 and isnull(g.ReprocessFlag, 0) = 0
		//				where a.Dlt != 1 and b.Dlt != 1 and g.PipeCladUniqNo is null and b.ProjectID = '{2}' and a.Result = 'C' and b.PipeCladLineNo = '{0}'
		//				) a0
		//				UNION

		//				select distinct 'CompletedRepair' as ItemOrigin, a.PipeCladUniqNo, 0 as ReportStatus,0 as ReprocessFlag, 'C' as Result
		//				from {1}.dbo.DNCReleasePipe a

		//				join {1}.dbo.MasterItemPipe d on d.PipeCladUniqNo = a.PipeCladUniqNo And d.Dlt != 1

		//				full outer join(select x.PipeCladUniqNo, x.ReportStatus, x.ReprocessFlag, x.Dlt FROM {1}.dbo.HydroExpanScadaDT x WHERE x.Dlt != 1
		//				) c on a.PipeCladUniqNo = c.PipeCladUniqNo  and isnull(c.ReportStatus,0) = 0 and isnull(c.ReprocessFlag,0) = 0 and c.Dlt != 1

		//				where a.Dlt != 1 and a.ProjectID = '{2}' and a.PipeCladLineNo = '{0}' and c.PipeCladUniqNo is null and substring(a.ProcessToDoOrig, len(a.ProcessToDoOrig) -1, 1) in ('') and a.ProcessToDo = '_' and a.IsMove = 1


		//			UNION

		//				select distinct 'InitialReleaseMRQ' as ItemOrigin, a.PipeCladUniqNo, a.ReportStatus,a.ReprocessFlag, 'C' as Result
		//				from(
		//					SELECT a.* FROM {1}.dbo.DNCReleasePipe a
		//					full outer join(
		//						select * from {1}.dbo.DNCReleasePipe b WHERE b.Dlt != 1
		//					) b on a.PipeCladUniqNo = b.PipeCladUniqNo and a.ReportStatus < b.ReportStatus
		//					WHERE a.Dlt != 1 and b.PipeCladUniqNo is null and a.ProjectID = '{2}' AND a.PipeCladLineNo = '{0}' AND a.ProcessToDo LIKE 'N%' AND a.ReleaseDest = 'HydroExpanScada' AND a.IsMove = 0
		//					) a
		//				left join {1}.dbo.MasterItemPipe d on d.PipeCladUniqNo = a.PipeCladUniqNo And d.Dlt != 1


		//			UNION

		//				select distinct 'OngoingRepairMRQ' as ItemOrigin, a.PipeCladUniqNo, a.ReportStatus,a.ReprocessFlag, 'C' as Result
		//				from(
		//					SELECT a.* FROM {1}.dbo.DNCReleasePipe a
		//					full outer join(
		//						select * from {1}.dbo.DNCReleasePipe b WHERE b.Dlt != 1
		//					) b on a.PipeCladUniqNo = b.PipeCladUniqNo and a.ReportStatus < b.ReportStatus
		//					WHERE a.Dlt != 1 and b.PipeCladUniqNo is null and a.ProjectID = '{2}' AND a.PipeCladLineNo = '{0}' AND a.ProcessToDo LIKE 'N%' AND a.ReleaseDest != 'HydroExpanScada' AND a.IsMove = 1
		//					) a
		//				left join {1}.dbo.MasterItemPipe d on d.PipeCladUniqNo = a.PipeCladUniqNo And d.Dlt != 1


		//			UNION

		//				select distinct 'OngoingProdTest' as ItemOrigin, a.PipeCladUniqNo, a.ReportStatus,a.ReprocessFlag, 'C' as Result
		//				from(
		//					SELECT a.* FROM {1}.dbo.ProdTestLiner a
		//					full outer join(
		//						select * from {1}.dbo.ProdTestLiner b WHERE b.Dlt != 1
		//					) b on a.PipeCladUniqNo = b.PipeCladUniqNo and a.ReprocessFlag < b.ReprocessFlag
		//					WHERE a.Dlt != 1 and b.PipeCladUniqNo is null and a.ProjectID = '{2}' AND a.PipeCladLineNo = '{0}' AND a.ProcessNo = 6 AND a.IsMove = 1
		//					) a
		//				left join {1}.dbo.MasterItemPipe d on d.PipeCladUniqNo = a.PipeCladUniqNo And d.Dlt != 1
							
		//			 ) zz where PipeCladUniqNo not in (select PipeCladUniqNo from {1}.dbo.HydroExpanScadaDT where ParentID = 0 and Dlt != 1)
		//		", AM.CladLineNo, AM.DBName, AM.ProjectID);


		//		var Data = db.GetDataTableSQL(qry);
		//		var lst = new List<string>();
		//		foreach (DataRow row in Data.Rows)
		//		{
		//			lst.Add(row.ItemArray[0].ToString());
		//		}
		//		return Request.CreateResponse(HttpStatusCode.OK, lst);
		//	}
		//	catch (Exception x)
		//	{
		//		return Request.CreateResponse(HttpStatusCode.BadRequest);
		//	}

		//}
		//[HttpGet]
		////public HttpResponseMessage GetItemForHydroScada(AuthModel AM)
		//public HttpResponseMessage GetItemForHydroFromNDT(string ProjectID, string CladLineNo, bool isLive = true)
		//{
		//	try
		//	{
		//		db.OpenConn();
		//		AuthModel AM = new AuthModel()
		//		{
		//			ProjectID = ProjectID,
		//			CladLineNo = CladLineNo,
		//			isLive = isLive
		//		};

		//		AM = db.CheckDB(AM);

		//		string qry = String.Format(@"
		//		Select a0.PipeCladUniqNo FROM (
		//			select a.PipeCladUniqNo, a.ReportStatus, a.ReprocessFlag, a.Result
		//			FROM {2}.dbo.UTAODT a
		//			Join {2}.dbo.UTAOHD b on a.ParentID = b.ID
		//			Where 
		//			a.Dlt != 1 and 
		//			b.Dlt != 1 and 
		//			a.ReportStatus = 0 and
		//			a.Result = 'C' and
		//			b.ProjectID = '{0}' and
		//			b.PipeCladLineNo = '{1}' and 
		//			a.PipeCladUniqNo not in(
								
		//				select z.PipeCladUniqNo from (select x.PipeCladUniqNo, x.ReportStatus, x.ReprocessFlag FROM {2}.dbo.HydroExpanScadaDT x 
		//				WHERE x.Dlt !=1 and x.ReportStatus = ReportStatus and x.ReprocessFlag = ReprocessFlag
								
		//			) z)
		//		) a0 
		//		Join(
		//			select a.PipeCladUniqNo, a.ReportStatus, a.ReprocessFlag, a.Result
		//			FROM {2}.dbo.LPTFinalMachDT a
		//			Join {2}.dbo.LPTFinalMachHD b on a.ParentID = b.ID
		//			Where 
		//			a.Dlt != 1 and 
		//			b.Dlt != 1 and 
		//			a.ReportStatus = 0 and
		//			a.Result = 'C' and
		//			b.ProjectID = '{0}' and
		//			b.PipeCladLineNo = '{1}' and 
		//			a.PipeCladUniqNo not in(
								
		//				select z.PipeCladUniqNo from (select x.PipeCladUniqNo, x.ReportStatus, x.ReprocessFlag FROM {2}.dbo.HydroExpanScadaDT x 
		//				WHERE x.Dlt !=1 and x.ReportStatus = ReportStatus and x.ReprocessFlag = ReprocessFlag
								
		//			) z)
		//		)a1 on a0.PipeCladUniqNo = a1.PipeCladUniqNo
		//		Join(
		//			select a.PipeCladUniqNo, a.ReportStatus, a.ReprocessFlag, a.Result
		//			FROM {2}.dbo.UTLamdisDT a
		//			Join {2}.dbo.UTLamdisHD b on a.ParentID = b.ID
		//			Where 
		//			a.Dlt != 1 and 
		//			b.Dlt != 1 and 
		//			a.ReportStatus = 0 and
		//			a.Result = 'C' and
		//			b.ProjectID = '{0}' and
		//			b.PipeCladLineNo = '{1}' and 
		//			a.PipeCladUniqNo not in(
								
		//				select z.PipeCladUniqNo from (select x.PipeCladUniqNo, x.ReportStatus, x.ReprocessFlag FROM {2}.dbo.HydroExpanScadaDT x 
		//				WHERE x.Dlt !=1 and x.ReportStatus = ReportStatus and x.ReprocessFlag = ReprocessFlag
								
		//			) z)
		//		)a2 on a0.PipeCladUniqNo = a2.PipeCladUniqNo"
		//			, AM.ProjectID, AM.CladLineNo, AM.DBName);

		//		var Data = db.GetDataTableSQL(qry);
		//		var lst = new List<string>();
		//		foreach (DataRow row in Data.Rows)
		//		{
		//			lst.Add(row.ItemArray[0].ToString());
		//		}
		//		return Request.CreateResponse(HttpStatusCode.OK, lst);
		//	}
		//	catch (Exception x)
		//	{
		//		return Request.CreateResponse(HttpStatusCode.BadRequest);
		//	}

		//}
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
