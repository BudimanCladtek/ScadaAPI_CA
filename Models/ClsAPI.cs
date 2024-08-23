using System.Collections.Generic;
using System.Web.Mvc;
using System.Data;
using System;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Linq;
using System.Reflection;
using iTextSharp.text.pdf;
using System.Configuration;
using SCADA_API;

namespace adminlte.Modul
{

    #region API
    public partial class ClsAPI
    {
        public string Address { get; set; }
        public string Service { get; set; }

        public HttpContent GetMethodUserValidation(string dta)
        {
            try
            {
                //List<ClsUserModel> listUser = new List<ClsUserModel>();
                //listUser.Add(dta);
                //var Datajson = "[{\"UserName\":\"3907\",\"Password\":\"123\" }]";
                //cjson tojson = new cjson();
                //string Djson = JsonConvert.SerializeObject(listUser);
                //tojson.json = Djson;
                //string json = JsonConvert.SerializeObject(tojson);


                Cjson tojson = new Cjson();
                tojson.json = dta;
                string json = JsonConvert.SerializeObject(tojson);
                //json = json.Replace("{}", "0");

                string APIPath = "";
                string _Address = "";
                if (ConfigurationManager.AppSettings["APIPath"] != null)
                {
                    APIPath = ConfigurationManager.AppSettings["APIPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/CheckUserSystem";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }
                else
                {
                    APIPath = ConfigurationManager.AppSettings["APIScadaPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/CheckUserSystemFromScada";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }

                var _Auth = "Basic aHlzZHJvZXhwYW5kdGVzdHNjYWRhOmNsYWR0ZWtiYXRhbTIwMjA=";

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 

                //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic aHlzZHJvZXhwYW5kdGVzdHNjYWRhOmNsYWR0ZWtiYXRhbTIwMjA=");

                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetMethodScadaUserValidation(string dta)
        {
            try
            {
                Cjson tojson = new Cjson();
                tojson.json = dta;
                string json = JsonConvert.SerializeObject(tojson);

                string APIPath = ConfigurationManager.AppSettings["APIScadaPath"].ToString();
                string _Address = APIPath + "/api/ClsScada/CheckUserSystemFromScada";

                var _Auth = "Basic aHlzZHJvZXhwYW5kdGVzdHNjYWRhOmNsYWR0ZWtiYXRhbTIwMjA=";

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };

                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetProject(string dta)
        {
            try
            {
                string APIPath =Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetProject";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 

                //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic aHlzZHJvZXhwYW5kdGVzdHNjYWRhOmNsYWR0ZWtiYXRhbTIwMjA=");

                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public HttpContent GetItemForWOLScada(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetItemForWOLScada";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetItemForMachiningScada(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetItemForMachining";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public HttpContent GetItemForPAWIIScada(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetItemForPAWIIScada";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetCladLineNo(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetCladLineNo";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public HttpContent GetWelderStamp(string dta)
        {
            try
            {
                string APIPath = "";
                string _Address = "";
                if (ConfigurationManager.AppSettings["APIPath"] != null)
                {
                    APIPath = ConfigurationManager.AppSettings["APIPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/GetWelderStamp";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }
                else
                {
                    APIPath = ConfigurationManager.AppSettings["APIScadaPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/GetWelderStampFromScada";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }

                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public HttpContent GetWPS(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetWPS";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public HttpContent GetWPSPAWII(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetWPSPAWII";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public HttpContent GetWPSByNo(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetWPSByNo";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetPipeDetail(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/GetPipeDetail";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        ////PostNewReportWOL
        public HttpContent PostNewReportWOL(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/PostNewReportWOL";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        ////PostNewReportWOL
        public HttpContent PostNewReportPAW(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/PostNewReportPAW";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        ////PostReportWOLFitting
        public HttpContent PostReportWOLFitting(string dta)
        {
            try
            {
                string APIPath = Global.APIPath;
                var _Address = APIPath + "/api/ClsScada/PostReportWOLFitting";
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetWire(string dta)
        {
            try
            {
                string APIPath = "";
                string _Address = "";
                if (ConfigurationManager.AppSettings["APIPath"] != null)
                {
                    APIPath = ConfigurationManager.AppSettings["APIPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/GetWire";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }
                else
                {
                    APIPath = ConfigurationManager.AppSettings["APIScadaPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/GetWireFromScada";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                    dta = String.Concat("{\"json\":\"", dta.Replace("\"", "'"), "\"", "}");
                }
                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public HttpContent GetGas(string dta)
        {
            try
            {
                string APIPath = "";
                string _Address = "";
                if (ConfigurationManager.AppSettings["APIPath"] != null)
                {
                    APIPath = ConfigurationManager.AppSettings["APIPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/GetGas";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }
                else
                {
                    APIPath = ConfigurationManager.AppSettings["APIScadaPath"].ToString();
                    _Address = APIPath + "/api/ClsScada/GetGasFromScada";
                    //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                }

                //var _Address = "http://192.168.0.200:8080"+ "/api/ClsScada/CheckUserSystem";
                //_Address = _Address.Replace("192.168.0.200", "172.16.202.202");
                var _Auth = Global._Auth;

                HttpClient client = new HttpClient
                {
                    BaseAddress = new Uri(_Address)
                };
                /// url base + services (with param)
                /// 
                client.DefaultRequestHeaders.Add("Authorization", _Auth);
                var data = new StringContent(dta, Encoding.UTF8, "application/json");

                var response = client.PostAsync(client.BaseAddress, data);
                response.Wait();

                var result = response.Result;
                if (result.IsSuccessStatusCode)
                {
                    var output = result.Content;
                    return output;
                }
                else
                {
                    return result.Content;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
    }


    public partial class Cjson
    {
        public string json { get; set; }
    }

    public partial class ClsUserModel
    {
        public string UserName { get; set; }
        public string PassCode { get; set; }
    }
    public partial class ClsPipe
    {
        public string UserName { get; set; }
        public string PassCode { get; set; }
    }
    #endregion

}