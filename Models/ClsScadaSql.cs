using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Web;
using Newtonsoft.Json;

namespace CORSYS_API.Models
{
    public class ClsScadaSql
    {
        SqlConnection con;
        SqlCommand cmd;
        SqlDataReader dr;
        SqlDataAdapter da;
        public void OpenConn(string database)
        {
            string connetionString = ConfigurationManager.AppSettings["ConnectionStr"].ToString();

            con = new SqlConnection(connetionString);
            try
            {
                con.Open();
                if (database.Length != 0)
                    con.ChangeDatabase(database);
            }
            catch (Exception x)
            {
                throw x;
            }
        }

        public SqlDataReader GetDataReaderSQL(string database, string query)
        {
            try
            {
                OpenConn(database);
                cmd = new SqlCommand(query, con);
                dr = cmd.ExecuteReader();
                return dr;
            }
            catch (Exception x)
            {
                throw x;
            }
            finally
            {
                con.Close();
            }
        }

        public DataTable GetDataTableSQL(string database, string query, int timeout=0)
        {
            try
            {
                OpenConn(database);
                DataTable dt = new DataTable();
                cmd = new SqlCommand(query, con);
                if(timeout==0)
                    cmd.CommandTimeout = 300;
                else
                    cmd.CommandTimeout = timeout;
                da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                return dt;
            }
            catch (Exception x)
            {
                throw x;
            }
            finally
            {
                da.Dispose();
                con.Close();
            }
        }
        public Object ToObject(DataTable dt)
        {
            try
            {
                var rowcount = dt.Rows.Count;
                var columncount = dt.Columns.Count;
                var obj = new object[rowcount, columncount];
                int rw = 0;
                int cl = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    cl = 0;
                    
                    foreach (DataColumn dc in dt.Columns)
                    {
                        var temp = dr[dc.ColumnName].ToString();
                        float test = 0;
                        if (float.TryParse(temp, out test))
                        {
                            if (temp.Length > 10 || temp.IndexOf('0')==0)
                            {
                                // text
                                obj[rw, cl] = "'" + temp;
                            }
                            else
                            {
                                // angka
                                obj[rw, cl] = test;
                            }
                        }
                        else
                        {
                            // general
                            obj[rw, cl] = dr[dc.ColumnName];
                        }
                        cl++;
                    }
                    rw++;
                }
                return obj;
            }
            catch (Exception x)
            {
                throw x;
            }
        }
        public Object ToObjectWithHeader(DataTable Data, List<string> colHeader,List<string> colHeaderName, bool addNo)
        {
            try
            {
                var rowcount = Data.Rows.Count;
                var columncount = colHeader.Count();
                var obj = new object[rowcount, columncount];
                int Temp =0;
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
                                Temp = Enumerable.Range(0, colHeaderName.Count) .Where(i => colHeaderName[i] == dc.ColumnName).FirstOrDefault();
                                obj[rw, Temp] = "'" + dr[dc.ColumnName];
                                cl++;
                            }
                        }
                        catch (Exception x)
                        {
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
        public DataTable ConvertToDataTable(Object[] array)
        {
            //for (int i = 0; i <= array.Length; i++)
            //{
            //    var aa = array.AsEnumerable();
            //    var aaa=  aa.[i];
            //    var a = new System.Collections.Generic.Mscorlib_DictionaryDebugView<string, object>(((object[])array.AsEnumerable())[i]).Items[0].Key;
            //    foreach (var a in o)
            //    {

            //    }
            //}
            PropertyInfo[] properties = array.GetType().GetElementType().GetProperties();
            DataTable dt = CreateDataTable(properties);
            if (array.Length != 0)
            {
                foreach (object o in array)
                    FillData(properties, dt, o);
            }
            return dt;
        }

        private DataTable CreateDataTable(PropertyInfo[] properties)

        {
            DataTable dt = new DataTable();
            DataColumn dc = null;
            foreach (PropertyInfo pi in properties)
            {
                dc = new DataColumn();
                dc.ColumnName = pi.Name;
                dc.DataType = pi.PropertyType;
                dt.Columns.Add(dc);
            }
            return dt;
        }

        private void FillData(PropertyInfo[] properties, DataTable dt, Object o)
        {
            DataRow dr = dt.NewRow();
            foreach (PropertyInfo pi in properties)
            {
                dr[pi.Name] = pi.GetValue(o, null);
            }
            dt.Rows.Add(dr);
        }


        public DataTable JsonToDatatable(string json)
        {
            DataTable Data = (DataTable)JsonConvert.DeserializeObject(json, (typeof(DataTable)));
            return Data;
        }
        public string JsonSerialize<T>(IEnumerable<T> data)
        {
            string json = JsonConvert.SerializeObject(data);
            return json;
        }

        public string RunQuery(string strQry, string db)
        {
            try
            {
                OpenConn(db);
                cmd = new SqlCommand(strQry, con);
                cmd.CommandTimeout = 300;
                cmd.CommandText = strQry;
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                return "Success";
            }
            catch(Exception e)
            {
                return e.Message;
            }
        }

    }

}