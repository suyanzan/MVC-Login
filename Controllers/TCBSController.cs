using DH_MVC.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace DH_MVC.Controllers
{
    public class TCBSController : Controller
    {
        string connString = @"Data Source=suyanzan-pc\mesdh2;Initial Catalog=MESDH2;Persist Security Info=True;User ID=dh2test;Password=Date1128";
        SqlConnection conn = new SqlConnection();
        // GET: TCBS
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(TCBSs data)
        {
            try
            {
                if (data.Product == null || data.Product == "" || data.Version == null || data.Version == "" || data.DPF == null || data.DPF == "" || data.FileList == null || data.FileList == "")
                {
                    //驗證密碼
                    ViewBag.Msg = "資料不完整再請檢查!!";
                    return View();
                }
                else
                {
                    if (AddTCBSData(data))
                    {
                        ViewBag.Msg = "新增資料成功";
                        Thread.Sleep(1);
                        Response.Redirect("~/TCBS/JobList");
                        return new EmptyResult();
                    }
                    else
                    {
                        //error訊息
                        ViewBag.Msg = "註冊失敗";
                        return View();
                    }
                }
                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Msg = ex.Message.ToString();
                return View();
            }
        }
        public void Connect()
        {
            conn.ConnectionString = connString;
            if (conn.State != ConnectionState.Open)
                conn.Open();
        }
        public void Disconnect()
        {
            //conn.ConnectionString = connString;
            if (conn.State != ConnectionState.Closed)
                conn.Close();
        }
        public bool AddTCBSData(TCBSs data)
        {
            try
            {
                Connect();
                string strSQL = @"INSERT INTO dbo.TCBS_Seamless (Product, Version, DPF, FileList,Account,InsertTime,Cancel)
                          VALUES (@Product, @Version, @DPF, @FileList, @Account, GETDATE(),'N')";
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                cmd.Parameters.Add("@Product", SqlDbType.NVarChar).Value = data.Product;
                cmd.Parameters.Add("@Version", SqlDbType.NVarChar).Value = data.Version;
                cmd.Parameters.Add("@DPF", SqlDbType.NVarChar).Value = data.DPF;
                cmd.Parameters.Add("@FileList", SqlDbType.NVarChar).Value = data.FileList;
                cmd.Parameters.Add("@Account", SqlDbType.NVarChar).Value = Session["account"].ToString();
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                throw;
                //return false;
            }
            finally
            {
                Disconnect();
            }
        }


        /*JobList*/
        // Get: JobList
        public ActionResult JobList()
        {
            return View();
        }
        [HttpPost]
        public ActionResult GetTCBSData()
        {
            try
            {
                Connect();
                string strSQL = @"select Product,Version,DPF,FileList from dbo.TCBS_Seamless where account = @Account and Cancel = 'N'";
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                cmd.Parameters.Add("@Account", SqlDbType.NVarChar).Value = Session["account"].ToString();
                List<TCBSs> TCBSs = new List<TCBSs>();
                string result = "";
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        TCBSs datalist = new TCBSs();
                        datalist.Product = dr["Product"].ToString();
                        datalist.Version = dr["Version"].ToString();
                        datalist.DPF = dr["DPF"].ToString();
                        datalist.FileList = dr["FileList"].ToString();
                        TCBSs.Add(datalist);
                    }
                }
                if (TCBSs == null)
                {
                    //讀取資料庫錯誤
                    return Json(result);
                }
                else
                {
                    result = JsonConvert.SerializeObject(TCBSs);
                    return Json(result);
                }
            }
            catch (Exception ex)
            {
                throw;
                //return false;
            }
            finally
            {
                Disconnect();
            }
        }
        [HttpGet]
        public ActionResult PutTCBSData(string Product, string Version, string DPF, string FileList)
        {
            try
            {
                Connect();
                string strSQL = @"update dbo.TCBS_Seamless set Cancel = 'Y',CancelTime = GETDATE() where account = @Account 
                                and Product=@Product and Version = @Version and DPF = @DPF and FileList = @FileList";
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                cmd.Parameters.Add("@Account", SqlDbType.NVarChar).Value = Session["account"].ToString();
                cmd.Parameters.Add("@Product", SqlDbType.NVarChar).Value = Product;
                cmd.Parameters.Add("@Version", SqlDbType.NVarChar).Value = Version;
                cmd.Parameters.Add("@DPF", SqlDbType.NVarChar).Value = DPF;
                cmd.Parameters.Add("@FileList", SqlDbType.NVarChar).Value = FileList;
                cmd.ExecuteNonQuery();
                //List<TCBSs> TCBSs = new List<TCBSs>();
                //TCBSs datalist = new TCBSs();
                //datalist.Product = Product;
                //datalist.Version = Version;
                //datalist.DPF = DPF;
                //datalist.FileList = FileList;
                //TCBSs.Add(datalist);
                //string result = JsonConvert.SerializeObject(TCBSs);
                //return Json(result);
                return Json(new { success = true }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                throw;
                //return false;
            }
            finally
            {
                Disconnect();
            }
        }
    }
}