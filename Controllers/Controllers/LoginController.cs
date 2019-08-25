using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DH_MVC.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(FormCollection post)
        {
            string domain = post["domain"];//抓前端id,name屬性也要有
            string account = post["username"];//抓前端id
            string password = post["Password"];//抓前端id
            //驗證密碼
            if (CheckPassword(account, password, domain))
            {
                Session["account"] = account;
                Session["UserName"] = "Yen";
                //Response.Redirect("~/TCBS/Index");
                //return new EmptyResult();
                //return View();
                return RedirectToAction("Index", "TCBS", new { account = Session["account"].ToString(), UserName = Session["UserName"].ToString()});
            }
            else
            {
                ViewBag.Msg = "登入失敗...";
                return View();
            }
        }
        public static  bool CheckPassword(string account,string password,string domain)
        {
            if (account == "123" && password == "123" && domain == "AS")
            {
                
                return true;
            }
            else
                return false;
        }
        public ActionResult LogOut()
        {
            Session.Abandon(); // it will clear the session at the end of request
            return RedirectToAction("Index", "Home");
        }
    }

}