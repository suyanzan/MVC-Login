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
            string account = post["username1"];//抓前端id
            string password = post["Password1"];//抓前端id
            //驗證密碼
            if (CheckPassword(account, password,domain))
            {
                Session["account"] = account;
                //Response.Redirect("~/TCBS/Index");
                //return new EmptyResult();
                return RedirectToAction("Index", "TCBS", new { account = Session["account"].ToString() });
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
    }

}
