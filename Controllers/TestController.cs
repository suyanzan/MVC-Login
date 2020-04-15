using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace CRUD_Inline.Controllers
{
    public class TestController : Controller
    {
        /// <summary>
        /// 自訂類別
        /// </summary>
        public class Person
        {
            public string Name { get; set; }

        }

        // GET: Test
        public ActionResult Index()
        {
            return View();
        }

        //給Ajax呼叫用 
        [HttpPost]
        public ActionResult GetData(int ID, Person person, List<Person> persons)
        {
            StringBuilder sb = new StringBuilder();
            if (persons != null && persons.Count > 0)
            {
                foreach (Person obj in persons)
                {
                    sb.Append(obj.Name + ",");
                }
            }
            return Content($"ID:{ID},person.Name:{person.Name},persons.Count:{persons.Count},persons.Names:{sb.ToString()}");
        }
    }
}