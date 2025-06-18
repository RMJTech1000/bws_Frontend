using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class BLQueryController : Controller
    {
        // GET: BLQuery
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult BLQueryView()
        {
            return View();
        }
        public ActionResult BLQueryDetails()
        {
            return View();
        }
        public ActionResult BLQueryInvoice()
        {
            return View();
        }
        public ActionResult BLQueryReceipt()
        {
            return View();
        }
    }
}