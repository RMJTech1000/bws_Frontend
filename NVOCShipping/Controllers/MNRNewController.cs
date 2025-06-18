using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class MNRNewController : Controller
    {
        // GET: MNRNew
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult MNRManagement()
        {
            ViewBag.Message = "Your MNRManagement.";

            return View();
        }

        public ActionResult MNRHistoryView()
        {
            ViewBag.Message = "Your MNRHistoryView.";

            return View();
        }
        public ActionResult MNRHistory()
        {
            ViewBag.Message = "Your MNRHistory.";

            return View();
        }

        public ActionResult MNRReport()
        {
            ViewBag.Message = "Your MNRReport.";

            return View();
        }
        public ActionResult MNRDepot()
        {
            ViewBag.Message = "Your MNRDepot.";

            return View();
        }

        public ActionResult MNRDepotHistoryView()
        {
            ViewBag.Message = "Your MNRDepotHistoryView.";

            return View();
        }
        public ActionResult MNRDepotHistory()
        {
            ViewBag.Message = "Your MNRDepotHistory.";

            return View();
        }
    }
}