using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class MastersController : Controller
    {
        // GET: Masters
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult GeneralMaster()
        {
            return View();
        }
        public ActionResult GeneralMasterView()
        {
            return View();
        }
        public ActionResult NotificationMaster()
        {
            return View();
        }
        public ActionResult NotificationMasterView()
        {
            return View();
        }

        public ActionResult MISExchangeRate()
        {
            return View();
        }

        public ActionResult MISExchangeRateView()
        {
            return View();
        }

        public ActionResult Certifitcate()
        {
            return View();
        }

        public ActionResult Certifitcateview()
        {
            return View();
        }
    }
}