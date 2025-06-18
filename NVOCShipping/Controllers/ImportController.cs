using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class ImportController : Controller
    {
        // GET: Import
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ShipmentDetails()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult PreAlert()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult CAN()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult DeliveryOrder()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult CustomsReports()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult DetDem()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult ContainerDetails()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult ImportView()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult ImpDestinationCharges()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult ImpTariffview()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult ImpTariffCast()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}