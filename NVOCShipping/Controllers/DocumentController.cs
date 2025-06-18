using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DataManager;
using DataTier;
namespace NVOCShipping.Controllers
{
    public class DocumentController : Controller
    {
        // GET: Document
        public ActionResult Booking()
        {
            return View();
        }

        public ActionResult BookingSearch()
        {
            return View();
        }
        public ActionResult CRO()
        {
            return View();
        }

        public ActionResult CROView()
        {
            return View();
        }

        public ActionResult BOLView()
        {
            return View();
        }

        public ActionResult BOL()
        {
            return View();
        }

        public ActionResult BLRelease()
        {
            return View();
        }

        public ActionResult RRBookingView()
        {
            return View();
        }
        public ActionResult BLTariff()
        {
            return View();
        }
        public ActionResult BLChargeCorrector()
        {
            return View();
        }

        public ActionResult BLChargeCorrectorSearch()
        {
            return View();
        }
        public ActionResult CorrectorView()
        {
            return View();
        }

        public ActionResult ContainerPickup()
        {
            return View();
        }

        public ActionResult ContainerPickupSearch()
        {
            return View();
        }

        public ActionResult Test()
        {
            return View();
        }
        public JsonResult GetCountryList(string searchTerm)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccessNew> st = cm.CustomerMasterNewEx(searchTerm);
            return Json(st, JsonRequestBehavior.AllowGet);
        }

        public ActionResult CROFinalView()
        {
            return View();
        }

        public ActionResult CROFinalNew()
        {
            return View();
        }

        public ActionResult ExpDet()
        {
            return View();
        }

        public ActionResult ExpTariffView()
        {
            return View();
        }

        public ActionResult ExptariffCast()
        {
            return View();
        }

        public ActionResult Corrector()
        {
            return View();
        }
        public ActionResult CorrectionMemo()
        {
            return View();
        }
        public ActionResult CorrectionMemoView()
        {
            return View();
        }
        public ActionResult BLExceptionsDashboard()
        {
            return View();
        }

        public ActionResult CommisionCorrcetion()
        {
            return View();
        }
    }
}