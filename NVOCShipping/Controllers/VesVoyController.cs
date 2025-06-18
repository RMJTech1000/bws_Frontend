using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class VesVoyController : Controller
    {
        // GET: VesVoy
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Vessel()
        {
            ViewBag.Message = "Your Vessel.";

            return View();
        }
        public ActionResult VesselView()
        {
            ViewBag.Message = "Your VesselView.";

            return View();
        }
        public ActionResult Voyage()
        {
            ViewBag.Message = "Your Voyage.";

            return View();
        }
        public ActionResult VoyageView()
        {
            ViewBag.Message = "Your VoyageView.";

            return View();
        }
        public ActionResult VoyageDetails()
        {
            ViewBag.Message = "Your VoyageDetails.";

            return View();
        }
        public ActionResult VoyageDetailsView()
        {
            ViewBag.Message = "Your VoyageDetails.";

            return View();
        }

        public ActionResult VoyageAllocation()
        {
            ViewBag.Message = "Your VoyageDetails.";

            return View();
        }

        public ActionResult VoyageAllocationView()
        {
            ViewBag.Message = "Your VoyageDetails.";

            return View();
        }

        public ActionResult ServiceSetup()
        {
            ViewBag.Message = "Your ServiceSetup.";

            return View();
        }
        public ActionResult ServiceSetupView()
        {
            ViewBag.Message = "Your ServiceSetup.";

            return View();
        }
        public ActionResult VoyageDetailsNew()
        {
            ViewBag.Message = "Your VoyageDetailsNew.";

            return View();
        }

        public ActionResult VoyageDetailsNewView()
        {
            ViewBag.Message = "Your VoyageDetailsNew.";

            return View();
        }

        public ActionResult VoyageOpeningView()
        {
            ViewBag.Message = "Your VoyageLocking.";

            return View();
        }
        public ActionResult VoyageOpening()
        {
            ViewBag.Message = "Your VoyageLocking.";

            return View();
        }

        public ActionResult VoyageLockingView()
        {
            ViewBag.Message = "Your VoyageLocking.";

            return View();
        }
        public ActionResult VoyageLocking()
        {
            ViewBag.Message = "Your VoyageLocking.";

            return View();
        }

        public ActionResult GatewayVoyage()
        {
            ViewBag.Message = "Your GatewayVoyage.";

            return View();
        }

        public ActionResult GatewayVoyageView()
        {
            ViewBag.Message = "Your GatewayVoyageView.";
            return View();
        }
    }
}