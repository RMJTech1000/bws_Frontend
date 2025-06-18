using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class ManageManifestController : Controller
    {
        // GET: ManageManifest
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ManageManifestView()
        {
            ViewBag.Message = "Your Vessel.";

            return View();
        }
        public ActionResult ManageManifest()
        {
            ViewBag.Message = "Your Vessel.";

            return View();
        }
    }
}