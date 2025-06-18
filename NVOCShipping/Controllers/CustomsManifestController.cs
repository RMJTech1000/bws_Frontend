using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class CustomsManifestController : Controller
    {
        // GET: CustomsManifest
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult CustomsManifest()
        {
            ViewBag.Message = "Your Vessel.";

            return View();
        }
    }
}