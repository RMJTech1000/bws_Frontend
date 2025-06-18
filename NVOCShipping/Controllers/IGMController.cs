using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class IGMController : Controller
    {
        // GET: IGM
        public ActionResult ManageManifest()
        {
            return View();
        }

        public ActionResult ManageManifestView()
        {
            return View();
        }

        public ActionResult CustomsManifest()
        {
            return View();
        }
    }
}