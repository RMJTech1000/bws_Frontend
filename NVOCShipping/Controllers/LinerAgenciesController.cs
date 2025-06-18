using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NVOCShipping.Models;

namespace NVOCShipping.Controllers
{
    public class LinerAgenciesController : Controller
    {
        // GET: LinerAgencies
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult LinerRateRequestView()
        {
            return View();
        }
        public ActionResult LinerRateRequest()
        {
            return View();
        }
        public ActionResult LinerBOLView()
        {
            return View();
        }
        public ActionResult LinerBOL()
        {
            return View();
        }
        public ActionResult LinerBLRelease()
        {
            return View();
        }

        #region ganesh (liner bl numbering)
        public ActionResult LinerBLNumbering()
        {
            return View();
        }
        public ActionResult LinerBLNumberingView()
        {
            return View();
        }
        #endregion
    }
}