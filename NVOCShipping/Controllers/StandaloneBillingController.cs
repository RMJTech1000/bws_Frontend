using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class StandaloneBillingController : Controller
    {
        // GET: StandaloneBilling
        public ActionResult StandaloneBillingView()
        {
            return View();
        }
        public ActionResult StandaloneBilling()
        {
            return View();
        }
        public ActionResult CRStandaloneBillingView()
        {
            return View();
        }
        public ActionResult CRStandaloneBilling()
        {
            return View();
        }
    }
}