using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class VendorBillController : Controller
    {
        // GET: VendorBill
        public ActionResult VendorBillCreate()
        {
            return View();
        }

        public ActionResult VendorBillCreateSearch()
        {
            return View();
        }
    }
}