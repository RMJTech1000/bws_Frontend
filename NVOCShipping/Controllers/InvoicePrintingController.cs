using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class InvoicePrintingController : Controller
    {
        // GET: InvoicePrinting
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult InvoicePrinting()
        {
            return View();
        }
    }
}