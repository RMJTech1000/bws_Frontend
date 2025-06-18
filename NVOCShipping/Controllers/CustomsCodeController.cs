using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class CustomsCodeController : Controller
    {
        // GET: CustomsCode
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult CustomsCode()
        {
            ViewBag.Message = "Your Vessel.";

            return View();
        }
    }
}