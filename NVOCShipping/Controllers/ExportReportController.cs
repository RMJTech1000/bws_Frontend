using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class ExportReportController : Controller
    {
        // GET: ExportReport
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult TerminalDepReport()
        {
            return View();
        }
        public ActionResult FreightManifestReport()
        {
            return View();
        }
        public ActionResult SurrenderNotice()
        {
            return View();
        }
        public ActionResult FreightManifestDetails()
        {
            return View();
        }
        public ActionResult TerminalDepReportDetails()
        {
            return View();
        }

        public ActionResult LoadListReport()
        {
            return View();
        }
        public ActionResult MultiPortsTDR()
        {
            return View();
        }

        public ActionResult MultiPortsTDRView()
        {
            return View();
        }
        public ActionResult FreightManifestGlobal()
        {
            return View();
        }
        public ActionResult FreightManifestGlobalDetails()
        {
            return View();
        }

        public ActionResult ScreenManifest()
        {
            return View();
        }
    }
}