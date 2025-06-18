using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class AllReportsController : Controller
    {
        // GET: AllReports
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult AccReportDashboard()
        {
            return View();
        }
        //public ActionResult AccReportCustomerView()
        //{
        //    return View();
        //}
        //public ActionResult AccReportCustomer()
        //{
        //    return View();
        //}
        public ActionResult ImportReportDashboard()
        {
            return View();
        }
        public ActionResult ExpReportDashboard()
        {
            return View();
        }
        public ActionResult InventoryReportDashboard()
        {
            return View();
        }
        public ActionResult SoaReportDashboard()
        {
            return View();
        }
        public ActionResult TradeReportsDashboard()
        {
            return View();
        }
    }
}