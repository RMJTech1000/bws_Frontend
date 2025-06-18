using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NVOCShipping.Controllers
{
    public class NotificationController : Controller
    {
        // GET: Notification
        public ActionResult ExemptionApprovalView()
        {
            return View();
        }
        public ActionResult ExemptionApproval()
        {
            return View();
        }

        public ActionResult TemplateDashboardUpload()
        {
            return View();
        }
        public ActionResult BLOpen()
        {
            return View();
        }

        public ActionResult SlotAssignRequestView()
        {
            return View();
        }
        public ActionResult SlotAssignRequest()
        {
            return View();
        }
    }
}