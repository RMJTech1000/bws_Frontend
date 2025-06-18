using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;

namespace NVOCShipping.Controllers
{
    public class OnlineController : Controller
    {
        // GET: Online
        public ActionResult SwitchBLRequest()
        {
            return View();
        }
        public ActionResult SwitchBL()
        {
            return View();
        }

        public ActionResult PaymentConfirmation()
        {
            return View();
        }
        public ActionResult PaymentConfirmationView()
        {
            return View();
        }

        public ActionResult ViewBLDocument()
        {
            return View();
        }

        public ActionResult BLDownloadAttachment(string fileName)
        {
            Attachment(fileName);
            return View();
        }
        private void Attachment(string fileName)
        {

            string path = Server.MapPath("~/BLFileAttached/" + fileName);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + fileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(path);
            Response.Flush();
            Response.End();
        }
    }
}
