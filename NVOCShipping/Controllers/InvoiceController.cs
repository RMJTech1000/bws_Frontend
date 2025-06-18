using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NVOCShipping.Models;
using System.Data.SqlClient;
using System.IO;
using DataTier;
using System.Net.Mail;

namespace NVOCShipping.Controllers
{
    public class InvoiceController : Controller
    {
        // GET: Invoice
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult InvoiceDashboard()
        {
            return View();
        }
        public ActionResult Invoice()
        {
            return View();
        }

        public ActionResult InvoiceCR()
        {
            return View();
        }

        public ActionResult BLTariffView()
        {
            return View();
        }

        public ActionResult BLTariffDetails()
        {
            return View();
        }

        public ActionResult InvoiceView()
        {
            return View();
        }
        public ActionResult FinalInvoice()
        {
            return View();
        }

        public ActionResult ImpInvoiceView()
        {
            return View();
        }

        public ActionResult ImpInvoiceDashboard()
        {
            return View();
        }
        public ActionResult SlotBillingView()
        {
            return View();
        }

        public ActionResult SlotBilling()
        {
            return View();
        }

        public ActionResult PortBillingView()
        {
            return View();
        }

        public ActionResult PortBilling()
        {
            return View();
        }

        public ActionResult VendorInvoiceApproval()
        {
            return View();
        }
        public ActionResult InvoiceDashboardViewPages()
        {
            return View();
        }
        [HttpPost]
        public ContentResult Upload()
        {
            string path = Server.MapPath("~/SlotBillAttachments/");


            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var FileNamev = "";

            foreach (string key in Request.Files)
            {
                Random rd = new Random();
                FileNamev += rd.Next(1000).ToString();

                HttpPostedFileBase postedFile = Request.Files[key];
                FileNamev += "_"  + postedFile.FileName;
                postedFile.SaveAs(path + FileNamev);


            }

            return Content(FileNamev);
        }

        public ActionResult DownloadAttachment(string fileName)
        {
            Attachment(fileName);
            return View();
        }
        private void Attachment(string fileName)
        {
            string path = Server.MapPath("~/SlotBillAttachments/" + fileName);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + fileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(path);
            Response.Flush();
            Response.End();
        }
    }
}