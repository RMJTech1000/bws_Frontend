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
    public class MNRController : Controller
    {
        // GET: MNR

        MyMNRRepair Data = new MyMNRRepair();
        public ActionResult Index()
        {
            return View();
        }

     
        public ActionResult MNRDamageView()
        {
            ViewBag.Message = "Your MNRDamageView.";

            return View();
        }

        public ActionResult MNRDamage()
        {
            ViewBag.Message = "Your MNRDamage.";

            return View();
        }

        public ActionResult MNRRepairView()
        {
            ViewBag.Message = "Your MNRRepairView.";

            return View();
        }

        public ActionResult MNRRepair()
        {
            ViewBag.Message = "Your MNRRepair.";

            return View();
        }

        public ActionResult MNRLocationView()
        {
            ViewBag.Message = "Your MNRLocationView.";

            return View();
        }

        public ActionResult MNRLocation()
        {
            ViewBag.Message = "Your MNRLocation.";

            return View();
        }

        public ActionResult MNRComponentView()
        {
            ViewBag.Message = "Your MNRComponentView.";

            return View();
        }

        public ActionResult MNRComponent()
        {
            ViewBag.Message = "Your MNRComponent.";

            return View();
        }
        public ActionResult MNRMaterialView()
        {
            ViewBag.Message = "Your MNRMaterialView.";

            return View();
        }

        public ActionResult MNRMaterial()
        {
            ViewBag.Message = "Your MNRMaterial";

            return View();
        }
        public ActionResult MNRMaintenanceRepairView()
        {
            ViewBag.Message = "Your MNRMaintenance.";

            return View();
        }
        public ActionResult MNRMaintenanceRepair()
        {
            ViewBag.Message = "Your MNRMaintenance.";

            return View();
        }
        public ActionResult MNRVendorRepair()
        {
            ViewBag.Message = "Your MNRVendorRepair.";

            return View();
        }

        public ActionResult MNRTariffView()
        {
            ViewBag.Message = "Your MNRTariffView.";

            return View();
        }

        public ActionResult MNRTariff()
        {
            ViewBag.Message = "Your MNRTariffView.";

            return View();
        }
        public ActionResult MNRWeeklyReport()
        {
            ViewBag.Message = "Your MNRWeeklyReport.";

            return View();
        }
       
        [HttpPost]
        public ContentResult Upload(string CntrNo)
        {
            string path = Server.MapPath("~/MNRAttachments/");
            //string path = Server.MapPath("http://oclattach.oceanus-lines.com/UploadFolder/");
           // string path = System.Web.Hosting.HostingEnvironment.MapPath("http://oclattach.oceanus-lines.com/UploadFolder/");
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
                FileNamev += "_" + CntrNo + "_" + postedFile.FileName;
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
            string path = Server.MapPath("~/MNRAttachments/" + fileName);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + fileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(path);
            Response.Flush();
            Response.End();
        }
    }
}
