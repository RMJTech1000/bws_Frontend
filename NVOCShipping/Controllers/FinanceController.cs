using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NVOCShipping.Models;
using System.Data.SqlClient;
using System.IO;


namespace NVOCShipping.Controllers

{
    public class FinanceController : Controller
    {
        #region ganesh (finance- chargecode,taxdecl,taxengine,credit control)
        public ActionResult ChargeCode()
        {
            

            return View();

        }
        public ActionResult ChargeCodeView()
        {
            
            return View();

        }

        public ActionResult TaxDeclaration()
        {


            return View();

        }

        public ActionResult TaxDeclarationView()
        {


            return View();

        }
        public ActionResult TaxEngine()
        {

            return View();

        }
        public ActionResult TaxNames()
        {

            return View();

        }
        public ActionResult TaxNamesView()
        {

            return View();

        }

        public ActionResult CreditControl()
        {

            return View();

        }
        public ActionResult CreditControlView()
        {

            return View();

        }

        public ActionResult BankMasterView()
        {

            return View();

        }
        public ActionResult BankMaster()
        {

            return View();

        }
        #endregion

        #region anand
        public ActionResult GLMaster()
        {
            return View();
        }
        public ActionResult GLMasterView()
        {
            return View();
        }
        public ActionResult GLMapping()
        {
            return View();
        }
        public ActionResult GLMappingView()
        {
            return View();
        }
        public ActionResult ReceiptLocal()
        {
            return View();
        }
        public ActionResult ReceiptOverseas()
        {
            return View();
        }
        public ActionResult InvoiceSetOff()
        {
            return View();
        }
        public ActionResult InvoiceSetOffView()
        {
            return View();
        }
       
        #endregion


        [HttpPost]
        public ContentResult Upload()
        {
            string path = Server.MapPath("~/FinanceUpload/");


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

                FileNamev += "_" + postedFile.FileName;
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

            string path = Server.MapPath("~/FinanceUpload/" + fileName);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + fileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(path);
            Response.Flush();
            Response.End();
        }
    }

}