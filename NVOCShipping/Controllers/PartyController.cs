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
    public class PartyController : Controller
    {
        SqlConnection con;
        SqlDataAdapter adapter;
        SqlCommand cmd;
        static String connectionString = @"Data Source=103.48.51.212,1232;Initial Catalog=NVODb;Integrated Security=false;User id=testmax;Password=max@2019;";

        // GET: Party
        public ActionResult Index()
        {
            return View();
        }
        #region anand
        public ActionResult Customer()
        {


            ViewBag.Message = "Your Customer.";

            return View();
        }
        public ActionResult CustomerView()
        {
            ViewBag.Message = "Your Customer View.";

            return View();
        }

        public ActionResult KYCApprovalView()
        {
            ViewBag.Message = "Your Customer View.";

            return View();
        }
        public ActionResult KYCApproval()
        {
            ViewBag.Message = "Your Customer View.";

            return View();
        }
        #endregion


        #region Ganesh
        public ActionResult Agency(string id)
        {
            MasterModel BusData = new MasterModel();
            BusData.MasterDetails = new Masters();
            if (id != null)
                BusData.MasterDetails.ID = Int32.Parse(id);
            else
                BusData.MasterDetails.ID = 0;
            ViewBag.Message = "Your Agency.";


            return View(BusData);

        }
        public ActionResult AgencyView()
        {
            ViewBag.Message = "Your Agency.";

            return View();
        }

        public ActionResult OnlinePortalView()
        {
            ViewBag.Message = "Your OnlinePortal.";

            return View();
        }

        public ActionResult OnlinePortal()
        {
            ViewBag.Message = "Your OnlinePortal.";

            return View();
        }

        #endregion

        [HttpPost]
        public ContentResult Upload()
        {
            string path = Server.MapPath("~/CustomerFileAttach/");


            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var FileNamev = "";
            foreach (string key in Request.Files)
            {
                Random rd = new Random();
                FileNamev += rd.Next(1000).ToString();
                
                HttpPostedFileBase postedFile =  Request.Files[key];
               
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

            string path = Server.MapPath("~/CustomerFileAttach/" + fileName);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + fileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(path);
            Response.Flush();
            Response.End();
        }
    }
}