using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NVOCShipping.Models;
using System.IO;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using DataManager;

namespace NVOCShipping.Controllers
{
    public class SalesController : Controller
    {
        ExportReportManager RegManage = new ExportReportManager();
        // GET: Sales
        public ActionResult Index()
        {
            return View();
        }
        #region Muthu
        public ActionResult MRG()
        {
            ViewBag.Message = "Your MRG.";

            return View();
        }
        public ActionResult MRGView()
        {
            ViewBag.Message = "Your MRGView.";

            return View();
        }
        public ActionResult Slot()
        {
            ViewBag.Message = "Your Slot.";

            return View();
        }
        public ActionResult SlotView()
        {
            ViewBag.Message = "Your Slot View.";

            return View();
        }


        public ActionResult TariffMaster()
        {
            ViewBag.Message = "Your Traiff View.";

            return View();
        }

        public ActionResult TariffView()
        {
            ViewBag.Message = "Your Traiff View.";

            return View();
        }

        public ActionResult RR()
        {

            ViewBag.Message = "Your Raterequest.";
            return View();
        }

        public ActionResult RRContainer()
        {

            ViewBag.Message = "Your Raterequest.";
            return View();
        }
        public ActionResult Raterequest()
        {

            ViewBag.Message = "Your Raterequest.";
            return View();
        }

        public ActionResult RRMRGSLOT()
        {

            ViewBag.Message = "Your Raterequest.";
            return View();
        }

        public ActionResult RRLocalCharges()
        {

            ViewBag.Message = "Your Raterequest.";
            return View();
        }

        public ActionResult RRRebate()
        {
            ViewBag.Message = "Your Raterequest.";
            return View();
        }
        public ActionResult RRConfirmation()
        {

            ViewBag.Message = "Your Raterequest.";
            return View();
        }


        public ActionResult RRNotification()
        {
            ViewBag.Message = "Your RRNotification.";
            return View();
        }



        public ActionResult btnPDF(string id)
        {
          
            string pdfpath = HttpRuntime.AppDomainAppPath + "\\RRFiles\\";
            string filePath = pdfpath + "Invoice.pdf";
            return File(filePath, "application/pdf", "Invoice.pdf");
        }
        public ActionResult RaterequestView()
        {
            ViewBag.Message = "Your RaterequestView.";

            return View();
        }
        public ActionResult WaiverRequest()
        {
            ViewBag.Message = "Your RaterequestView.";

            return View();
        }

        public ActionResult WaiverRequestView()
        {
            ViewBag.Message = "Your RaterequestView.";

            return View();
        }


        public ActionResult RRForm()
        {
            ViewBag.Message = "Your RRForm.";
            return View();
        }

        #endregion

        [HttpPost]
        public ContentResult Upload()
        {
            string path = Server.MapPath("~/RRFile/");


            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var FileNamev = "";
            foreach (string key in Request.Files)
            {
                Random rd = new Random();
                FileNamev += rd.Next(10000).ToString();

                HttpPostedFileBase postedFile = Request.Files[key];
               
                FileNamev += "_" + postedFile.FileName;
                postedFile.SaveAs(path + FileNamev);

            }

            return Content(FileNamev);
        }

        [HttpPost]
        public ContentResult UploadWaiver()
        {
            string path = Server.MapPath("~/WaiverFileAttach/");


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
                postedFile.SaveAs(path + postedFile.FileName);
                FileNamev += "_" + postedFile.FileName;

            }

            return Content(FileNamev);
        }

        public ActionResult CommissionContractView()
        {
            ViewBag.Message = "Your CommissionContractView.";

            return View();
        }
        public ActionResult CommissionContract()
        {
            ViewBag.Message = "Your CommissionContract.";

            return View();
        }
        public ActionResult IHCHaulageTariff()
        {
            ViewBag.Message = "Your IHCHaulageTariff.";

            return View();
        }
        public ActionResult IHCHaulageTariffView()
        {
            ViewBag.Message = "Your IHCHaulageTariffView.";

            return View();
        }
        public ActionResult EmptyYardCntrCostView()
        {
            ViewBag.Message = "Your EmptyYardCntrCostView.";

            return View();
        }
        public ActionResult EmptyYardCntrCost()
        {
            ViewBag.Message = "Your EmptyYardCntrCost.";

            return View();
        }
        public ActionResult DownloadAttachment(string fileName)
        {
            Attachment(fileName);
            return View();
        }
      
        private void Attachment(string fileName)
        {
            var fileNamev = fileName.Replace("Ash_value", "#");
            string path = Server.MapPath("~/RRFile/" + fileNamev);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + fileNamev);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(path);
            Response.Flush();
            Response.End();
        }


        public ActionResult BOLCntrDownloadAttachment(string BkgId, string BLID)
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("BLContainers");
            ExcelRange r;
            ws.Cells["A1"].Value = "S.No.";
            ws.Cells["B1"].Value = "CONTAINERNO";
            ws.Cells["C1"].Value = "SIZE";
            ws.Cells["D1"].Value = "ISOCODE";
            ws.Cells["E1"].Value = "SEALNUMBER";
            ws.Cells["F1"].Value = "NOOFPACKAGE";
            ws.Cells["G1"].Value = "PKGTYPE";
            ws.Cells["H1"].Value = "GROSSWT";
            ws.Cells["I1"].Value = "GW_UOM";
            ws.Cells["J1"].Value = "NETWT";
            ws.Cells["K1"].Value = "NW_UOM";
            ws.Cells["L1"].Value = "VGM";
            ws.Cells["M1"].Value = "CBM";

            r = ws.Cells["A1:M1"];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(17, 188, 183));
            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int rw = 2;
            DataTable dtx = GetBOLCntrExistingValus(BkgId, BLID);
            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                ws.Cells["A" + rw].Value = rw - 1;
                ws.Cells["B" + rw].Value = dtx.Rows[i]["CntrNo"].ToString();
                ws.Cells["C" + rw].Value = dtx.Rows[i]["Size"].ToString();
                ws.Cells["D" + rw].Value = dtx.Rows[i]["ISOCode"].ToString();
                ws.Cells["E" + rw].Value = dtx.Rows[i]["SealNo"].ToString();
                ws.Cells["F" + rw].Value = dtx.Rows[i]["NoOfPkg"].ToString();
                ws.Cells["G" + rw].Value = dtx.Rows[i]["PakgTypeName"].ToString();
                ws.Cells["H" + rw].Value = dtx.Rows[i]["GrsWt"].ToString();
                ws.Cells["I" + rw].Value = dtx.Rows[i]["GrsWtType"].ToString();
                ws.Cells["J" + rw].Value = dtx.Rows[i]["NtWt"].ToString();
                ws.Cells["K" + rw].Value = dtx.Rows[i]["NtWtType"].ToString();
                ws.Cells["L" + rw].Value = dtx.Rows[i]["VGM"].ToString();
                ws.Cells["M" + rw].Value = dtx.Rows[i]["CBM"].ToString();
                rw++;

            }

            r = ws.Cells["A2:D" + (rw-1)];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218));
            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A1:M" + (rw - 1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:M" + (rw - 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:M" + (rw - 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:M" + (rw - 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:M" + (rw - 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            //ws.Cells[1, 1, rw, 14].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename= "+ dtx.Rows[0]["BLNumber"].ToString() + "_BLContainerDetails.xlsx");
            Response.End();
            return View();
        }


        public DataTable GetBOLCntrExistingValus(string BkgID, string BLID)
        {

            string _Query = "  select distinct 0 as BCNID,  (select top(1) BookingNo from NVO_Booking where ID = BLNumber) as BLNumber,NVO_Containers.Id as CntrID,CntrNo,TypeID,(select top(1) ISOCode  from NVO_tblCntrTypes where ID = TypeID) as ISOCode, " +
                            " (select top(1) Size  from NVO_tblCntrTypes where ID = TypeID) as Size, " +
                            " (select top(1) SealNo from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as SealNo, " +
                            " (select top(1) NoOfPkg from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as NoOfPkg,  " +
                            " (select top(1) PakgType from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as PakgType, " +
                            " (select top(1) (select top(1) PkgDescription from NVO_CargoPkgMaster where NVO_CargoPkgMaster.ID = NVO_BOLCntrDetails.PakgType) from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as PakgTypeName, " +
                            " (select top(1) GrsWt from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as GrsWt," +
                            " (select top(1) NtWt from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as NtWt,  " +
                            " (select top(1) VGM from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as VGM, " +
                            " (select top(1) CBM from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as CBM, " +
                            " (select top(1) case  when GrsWtType= 1 then 'KGS' else case when GrsWtType= 2  then'MT' end end from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as GrsWtType, " +
                            " (select top(1) case  when NtWtType= 1 then 'KGS' else case when NtWtType= 2  then'MT' end end from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BkgId = NVO_ContainerTxns.BLNumber and BLID=" + BLID + " and NVO_ContainerTxns.ContainerID=NVO_BOLCntrDetails.CntrID) as NtWtType " +
                            " from NVO_ContainerTxns " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
                            " where BLNumber = " + BkgID;

            if (BLID != "0")
            {
                _Query += " and BLNumber=(select BkgId from NVO_BOLCntrDetails where BkgId=" + BkgID + " and BLID= " + BLID + " and CntrID=NVO_Containers.ID)";

            }

            return RegManage.GetViewData(_Query, "");
        }



        public ActionResult BOLPartyDownloadAttachment(string BkgId)
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("BLParty");
            ExcelRange r;
            ws.Cells["A1"].Value = "S.No.";
            ws.Cells["B1"].Value = "Party Type";
            ws.Cells["C1"].Value = "Party Name";
            ws.Cells["D1"].Value = "Address";
           

            r = ws.Cells["A1:D1"];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(17, 188, 183));
            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


            ws.Cells["A1:D1"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:D1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:D1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:D1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            //ws.Cells[1, 1, rw, 14].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename= BLPatrtyDetails.xlsx");
            Response.End();
            return View();
        }


        [HttpPost]
        public ContentResult UploadBLCntr()
        {
            string path = Server.MapPath("~/BLCntrUpload/");


            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var FileNamev = "";
            foreach (string key in Request.Files)
            {
                Random rd = new Random();
                FileNamev += rd.Next(10000).ToString();

                HttpPostedFileBase postedFile = Request.Files[key];

                FileNamev += "_" + postedFile.FileName;
                postedFile.SaveAs(path + FileNamev);

            }

            return Content(FileNamev);
        }
    }
}