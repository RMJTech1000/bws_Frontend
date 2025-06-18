using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DataManager;
using System.ComponentModel;
using System.Configuration;
using System.Text.RegularExpressions;

namespace NVOCShipping.Controllers
{
    public class ExportPDFController : Controller
    {
        ExportReportManager RegMng = new ExportReportManager();
        // GET: ExportPDF
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportTerminalPDF(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID)
        {
            BindTerminalPDF(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);
            return View();
        }

        public void BindTerminalPDF(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID)
        {
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 870);
            doc = new Document(rec);
           // doc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            Paragraph para = new Paragraph();


            PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();
           
            PdfContentByte cb = writer.DirectContent;
            cb.SetColorStroke(Color.BLACK);
            int _Xp = 10, _Yp = 785, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
            png1.SetAbsolutePosition(15, 810);
            png1.ScalePercent(18f);
            doc.Add(png1);

            //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddress.png"));
            //png2.SetAbsolutePosition(80, 810);
            //png2.ScalePercent(25f);
            //doc.Add(png2);

            BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader21, 23);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLUE WAVE SHIPPING PVT LTD", 280, 840, 0);


            BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader22, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 280, 820, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 280, 810, 0);
            cb.SetColorStroke(new Color(0, 0, 128));

            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            cb.SetFontAndSize(bfheader2, 10);
            cb.SetColorFill(new Color(0, 0, 128));


            cb.BeginText();

            cb.MoveTo(10, 795);
            cb.LineTo(660, 795);

            cb.MoveTo(10, 770);
            cb.LineTo(660, 770);

            cb.MoveTo(10, 795);
            cb.LineTo(10, 770);

            cb.MoveTo(660, 795);
            cb.LineTo(660, 770);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMINAL DEPARTURE REPORT", 220, 780, 0);
            cb.SetRGBColorFill(0x00, 0x00, 0xFF);           
            cb.EndText();

            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.BeginText();
            cb.MoveTo(10, 755);
            cb.LineTo(660, 755);
            cb.MoveTo(10, 735);
            cb.LineTo(660, 735);
            cb.MoveTo(10, 715);
            cb.LineTo(660, 715);
            cb.MoveTo(10, 695);
            cb.LineTo(660, 695);
            cb.MoveTo(10, 675);
            cb.LineTo(660, 675);
            cb.MoveTo(10, 655);
            cb.LineTo(660, 655);
            cb.MoveTo(10, 635);
            cb.LineTo(660, 635);
            //cb.MoveTo(10, 615);
            //cb.LineTo(660, 615);
            cb.MoveTo(10, 615);
            cb.LineTo(660, 615);
            cb.MoveTo(10, 755);
            cb.LineTo(10, 615);
            cb.MoveTo(180, 755);
            cb.LineTo(180, 615);
            cb.MoveTo(660, 755);
            cb.LineTo(660, 615);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING", 15, 740, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NAME OF VESSEL", 15, 720, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA POL", 15, 700, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEXT PORT", 15, 680, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL OPERATOR", 15, 660, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER OPERATOR", 15, 640, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD Agent", 15, 620, 0);
            // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "T/S Agent", 15, 620, 0)
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOYAGE NUMBER", 400, 720, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD POL", 400, 700, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA NEXT PORT", 400, 680, 0);

            DataTable _dtv = GetTerminalDepPdfValues(VesVoyID, PODID,TSPORTID,AgentID,VesselOpID);
            if(_dtv.Rows.Count > 0)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["POL"].ToString(), 190, 740, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VesselName"].ToString(), 190, 720, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["ETA"].ToString(), 190, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["NextPort"].ToString(), 190, 680, 0);
              
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLUE WAVE SHIPPING PVT LTD", 190, 640, 0);
                

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VoyageNo"].ToString(), 500, 720, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["ETD"].ToString(), 500, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT,  _dtv.Rows[0]["NextPortETA"].ToString(), 500, 680, 0);
           

            
            
            cb.EndText();

                ////TABLE
                cb.BeginText();
                cb.MoveTo(10, 580);
                cb.LineTo(660, 580);
                cb.MoveTo(10, 560);
                cb.LineTo(660, 560);
                cb.MoveTo(10, 580);
                cb.LineTo(10, 400);
                cb.MoveTo(28, 580);
                cb.LineTo(28, 400);
                cb.MoveTo(75, 580);
                cb.LineTo(75, 400);
                cb.MoveTo(110, 580);
                cb.LineTo(110, 400);
                cb.MoveTo(140, 580);
                cb.LineTo(140, 400);
                cb.MoveTo(230, 580);
                cb.LineTo(230, 400);
                cb.MoveTo(325, 580);
                cb.LineTo(325, 400);
                cb.MoveTo(415, 580);
                cb.LineTo(415, 400);
                cb.MoveTo(459, 580);
                cb.LineTo(459, 400);
                //cb.MoveTo(500, 580);
                //cb.LineTo(500, 400);
                //cb.MoveTo(470, 580);
                //cb.LineTo(470, 400);
                cb.MoveTo(487, 580);
                cb.LineTo(487, 400);
                cb.MoveTo(520, 580);
                cb.LineTo(520, 400);
                cb.MoveTo(590, 580);
                cb.LineTo(590, 400);
                cb.MoveTo(660, 580);
                cb.LineTo(660, 400);
                cb.MoveTo(10, 400);
                cb.LineTo(660, 400);
                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 6);
                cb.SetColorFill(new Color(0, 0, 128));
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "S.NO", 12, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER #", 31, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SIZE TYPE", 77, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SERVICE", 112, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TS PORT", 142, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 232, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FINAL DEST", 326, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PICKUP DATE", 417, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GW.KGS", 461, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VSL OPERATOR", 505, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL NUMBER", 540, 568, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD AGENT", 580, 568, 0);
                int RowGrd = 550;
                DataTable _dtC = GetCntrDtlsTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);

                for (int k = 0; k < _dtC.Rows.Count; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["SNo"].ToString(), 12, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["CntrNo"].ToString(), 31, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["Size"].ToString(), 77, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["ServiceType"].ToString(), 112, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["TSPORT"].ToString(), 142, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["POD"].ToString(), 232, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["FPOD"].ToString(), 326, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["PickUpDate"].ToString(), 417, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["GrsWt"].ToString(), 461, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["Operator"].ToString(), 505, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["BLNumber"].ToString(), 540, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["DestinationAgent"].ToString(), 592, RowGrd, 0);
                    RowGrd -= 10;

                }



            }
           cb.EndText();

           
            cb.Stroke();


            writer.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=TerminalDepartureReport.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
           // Response.Write(doc);
            Response.End();
        }

        public ActionResult ExportFreightManifestPDF(string VesVoyID, string PODID, string BLNumber)
        {
            BindFreightManifestPDF(VesVoyID,PODID,BLNumber);
            return View();
        }

        public void BindFreightManifestPDF(string VesVoyID, string PODID, string BLNumber)
        {
            Document doc = new Document();
            Rectangle rec = new Rectangle(840, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();
            doc.SetPageSize(iTextSharp.text.PageSize.A3.Rotate());
            PdfContentByte cb = writer.DirectContent;
            cb.SetColorStroke(Color.BLACK);
            int _Xp = 10, _Yp = 820, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 700, 0);
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
            png1.SetAbsolutePosition(40, 810);
            png1.ScalePercent(20f);
            doc.Add(png1);

            //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/header.JPG"));
            //png2.SetAbsolutePosition(200, 810);
            //png2.ScalePercent(80f);
            //doc.Add(png2);

            BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader21, 23);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLUE WAVE SHIPPING PVT LTD", 280, 850, 0);


            BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader22, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 280, 830, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 280, 815, 0);

            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader2, 10);
            cb.SetColorFill(new Color(0, 0, 128));

            cb.BeginText();

            cb.MoveTo(10, 795);
            cb.LineTo(830, 795);

            cb.MoveTo(10, 770);
            cb.LineTo(830, 770);

            cb.MoveTo(10, 795);
            cb.LineTo(10, 770);

            cb.MoveTo(830, 795);
            cb.LineTo(830, 770);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREIGHT MANIFEST", 300, 780, 0);
            cb.EndText();

            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.BeginText();
            cb.MoveTo(10, 755);
            cb.LineTo(830, 755);
            cb.MoveTo(10, 690);
            cb.LineTo(830, 690);
            cb.MoveTo(10, 755);
            cb.LineTo(10, 690);
            cb.MoveTo(830, 755);
            cb.LineTo(830, 690);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL", 15, 740, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD", 15, 720, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 15, 700, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOYAGE", 320, 740, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA", 320, 720, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 320, 700, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN NO", 550, 740, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOAD TERMINAL", 550, 720, 0);


           
            cb.EndText();

            //TABLE
            cb.BeginText();
            
            BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader4, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.MoveTo(10, 672);
            cb.LineTo(840, 672);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill Of Ladding No", 15, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper/Consignee/Notify", 125, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks & Numbers", 320, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No Of Package", 480, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Goods", 545, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gross.Weight", 730, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CBM", 800, 675, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper", 125, 660, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Consignee", 125, 540, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify", 125, 450, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Size/Type/Seal", 15, 360, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Collection Mode :", 630, 360, 0);
            int RowsColumn = 0;
            int ColumnRows = 0;
            int ColRow = 660;
            RowsColumn = 0;
            DataTable _dtf = GetFreightManifestPdfValues(VesVoyID, PODID, BLNumber);

            BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader5, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            string Address = "";
            if (_dtf.Rows.Count > 0)
            {

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["VesselName"].ToString(), 80, 740, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["VoyageNo"].ToString(), 400, 740, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["SCNNo"].ToString(), 650, 740, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["ETD"].ToString(), 80, 720, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["ETA"].ToString(), 400, 720, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["LoadTerminal"].ToString(), 650, 720, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["POL"].ToString(), 80, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["POD"].ToString(), 400, 700, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["BLNumber"].ToString(), 15, 660, 0);

                var splitShipper = _dtf.Rows[0]["Shipper"].ToString().Split('\n');
                int RowInxS = 640;
                for (int i = 0; i < splitShipper.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitShipper[i].ToString(), 125, RowInxS, 0);
                    RowInxS -= 12;
                }

                var split = _dtf.Rows[0]["ShipperAddress"].ToString().Split('\n');
                int RowInx = 630;
                for (int i = 0; i < split.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[i].ToString(), 125, RowInx, 0);
                    RowInx -= 12;
                }

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["Consignee"].ToString(), 125, 525, 0);
                var splitx = _dtf.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                int RowInxV = 515;
                for (int i = 0; i < splitx.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitx[i].ToString(), 125, RowInxV, 0);
                    RowInxV -= 12;
                }
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["Notify1"].ToString(), 125, 435, 0);
                var splity = _dtf.Rows[0]["Notify1Address"].ToString().Split('\n');
                int RowInxY = 425;
                for (int i = 0; i < splity.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splity[i].ToString(), 125, RowInxY, 0);
                    RowInxY -= 12;
                }
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["MarkNo"].ToString(), 320, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["PakgType"].ToString(), 480, 660, 0);
                
                Address = _dtf.Rows[0]["CagoDescription"].ToString().ToUpper();
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["CagoDescription"].ToString(), 545, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["GRWT"].ToString(), 730, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["CBM"].ToString(), 800, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["CntrNo"].ToString() + '/' + _dtf.Rows[0]["Seal"].ToString(), 15, 345, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight :" +  _dtf.Rows[0]["FreightPayment"].ToString(), 630, 345, 0);
                
            }

            string[] ArrayAddress = Regex.Split(Address, char.ConvertFromUtf32(10));

            for (int x = 0; x < ArrayAddress.Length; x++)
            {
                string[] Aaddsplit = SplitByLenght(ArrayAddress[x].ToString(), 35);
                for (int k = 0; k < Aaddsplit.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 545, ColRow - ColumnRows, 0);
                    ColumnRows += 13;
                    RowsColumn++;
                }

            }



            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "14 DAYS FREE DETENTION AT POD", 545, 450, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Frt / 40HC 40.00 USD", 315, 345, 0);
            

            BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader6, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Contuinuity as per annexure attached", 545, 435, 0);
           
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHARGES BREAKUP", 15, 290, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Details", 320, 360, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container No/Seal", 15, 260, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Type", 155, 260, 0);

            int inxColumn = 205;
            DataTable _dtR = GetCntrChargesBreakup(_dtf.Rows[0]["RRID"].ToString());
            for (int x = 0; x < _dtR.Rows.Count; x++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtR.Rows[x]["ChgCode"].ToString(), inxColumn, 260, 0);
                inxColumn += 50;
 
            }
            
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Total", inxColumn, 260, 0);
            //DataTable _dtR = GetCntrChargesBreakup(_dtf.Rows[0]["RRID"].ToString());
            string Totalv = "";
            //string BAF = _dtR.Rows[0]["BAF"].ToString();
            //string ECRS = _dtR.Rows[0]["ECRS"].ToString();
            //string FRT = _dtR.Rows[0]["FRT"].ToString();
            ////string[] val = Totalv.Split(',');
            //Totalv = (double.Parse(BAF) + double.Parse(ECRS) + double.Parse(FRT)).ToString();
            //int Totalv;
            //int BAF;
            //int ECRS;
            //int FRT;
            //if (_dtR.Rows[0]["BAF"].ToString() != "")
            //{
            //    BAF = Int32.Parse(_dtR.Rows[0]["BAF"].ToString());
            //}
            //else
            //{
            //    BAF = 0;
            //}
            //if(_dtR.Rows[0]["ECRS"].ToString() != "")
            //{
            //    ECRS = Int32.Parse(_dtR.Rows[0]["ECRS"].ToString());
            //}
            //else
            //{
            //    ECRS = 0;
            //}
            //if (_dtR.Rows[0]["FRT"].ToString() != "")
            //{
            //    FRT = Int32.Parse(_dtR.Rows[0]["FRT"].ToString());
            //}
            //else
            //{
            //    FRT = 0;
            //}


            //Totalv = (BAF + ECRS + FRT);
            decimal Totalv1 = 0;
            if (_dtR.Rows.Count > 0)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtf.Rows[0]["Seal"].ToString(), 15, 240, 0);
                int ColumnRow = 205;
                
                for (int y = 0; y < _dtR.Rows.Count; y++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtR.Rows[y]["ManifRate"].ToString(), ColumnRow, 240, 0);
                    Totalv += decimal.Parse(_dtR.Rows[y]["ManifRate"].ToString()).ToString();

                    ColumnRow += 50;
                }

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Totalv.ToString(), ColumnRow, 240, 0);
            }
            cb.MoveTo(10, 275);
            cb.LineTo(400, 275);
            cb.MoveTo(10, 255);
            cb.LineTo(400, 255);
            cb.MoveTo(10, 235);
            cb.LineTo(400, 235);
            cb.MoveTo(10, 275);
            cb.LineTo(10, 235);
            cb.MoveTo(150, 275);
            cb.LineTo(150, 235);
            cb.MoveTo(200, 275);
            cb.LineTo(200, 235);
            cb.MoveTo(300, 275);
            cb.LineTo(300, 235);
            cb.MoveTo(250, 275);
            cb.LineTo(250, 235);
            cb.MoveTo(350, 275);
            cb.LineTo(350, 235);
            cb.MoveTo(400, 275);
            cb.LineTo(400, 235);
            
            cb.EndText();
            cb.Stroke();

            writer.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=FreightManifestReport.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();
        }


        //public void BindFreightManifestPDF()
        //{
        //    Document doc = new Document();
        //    Rectangle rec = new Rectangle(670, 870);
        //    doc = new Document(rec);
        //    Paragraph para = new Paragraph();


        //    PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
        //    doc.Open();

        //    PdfContentByte cb = writer.DirectContent;
        //    cb.SetColorStroke(Color.BLACK);
        //    int _Xp = 10, _Yp = 785, YDiff = 10;

        //    BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader, 14);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);
        //    iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/logo.png"));
        //    png1.SetAbsolutePosition(15, 810);
        //    png1.ScalePercent(45f);
        //    doc.Add(png1);

        //    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/header.JPG"));
        //    png2.SetAbsolutePosition(130, 810);
        //    png2.ScalePercent(50f);
        //    doc.Add(png2);



        //    BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader2, 10);
        //    cb.SetColorFill(Color.BLUE);

        //    cb.BeginText();

        //    cb.MoveTo(10, 795);
        //    cb.LineTo(660, 795);

        //    cb.MoveTo(10, 770);
        //    cb.LineTo(660, 770);

        //    cb.MoveTo(10, 795);
        //    cb.LineTo(10, 770);

        //    cb.MoveTo(660, 795);
        //    cb.LineTo(660, 770);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREIGHT MANIFEST", 280, 780, 0);
        //    cb.EndText();

        //    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader3, 8);
        //    cb.SetColorFill(Color.BLUE);
        //    cb.BeginText();
        //    cb.MoveTo(10, 755);
        //    cb.LineTo(660, 755);
        //    cb.MoveTo(10, 690);
        //    cb.LineTo(660, 690);
        //    cb.MoveTo(10, 755);
        //    cb.LineTo(10, 690);
        //    cb.MoveTo(660, 755);
        //    cb.LineTo(660, 690);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL", 15, 740, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD", 15, 720, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 15, 700, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOYAGE", 280, 740, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA", 280, 720, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 280, 700, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN NO", 470, 740, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOAD TERMINAL", 470, 720, 0);


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ISEACO WISDOM", 80, 740, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "031S", 330, 740, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "3EXP9", 550, 740, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "04-Jul-2021", 80, 720, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "04-Jul-2021", 330, 720, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PSA TERMINAL", 550, 720, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SINGAPORE,SINGAPORE", 80, 700, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANGKOK,THAILAND", 330, 700, 0);
        //    cb.EndText();

        //    //TABLE
        //    cb.BeginText();

        //    BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader4, 8);
        //    cb.SetColorFill(Color.BLUE);
        //    cb.MoveTo(10, 672);
        //    cb.LineTo(660, 672);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill Of Ladding No", 15, 675, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper/Consignee/Notify", 100, 675, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks & Numbers",245, 675, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No Of Package", 330, 675, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Goods", 405, 675, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gross.Weight", 580, 675, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CBM",640 , 675, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OCLSG210421BK", 15, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper", 100, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 220, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "20 Pallets", 330 , 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SHIPPERS STOW ,COUNT ,LOAD & SEALED", 405, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "12005.31", 580, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "52.4", 640, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GEODIS SINGAPORE PTE LTD", 100, 630, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VIVA PREMIUM", 245, 640, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1X40'HC CONTAINER S.T.C:", 405, 640, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Ref No: TH058-2DO21/002", 405, 570, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Consignee", 100, 540, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GEODIS THAI LTD.", 100, 525, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify", 100, 450, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SAME AS CONSIGNEE", 100, 435, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "14 DAYS FREE DETENTION AT POD", 430, 450, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Contuinuity as per annexure attached", 430, 435, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Size/Type/Seal", 15, 360, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Details", 315, 360, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Collection Mode :", 560, 360, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ECMU9733011/40HC/001001", 15, 345, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Frt / 40HC 40.00 USD", 315, 345, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight : Prepaid", 560, 345, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHARGES BREAKUP", 15, 290, 0);
        //    cb.MoveTo(10, 275);
        //    cb.LineTo(340, 275);
        //    cb.MoveTo(10, 255);
        //    cb.LineTo(340, 255);
        //    cb.MoveTo(10, 235);
        //    cb.LineTo(340, 235);
        //    cb.MoveTo(10, 275);
        //    cb.LineTo(10, 235);
        //    cb.MoveTo(100, 275);
        //    cb.LineTo(100, 235);
        //    cb.MoveTo(145, 275);
        //    cb.LineTo(145, 235);
        //    cb.MoveTo(190, 275);
        //    cb.LineTo(190, 235);
        //    cb.MoveTo(235, 275);
        //    cb.LineTo(235, 235);
        //    cb.MoveTo(275, 275);
        //    cb.LineTo(275, 235);
        //    cb.MoveTo(340, 275);
        //    cb.LineTo(340, 235);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container No/Seal", 15, 260, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Type", 105, 260, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BAFP", 150, 260, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ECRS", 195, 260, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FRT", 240, 260, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Total", 280, 260, 0);
        //    cb.EndText();
        //    cb.Stroke();

        //    writer.CloseStream = false;
        //    doc.Close();
        //    Response.Buffer = true;
        //    Response.ContentType = "application/pdf";
        //    Response.AddHeader("content-disposition", "attachment;filename=FreightManifestReport.pdf");
        //    Response.Cache.SetCacheability(HttpCacheability.NoCache);
        //    Response.Write(doc);
        //    Response.End();
        //}


        public ActionResult ExportSurrenderNoticePDF()
        {
            BindSurrenderNoticePDF();
            return View();
        }

        public void BindSurrenderNoticePDF()
        {
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 870);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();

            PdfContentByte cb = writer.DirectContent;
            cb.SetColorStroke(Color.BLACK);
            int _Xp = 10, _Yp = 785, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/logo.png"));
            png1.SetAbsolutePosition(80, 810);
            png1.ScalePercent(45f);
            doc.Add(png1);

            iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/header.JPG"));
            png2.SetAbsolutePosition(200, 810);
            png2.ScalePercent(50f);
            doc.Add(png2);



            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader2, 11);
            cb.SetColorFill(Color.BLACK);

            cb.BeginText();

            //cb.MoveTo(10, 795);
            //cb.LineTo(660, 795);

            //cb.MoveTo(10, 770);
            //cb.LineTo(660, 770);

            //cb.MoveTo(10, 795);
            //cb.LineTo(10, 770);

            //cb.MoveTo(660, 795);
            //cb.LineTo(660, 770);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Telex Surrender Release - Notice", 140, 780, 0);
            cb.EndText();

            
            //cb.MoveTo(10, 755);
            //cb.LineTo(660, 755);            
            //cb.MoveTo(10, 690);
            //cb.LineTo(660, 690);
            //cb.MoveTo(10, 755);
            //cb.LineTo(10, 690);
            //cb.MoveTo(660, 755);
            //cb.LineTo(660, 690);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "To,", 40, 740, 0);
            DataTable _dtn = GetSurrenderNoticePdfValues();
            if(_dtn.Rows.Count > 0)
            {
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 10);
                cb.SetColorFill(Color.BLACK);
                cb.BeginText();
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["DesignationAgent"].ToString(), 40, 720, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT,"BL NO:" + _dtn.Rows[0]["BLNumber"].ToString(), 40, 638, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["DesignationAgentAddress"].ToString(), 40, 705, 0);
                cb.EndText();

                cb.BeginText();
                BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader5, 6);
                cb.SetColorFill(Color.BLACK);

                cb.MoveTo(40, 610);
                cb.LineTo(640, 610);
                cb.MoveTo(40, 590);
                cb.LineTo(640, 590);
                cb.MoveTo(40, 570);
                cb.LineTo(640, 570);

                cb.MoveTo(40, 610);
                cb.LineTo(40, 570);
                cb.MoveTo(120, 610);
                cb.LineTo(120, 570);
                cb.MoveTo(200, 610);
                cb.LineTo(200, 570);
                cb.MoveTo(280, 610);
                cb.LineTo(280, 570);
                cb.MoveTo(360, 610);
                cb.LineTo(360, 570);
                cb.MoveTo(445, 610);
                cb.LineTo(445, 570);
                cb.MoveTo(540, 610);
                cb.LineTo(540, 570);
                //cb.MoveTo(520, 610);
                //cb.LineTo(520, 570);
                //cb.MoveTo(590, 610);
                //cb.LineTo(590, 570);
                cb.MoveTo(640, 610);
                cb.LineTo(640, 570);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL VOYAGE", 45, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL NUMBER", 125, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER QUANTITY", 205, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLACE OF ORIGIN", 285, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING", 365, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF DISCHARGE", 450, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FINAL DESTINATION", 545, 596, 0);
                

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["VesVoy"].ToString(), 45, 580, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["BLNumber"].ToString(), 125, 580, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 205, 596, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["POO"].ToString(), 285, 580, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["POL"].ToString(), 365, 580, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["POD"].ToString(), 450, 580, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["FPOD"].ToString(), 545, 580, 0);
               
                cb.EndText();

                cb.BeginText();
                BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader6, 8);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SHIPPER", 45, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["Shipper"].ToString(), 45, 535, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONSIGNEE", 45, 500, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtn.Rows[0]["Consignee"].ToString(), 45, 485, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "The Full set of Original BL's has been surrendered at POL,please release cargot to designated consignee without", 45, 465, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "presentation of bwl BL's", 45, 456, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Remarks:", 45, 436, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "The Discharge port or Destination agent may then release cargo to the verified consignee against following conditions:", 45, 427, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1. Collecting all charges from concerned parties", 45, 407, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "2. Collection signatory LOI from consignee", 45, 398, 0);
                cb.EndText();


            }
           
           
            //TABLE
            cb.BeginText();

            BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader4, 8);
            cb.SetColorFill(Color.BLACK);
           
            
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SELENGAR,MALAYSIA", 40, 696, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TEL: +603-30008779; FAX: +603-33192031", 40, 687, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL: cs-bws@bluewave.com.my", 40, 678, 0);

            

            
            cb.EndText();

            cb.Stroke();

            writer.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=SurrenderNotice.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.Write(doc);
            Response.End();
        }

        public DataTable GetTerminalDepPdfValues(string VesVoy,string POD, string TsPort, string Agent, string VesselOperator)
        {
            string strWhere = "";
            //string _Query = " Select Top(1)  B.ID,BL.ID AS BLId,vesname.VesVoy as VeslName,VoyageNo,B.POL,convert(varchar, ETA, 106) as ETA,convert(varchar, ETD, 106) as ETD,convert(varchar, ATA, 106) as ATA,convert(varchar, ATD, 106) as ATD," + " (select top(1) PortName from NVO_PortMaster where ID = NVO_VoyPortDtls.NextPortID) as NextPort," +
            //    " convert(varchar, (select top(1)  NextPortETA from NVO_VoyPortDtls where NVO_VoyPortDtls.VoydtID = VD.ID),106) as NextPortETA,(select top(1) CustomerName from NVO_CustomerMaster where ID = NVO_VoyPortDtls.VesOperatorID) as VesselOperator,"+
            //    " (select top(1) AgencyName from NVO_AgencyMaster where ID = BL.DesAgentID) as PODAgent from NVO_Booking B inner join NVO_BOL BL On BL.BkgID = B.ID inner join NVO_VoyageDetails VD On VD.ID = B.VesVoyID inner join NVO_VoyPortDtls On NVO_VoyPortDtls.VoydtID = VD.ID " +
            // " outer apply(select ID, (select top(1) VesselName from NVO_VesselMaster where ID = NVO_VoyageDetails.VesID)  as VesVoy from NVO_VoyageDetails where NVO_VoyageDetails.ID = B.VesVoyID) vesname ";

            string _Query = "Select Top(1)  NVO_Booking.ID,NVO_BOL.ID as BLID, " +
                            " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID)+ '-' +  (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = NVO_Booking.VesVoyID)as VesselName, " +
                            " (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as VoyageNo, " +
                            " NVO_Booking.POL,(select Convert(varchar,ETA,106) from NVO_VoyageRoute where NVO_VoyageRoute.RID = NVO_Booking.ID)as ETA,(select Convert(varchar,ETD,106) from NVO_VoyageRoute where NVO_VoyageRoute.RID = NVO_Booking.ID)as ETD, " +
                            " (select top(1) PortName from NVO_VoyageRoute " +
                            " inner join NVO_PortMaster On NVO_PortMaster.ID = NVO_VoyageRoute.PortID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as NextPort, " +
                            " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc),106) as NextPortETA, " +
                            " (select top(1) Operator from NVO_VoyageOpertaors  where NVO_VoyageOpertaors.VoyageID = NVO_Booking.VesVoyID) as VesselOperator, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster " +
                            " where NVO_AgencyMaster.ID = NVO_Booking.DestinationAgentID) as PODAgent from NVO_Booking " +
                            " inner join NVO_Voyage On NVO_Voyage.ID = NVO_Booking.VesVoyID" +
                            " inner join NVO_BOL On NVO_BOL.BkgID = NVO_Booking.ID";

            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.VesVoyID=" + VesVoy;
                else

                    strWhere += " AND NVO_Booking.VesVoyID=" + VesVoy;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.PODID=" + POD;
                else
                    strWhere += " AND NVO_Booking.PODID=" + POD;

            if (TsPort != "" && TsPort != "0" && TsPort != "null" && TsPort != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.TSPORTID=" + TsPort;
                else
                    strWhere += " AND NVO_Booking.TSPORTID=" + TsPort;

            if (Agent != "" && Agent != "0" && Agent != "null" && Agent != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.DestinationAgentID=" + Agent;
                else
                    strWhere += " AND NVO_Booking.DestinationAgentID=" + Agent;

            if (VesselOperator != "" && VesselOperator != "0" && VesselOperator != "null" && VesselOperator != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.SlotOperatorID=" + VesselOperator;
                else
                    strWhere += " AND NVO_Booking.SlotOperatorID=" + VesselOperator;

            if (strWhere == "")
                strWhere = _Query;

            return RegMng.GetViewData(strWhere, "");
        }

        public DataTable GetFreightManifestPdfValues(string VesVoy, string POD, string BLNumber)
        {
            string strWhere = "";
            //string _Query = "Select Top(1) " +
            //                " NVO_BOL.Id, vesname.VesVoy as VeslName,NVO_VoyageDetails.VoyageNo,NVO_voyageDetails.SCNNo, " +
            //                " convert(varchar, ETA, 106) as ETA,convert(varchar, ETD, 106) as ETD, " +
            //                " (select(select Top(1) TerminalName From NVO_TerminalMaster where NVO_TerminalMaster.ID = NVO_VoyageDetails.LoadingTerminalID)  from NVO_VoyageDetails where NVO_VoyageDetails.ID = NVO_BOL.BkgID)as LoadTerminal, " +
            //                " POL,POD, BLNumber,Shipper, " +
            //                " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 1 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Shipper,  " +
            //                " (select top(1) PartyAddress from NVO_BOLCustomerDetails BC where PartyTypeID = 1 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as ShipperAddress,  " +
            //                " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 2 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Consignee,  " +
            //                " (select top(1) PartyAddress from NVO_BOLCustomerDetails BC where PartyTypeID = 2 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as ConsigneeAddress,  " +
            //                " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 3 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Notify1,  " +
            //                " (select top(1) PartyAddress from NVO_BOLCustomerDetails BC where PartyTypeID = 3 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Notify1Address,  " +
            //                " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 12 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Notify2,  " +
            //                " (select top(1) PartyAddress from NVO_BOLCustomerDetails BC where PartyTypeID = 12 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Notify2Address,  " +
            //                " MarkNo,CagoDescription, " +
            //                " (select sum(Grswt) from NVO_BOLCntrDetails Cd where Cd.BLID = NVO_BOL.Id AND Cd.BkgId = NVO_BOL.BkgID) as GRWT,  " +
            //                " (select sum(CBM) from NVO_BOLCntrDetails Cd where Cd.BLID = NVO_BOL.Id  AND Cd.BkgId = NVO_BOL.BkgID) as CBM, " +
            //                " (select sum(PakgType) from NVO_BOLCntrDetails Cd where CD.BLID = NVO_BOL.Id  AND Cd.BkgId = NVO_BOL.BkgID) as PakgType, " +
            //                " (CntrNo + '/' + Size + '/' + SealNo)as Seal, " +
            //                " (select top(1) GeneralName  from NVO_GeneralMaster where Id = FreightPaymentID) as FreightPayment " +
            //                " from NVO_BOL inner join NVO_Booking on NVO_Booking.Id = NVO_BOL.BkgId " +
            //                " outer apply(select ID, (select top(1) VesselName from NVO_VesselMaster where ID = NVO_VoyageDetails.VesID) as VesVoy from NVO_VoyageDetails " +
            //                " where NVO_VoyageDetails.ID = NVO_Booking.VesVoyID )vesname " +
            //                " inner join NVO_VoyageDetails On NVO_VoyageDetails.ID = NVO_Booking.VesVoyID " +
            //                " inner join NVO_VoyPortDtls On NVO_VoyPortDtls.VoydtID = NVO_VoyageDetails.ID " +
            //                " inner join NVO_BOLCntrDetails On NVO_BOLCntrDetails.BkgId = NVO_BOL.BkgID ";

            string _Query = "select Top(1)  NVO_Booking.ID,NVO_BOL.ID as BLID,nvo_booking.RRID, " +
                            " (select top 1 NVO_Containers.CntrNo from NVO_BOLCntrDetails Cd " +
                            " inner join NVO_Containers On NVO_Containers.ID = Cd.CntrID where CD.BLID = NVO_BOL.Id " +
                            " AND Cd.BkgId = NVO_BOL.BkgID) as CntrNo, " +
                            " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + '-' +(select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from " +
                            " NVO_Voyage V where V.ID = NVO_Booking.VesVoyID)as VesselName," +
                            " (select top(1)  ExportVoyageCd from NVO_VoyageRoute " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as VoyageNo, " +
                            " '' as SCNNo , " +
                            " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID),106) as ETA, " +
                            " (select top(1)  TerminalName from NVO_VoyageRoute inner join NVO_TerminalMaster On NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID)as LoadTerminal, " +
                            " convert(varchar, (select top(1)  ETD from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID " +
                            " order by NVO_VoyageRoute.RID desc),106) as ETD , " +
                            " NVO_Booking.POL,NVO_Booking.POD,NVO_BOL.BLNumber, " +
                            " (select Top (1)(CustomerName  + '-' + Branch) as CustomerName from NVO_CustomerMaster CM " +
                            " inner join NVO_CusBusinessTypes on NVO_CusBusinessTypes.CustomerID = CM.ID " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = CM.Id " +
                            " inner join NVO_BOLCustomerDetails BC on BC.PartID = CID " +
                            " where CID = BC.PartID  and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Shipper, " +
                            " (select top(1) PartyAddress from NVO_BOLCustomerDetails BC " +
                            " where PartyTypeID = 1 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as ShipperAddress, " +
                            " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 2 and BC.BLID = NVO_BOL.Id " +
                            " AND BC.BkgId = NVO_BOL.BkgID) as Consignee, " +
                            " (select top(1) PartyAddress from NVO_BOLCustomerDetails " +
                            " BC where PartyTypeID = 2 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as ConsigneeAddress, " +
                            " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 3 and BC.BLID = NVO_BOL.Id " +
                            " AND BC.BkgId = NVO_BOL.BkgID) as Notify1, " +
                            " (select top(1) PartyAddress from NVO_BOLCustomerDetails BC " +
                            " where PartyTypeID = 3 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Notify1Address, " +
                            " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 12 and BC.BLID = NVO_BOL.Id " +
                            " AND BC.BkgId = NVO_BOL.BkgID) as Notify2,   (select top(1) PartyAddress from NVO_BOLCustomerDetails BC " +
                            " where PartyTypeID = 12 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Notify2Address, " +
                            " MarkNo,CagoDescription, " +
                            " (select sum(Grswt) from NVO_BOLCntrDetails Cd where Cd.BLID = NVO_BOL.Id AND " +
                            " Cd.BkgId = NVO_BOL.BkgID) as GRWT, " +
                            " (select sum(CBM) from NVO_BOLCntrDetails Cd where Cd.BLID = NVO_BOL.Id " +
                            " AND Cd.BkgId = NVO_BOL.BkgID) as CBM, " +
                            " (select sum(PakgType) from NVO_BOLCntrDetails Cd where CD.BLID = NVO_BOL.Id " +
                            " AND Cd.BkgId = NVO_BOL.BkgID) as PakgType, " +
                            " (CntrNo + '/' + Size + '/' + SealNo) as Seal, " +
                            " (select top(1) GeneralName from NVO_GeneralMaster where Id = FreightPaymentID) as FreightPayment " +
                            " from NVO_Booking " +
                            " inner join NVO_Voyage On NVO_Voyage.ID = NVO_Booking.VesVoyID " +
                            " inner join NVO_BOL On NVO_BOL.BkgID = NVO_Booking.ID " +
                            //" inner join NVO_VoyageDetails On NVO_VoyageDetails.ID = NVO_Booking.VesVoyID " +
                            " inner join NVO_BOLCntrDetails On NVO_BOLCntrDetails.BkgId = NVO_BOL.BkgID";
            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.VesVoyID=" + VesVoy;
                else
                    strWhere += " AND NVO_Booking.VesVoyID=" + VesVoy;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.PODID=" + POD;
                else
                    strWhere += " AND NVO_Booking.PODID=" + POD;

            if (BLNumber != "" && BLNumber != "0" && BLNumber != "null" && BLNumber != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_BOL.ID=" + BLNumber;
                else
                    strWhere += " AND NVO_BOL.ID=" + BLNumber;

           

            if (strWhere == "")
                strWhere = _Query;

            return RegMng.GetViewData(strWhere, "");
        }

        public DataTable GetSurrenderNoticePdfValues()
        {
            string _Query = "Select Top(1) " +
                            " NVO_BOL.Id, NVO_Booking.VesVoy,BLNumber,POO,POL,POD,FPOD, " +
                            " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 1 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Shipper, " +
                            " (select top(1) PartyName from NVO_BOLCustomerDetails BC where PartyTypeID = 2 and BC.BLID = NVO_BOL.Id  AND BC.BkgId = NVO_BOL.BkgID) as Consignee, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where NVO_AgencyMaster.Id = DesAgentID) as DesignationAgent, " +
                            " (select top(1) Address from NVO_AgencyMaster where NVO_AgencyMaster.Id = DesAgentID) as DesignationAgentAddress " +
                            " from NVO_BOL inner join NVO_Booking on NVO_Booking.Id = NVO_BOL.BkgId " +
                            " inner join NVO_VoyageDetails On NVO_VoyageDetails.ID = NVO_Booking.VesVoyID " +
                            " inner join NVO_VoyPortDtls On NVO_VoyPortDtls.VoydtID = NVO_VoyageDetails.ID " +
                            " inner join NVO_BOLCntrDetails On NVO_BOLCntrDetails.BkgId = NVO_BOL.ID";
            return RegMng.GetViewData(_Query, "");
        }

        public DataTable GetCntrDtlsTerminalDepPdfValues(string VesVoy, string POD, string TsPort, string Agent, string VesselOperator)
        {
            string strWhere = "";
            string _Query = "Select * from NVO_ViewTerminalDepReport WHERE  VesVoyID=" + VesVoy + "";

            if (POD != "" && POD != "0" && POD != null && POD != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) PODID   from NVO_Booking where ID = BkgId) =" + POD;
                else
                    strWhere += " and (select top(1) PODID   from NVO_Booking where ID = BkgId) =" + POD;


            if (TsPort != "" && TsPort != "0" && TsPort != null && TsPort != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) TSPORTID   from NVO_Booking where ID = BkgId)=" + TsPort;
                else
                    strWhere += " and (select top(1) TSPORTID   from NVO_Booking where ID = BkgId)=" + TsPort;

            if (Agent != "" && Agent != "0" && Agent != null && Agent != "?")

                if (strWhere == "")

                    strWhere += _Query + " and (select top(1) DestinationAgentID   from NVO_Booking where ID = BkgId)=" + Agent;
                else
                    strWhere += " and (select top(1) DestinationAgentID   from NVO_Booking where ID = BkgId)=" + Agent;


            if (VesselOperator != "" && VesselOperator != "0" && VesselOperator != null && VesselOperator != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) SlotOperatorID   from NVO_Booking where ID = BkgId)=" + VesselOperator;
                else
                    strWhere += " and (select top(1) SlotOperatorID   from NVO_Booking where ID = BkgId)=" + VesselOperator;


            if (strWhere == "")
                strWhere = _Query;

            return RegMng.GetViewData(strWhere, "");
        }

        public DataTable GetCntrChargesBreakup(string RID)
        {

            //string _Query = "select Id,RatesheetNo,(select top(1) Size + '-' + Type from NVO_tblCntrTypes where ID = CntrTypeID) AS Size,(select top(1) GeneralName from NVO_GeneralMaster where ID = CommodityTypeID) AS Commodity," +

            //    " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 1),0) as FRT, " +

            //    " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID  where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 1),'') as FRTCurr," +

            //  " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 22),0) as BAF, " +

            // " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 22),'') as BAFCurr, " +

            //  " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 27),0) as DGS, " +

            // " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 27),0) as DGSCurr," +

            // " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 15),0) as ECRS, " +

            // " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 15),'') as ECRSCurr," +

            //  " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 46),0) as CAF, " +

            //  " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 46),'') as CAFCurr," +

            //   " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 35),0) as EWRS, " +
            //   " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 35),0) as EWRSCurr," +

            //   " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 23),0) as LSS," +

            //   " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID =CurrencyID where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 23),'') as LSSCurr " +
            //   "  from NVO_Ratesheet inner join NVO_RatesheetCntrTypes on NVO_RatesheetCntrTypes.RRID=NVO_Ratesheet.ID where ID =" + RID;

            string _Query = "select * from NVO_V_FreightManifest where RRID=" + RID;
            return RegMng.GetViewData(_Query, "");
        }

        private string[] SplitByLenght(string Values, int split)
        {
            System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>();
            int SplitTheLoop = Values.Length / split;
            for (int i = 0; i < SplitTheLoop; i++)
                list.Add(Values.Substring(i * split, split));
            if (SplitTheLoop * split != Values.Length)
                list.Add(Values.Substring(SplitTheLoop * split));

            return list.ToArray();
        }

    }
   
}