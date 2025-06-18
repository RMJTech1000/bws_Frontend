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
    public class OnBoardConfirmController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: OnBoardConfirm
        public ActionResult DashBoard_onboardConfirm()
        {
            return View();
        }
        public ActionResult OnBoardTransacation()
        {
            return View();
        }

        public ActionResult VesselCertificate()
        {
            return View();
        }

        public ActionResult VesselCertificateSearch()
        {
            return View();
        }

        public ActionResult DelayCertificateSearch()
        {
            return View();
        }

        public ActionResult DelayCertificate()
        {
            return View();
        }



        public ActionResult VesselCertificatePrintPDF(string id, string AgencyID)
        {
            VesselCertificatePrint(id, AgencyID);
            return View();

        }
        public void VesselCertificatePrint(string BLID, string AgencyID)
        {
            DataTable dtv = GetVesselData(BLID);
            if (dtv.Rows.Count > 0)
            {
                //create pdf
                Document doc = new Document();
                //
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();

                //
                PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/LOGO1.png"));
                png2.SetAbsolutePosition(40, 830);
                png2.ScalePercent(75f);
                doc.Add(png2);


                DataTable dtc = GetAgencyDetails(AgencyID);
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 14);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 250, 855, 0);
                cb.SetFontAndSize(bfheader3, 9);
                int AddRow = 840;
                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, LogoAddresss[a].ToString(), 250, AddRow, 0);
                    AddRow -= 10;
                }


                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.DARK_GRAY);

                cb.BeginText();
                //Border-Top//
                //cb.MoveTo(10, 935);
                //cb.LineTo(695, 935);

                //Top1//
                cb.MoveTo(10, 805);
                cb.LineTo(660, 805);
                //Top2//
                //cb.MoveTo(10, 770);
                //cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Date : " + dtv.Rows[0]["Date"].ToString(), 20, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME / VOYAGE : " + dtv.Rows[0]["VSLVoyage"].ToString(), 20, 767, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLNUMBER  : " + dtv.Rows[0]["BLNumber"].ToString(), 20, 753, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL  CERTIFICATE", 20, 730, 0);

                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THIS IS TO  CERTIFY THAT", 20, 718, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CERTIFICATE FROM  THE OWNER, CARRIER  OR CAPTAIN  OF THE CARRYING", 20, 680, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL  OT THEIR  AGENT SHOWING  ITS NAME FLAG AND NATIONALITY  AND ", 20, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONFIRMING  THAT ", 20, 660, 0);



                string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                string[] Aaddsplit;
                int ColumnRows = 630;
                for (int x = 0; x < subjetv.Length; x++)
                {
                    Aaddsplit = subjetv[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 20, ColumnRows, 0);
                        ColumnRows -= 13;
                    }
                }

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YOURS FAITHFULLY", 20, 400, 0);


                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AS AGENT FOR THE CARRIER -", 20, 300, 0);
                cb.Stroke();



                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=BLPrintPDF.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }
        }

        public ActionResult DelayCertificatePrintPDF(string id, string AgencyID)
        {
            DelayCertificatePrint(id, AgencyID);
            return View();

        }

        public void DelayCertificatePrint(string ID, string AgencyID)
        {
            // DataTable dtv = GetdelayData(BLID);
            DataTable dtv = getDelayCertificate(ID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


                BaseFont bfheader9 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader9, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                png2.SetAbsolutePosition(45, 750);
                png2.ScalePercent(17f);
                doc.Add(png2);


                DataTable dtc = GetAgencyDetails(dtv.Rows[0]["AgencyID"].ToString());
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 14);
                cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 300, 855, 0);
                cb.SetFontAndSize(bfheader3, 9);

                int AddRow = 840;


                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.DARK_GRAY);

                cb.BeginText();



                cb.SetFontAndSize(bfheader2, 10);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                cb.EndText();
                cb.BeginText();

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL DELAY NOTICE ", 250, 730, 0);

                cb.MoveTo(250, 728);
                cb.LineTo(370, 728);

                //--------------------
                cb.SetFontAndSize(bfheader2, 9);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 65, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 95, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 65, 680, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 95, 680, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 100, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO VALUABLE CUSTOMER", 100, 680, 0);

                BaseFont BHHeaderSubject = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(BHHeaderSubject, 11);
                Font underlineFont = new Font(BHHeaderSubject, 11, Font.UNDERLINE);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUBJECT: VESSEL ARRIVAL DELAYED – " +
                    dtv.Rows[0]["VesVoy"].ToString() + " ETA:  " + dtv.Rows[0]["ETA"].ToString(), 65, 650, 0);


                cb.MoveTo(65, 648);
                cb.LineTo(500, 648);
                BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);




                int ColumnRows = 620;

                cb.SetFontAndSize(bfheader7, 9);
                string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                string[] Aaddsplit;

                for (int x = 0; x < subjetv.Length; x++)
                {
                    Aaddsplit = subjetv[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 65, ColumnRows, 0);
                        ColumnRows -= 14;
                    }
                }
                //box
                BaseFont bfheader11 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetColorFill(Color.BLACK);
                cb.SetFontAndSize(bfheader11, 9);

                ColumnRows -= 20;
                int LineVLeft = ColumnRows;
                cb.MoveTo(65, ColumnRows);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME/VOYAGE NO", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, ColumnRows, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, ColumnRows, 0);

                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOAD PORT ETA", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 260, ColumnRows, 0);
                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YARD CLOSING", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["YardClosing"].ToString(), 260, ColumnRows, 0);

                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VoyageNotes"].ToString(), 260, ColumnRows, 0);


                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesselID"].ToString(), 260, ColumnRows, 0);

                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEXT PORT ETA", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["NextPortETA"].ToString(), 260, ColumnRows, 0);

                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REVISED ETA", 70, ColumnRows -= 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 260, ColumnRows, 0);

                cb.MoveTo(65, ColumnRows -= 5);
                cb.LineTo(500, ColumnRows);


                //vertical
                cb.MoveTo(65, LineVLeft);
                cb.LineTo(65, ColumnRows);

                cb.MoveTo(250, LineVLeft);
                cb.LineTo(250, ColumnRows);

                cb.MoveTo(500, LineVLeft);
                cb.LineTo(500, ColumnRows);

                //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                //png2.SetAbsolutePosition(45, 780);
                //png2.ScalePercent(17f);
                //doc.Add(png2);



                //DataTable dtc = GetAgencyDetails(AgencyID);

                //BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                //cb.SetFontAndSize(bfheader2, 12);
                //cb.SetColorFill(Color.DARK_GRAY);

                //cb.BeginText();

                ////Top1//
                //cb.MoveTo(10, 780);
                //cb.LineTo(660, 780);

                //cb.SetFontAndSize(bfheader2, 12);
                //cb.SetColorFill(Color.BLACK);

                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                //cb.EndText();
                //cb.BeginText();


                ////--------------------

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL DELAY NOTICE ", 250, 730, 0);

                //cb.MoveTo(250, 720);
                //cb.LineTo(400, 720);



                //cb.SetFontAndSize(bfheader2, 9);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 65, 700, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 95, 700, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 65, 680, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 95, 680, 0);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 100, 700, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO VALUABLE CUSTOMER", 100, 680, 0);

                //BaseFont BHHeaderSubject = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //cb.SetFontAndSize(BHHeaderSubject, 11);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUBJECT: VESSEL ARRIVAL DELAYED – " +
                //    "ZHONG GU CHANG SHA V.2428S ETA MYPKG 26/07/24", 65, 650, 0);
                //cb.MoveTo(65, 648);
                //cb.LineTo(500, 648);
                //BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);


                //int ColumnRows = 620;
                //cb.SetFontAndSize(bfheader7, 9);
                //string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                //string[] Aaddsplit;

                //for (int x = 0; x < subjetv.Length; x++)
                //{
                //    Aaddsplit = subjetv[x].Split('\n');

                //    for (int k = 0; k < Aaddsplit.Length; k++)
                //    {

                //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 65, ColumnRows, 0);
                //        ColumnRows -= 14;
                //    }
                //}
                ////box
                //BaseFont bfheader11 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //cb.SetColorFill(Color.BLACK);
                //cb.SetFontAndSize(bfheader11, 9);

                //ColumnRows -= 20;
                //cb.MoveTo(65, ColumnRows);
                //cb.LineTo(500, ColumnRows);

                //ColumnRows -= 13;
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME/VOYAGE NO", 70, ColumnRows, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOAD PORT ETA", 70, ColumnRows -= 20, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YARD CLOSING", 70, ColumnRows -= 20, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN", 70, ColumnRows -= 20, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 70, ColumnRows -= 20, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEXT PORT ETA", 70, ColumnRows -= 20, 0);
                //cb.MoveTo(65, ColumnRows -= 20);
                //cb.LineTo(500, ColumnRows);

                //cb.MoveTo(65, ColumnRows -= 20);
                //cb.LineTo(500, ColumnRows);

                //cb.MoveTo(65, ColumnRows -= 20);
                //cb.LineTo(500, ColumnRows);

                //cb.MoveTo(65, ColumnRows -= 20);
                //cb.LineTo(500, ColumnRows);

                //cb.MoveTo(65, ColumnRows -= 20);
                //cb.LineTo(500, ColumnRows);

                //cb.MoveTo(65, ColumnRows -= 20);
                //cb.LineTo(500, ColumnRows);


                //////vertical
                ////cb.MoveTo(30, 660);
                ////cb.LineTo(30, 540);

                ////cb.MoveTo(250, 660);
                ////cb.LineTo(250, 540);

                ////cb.MoveTo(500, 660);
                ////cb.LineTo(500, 540);


                //////----------------------------------------


                //////
                ////BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                ////cb.SetFontAndSize(bfheader5, 9);
                ////cb.SetColorFill(Color.BLACK);

                //////




                ////BaseFont bfheader10 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                ////cb.SetFontAndSize(bfheader10, 9);
                ////cb.SetColorFill(Color.BLACK);
                //////VESSELVOYAGE
                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 280, 650, 0);

                //////LOAD PORT ETA
                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 280, 630, 0);

                //////yard closing
                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["YardClosing"].ToString(), 280, 610, 0);


                //////SCN
                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VoyageNotes"].ToString(), 280, 590, 0);


                //////VESSEL ID

                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesselID"].ToString(), 280, 570, 0);

                //////NEXT PORT ETA

                ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["NextPortETA"].ToString(), 280, 550, 0);


                //////Text


                cb.Stroke();
                cb.EndText();


                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.End();
            }
        }



        public ActionResult ShippingCertificatePrintPDF(string id, string AgencyID)
        {
            ShippingCertificatePrint(id, AgencyID);
            return View();

        }

        public void ShippingCertificatePrint(string ID, string AgencyID)
        {
            // DataTable dtv = GetdelayData(BLID);
            DataTable dtv = getDelayCertificate(ID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(700, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

               
                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                png.SetAbsolutePosition(25, 780);     //logo fixed location
                png.ScalePercent(15f);
                doc.Add(png);

                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 15);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SHIPPING CERTIFICATE", 260, 700, 0);//right

                cb.MoveTo(260, 698);
                cb.LineTo(440, 698);  //BORDER LINE CODE




                BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader5, 14);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO WHOM IT MAY CONCERN", 260, 780, 0);//right


                BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader7, 12);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE  :", 70, 730, 0);//right


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 660, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 640, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 620, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 600, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 580, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING NUMBER  ", 70, 660, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE OF SHIPMENT", 70, 640, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ", 70, 620, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING  ", 70, 600, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF DISCHARGE  ", 70, 580, 0);//right

                cb.MoveTo(60, 560);
                cb.LineTo(650, 560);


                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 12);
                cb.SetColorFill(Color.BLACK);
                cb.BeginText();

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 130, 730, 0);



                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 260, 660, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Bldate1"].ToString(), 260, 640, 0);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 260, 640, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, 620, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POL"].ToString(), 260, 600, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POD"].ToString(), 260, 580, 0);




                int ColumnRows = 530;
                cb.SetFontAndSize(bfheader3, 9);
                string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                string[] Aaddsplit;

                for (int x = 0; x < subjetv.Length; x++)
                {
                    Aaddsplit = subjetv[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 70, ColumnRows, 0);
                        ColumnRows -= 14;
                    }
                }





                cb.Stroke();
                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.End();
            }
        }

        public ActionResult FreeTimeShippingCertificatePrintPDF(string id, string AgencyID)
        {
            FreeTimeShippingCertificatePrint(id, AgencyID);
            return View();

        }

        public void FreeTimeShippingCertificatePrint(string ID, string AgencyID)
        {
            // DataTable dtv = GetdelayData(BLID);
            DataTable dtv = getDelayCertificate(ID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(700, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                png.SetAbsolutePosition(25, 780);     //logo fixed location
                png.ScalePercent(15f);
                doc.Add(png);

                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 15);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREE TIME SHIPPING CERTIFICATE", 220, 780, 0);//right

                cb.MoveTo(220, 778);
                cb.LineTo(490, 778);  //BORDER LINE CODE


                BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader5, 14);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO WHOM IT MAY CONCERN", 260, 740, 0);//right


                BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader7, 12);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 680, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 660, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 640, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 620, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 600, 0);//right


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE  ", 70, 680, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING NUMBER  ", 70, 660, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL & VOYAGE  ", 70, 640, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING  ", 70, 620, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF DISCHARGE  ", 70, 600, 0);//right



                cb.MoveTo(60, 590);
                cb.LineTo(650, 590);


                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 12);
                cb.SetColorFill(Color.BLACK);
                cb.BeginText();


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Bldate1"].ToString(), 260, 680, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 260, 660, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, 640, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POL"].ToString(), 260, 620, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POD"].ToString(), 260, 600, 0);


                int ColumnRows = 570;
                cb.SetFontAndSize(bfheader3, 9);
                string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                string[] Aaddsplit;

                for (int x = 0; x < subjetv.Length; x++)
                {
                    Aaddsplit = subjetv[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 70, ColumnRows, 0);
                        ColumnRows -= 14;
                    }
                }





                cb.Stroke();
                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.End();
            }
        }

        public ActionResult VesselEarlyArrivalNoticePDF(string id, string AgencyID)
        {
            VesselEarlyArrivalNoticePrint(id, AgencyID);
            return View();

        }

        public void VesselEarlyArrivalNoticePrint(string ID, string AgencyID)
        {
            // DataTable dtv = GetdelayData(BLID);
            DataTable dtv = getDelayCertificate(ID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(700, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                png.SetAbsolutePosition(25, 780);     //logo fixed location
                png.ScalePercent(15f);
                doc.Add(png);

                cb.BeginText();


                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 20);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL EARLY ARRIVAL NOTICE", 220, 750, 0);//right

                cb.MoveTo(220, 748);
                cb.LineTo(560, 748);  //BORDER LINE     CODE


                BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader7, 13);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE  ", 70, 680, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO    ", 70, 660, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 680, 0);//right

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 660, 0);//right


                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 12);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME/VOYAGE NO", 110, 235, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REVISED ETA", 110, 215, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YARD CLOSING", 110, 195, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN", 110, 175, 0);//right
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 110, 155, 0);//right
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA INNSA", 120, 85, 0);//right


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 235, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["RevisedETA"].ToString(), 310, 215, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["YardClosing"].ToString(), 310, 195, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VoyageNotes"].ToString(), 310, 175, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesselID"].ToString(), 310, 155, 0);








                BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader5, 14);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUBJECT : VESSEL ARRIVE EARLY  -  ", 70, 600, 0);//right

                cb.MoveTo(70, 598);
                cb.LineTo(560, 598);  //BORDER LINE     CODE

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA MYPKG", 70, 580, 0);//right


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 130, 680, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO VALUABLE CUSTOMER", 130, 660, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 330, 600, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 175, 580, 0);


               
                cb.MoveTo(70, 578);
                cb.LineTo(300, 578);  //BORDER LINE     CODE

                int ColumnRows = 550;
                cb.SetFontAndSize(bfheader3, 9);
                string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                string[] Aaddsplit;

                for (int x = 0; x < subjetv.Length; x++)
                {
                    Aaddsplit = subjetv[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 70, ColumnRows, 0);
                        ColumnRows -= 14;
                    }
                }


                //Line

                cb.MoveTo(100, 250);
                cb.LineTo(500, 250);  //

                cb.MoveTo(100, 230);
                cb.LineTo(500, 230);  //

                cb.MoveTo(100, 210);
                cb.LineTo(500, 210);  

                cb.MoveTo(100, 190);
                cb.LineTo(500, 190);  

                cb.MoveTo(100, 170);
                cb.LineTo(500, 170);  

                cb.MoveTo(100, 150);
                cb.LineTo(500, 150);

                //cb.MoveTo(100, 80);
                //cb.LineTo(500, 80);  

                //cb.MoveTo(100, 60);
                //cb.LineTo(500, 60);

                //vertical

                cb.MoveTo(500, 250);
                cb.LineTo(500, 150);

                cb.MoveTo(300, 250);
                cb.LineTo(300, 150);


                cb.MoveTo(100, 250);
                cb.LineTo(100, 150);

                //cb.MoveTo(500, 660);
                //cb.LineTo(500, 540);

                cb.Stroke();
                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.End();
            }
        }


        public DataTable GetVesselData(string BLID)
        {
            string _Query = " Select BLID,(SELECT TOP(1) BLNumber FROM nvo_BOL where Id =NVO_VesselCertificate.BLID) as BLNumber," +
                            " CertificateType, CertificateTitle, VSLVoyage, Subject, convert(varchar,Date,106) as Date  from NVO_VesselCertificate where BLID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetdelayData(string BLID)
        {
            string _Query = " Select BLID,(SELECT TOP(1) BLNumber FROM nvo_BOL where Id =NVO_DelayCertificate.BLID) as BLNumber," +
                            " (select (select top(1) PortName from NVO_PortMaster where NVO_PortMaster.Id =nvo_BOL.POLID ) from  nvo_BOL where nvo_BOL.ID=BLID) as POL,"+
                            " CertificateType, CertificateTitle, VSLVoyage, Subject, convert(varchar,Date,106) as Date  from NVO_DelayCertificate where BLID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }

        //public DataTable getDelayCertificate(string ID)
        //{
        //    string _Query = " Select NVO_Booking.ID as BkgID,BookingNo as BLNumber,AgentID as AgencyID, CertificateType,CertificateTitle,VSLVoyage,Subject,  convert(varchar,Date,106) as Date ,BkgParty,POL, " +
        //                    " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 282 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as VoyageNotes, " +
        //                    " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 333 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as YardClosing, " +
        //                    " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 335 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as NextPortETA, " +
        //                    " (select (select top(1) VesselID from NVO_VesselMaster where NVO_VesselMaster.ID=NVO_Voyage.VesselID ) from NVO_Voyage where NVO_Voyage.Id=NVO_Booking.VesVoyID) as VesselID, " +
        //                    " (select  top(1) convert(varchar,ETA, 103) as ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID) as ETA, " +
        //                    " (select  top(1) convert(varchar,ETA, 103) as ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID and NVO_VoyageRoute.PortID=NVO_Booking.PODID) as ETA1, " +
        //                    " (select top(1) (select top(1) PortName from NVO_PortMainMaster where NVO_PortMainMaster.ID =NVO_VoyageRoute.PortID)   from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by RID asc) as Port1, " +
        //                    " (select top(1)(select top(1) PortName from NVO_PortMainMaster where NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID)  from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by RID desc) as Port2 " +
        //                    " from NVO_DelayCertificate " +
        //                    " inner join NVO_Booking on NVO_Booking.VesVoyID = NVO_DelayCertificate.VSLVoyage where NVO_DelayCertificate.Id=" + ID;
        //    return Manag.GetViewData(_Query, "");
        //}


        //POD,BLdate,RevisedETA
        public DataTable getDelayCertificate(string ID)
        {
            string _Query = " Select NVO_Booking.ID as BkgID,NVO_Booking.VesVoy,BookingNo as BLNumber,AgentID as AgencyID, CertificateType,CertificateTitle,VSLVoyage, Subject, convert(varchar,Date,106) as Date ,BkgParty,POL,POD, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 282 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as VoyageNotes, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 333 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as YardClosing, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 335 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as NextPortETA, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 340 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as RevisedETA, " +
                            " (select (select top(1) VesselID from NVO_VesselMaster where NVO_VesselMaster.ID=NVO_Voyage.VesselID ) from NVO_Voyage where NVO_Voyage.Id=NVO_Booking.VesVoyID) as VesselID, " +
                            " (select  top(1) convert(varchar,ETA, 103) as ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID) as ETA, " +
                            " (select  top(1) convert(varchar,ETA, 103) as ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID and NVO_VoyageRoute.PortID=NVO_Booking.PODID) as ETA1, " +
                            "(select  top(1) convert(varchar, BLDate, 103) as BLDate from NVO_BLRelease where NVO_BLRelease.BkgID = NVO_Booking.ID) as Bldate1, "+
                            "(select top(1) (select top(1) PortName from NVO_PortMainMaster where NVO_PortMainMaster.ID =NVO_VoyageRoute.PortID)   from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by RID asc) as Port1, " +
                            " (select top(1)(select top(1) PortName from NVO_PortMainMaster where NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID)  from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by RID desc) as Port2 " +
                            " from NVO_DelayCertificate " +
                            " inner join NVO_Booking on NVO_Booking.VesVoyID = NVO_DelayCertificate.VSLVoyage where NVO_DelayCertificate.Id=" + ID;
           
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }
    }
}