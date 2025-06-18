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
    public class ImportPDFController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: ImportPDF
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult CustomsReportPDF(string id, string bkgid, string DoId, string LocID)
        {
            BindCustomsReportPDF(id,bkgid,DoId,LocID);
            return View();
        }

        public void BindCustomsReportPDF(string id, string bkgid, string DoId, string LocID)
        {
            DataTable dtv = GetDoReportsPrint(id, bkgid, DoId);
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
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
                png1.SetAbsolutePosition(100, 800);
                png1.ScalePercent(28f);
                doc.Add(png1);

                if(LocID == "1") { 
                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaseal.png"));
                png2.SetAbsolutePosition(15, 230);
                png2.ScalePercent(70f);
                doc.Add(png2);
                }
                if (LocID == "3")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/mundraseal.png"));
                    png2.SetAbsolutePosition(15, 230);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
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
                cb.MoveTo(10, 770);
                cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 13);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CY LETTER", 300, 783, 0);
                

                int RowIndex = 719;
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE MANAGER", 15, 743, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CFS"].ToString(), 15, 731, 0);
                var Addresss4 = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n");
                for (int k = 0; k < Addresss4.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addresss4[k].ToString(), 15, RowIndex, 0);
                    RowIndex -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "D.NO. 1088 VILLAGE KHOPTA TALUKA URAN DIST RAIGAD NAVI MUMBAI 400702", 15, 719, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 530, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IssueDate"].ToString(), 580, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RE : MOVEMENT PERMISSION OF CONTAINERS", 120, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUB:-", 15, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL V. VOYAGE", 100, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ARRIVAL PORT & DATE", 100, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Port"].ToString(), 310, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA1"].ToString(), 440, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL # & DATE", 100, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 310, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLDate"].ToString(), 440, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM # & DATE", 100, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMNo"].ToString(), 310, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMDate"].ToString(), 440, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LINE #", 100, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["LineNumber"].ToString(), 310, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DO # & DATE", 100, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DONo"].ToString(), 310, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoDate"].ToString(), 440, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Please find below the Container(s) List", 15, 450, 0);
                int RowIndexV = 510;
                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Please permit to", 100, 560, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Goods Landed to A/C", 100, 525, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Consignee"].ToString(), 200, 525, 0);
                var split = dtv.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                for (int k = 0; k < split.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 200, RowIndexV, 0);
                    RowIndexV -= 12;
                }

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "to remove the following loaded containers to the premises for destuffing", 100, 470, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CntrDetails"].ToString(), 15, 435, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As we have already filed a bond with commissioner of customs under S/43-CONT (B)/ NS/07/2018 container ( bond ). Kindly permit removal if all", 15, 410, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "documents are found in order at your end.", 15, 401, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PERMIT TO TAKE UPTO", 15, 381, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ReturnDate"].ToString(), 200, 381, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Kindly do the needful & oblige.", 15, 356, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Thanking You,", 15, 331, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your's Faithfully", 15, 306, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR NERIDA SHIPPING PVT LTD", 15, 281, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 15, 210, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note : This is computer generated document Hence No signature required.", 100, 210, 0);

                cb.MoveTo(100, 205);
                cb.LineTo(370, 205);
                cb.Stroke();
                //cb.BeginText();




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

        public ActionResult CustomsReportDestuffing(string id, string bkgid, string DoId, string LocID)
        {
            BindCustomsReportDestuffing(id,bkgid,DoId,LocID);
            return View();
        }

        public void BindCustomsReportDestuffing(string id, string bkgid, string DoId, string LocID)
        {
            DataTable dtv = GetDoReportsPrint(id,bkgid,DoId);
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
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
                png1.SetAbsolutePosition(100, 800);
                png1.ScalePercent(28f);
                doc.Add(png1);

                if(LocID == "1") { 
                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaseal.png"));
                png2.SetAbsolutePosition(15, 110);
                png2.ScalePercent(70f);
                doc.Add(png2);
                }

                if (LocID == "3")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/mundraseal.png"));
                    png2.SetAbsolutePosition(15, 110);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
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
                cb.MoveTo(10, 770);
                cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 13);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DOC DESTUFFING LETTER", 250, 783, 0);
                

                int RowIndex = 719;
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE MANAGER", 15, 743, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CFS"].ToString(), 15, 731, 0);
                var Addresss4 = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n");
                for (int k = 0; k < Addresss4.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addresss4[k].ToString(), 15, RowIndex, 0);
                    RowIndex -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTINENTAL WAREHOUSING CORPORATION (NHAVA SHEVA)", 15, 731, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "D.NO. 1088 VILLAGE KHOPTA TALUKA URAN DIST RAIGAD NAVI MUMBAI 400702", 15, 719, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 530, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IssueDate"].ToString(), 580, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RE : Please allow to DOCK Destuff the Import full container(s)", 120, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUB:-", 15, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL V. VOYAGE", 100, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ARRIVAL PORT & DATE", 100, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Port"].ToString(), 310, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA1"].ToString(), 440, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL # & DATE", 100, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 310, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLDate"].ToString(), 440, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM # & DATE", 100, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMNo"].ToString(), 310, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMDate"].ToString(), 440, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LINE #", 100, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["LineNumber"].ToString(), 310, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DO # & DATE", 100, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoNo"].ToString(), 310, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoDate"].ToString(), 440, 600, 0);

                int RowIndexV = 545;
                int RowIndexK = 456;
                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 8);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONSIGNEE", 15, 560, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Consignee"].ToString(), 100, 560, 0);
                var split = dtv.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                for (int k = 0; k < split.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 100, RowIndexV, 0);
                    RowIndexV -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 100, 560, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 100, 545, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 100, 536, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 100, 527, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 100, 518, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 100, 509, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 100, 500, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHA", 15, 465, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CHA"].ToString(), 100, 465, 0);
                var splitv = dtv.Rows[0]["CHAAddress"].ToString().Split('\n');
                for (int k = 0; k < splitv.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitv[k].ToString(), 100, RowIndexK, 0);
                    RowIndexK -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 100, 465, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 100, 456, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 100, 447, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 100, 438, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 100, 429, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 100, 420, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 100, 411, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Please find below the Container(s) List", 15, 381, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CntrDetails"].ToString(), 15, 369, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VALID UPTO", 15, 330, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ReturnDate"].ToString(), 150, 330, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Thanking You,", 15, 300, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your's Faithfully", 15, 270, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR NERIDA SHIPPING PVT LTD", 15, 150, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 15, 100, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note : This is computer generated document Hence No signature required.", 100, 100, 0);

                cb.MoveTo(100, 98);
                cb.LineTo(365, 98);
                cb.Stroke();
                //cb.BeginText();




                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=CustomsReportDestuffing.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }
        }

        public ActionResult FactoryDestuffing(string id, string bkgid, string DoId, string LocID)
        {
            BindFactoryDestuffing(id,bkgid,DoId,LocID);
            return View();
        }

        public void BindFactoryDestuffing(string id, string bkgid, string DoId, string LocID)
        {
            DataTable dtv = GetDoReportsPrint(id, bkgid, DoId);
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
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
                png1.SetAbsolutePosition(100, 800);
                png1.ScalePercent(28f);
                doc.Add(png1);

                if(LocID == "1") { 
                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaseal.png"));
                png2.SetAbsolutePosition(15, 100);
                png2.ScalePercent(70f);
                doc.Add(png2);
                }
                if (LocID == "3")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/mundraseal.png"));
                    png2.SetAbsolutePosition(15, 100);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
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
                cb.MoveTo(10, 770);
                cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 13);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FACTORY DESTUFFING LETTER", 250, 783, 0);
                

                int RowIndex = 719;
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE MANAGER", 15, 743, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CFS"].ToString(), 15, 731, 0);
                var Addresss4 = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n");
                for (int k = 0; k < Addresss4.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addresss4[k].ToString(), 15, RowIndex, 0);
                    RowIndex -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTINENTAL WAREHOUSING CORPORATION (NHAVA SHEVA)", 15, 731, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "D.NO. 1088 VILLAGE KHOPTA TALUKA URAN DIST RAIGAD NAVI MUMBAI 400702", 15, 719, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 530, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IssueDate"].ToString(), 580, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUB:-", 15, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Movement / Removal of Import full container(s) outside . The PORT/CFS premises for factory destuffing", 120, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REF:-", 15, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL V. VOYAGE", 100, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ARRIVAL PORT & DATE", 100, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Port"].ToString(), 310, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA1"].ToString(), 440, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL # & DATE", 100, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 310, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLDate"].ToString(), 440, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM # & DATE", 100, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMNo"].ToString(), 310, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMDate"].ToString(), 440, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LINE #", 100, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["LineNumber"].ToString(), 310, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DO # & DATE", 100, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoNo"].ToString(), 310, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoDate"].ToString(), 440, 600, 0);

                int RowIndexK = 545;
                int RowIndexL = 373;
                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 8);
                cb.SetColorFill(Color.BLACK);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONSIGNEE", 15, 560, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Consignee"].ToString(), 100, 560, 0);
                var split = dtv.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                for (int k = 0; k < split.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 100, RowIndexK, 0);
                    RowIndexK -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 100, 560, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 100, 545, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 100, 536, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 100, 527, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 100, 518, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 100, 509, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 100, 500, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As per given consignee's request please permit the following container(s) to be moved outside the PORT/CFS premises (FOR FACTORY", 15, 485, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DESTUFFING) at consignee's premises on completion of the required customs and port formalities under our S/43-CONT (B)/", 15, 476, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NS/07/2018 container (Bond) executd with the commissioner of customs for moving IMPORT/EXPORT containers and would request", 15, 467, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "you to advise the gate and shed staff accordingly.", 15, 458, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Please find below the Container(s) List", 15, 438, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CntrDetails"].ToString(), 15, 423, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHA", 15, 403, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CHA"].ToString(), 15, 388, 0);
                var splitv = dtv.Rows[0]["CHAAddress"].ToString().Split('\n');
                for (int k = 0; k < splitv.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitv[k].ToString(), 15, RowIndexL, 0);
                    RowIndexL -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 15, 388, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 15, 373, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 15, 364, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 15, 355, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 15, 346, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 15, 337, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 15, 328, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "After destuffing please allow to store the above mentioned containers at out empty container yard. Please find enclosed copy of the permission", 15, 308, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "granted by the container cell of customs.", 15, 299, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VALID UPTO", 100, 279, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ReturnDate"].ToString(), 250, 279, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Thanking You,", 15, 259, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your's Faithfully", 15, 229, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR NERIDA SHIPPING PVT LTD", 15, 140, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 15, 80, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note : This is computer generated document Hence No signature required.", 100, 80, 0);

                cb.MoveTo(100, 78);
                cb.LineTo(365, 78);
                cb.Stroke();
                //cb.BeginText();




                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=FactoryDestuffing.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }
        }

        public ActionResult GangLetter(string id, string bkgid, string DoId, string LocID)
        {
            BindGangLetter(id,bkgid,DoId, LocID);
            return View();
        }

        public void BindGangLetter(string id, string bkgid, string DoId, string LocID)
        {
            DataTable dtv = GetDoReportsPrint(id, bkgid, DoId);
            if(dtv.Rows.Count > 0)
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
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
            png1.SetAbsolutePosition(100, 800);
            png1.ScalePercent(28f);
            doc.Add(png1);
            
                if(LocID == "1")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaseal.png"));
                    png2.SetAbsolutePosition(15, 290);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
                }
                if (LocID == "3")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/mundraseal.png"));
                    png2.SetAbsolutePosition(15, 290);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
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
            cb.MoveTo(10, 770);
            cb.LineTo(660, 770);

            cb.SetFontAndSize(bfheader2, 9);
            cb.SetColorFill(Color.BLACK);

            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
            cb.EndText();
            cb.BeginText();
            cb.SetFontAndSize(bfheader2, 13);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GANG LETTER", 280, 783, 0);
            

            int RowIndex = 719;
            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 8);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE MANAGER", 15, 743, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CFS"].ToString(), 15, 731, 0);
            var Addresss4 = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n");
            for (int k = 0; k < Addresss4.Length; k++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addresss4[k].ToString(), 15, RowIndex, 0);
                RowIndex -= 12;
            }
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 530, 755, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IssueDate"].ToString(), 580, 755, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLEASE ALLOW GANG TO DESTUFF THE BELOW MENTIONED FCL CONTAINER(S)", 15, 690, 0);

            int RowIndexV = 550;
            int RowIndexK = 635;
            BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader4, 8);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "M/S", 15, 650, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CHA"].ToString(), 100, 650, 0);
            var splitv = dtv.Rows[0]["CHAAddress"].ToString().Split('\n');
            for (int k = 0; k < splitv.Length; k++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitv[k].ToString(), 100, RowIndexK, 0);
                RowIndexK -= 12;
            }
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 100, 650, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 100, 635, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 100, 626, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 100, 617, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 100, 608, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 100, 599, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 100, 590, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AC.M/S", 15, 565, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Consignee"].ToString(), 100, 565, 0);
            var split = dtv.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
            for (int k = 0; k < split.Length; k++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 100, RowIndexV, 0);
                RowIndexV -= 12;
            }
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 100, 565, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 100, 550, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 100, 541, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 100, 532, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 100, 523, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 100, 514, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 100, 505, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL", 80, 480, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 270, 480, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 300, 480, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL#", 80, 466, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 270, 466, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 300, 466, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM", 80, 452, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 270, 452, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMNo"].ToString(), 300, 452, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LINE #", 80, 438, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 270, 438, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["LineNumber"].ToString(), 300, 438, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE CONTAINER NOS ARE FOLLOWS", 15, 410, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CntrDetails"].ToString(), 15, 395, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VALID UPTO:", 15, 370, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ReturnDate"].ToString(), 150, 370, 0);
          
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR NERIDA SHIPPING PVT LTD", 15, 330, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 15, 270, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note : This is computer generated document Hence No signature required.", 100, 270, 0);

            cb.MoveTo(100, 268);
            cb.LineTo(365, 268);
            cb.Stroke();
                //cb.BeginText();




                cb.EndText();

                writer.CloseStream = false;
           
            doc.Close();
            
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=GangLetter.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();
            }
        }

        public ActionResult OffLoadingLetter(string id, string bkgid, string DoId, string LocID)
        {
            BindOffLoadingLetter(id,bkgid,DoId, LocID);
            return View();
        }

        public void BindOffLoadingLetter(string id, string bkgid, string DoId, string LocID)
        {
            DataTable dtv = GetDoReportsPrint(id, bkgid, DoId);
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
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
                png1.SetAbsolutePosition(100, 800);
                png1.ScalePercent(28f);
                doc.Add(png1);

                if(LocID == "1") { 
                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaseal.png"));
                png2.SetAbsolutePosition(15, 100);
                png2.ScalePercent(70f);
                doc.Add(png2);
                }
                if (LocID == "3")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/mundraseal.png"));
                    png2.SetAbsolutePosition(15, 100);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
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
                cb.MoveTo(10, 770);
                cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
               
                cb.SetFontAndSize(bfheader2, 13);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OFF LOADING LETTER", 250, 783, 0);
                

                int RowIndex = 719;
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE MANAGER", 15, 743, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CFS"].ToString(), 15, 731, 0);
                var Addresss4 = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n");
                for (int k = 0; k < Addresss4.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addresss4[k].ToString(), 15, RowIndex, 0);
                    RowIndex -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTINENTAL WAREHOUSING CORPORATION (NHAVA SHEVA)", 15, 731, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "D.NO. 1088 VILLAGE KHOPTA TALUKA URAN DIST RAIGAD NAVI MUMBAI 400702", 15, 719, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 530, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IssueDate"].ToString(), 580, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUB:-", 15, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Kindly arrange to off load below container(s) in your depot on or before date -    2020-11-17", 120, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REF:-", 15, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL V. VOYAGE", 100, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ARRIVAL PORT & DATE", 100, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Port"].ToString(), 310, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA1"].ToString(), 440, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL # & DATE", 100, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 310, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLDate"].ToString(), 440, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM # & DATE", 100, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMNo"].ToString(), 310, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMDate"].ToString(), 440, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LINE #", 100, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["LineNumber"].ToString(), 310, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DO # & DATE", 100, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoNo"].ToString(), 310, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 600, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["DoDate"].ToString(), 440, 600, 0);

                int RowIndexK = 545;
                int RowIndexL = 373;
                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONSIGNEE", 15, 560, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Consignee"].ToString(), 100, 560, 0);
                var split = dtv.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                for (int k = 0; k < split.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 100, RowIndexK, 0);
                    RowIndexK -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 100, 560, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 100, 545, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 100, 536, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 100, 527, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 100, 518, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 100, 509, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 100, 500, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Please find below the Container(s) List", 15, 438, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CntrDetails"].ToString(), 15, 423, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHA", 15, 403, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CHA"].ToString(), 15, 388, 0);
                var splitv = dtv.Rows[0]["CHAAddress"].ToString().Split('\n');
                for (int k = 0; k < splitv.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitv[k].ToString(), 15, RowIndexL, 0);
                    RowIndexL -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OECL SHIPPING AND LOGISTICS PVT LTD", 15, 388, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MIDAS HOUSE, 5TH FLOOR, OFFICE NBR:503 SAHAR PLAZA,", 15, 373, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PLAZA, ANDHERI KURLA ROAD, JB NAGAR, ANDHERI", 15, 364, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EAST", 15, 355, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAIL : OPS.MUM@OECL.SG", 15, 346, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAN NO. AABCO3217H TEL : +919892879716 IEC:", 15, 337, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0411020561 MUMBAI 400059 INDIA", 15, 328, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VALID UPTO", 15, 279, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ReturnDate"].ToString(), 150, 279, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Thanking You,", 15, 259, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your's Faithfully", 15, 229, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR NERIDA SHIPPING PVT LTD", 15, 140, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 15, 80, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note : This is computer generated document Hence No signature required.", 100, 80, 0);

                cb.MoveTo(100, 78);
                cb.LineTo(365, 78);
                cb.Stroke();
                //cb.BeginText();




                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=OffLoadingLetter.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }
        }


        public ActionResult SurveyRequestLetterPDF(string id, string bkgid, string DoId, string LocID)
        {
            BindSurveyRequestLetterPDF(id, bkgid, DoId, LocID);
            return View();
        }

        public void BindSurveyRequestLetterPDF(string id, string bkgid, string DoId, string LocID)
        {
            DataTable dtv = GetDoReportsPrint(id, bkgid, DoId);
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
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
                png1.SetAbsolutePosition(100, 800);
                png1.ScalePercent(28f);
                doc.Add(png1);

                if(LocID == "1") { 
                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaseal.png"));
                png2.SetAbsolutePosition(15, 230);
                png2.ScalePercent(70f);
                doc.Add(png2);
                }
                if (LocID == "3")
                {
                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/mundraseal.png"));
                    png2.SetAbsolutePosition(15, 230);
                    png2.ScalePercent(70f);
                    doc.Add(png2);
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
                cb.MoveTo(10, 770);
                cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                
                cb.SetFontAndSize(bfheader2, 13);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SURVEY REQUEST LETTER", 300, 783, 0);
                

                int RowIndex = 719;
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THE MANAGER", 15, 743, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CFS"].ToString(), 15, 731, 0);
                var Addresss4 = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n");
                for (int k = 0; k < Addresss4.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addresss4[k].ToString(), 15, RowIndex, 0);
                    RowIndex -= 12;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "D.NO. 1088 VILLAGE KHOPTA TALUKA URAN DIST RAIGAD NAVI MUMBAI 400702", 15, 719, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 530, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IssueDate"].ToString(), 580, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RE : MOVEMENT PERMISSION OF CONTAINERS", 120, 690, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUB:-", 15, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL V. VOYAGE", 100, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ARRIVAL PORT & DATE", 100, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Port"].ToString(), 310, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA1"].ToString(), 440, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL # & DATE", 100, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 310, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLDate"].ToString(), 440, 642, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM # & DATE", 100, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMNo"].ToString(), 310, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "&", 420, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IGMDate"].ToString(), 440, 628, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LINE #", 100, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 280, 614, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["LineNumber"].ToString(), 310, 614, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER DETAILS", 15, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CntrDetails"].ToString(), 15, 535, 0);
                int RowIndexV = 425;
                BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader4, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "We are granting you the survey as requested without prejudice to rights of the carrier", 15, 510, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your request for survey is time barred under the terms and conditions of the bill of lading", 15, 490, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "However without prejudice to the rights of carrier, we are requesting our surveyors to conduct the survey", 15, 480, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SURVEYOR:", 15, 440, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Surveyor"].ToString(), 100, 440, 0);
                var split = dtv.Rows[0]["SurveyorAddress"].ToString().Split('\n');
                for (int k = 0; k < split.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 100, RowIndexV, 0);
                    RowIndexV -= 12;
                }


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Thanking You,", 15, 331, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your's Faithfully", 15, 306, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR NERIDA SHIPPING PVT LTD", 15, 281, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 15, 210, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note : This is computer generated document Hence No signature required.", 100, 210, 0);

                cb.MoveTo(100, 205);
                cb.LineTo(370, 205);
                cb.Stroke();
                //cb.BeginText();




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


        public ActionResult ArrivalNotice(string id, string AgencyID, string LocID)
        {
            BindArrivalNotice(id, AgencyID, LocID);
            return View();
        }

        public void BindArrivalNotice(string id, string AgencyID, string LocID)
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
            cb.SetFontAndSize(bfheader, 13);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

            DataTable dtc = GetAgencyDetails(AgencyID);
            if (dtc.Rows.Count > 0)
            {

                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/logoBWS.png"));
              
                png1.SetAbsolutePosition(25, 835);     //logo fixed location
                png1.ScalePercent(10f);
                doc.Add(png1);

            }

            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CARGO ARRIVAL NOTICE", 230, 860, 0);//right
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLUEWAVE SHIPPING & LOGISTIC PVT LTD", 390, 880, 0);//right

            BaseFont bfheader51 = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader51, 9);
            cb.SetColorFill(Color.BLACK);


            int ColumnRows12 = 870;
            string[] Aaddsplit = dtc.Rows[0]["Address"].ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int k = 0; k < Aaddsplit.Length; k++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString().ToUpper(), 390, ColumnRows12, 0);
                ColumnRows12 -= 9;
               
            }

            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NO 31-1, JALAN RAMIN 1, BANDAR AMBANG BOTANIC,", 450, 860, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "41200 KLANG SELANGOR, MALAYSIA ,", 450, 850, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TEL 603-3000 8778, FAX : 603-3319 2031 ", 450, 840, 0);
            BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader7, 9);
            cb.SetColorFill(Color.BLACK);

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
            cb.MoveTo(10, 770);
            cb.LineTo(660, 770);

            cb.SetFontAndSize(bfheader2, 9);
            cb.SetColorFill(Color.BLACK);

            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
            
            cb.SetFontAndSize(bfheader2, 13);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOTICE OF ARRIVAL", 250, 783, 0);
           
            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 9);
            cb.SetColorFill(Color.BLACK);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 15, 755, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CFS", 300, 743, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOA DATE :", 530, 755, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, System.DateTime.Now.Date.ToString("dd/MM/yyyy"), 600, 755, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Dear Sir,", 15, 655, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL Number", 15, 640, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel/Voyage", 250, 640, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA", 490, 640, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 15, 625, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 250, 625, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Item No", 490, 625, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Packages", 15, 610, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM No", 250, 610, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGM Date ", 490, 610, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gross.Wt", 15, 595, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Volume :", 250, 595, 0);

            int ColumnRows = 0;
            int RowsColumn = 0;
            DataTable _dtx = GetCanPrint(id);
            if (_dtx.Rows.Count > 0)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["CANRelease"].ToString(), 15, 743, 0);

                ColumnRows = 734;
                RowsColumn = 0;



                string[] Aaddsplit2 = _dtx.Rows[0]["CANReleaseAddress"].ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int k = 0; k < Aaddsplit2.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit2[k].ToString().ToUpper(), 15, ColumnRows, 0);
                    ColumnRows -= 9;
                    RowsColumn++;
                }

                //string[] Aaddsplit2 = _dtx.Rows[0]["CANReleaseAddress"].ToString().Trim().Split('\n');
                //for (int k = 0; k < Aaddsplit2.Length; k++)
                //{
                //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit2[k].ToString().ToUpper(), 15, ColumnRows, 0);
                //    ColumnRows -= 9;
                //    RowsColumn++;
                //}
                ColumnRows = 720;
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["CFSName"].ToString(), 300, 731, 0);
                string[] Aaddsplit3 = _dtx.Rows[0]["CFSAddress"].ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int l = 0; l < Aaddsplit3.Length; l++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit3[l].ToString().ToUpper(), 300, ColumnRows, 0);
                    ColumnRows -= 9;
                    RowsColumn++;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["CFSAddress"].ToString(), 350, 720, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["ETA1"].ToString(), 580, 755, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["BLNumber"].ToString(), 100, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["VesVoy"].ToString(), 340, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["ETA"].ToString(), 530, 640, 0);
                var PODV = _dtx.Rows[0]["POD"].ToString().Split('-');
                var POLV = _dtx.Rows[0]["POL"].ToString().Split('-');
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, POLV[0].ToString(), 100, 625, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, PODV[0].ToString(), 340, 625, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["LineNumber"].ToString(), 530, 625, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["Pakage"].ToString() + " " + _dtx.Rows[0]["NoofPak"].ToString(), 100, 610, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["IGMNo"].ToString(), 340, 610, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["IGMDate"].ToString(), 550, 610, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["Grwt"].ToString(), 100, 595, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "20 X " + _dtx.Rows[0]["GP20"].ToString() + "    40  X  " + _dtx.Rows[0]["GP40"].ToString() + "", 340, 595, 0);
            }


            BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader4, 9);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Details/Seal No", 15, 570, 0);
            int IndexRows = 555;
            int ColumnIndex = 15;
            DataTable dtcntr = GetCanContainerPrint(id);
            
            for (int i = 0; i < dtcntr.Rows.Count; i++)
            {
            
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows[i]["CntrDtls"].ToString() + "", ColumnIndex, IndexRows, 0);

                ColumnIndex += 140;
            }
            IndexRows -= 12;
            // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows.Count + "", ColumnIndex, IndexRows, 0);


            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1.This is to inform you that the above consignment is expected to arrive on above vessel. Kindly arrange to present original bills of lading duly", 15, 511, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "  discharged and obtain delivery order to clear the goods from the Port / CFS / ICD premises on payment of all relevant charges as applicable", 18, 499, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "  with in nomal / granted free days from arrival of vsl at Port or else detention charges will applicable as per prevaling tarrff.", 18, 488, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "2.If the cargo is not cleared within 2 months from the date of arrival the container will be destuffed in accordance with the provisions of the", 15, 470, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "  port regulations and the cargo would be lying in the custody of port at your sole risk as to cost and consequences which please note", 18, 461, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "3.You may contact our office, in case you require any further information with regards to the applicable charges and other necessary", 15, 444, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "  informatiion (Such as IGM no., free time , other charges etc, if any) about your shipment. The Payment of applicable charges must only be in", 18, 435, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAY ORDER or DEMAND DRAFT.", 15, 400, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "4.Delivery Counter will be open from 10.00 am - 5 .00 pm from Monday - Friday. DO & Revalidation Delivery Order will not be released on", 15, 388, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Saturday / Sunday due to weekly off and on PUBLIC HOLIDAY", 18, 379, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Neither it will be released during week days after office hours.", 18, 370, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5.All KYC formalities has to be completed before approaching for Invoice / Delivery order.", 15, 357, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "6.For KYC / INVOICE / DELIVERY ORDER please send mail to contact details", 15, 344, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Mr. Vishal - maa-sales1@blue-wave.in", 15, 310, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Mr. Dhinesh - maa-cs1@blue-wave.in", 15, 290, 0);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Thanking You,", 15, 270, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Your's Faithfully", 15, 245, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR BLUEWAVE SHIPPING & LOGISTIC PVT LTD", 15, 220, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agents((BLUE WAVE SHIPPING & LOGISTIC PTE LTD))", 15, 200, 0);
            BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader5, 8);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*** This is system generated file, doesn't required any signature ****", 150, 80, 0);

            //cb.MoveTo(100, 78);
            //cb.LineTo(365, 78);
            cb.MoveTo(10, 535);
            cb.LineTo(660, 535);
            cb.Stroke();
            //cb.BeginText();




            cb.EndText();

            writer.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=ArrivalNotice.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();
        }

        public ActionResult NOVCANPDF(string id, string AgencyID, string LocID)
        {
            
            BindNOVCANPDF(id, AgencyID, LocID);
            return View();
        }

        public void BindNOVCANPDF(string id, string AgencyID, string LocID)
        {
            Document doc = new Document();
            //Rectangle rec = new Rectangle(670, 900);
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

            DataTable dtc = GetAgencyDetails(AgencyID);
            if (dtc.Rows.Count > 0)
            {
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/logoBWS.png"));
                //png1.SetAbsolutePosition(300, 840);
                //png1.ScalePercent(80f);
                //doc.Add(png1);
                png1.SetAbsolutePosition(25, 835);     //logo fixed location
                png1.ScalePercent(10f);
                doc.Add(png1);

            }

            //Box
            cb.MoveTo(20, 830);
            cb.LineTo(675, 830);

            cb.MoveTo(20, 830);
            cb.LineTo(20, 30);

            //line  bottom box
            cb.MoveTo(20, 30);
            cb.LineTo(675, 30);

            //Line right box code

            cb.MoveTo(675, 830);
            cb.LineTo(675, 30);


            cb.MoveTo(20, 790);
            cb.LineTo(675, 790);


            cb.MoveTo(20, 680);
            cb.LineTo(675, 680);


            cb.MoveTo(20, 570);
            cb.LineTo(675, 570);

            cb.MoveTo(20, 550);
            cb.LineTo(675, 550);

            cb.MoveTo(20, 390);
            cb.LineTo(675, 390);

            cb.Stroke();
            cb.BeginText();

            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 13);
            cb.SetColorFill(Color.BLACK);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CARGO ARRIVAL NOTICE", 250, 860, 0);//right
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLUE WAVE SHIPPING (M) SDN BHD", 450, 880, 0);//right

            BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader5, 8);
            cb.SetColorFill(Color.BLACK);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NO 31-1, JALAN RAMIN 1, BANDAR AMBANG BOTANIC,", 450, 860, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "41200 KLANG SELANGOR, MALAYSIA ,", 450, 850, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TEL 603-3000 8778, FAX : 603-3319 2031 ", 450, 840, 0);
            BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader7, 9);
            cb.SetColorFill(Color.BLACK);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MBL #", 25, 815, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "HBL #", 150, 815, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FEEDER VESSEL / VOYAGE", 300, 815, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOY ETA", 480, 815, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CAN DATE", 600, 815, 0);//right



            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONSIGNEE", 25, 775, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SHIPPER", 380, 775, 0);//right


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BS CODE", 25, 665, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 200, 665, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 400, 665, 0);//right

            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POF", 480, 665, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOLUME", 600, 665, 0);//right


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN", 25, 620, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 200, 620, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DISCHARGE TERMINAL", 400, 620, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GROSS WEIGHT", 600, 620, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER DETAILS", 25, 560, 0);//right CMD PRIYA




            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CBM", 25, 590, 0);//right




            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CARGO DESCRIPTION", 400, 510, 0);//right

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMS & CONDITIONS:", 25, 375, 0);//right  BLUE WAVE SHIPPING (M) SDN BHD (1024530U)

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorized Signatory ", 25, 75, 0);//

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for BLUE WAVE SHIPPING (M) SDN BHD (1024530U) ", 25, 60, 0);//

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "* THIS IS A COMPUTER-GENERATED DOCUMENT AND DOES NOT REQUIRE SIGNATURE *", 200, 40, 0);//right


            BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader1, 10);
            cb.SetColorFill(Color.BLACK);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  Thank you for Importing your Goods through our Line.", 25, 360, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  We are Pleased to advise you that the above goods are expected to arrive Port Klang and You are requested to ", 25, 340, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   present Original Bill of Lading and Collect Delivery Orders against payment of relevant charges by cash or DD,", 25, 320, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  Office Timing for EDO release: Between 10.00 am - 1.00pm & 2.00 pm - 5.00pm from Monday – Friday,", 25, 300, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  Free days - subject to approval / confirmation from our principals, though incorporated in the original B/L", 25, 280, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  If goods are not cleared within the stipulated time same will be abandoned and we will not be responsible for any consequences", 25, 260, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  All DD / Pay order to be made in favour of ''BLUE WAVE SHIPPING (M) SDN BHD'',", 25, 240, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  SFFLA-NCD member / CLA member are not required to pay for container deposit, should submit the membership proof instead,", 25, 220, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*  Non SFFLA-NCD member / Non CLA member are mandatory to pay container deposit of RM500/20GP & RM1000/40HC.", 25, 200, 0);




            DataTable _dtx = GetCanPrint(id);
            if (_dtx.Rows.Count > 0)
            {

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["Grwt"].ToString() + " KGS", 590, 605, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["Cbm"].ToString(), 25, 575, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "20 X " + _dtx.Rows[0]["GP20"].ToString() + "    40  X  " + _dtx.Rows[0]["GP40"].ToString() + "", 590, 650, 0);
            }

            DataTable dt = GetCanInvoice(id);
            if (dt.Rows.Count > 0)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BLNumber"].ToString(), 25, 800, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["HBLNo"].ToString(), 150, 800, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["NOADate"].ToString(), 600, 800, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesVoy"].ToString(), 300, 800, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["Terminal"].ToString(), 400, 605, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["ETA1"].ToString(), 480, 800, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["POL"].ToString(), 200, 650, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["POD"].ToString(), 400, 650, 0);
                if (LocID == "107")
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["UENCode"].ToString(), 360, 640, 0);
                else
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BSCode"].ToString(), 25, 650, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["SCNCode"].ToString(), 25, 605, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesselID"].ToString(), 200, 605, 0);

            }


            cb.SetFontAndSize(bfheader1, 9);
            cb.SetColorFill(Color.BLACK);
            int ColumnRows = 760; int RowsColumn = 0;
            RowsColumn = 0;
            string[] ArrayAddress1 = Regex.Split(dt.Rows[0]["CONSIGNEE"].ToString().Trim().ToUpper() + "\r" + dt.Rows[0]["CANReleaseAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));

            //string[] ArrayAddress1 = Regex.Split(dt.Rows[0][""].ToString().Trim().ToUpper() + "\r" + dt.Rows[0][""].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
            string[] Aaddsplit1;

            for (int x = 0; x < ArrayAddress1.Length; x++)
            {
                Aaddsplit1 = ArrayAddress1[x].Split('\n');

                for (int k = 0; k < Aaddsplit1.Length; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit1[k].ToString(), 25, ColumnRows, 0);
                    ColumnRows -= 10;
                    RowsColumn++;
                }
            }

            cb.SetFontAndSize(bfheader1, 9);
            cb.SetColorFill(Color.BLACK);
            //SHIPPER
            int ColumnRows1 = 760;
            int RowsColumn1 = 0;
            RowsColumn = 0;
            string[] splitv = Regex.Split(dt.Rows[0]["Shipper"].ToString().Trim().ToUpper() + "\r" + dt.Rows[0]["ShipperAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
            string[] splitv1;
            for (int x = 0; x < splitv.Length; x++)
            {
                splitv1 = splitv[x].Split('\n');

                for (int k = 0; k < splitv1.Length; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitv1[k].ToString(), 380, ColumnRows1, 0);
                    ColumnRows1 -= 10;
                    RowsColumn++;
                }
            }

            cb.SetFontAndSize(bfheader1, 10);
            cb.SetColorFill(Color.BLACK);


            int IndexRows = 540;
            int ColumnIndex = 25;
            int PageWidth = 595;
            int ColumnWidth = 140;
            int MaxColumnsPerRow = (PageWidth - ColumnIndex) / ColumnWidth;
            int RowHeight = 15;
            int ColumnSpacing = 160;

            DataTable dtcntr = GetCanContainerPrint(id);

            for (int i = 0; i < dtcntr.Rows.Count; i++)
            {
                if ((i % MaxColumnsPerRow) == 0 && i != 0)
                {
                    ColumnIndex = 25;
                    IndexRows -= RowHeight;
                }
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows[i]["CntrDtls"].ToString() + "", ColumnIndex, IndexRows, 0);
                ColumnIndex += ColumnSpacing;
            }


            //int IndexRows = 540; // Initial Y-coordinate for the row
            //int ColumnIndex = 25; // Initial X-coordinate for the first column
            //int PageWidth = 595; // Standard A4 page width
            //int ColumnWidth = 140; // Width allocated for each column
            //int MaxColumnsPerRow = (PageWidth - ColumnIndex) / ColumnWidth; // Max columns per row based on available width
            //int RowHeight = 15; // Adjust this for more space between rows
            //int ColumnSpacing = 160; // Adjust this for more space between columns (column width + some space)

            //// Get container data
            //DataTable dtcntr = GetCanContainerPrint(id);

            //for (int i = 0; i < dtcntr.Rows.Count; i++)
            //{
            //    // Move to the next row when max columns per row are reached
            //    if ((i % MaxColumnsPerRow) == 0 && i != 0)
            //    {
            //        ColumnIndex = 25; // Reset column index to the start of the row
            //        IndexRows -= RowHeight; // Move to the next row with adjusted row height
            //    }

            //    // Print the container details in the specified column and row position
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows[i]["CntrDtls"].ToString(), ColumnIndex, IndexRows, 0);

            //    // Move to the next column with adjusted spacing
            //    ColumnIndex += ColumnSpacing;
            //}



            cb.EndText();

            writer.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=ArrivalNotice.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();
        }

        //public FileResult NOVCANGNLPDF(string id, string AgencyID, string CountryID, string UserName)
        //{
        //    DataTable _dtCom = GetCompnayDetails();
        //    string FileHidpath = "";
        //    string str = FileHidpath;
        //    string mime = "";
        //    MergeEx pdfmp = new MergeEx();

        //    Document doc = new Document();
        //    Rectangle rec = new Rectangle(670, 870);
        //    doc = new Document(rec);
        //    Paragraph para = new Paragraph();



        //    string pdfpath = Server.MapPath("./pdfpath/");

        //    pdfmp.SourceFolder = pdfpath;

        //    pdfmp.DestinationFile = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";
        //    FileHidpath = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";

        //    string _FileName = Session.SessionID.ToString() + id + 1;
        //    PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));
        //    doc.Open();

        //    PdfContentByte cb = writer.DirectContent;
        //    cb.SetColorStroke(Color.BLACK);
        //    int _Xp = 10, _Yp = 785, YDiff = 10;

        //    BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader, 14);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 830, 0);

        //    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/LOGO1.png"));
        //    png2.SetAbsolutePosition(20, 790);
        //    png2.ScalePercent(55f);
        //    doc.Add(png2);

        //    cb.BeginText();
        //    BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader2, 15);
        //    cb.SetColorFill(Color.BLACK);

        //    cb.MoveTo(10, 850);
        //    cb.LineTo(655, 850);

        //    cb.MoveTo(10, 850);
        //    cb.LineTo(10, 40);

        //    cb.MoveTo(655, 850);
        //    cb.LineTo(655, 40);

        //    cb.MoveTo(10, 780);
        //    cb.LineTo(655, 780);

        //    cb.MoveTo(10, 760);
        //    cb.LineTo(655, 760);

        //    //cb.MoveTo(180, 780);
        //    //cb.LineTo(180, 740);
        //    cb.MoveTo(260, 780);
        //    cb.LineTo(260, 740);
        //    cb.MoveTo(420, 780);
        //    cb.LineTo(420, 740);
        //    cb.MoveTo(580, 780);
        //    cb.LineTo(580, 740);

        //    cb.MoveTo(10, 740);
        //    cb.LineTo(655, 740);
        //    cb.MoveTo(10, 720);
        //    cb.LineTo(655, 720);
        //    cb.MoveTo(10, 670);
        //    cb.LineTo(655, 670);



        //    cb.MoveTo(10, 650);
        //    cb.LineTo(655, 650);
        //    cb.MoveTo(10, 630);
        //    cb.LineTo(655, 630);
        //    cb.MoveTo(10, 610);
        //    cb.LineTo(655, 610);

        //    cb.MoveTo(180, 670);
        //    cb.LineTo(180, 610);

        //    //Muthu
        //    cb.MoveTo(260, 630);
        //    cb.LineTo(260, 610);
        //    //end

        //    cb.MoveTo(320, 670);
        //    cb.LineTo(320, 610);
        //    cb.MoveTo(420, 670);
        //    cb.LineTo(420, 610);

        //    cb.MoveTo(500, 670);
        //    cb.LineTo(500, 610);

        //    //muthu
        //    cb.MoveTo(580, 630);
        //    cb.LineTo(580, 610);
        //    //end

        //    cb.MoveTo(10, 590);
        //    cb.LineTo(655, 590);

        //    cb.MoveTo(10, 560);
        //    cb.LineTo(655, 560);

        //    cb.MoveTo(10, 540);
        //    cb.LineTo(655, 540);

        //    cb.MoveTo(10, 420-180);
        //    cb.LineTo(655, 420-180);

        //    cb.MoveTo(50, 560);
        //    cb.LineTo(50, 420-180);


        //    cb.MoveTo(180, 560);
        //    cb.LineTo(180, 420-180);

        //    cb.MoveTo(260, 560);
        //    cb.LineTo(260, 420-180);

        //    cb.MoveTo(320, 560);
        //    cb.LineTo(320, 420-180);

        //    cb.MoveTo(420, 560);
        //    cb.LineTo(420, 380 - 180);

        //    cb.MoveTo(480, 560);
        //    cb.LineTo(480, 380 - 180);

        //    cb.MoveTo(580, 560);
        //    cb.LineTo(580, 400 - 180);

        //    cb.MoveTo(10, 400-180);
        //    cb.LineTo(655, 400 - 180);
        //    cb.MoveTo(10, 380 - 180);
        //    cb.LineTo(655, 380 - 180);
        //    cb.MoveTo(10, 360 - 180);
        //    cb.LineTo(655, 360 - 180);
        //    cb.MoveTo(10, 320 - 180);
        //    cb.LineTo(655, 320 - 180);






        //    cb.MoveTo(170, 300-180);
        //    cb.LineTo(170, 260-180);
        //    cb.MoveTo(320, 300-180);
        //    cb.LineTo(320, 260-180);
        //    cb.MoveTo(420, 300-180);
        //    cb.LineTo(420, 260-180);

        //    cb.MoveTo(10, 120);
        //    cb.LineTo(655, 120);

        //    cb.MoveTo(10, 100);
        //    cb.LineTo(655, 100);

        //    cb.MoveTo(10, 80);
        //    cb.LineTo(655, 80);

        //    cb.MoveTo(10, 60);
        //    cb.LineTo(655, 60);
        //    cb.MoveTo(10, 40);
        //    cb.LineTo(655, 40);
        //    DataTable dt = GetCanInvoice(id);
        //    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    if (CountryID == "107")
        //    {
        //        cb.SetFontAndSize(bfheader3, 9);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GLOBAL NETWORK LINES PTE LTD", 180, 835, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOTICE OF ARRIVAL WITH INVOICE", 450, 820, 0);
        //        BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader4, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "81 ANSON ROAD, LEVEL 8", 180, 820, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "8.26, 079908 SINGAPORE,", 180, 810, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TEL: +65 6500 6312 / 13 / 14", 180, 800, 0);

        //    }
        //    else
        //    {

        //        cb.SetFontAndSize(bfheader3, 9);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GLOBAL NETWORK LINES SDN BHD", 180, 835, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOTICE OF ARRIVAL WITH INVOICE", 450, 820, 0);
        //        BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader4, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No.3-12,MTBB2,JALAN BATU NILAM 16.", 180, 820, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANDAR BUKIT TINGGI 2, 41200 KLANG,", 180, 810, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SELANGOR D.E", 180, 800, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Office: +603-33193806 / +603-33195806 / +603-33195500", 180, 790, 0);
        //    }


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOA DATE", 450, 800, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["NOADate"].ToString(), 550, 800, 0);

        //    BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader6, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOA DATE", 480, 800, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL NUMBER", 80, 770, 0);
        //    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL DATE", 210, 770, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL/VOYAGE", 330, 770, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMINAL", 480, 770, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA", 590, 770, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONSIGNEE", 15, 730, 0);




        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 80, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 240, 660, 0);
        //    if (CountryID == "107")
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "UEN NO", 330, 660, 0);
        //    else
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BS CODE", 330, 660, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SHIP CALL NO", 435, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 540, 660, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREE DAYS", 80, 620, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMBINED", 190, 620, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DETENTION", 340, 620, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DEMMURAGE", 510, 620, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER DETAILS", 15, 600, 0);



        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "S.NO", 20, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHARGES DESCRIPTION", 70, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RATE PER UNIT", 190, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "QUANTITY", 270, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL AMOUNT", 340, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EX.RATE", 440, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOCAL AMOUNT", 500, 550, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TAX", 600, 550, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORD", 15, 410-180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL", 440, 410-180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GRAND TOTAL", 422, 390 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "E & OE", 15, 370 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANK NAME & DETAILS", 15, 310 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BENEFICIARY", 15, 290 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANK NAME", 15, 270 - 180, 0);





        //    BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader7, 8);
        //    cb.SetColorFill(Color.BLACK);


        //    if (dt.Rows.Count > 0)
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BLNumber"].ToString(), 50, 750, 0);
        //       // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "20‐08‐2023", 210, 750, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesVoy"].ToString(), 290, 750, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["Terminal"].ToString(), 450, 750, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["ETA"].ToString(), 600, 750, 0);


        //    }
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["CONSIGNEE"].ToString(), 15, 710, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["CANReleaseAddress"].ToString(), 15, 690, 0);




        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["POL"].ToString(), 40, 640, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["POD"].ToString(), 190, 640, 0);
        //    if (CountryID == "107")
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["UENCode"].ToString(), 360, 640, 0);
        //    else
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BSCode"].ToString(), 360, 640, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["SCNCode"].ToString(), 430, 640, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesselID1"].ToString(), 530, 640, 0);
        //    DataTable dtF = GetCanFreedaysPrint(id);
        //    if (dtF.Rows.Count > 0)
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtF.Rows[0]["Comp"].ToString(), 280, 620, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtF.Rows[0]["Det"].ToString(), 460, 620, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtF.Rows[0]["Dum"].ToString(), 600, 620, 0);
        //    }

        //    int IndexRows = 580;
        //    int ColumnIndex = 15;
        //    DataTable dtcntr = GetCanContainerPrint(id);
        //    for (int i = 0; i < dtcntr.Rows.Count; i++)
        //    {

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows[i]["CntrDtls"].ToString() + "", ColumnIndex, IndexRows, 0);
        //        ColumnIndex += 140;
        //    }
        //    IndexRows -= 12;

        //    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows.Count + "", ColumnIndex, IndexRows, 0);
        //    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RSTU20201910/20GP", 15, 580, 0);

        //    IndexRows = 530;
        //    int RowsId = 1;
        //    decimal TotalValue = 0;
        //    DataTable dtx = GetTariffExistingImp(id, AgencyID);
        //    for( int z=0; z<dtx.Rows.Count; z++)
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, (RowsId++).ToString(), 20, IndexRows, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtx.Rows[z]["ChargeCode"].ToString(), 70, IndexRows, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtx.Rows[z]["CustomerRate"].ToString() + " " + dtx.Rows[z]["Currency"].ToString(), 190, IndexRows, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtx.Rows[z]["Qty"].ToString(), 270, IndexRows, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtx.Rows[z]["BillAmt"].ToString() + " " + dtx.Rows[z]["Currency"].ToString(), 340, IndexRows, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtx.Rows[z]["ExRate"].ToString(), 440, IndexRows, 0);

        //        decimal Totalamt = (decimal.Parse(dtx.Rows[z]["BillAmt"].ToString()) * decimal.Parse(dtx.Rows[z]["ExRate"].ToString()));
        //        TotalValue += Totalamt;
        //        if (CountryID == "107")
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Totalamt.ToString("#,#0.00") + " " + dtx.Rows[z]["Currency"].ToString(), 500, IndexRows, 0);
        //        else
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Totalamt.ToString("#,#0.00") + " " + "MYR", 500, IndexRows, 0);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, IndexRows, 0);
        //        IndexRows -= 15;
        //    }




        //    decimal TotalW = Convert.ToDecimal(TotalValue.ToString("#,#0.00"));
        //    string Rupes = NumberConverWords.changeCurrencyToWords(TotalW.ToString());

        //    DataTable DtCu = GetCurrency(CountryID);
        //    if (DtCu.Rows.Count > 0)
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Rupes.Replace("Rupees", DtCu.Rows[0]["CurrencyName"].ToString()).ToUpper(), 15, 390 - 180, 0);
        //    }
        //    else
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Rupes.ToUpper(), 15, 390 - 180, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, TotalValue.ToString("#,#0.00"), 500, 410 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, 410 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, TotalValue.ToString("#,#0.00") + " " + DtCu.Rows[0]["CurrencyCode"].ToString(), 500, 390 - 180, 0);


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1.A FEE OF RM50 WILL BE LEVIED ON ALL RETURNED CHEQUES", 15, 350 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "2.ALL LEGAL COST WILL BE ACCRUED AGAINST YOU IF ACTION IS NECESSARY.", 15, 335 - 180, 0);
        //    if (CountryID == "107")
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GLOBAL NETWORK LINES PTE LTD", 180, 290 - 180, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OCBC BANK", 180, 270 - 180, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "63, CHULIA STREET, #11-01, OCBC CENTRE EAST, 049514 SINGAPORE", 180, 250 - 180, 0);

        //    }
        //    else
        //    {


        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GLOBAL NETWORK LINES SDN-BHD", 180, 290 - 180, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MALAYAN BANKING BERHAD", 180, 270 - 180, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "JALAN BUNUS,KUALA LUMPUR", 180, 250 - 180, 0);
        //    }
        //    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["UserName"].ToString(), 180, 230 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, UserName, 180, 230 - 180, 0);
        //    if (CountryID == "107")
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR GLOBAL NETWORK LINES PTE LTD", 240, 230 - 180, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "601622160001", 500, 270 - 180, 0);

        //    }
        //    else
        //    {
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR GLOBAL NETWORK LINES SDN BHD", 240, 230 - 180, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "601622160001", 500, 270 - 180, 0);

        //    }
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT REF.NO", 325, 290 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "INVOICE NO", 500, 290 - 180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ACCOUNT NO", 325, 270 - 180, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AUTHORISED SIGNATORY", 550, 230-180, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANK ADDRESS", 15, 250-180, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PREPARED BY", 15, 230-180, 0);


        //    cb.EndText();


        //    cb.Stroke();
        //    doc.Close();
        //    pdfmp.AddFile(_FileName + ".pdf");

        //    int SheetNo = 1;
        //    string Filesv = "Attach" + id;
        //    string _AttFileName = Filesv;
        //    Document Attdocument = new Document(rec);
        //    PdfWriter Attwriter = PdfWriter.GetInstance(Attdocument, new FileStream(pdfpath + (_AttFileName + SheetNo) + ".pdf", FileMode.Create));
        //    Attdocument.Open();
        //    PdfContentByte Attcb = Attwriter.DirectContent;
        //    Attcb.SetColorStroke(Color.BLACK);

        //    Attcb.BeginText();
        //    if (CountryID == "107")
        //    {
        //        iTextSharp.text.Image png3 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/amendsecond.jpg"));
        //        png3.SetAbsolutePosition(20, 100);
        //        png3.ScalePercent(27f);
        //        Attdocument.Add(png3);
        //    }
        //    else
        //    {
        //        iTextSharp.text.Image png3 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/GNLCANSecond.jpg"));
        //        png3.SetAbsolutePosition(20, 100);
        //        png3.ScalePercent(47f);
        //        Attdocument.Add(png3);
        //    }



        //    Attcb.EndText();

        //    Attdocument.Close();
        //    pdfmp.AddFile(_AttFileName + SheetNo + ".pdf");





        //    pdfmp.Execute();
        //    mime = MimeMapping.GetMimeMapping(FileHidpath);
        //    return File(FileHidpath, mime);
        //}


        public DataTable GetCompnayDetails()
        {
            string _Query = "select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetTariffExistingImp(string ID, string AgencyID)
        {
            
            string _Query = " Select NVO_BLCharges.ID as RCID,CntrType, (select top(1) Size from NVO_tblCntrTypes where ID =CntrType) as CntrTypes, " +
               " ChargeCodeID,(select top(1) ChgDesc from NVO_ChargeTB where ID = ChargeCodeID) as ChargeCode,ChargeCodeID, " +
               " (select top(1) GeneralName from NVO_GeneralMaster where NVO_GeneralMaster.Id = NVO_BLCharges.TariffTypeID) as ChargeType, TariffTypeID as ChargeTypeID," +
               " Case when BasicID = 2 then 'Container' else 'BL' end as Basic,BasicID," +
               " CurrencyID, (select top(1) CurrencyCode from NVO_CurrencyMaster where ID = CurrencyID)  as Currency, " +
               " PaymentModeID, (select top(1) GeneralName from NVO_GeneralMaster where ID = PaymentModeID) as PaymentMode, " +
               " ReqRate,ManifRate,CustomerRate,RateDiff, " +
               " Case when CntrType = 17 then 1 else (select count(CntrID) from NVO_BOLCntrDetails inner join NVO_Containers on NVO_Containers.ID=NVO_BOLCntrDetails.CntrID where NVO_BOLCntrDetails.BLID = NVO_BLCharges.BLID and NVO_BOLCntrDetails.BkgId = NVO_BLCharges.BkgID and TypeID =CntrType)  end as Qty, " +
               " ((Case when CntrType = 17 then 1 else (select count(CntrID) from NVO_BOLCntrDetails inner join NVO_Containers on NVO_Containers.ID=NVO_BOLCntrDetails.CntrID where NVO_BOLCntrDetails.BLID = NVO_BLCharges.BLID and NVO_BOLCntrDetails.BkgId = NVO_BLCharges.BkgID and TypeID =CntrType)  end) * CustomerRate) as BillAmt, " +
               " (select (select top(1) FinalInvoice from NVO_InvoiceCusBilling where Id =InvCusBillingID ) from NVO_InvoiceCusBillingdtls where BLInvID =NVO_BLCharges.ID) as InvoiceNo, " +
               " (select(select top(1) ID from NVO_InvoiceCusBilling where Id = InvCusBillingID ) from NVO_InvoiceCusBillingdtls where BLInvID = NVO_BLCharges.ID) as InvID,isnull(intQty,1) as intQty, " +
               " (select top(1) Rate from NVO_view_DailyExRate where AgencyID= "+ AgencyID + " and FCurrencyID= NVO_BLCharges.CurrencyID) as ExRate " +
               " from NVO_BLCharges " +
               " inner join NVO_Booking on NVO_Booking.ID = NVO_BLCharges.BkgID " +
               " where  NVO_BLCharges.BLID= " + ID + " and  PaymentModeID=19";

            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCurrency(string CountryID)
        {
            string _Query = "select * from NVO_CurrencyMaster where CountryID = " + CountryID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRFreeDays(string RRID)
        {
            string _Query = " select distinct isnull((select ImpFreeDays from NVO_RatesheetMode where RRID = " + RRID + " and ModeID = 1),0) as Comp, " +
                            " isnull((select ImpFreeDays from NVO_RatesheetMode where RRID = " + RRID + " and ModeID = 2),0) as Det, " +
                            " isnull((select ImpFreeDays from NVO_RatesheetMode where RRID = " + RRID + " and ModeID = 3),0) as Demu " +
                            " from NVO_RatesheetMode where RRId = " + RRID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCanPrint(string id)
        {
           // string _Query = "select * from V_NVOImpPrintValue where Id =" + id;
            string _Query = "select * from Nvo_V_ImportCanPrint where Id =" + id;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCanInvoice(string id)
        {
            string _Query = " select BLNumber, RRID, convert(varchar,getdate(), 103) as NOADate,ImpFreeDays, POL, POD, VesVoy,(select top(1)(select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID = NVO_BOLImpVoyageDetails.VesVoyID) " +
                            " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as VesVoy,BLVesVoyID, " +
                            " (select(select top(1) VesselCallSign from NVO_VesselMaster where NVO_VesselMaster.ID = NVO_Voyage.VesselID) " +
                            " from NVO_Voyage where NVO_Voyage.ID = NVO_BOL.BLVesVoyID) as VesselID, " +
                            " (select top(1) convert(varchar, ETA, 103) from NVO_VoyageRoute where VoyageID = NVO_BOL.BLVesVoyID order by RID Desc) as ETA, "+
                            " (select top(1) convert(VARCHAR, ETA, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as ETA1, " +
                            " ImpFreeDays, (SELECT top(1) UserName FROM NVO_UserDetails WHERE NVO_UserDetails.Id = ConfirmedBy) as UserName, " +
                            " (select top(1) (select top(1) (select top(1)TerminalName from  NVO_TerminalMaster where NVO_TerminalMaster.ID=NVO_VoyageRoute.TerminalID)  " +
                            " from NVO_VoyageRoute  where NVO_VoyageRoute.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails  where " +
                            " NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID)  as Terminal, " +
                            " (select top(1) Address from  NVO_CusBranchLocation where CID=NVO_ImpCAN.ReleaseTo) as CANReleaseAddress,  " +
                            " (select Top(1) CustomerName from NVO_view_CustomerDetails where NVO_view_CustomerDetails.CID = NVO_ImpCAN.ReleaseTo) as CONSIGNEE, " +
                            " (select top(1)(select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 283 and NVO_VoyageNotesDtls.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) " +
                            " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as BSCode, " +
                            " (select top(1)(select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 282 and NVO_VoyageNotesDtls.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) " +
                            " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as SCNCode, " +
                             " (select top(1)(select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 332 and NVO_VoyageNotesDtls.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) " +
                            " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as UENCode, " +
                            " (select (select top(1) (select top(1) VesselID from NVO_VesselMaster where NVO_VesselMaster.ID=NVO_Voyage.VesselID) from NVO_Voyage where NVO_Voyage.ID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails " +
                            " where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) VesselID1, " +
                            " (select(select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_BOLCustomerDetails.PartID) from NVO_BOLCustomerDetails where PartyTypeID = 1 and BLID =NVO_BOL.ID) as Shipper, "+
                            " (select(select top(1) Address from  NVO_CusBranchLocation where CID = NVO_BOLCustomerDetails.PartID) from NVO_BOLCustomerDetails where PartyTypeID = 1 and BLID =NVO_BOL.ID) as ShipperAddress,NVO_ImpCAN.HBLNo " +
                            " from NVO_BOL " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_BOL.BkgID " +
                            " left outer join NVO_ImpCAN on NVO_ImpCAN.BLID = NVO_BOL.ID " +
                            " where NVO_BOL.Id=" + id;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCanContainerPrint(string id)
        {
            string _Query = " Select(select top(1) CntrNo from NVO_Containers where Id = NVO_BOLCntrDetails.CntrID) + '/ ' + size + '/ ' + SealNo  as CntrDtls from NVO_BOLCntrDetails where BLID =" + id;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCanFreedaysPrint(string id)
        {
            string _Query = " select distinct " +
                            " isnull((select ImpFreeDays from NVO_BOLMode BM where ModeID = 1 and BM.BLID = NVO_BOLMode.BLID),0) as Comp, " +
                            " isnull((select ImpFreeDays from NVO_BOLMode BM where ModeID = 2 and BM.BLID = NVO_BOLMode.BLID),0) as Det, " +
                            " isnull((select ImpFreeDays from NVO_BOLMode BM where ModeID = 3 and BM.BLID = NVO_BOLMode.BLID),0) as Dum " +
                            " from NVO_BOLMode where BLID = " + id;
            return Manag.GetViewData(_Query, "");
        }

        public IEnumerable<string> GetChunks(string sourceString, int chunkLength)
        {
            using (var sr = new StringReader(sourceString))
            {
                var buffer = new char[chunkLength];
                int read;
                while ((read = sr.Read(buffer, 0, chunkLength)) == chunkLength)
                {
                    yield return new string(buffer, 0, read);
                }
            }
        }

        public DataTable GetDoReportsPrint(string id, string bkgid, string DoId)
        {
            string _Query = "select * from V_NVOCustomsReportDoPrint where Id =" + id + " and BkgID=" + bkgid + " and DOID=" + DoId;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }

    }
}