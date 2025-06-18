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
    public class BLPrintPDFController : Controller
    {
        // GET: BLPrintPDF
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult BLPrintPDF()
        {
            CreatePDF();
            return View();
        }

        public void CreatePDF(string BkgID)
        {
            string pdfpath = Server.MapPath("./pdfpath/");
            MergeEx pdfmp = new MergeEx();
            pdfmp.SourceFolder = pdfpath;
            string Datev = System.DateTime.Now.ToString();
            var Datevt = Datev.Replace("/", "-");
            Datevt = Datevt.Replace(":", "-");


            pdfmp.DestinationFile = pdfpath + "Multiple_Invoice.pdf";
            string Values = "./pdfpath/Multiple_Invoice.pdf";
            string _FileName = "";
            //int _Xp = 10, _Yp = 785, YDiff = 10;
            DataTable _dt = new DataTable();
            //STARINDIAHBL
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));

            document.Open();
            PdfContentByte cb = writer.DirectContent;

            PdfContentByte cb1 = writer.DirectContent;
            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb1.BeginText();
            cb1.SetFontAndSize(bfheader2, 25);
            cb1.SetColorFill(Color.BLACK);

            //cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MAX XPRESS.", 30, 780, 0);
            cb1.SetFontAndSize(bfheader2, 12);
            cb1.SetColorFill(Color.BLUE);
            //cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TAX INVOICE", 470, 720, 0);
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/taximg.png"));
            png1.SetAbsolutePosition(470, 710);
            png1.ScalePercent(21f);
            document.Add(png1);




            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/invhead.png"));
            png.SetAbsolutePosition(25, 750);
            png.ScalePercent(25f);
            document.Add(png);
            cb1.Stroke();
            cb1.EndText();

            iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/shipmentblue.png"));
            png2.SetAbsolutePosition(30, 560);
            png2.ScalePercent(24.5f);
            document.Add(png2);

            iTextSharp.text.Image png3 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/containerimg.png"));
            png3.SetAbsolutePosition(30, 500);
            png3.ScalePercent(24.5f);
            document.Add(png3);

            iTextSharp.text.Image png4 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/itemsimg.png"));
            png4.SetAbsolutePosition(30, 400);
            png4.ScalePercent(24.5f);
            document.Add(png4);

            iTextSharp.text.Image png5 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/totalimg.png"));
            png5.SetAbsolutePosition(30, 260);
            png5.ScalePercent(24.5f);
            document.Add(png5);

            iTextSharp.text.Image png6 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/totalimg.png"));
            png6.SetAbsolutePosition(30, 242);
            png6.ScalePercent(24.5f);
            document.Add(png6);

            iTextSharp.text.Image png7 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/taxesimg.png"));
            png7.SetAbsolutePosition(320, 200);
            png7.ScalePercent(24.5f);
            document.Add(png7);

            iTextSharp.text.Image png8 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/ImageIndex/formaximg.png"));
            png8.SetAbsolutePosition(400, -100);
            png8.ScalePercent(24.5f);
            document.Add(png8);

            cb.MoveTo(320, 213);
            cb.LineTo(320, 155);
            cb.MoveTo(551, 213);
            cb.LineTo(551, 155);

            cb.MoveTo(320, 155);
            cb.LineTo(551, 155);


            //#region Line Draw
            BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader1, 13);
            cb.SetColorFill(Color.BLACK);

            cb.BeginText();


            //Top Header
            cb.MoveTo(15, 800);
            cb.LineTo(575, 800);

            cb.SetFontAndSize(bfheader2, 8);
            cb.SetColorFill(Color.DARK_GRAY);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MAXXPRESS Global Solutions Private Limited", 235, 739, 0);

            cb.MoveTo(15, 740);
            cb.LineTo(575, 740);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Party Details", 30, 710, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice No", 340, 650, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Date", 470, 650, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Party GSTNO", 30, 600, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Is Tax Exempted", 200, 600, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "State", 340, 600, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "State Code", 470, 600, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL No", 30, 545, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Receipt", 200, 545, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port of Discharge", 340, 545, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Reference No", 470, 545, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Type / Nos", 30, 485, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Measurement", 200, 485, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port of Loading", 340, 485, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Delivery", 470, 485, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel", 340, 445, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Qty", 470, 445, 0);
            cb.SetFontAndSize(bfheader2, 9);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MAXXPRESS GLOBAL SOLUTIONS PRIVATE LIMITED", 30, 695, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EMAA210313111", 340, 635, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "01/03/2021", 470, 635, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "33AAFCA8294G1ZK", 30, 585, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No", 200, 585, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TAMIL NADU", 340, 585, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "33", 470, 585, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MAA120110588", 30, 530, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COCHIN, INDIA", 200, 530, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ONNE, NIGERIA", 340, 530, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NULL", 470, 530, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MSKU0540160", 30, 470, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0.000", 200, 470, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COCHIN, INDIA", 340, 470, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ONNE, NIGERIA", 470, 470, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SARAYU-101S", 340, 430, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "40'HQ X 2,", 470, 430, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ADD TEXT HERE", 30, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "996713", 210, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "INR", 257, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "00", 290, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5000.00", 320, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,00,000.00", 370, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,00,000.00", 434, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,00,000.00", 500, 385, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL", 320, 248, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,00,000.00", 370, 248, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,00,000.00", 434, 248, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,00,000.00", 500, 248, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SGST", 324, 185, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "9%", 395, 185, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,000.00", 440, 185, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,000.00", 505, 185, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CGST", 324, 174, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "9%", 395, 174, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,000.00", 440, 174, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5,000.00", 505, 174, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGST", 324, 163, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "18%", 395, 163, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 440, 163, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 505, 163, 0);
            cb.SetFontAndSize(bfheader2, 8);
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL IN WORDS : FOUR THOUSAND ONE HUNDRED THIRTY ONLY", 30, 120, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REMARKS", 30, 80, 0);
            cb.SetFontAndSize(bfheader2, 7);
            cb.SetColorFill(Color.GRAY);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "All payments to be made only by DD/PO/NEFT in favour of “MAXXPRESS GLOBAL SOLUTIONS PRIVATE LIMITED”", 30, 70, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Above rates are as per agreement, any discrepancy in the above should be brought to our notice within 5 days of issue of this invoice.", 30, 60, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Kindly quote the bill number when making the payment.", 30, 50, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Interest 24% per annum will be charged on bills remaining unpaid after agreed due date from the date of invoice.", 30, 40, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Dispute relating to this invoice must be submitted in writing within 5 days from the receipt of invoice, thereafter no claim will be entertained.", 30, 30, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "All dispute subject to chennai jurisdiction only.", 30, 20, 0);





            //line left 
            cb.MoveTo(15, 800);
            cb.LineTo(15, 5);

            //Line right
            cb.MoveTo(575, 800);
            cb.LineTo(575, 5);

            cb.Stroke();

            cb.SetFontAndSize(bfheader1, 12);
            BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bf, 7);


            cb.EndText();
            document.Close();
            pdfmp.AddFile(_FileName + ".pdf");
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=BookingConfirmation.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            pdfmp.Execute();
            
            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "_POPUP_Window", "<script>window.open('" + Values + "','new','left=10,top=10,width=1200,height=600,scrollbars=yes')</script>", false);        
    }
        }
}