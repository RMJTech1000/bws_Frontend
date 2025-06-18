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
    public class FinanceReportInvoiceController : Controller
    {
        FinanceReportManager Manag = new FinanceReportManager();
        NumberConverWords Convertwords = new NumberConverWords();
        // GET: FinanceReportInvoice
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ExportFreightManifestPDF(string idv)
        {
            BindPDF1(idv);
            return View();
        }
        public void BindPDF1(string idv)
        {
            DataTable dtv = GetInvPDFValus(idv);
            if (dtv.Rows.Count > 0)
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
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/logo.png"));
                png1.SetAbsolutePosition(40, 810);
                png1.ScalePercent(45f);
                doc.Add(png1);

                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddress.png"));
                png2.SetAbsolutePosition(400, 810);
                png2.ScalePercent(25f);
                doc.Add(png2);

                //Top Header
                //cb.MoveTo(15, 835);
                //cb.LineTo(650, 835);

                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 10);
                cb.SetColorFill(Color.BLUE);
                //center
                //cb.MoveTo(345, 835);
                //cb.LineTo(345, 600);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "TAX INVOICE", 160, 830, 0);

                //BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //cb.SetFontAndSize(bfheader3, 8);
                //cb.SetColorFill(Color.DARK_GRAY);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ANCA INTERNATIONAL LLP", 20, 805, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "B - 210A, PHASE - II, NOIDA U.P.INDIA", 20, 795, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ANCA INTERNATIONAL LLP", 20, 785, 0);
                cb.BeginText();
                //Border-Top//
                cb.MoveTo(10, 900);
                cb.LineTo(695, 900);

                //left//
                cb.MoveTo(10, 805);
                cb.LineTo(10, 700);
                //right//  
                //Top//
                cb.MoveTo(10, 805);
                cb.LineTo(660, 805);
                //Bottom//
                cb.MoveTo(10, 115);
                cb.LineTo(660, 115);
                //left//
                cb.MoveTo(10, 805);
                cb.LineTo(10, 115);
                //right//      
                cb.MoveTo(660, 805);
                cb.LineTo(660, 115);

                //center
                //cb.MoveTo(330, 805);
                //cb.LineTo(330, 680);

                //cb.MoveTo(695, 935);
                //cb.LineTo(695, 842);
                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);


                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL TO", 15, 760, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 760, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "State Name", 15, 700, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 700, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GSTIN#:", 15, 688, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 688, 0);
                cb.EndText();

                //cb.BeginText();
                //cb.SetFontAndSize(bfheader2, 8);
                //cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLACK WAVE SHIPPING (M) SDN BHD (1024530 - U)", 455, 756, 0);
                //cb.EndText();

                //cb.BeginText();
                //cb.SetFontAndSize(bfheader2, 8);
                //cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "F/ Agent Name & Ref", 345, 740, 0);
                //cb.EndText();
                //cb.BeginText();
                //cb.SetFontAndSize(bfheader2, 8);
                //cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify Party (No claim shall attach for failure to notify)", 15, 712, 0);
                //cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice No", 15, 795, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice Date", 200, 795, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Due Date", 345, 795, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL Number", 500, 795, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel/Voyage", 15, 670, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 170, 670, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 250, 670, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 345, 670, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POO", 470, 670, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POF", 580, 670, 0);
                cb.EndText();


                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage ETA", 15, 641, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage ETD", 170, 641, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Liner BL", 345, 641, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Volume", 470, 641, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gr.Wt (KGS/MT)", 560, 641, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Details :", 15, 602, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Services", 15, 545, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SAC", 170, 545, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Rate Per", 220, 545, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Quantity", 290, 545, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Ex.Rate", 350, 545, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Taxable", 390, 545, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Amount In    INR", 500, 555, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 445, 535, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SGST", 480, 535, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 515, 535, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CGST", 550, 535, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 595, 535, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGST", 630, 535, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORDS", 15, 389, 0);
                cb.EndText();




                cb.BeginText();
                //Center line Small
                //cb.MoveTo(470, 836);
                //cb.LineTo(470, 798);
                //Horizontal line Small
                //cb.MoveTo(330, 798);
                //cb.LineTo(470, 798);
                //horizontal line1 big
                //cb.MoveTo(10, 775);
                //cb.LineTo(660, 775);
                DataTable dtvs = GetInvPDFDtlValues(dtv.Rows[0]["BLID"].ToString());
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["PartyName"].ToString(), 15, 745, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Address"].ToString(), 15, 736, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 745, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 736, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["StateCode"].ToString(), 100, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["GSTIN"].ToString(), 100, 688, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 688, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvoiceNo"].ToString(), 15, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDate"].ToString(), 200, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDueDate"].ToString(), 345, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["BookingNo"].ToString(), 500, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["VesVoy"].ToString(), 15, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["POL"].ToString(), 170, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["POD"].ToString(), 345, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["POO"].ToString(), 450, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["FPOD"].ToString(), 580, 656, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["ETA"].ToString(), 15, 631, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["ETD"].ToString(), 170, 631, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 250, 631, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["CntrCount"].ToString(), 470, 631, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["GrsWt"].ToString(), 570, 631, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["CntrSizeService"].ToString(), 15, 586, 0);

                DataTable dtInvDtls = GetInvCusBillingdtls(dtv.Rows[0]["Id"].ToString());
                int RowGrd = 520;
                for (int i = 0; i < dtInvDtls.Rows.Count; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["NarrationDescription"].ToString(), 15, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["SACCode"].ToString(), 170, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["RatePerUnit"].ToString(), 220, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["Qty"].ToString(), 290, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["ROE"].ToString(), 350, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["LocalAmount"].ToString(), 390, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "18", 595, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0.00", 630, RowGrd, 0);
                    RowGrd -= 10;

                }
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IF ANY DISCREPENCY IN THE INVOICE BRING IT TO THE NOTICE WITHIN 7 DAYS FROM THE INVOICE DATE, OTHERWISE IT WILL BE PRESUMED THAT THE", 15, 309, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT REFLECTED ON THE BILL IS CORRECT AND HAVE BEEN VERIFIED AT YOUR END", 15, 300, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT MUST BE RECEIVED WITHIN THE AGREED CREDIT PERIOD,FAILING WHICH INTEREST @ 24% PER ANNUM WILL BE CHARGED ON OVERDUE INVOICES", 15, 291, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ALL OBJECTIONS/CLAIMS ARE SUBJECTED TO MUMBAI JURISDICTION", 15, 282, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NERIDA SHIPPING PVT LTD", 100, 220, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvoiceNo"].ToString(), 550, 220, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ICICI BANK", 100, 205, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ICIC0001199", 100, 190, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 100, 175, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "119905000426", 550, 205, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Mumbai", 550, 190, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 130, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CreatedBy"].ToString(), 120, 130, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorised Signatory", 450, 130, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "# 909, 1The Avenue', International Airport Road,Opp. The Leela, Andheri (East) Mumbai - 400 059 India.", 250, 105, 0);

                BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader5, 8);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Note- For Security Deposit - Please deposit in our 'Kotak Mahindra Bank' Account, we will not accept the security deposit in ICICI BANK", 15, 270, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Security Deposit Bank Name", 350, 205, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "KOTAK MAHINDRA BANK", 550, 205, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Security Deposit Acc No", 350, 190, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "5045046761", 550, 190, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IFSC", 350, 175, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "KKBK0001400", 550, 170, 0);
                //BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 270, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Acc No", 350, 205, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Branch", 350, 190, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 350, 175, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 550, 170, 0);
                BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader6, 8);
                cb.SetColorFill(Color.BLACK);
                decimal TotalAmount = 0;

                TotalAmount = decimal.Parse(dtv.Rows[0]["InvTotal"].ToString());

                decimal TotalW = Math.Round(Convert.ToDecimal(TotalAmount.ToString("#,#0.00")));
                string Rupes = NumberConverWords.changeCurrencyToWords(TotalW.ToString());

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Rupes.ToUpper(), 15, 374, 0);

                BaseFont bfheader8 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader8, 8);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvAmount"].ToString(), 390, 389, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvTax"].ToString(), 625, 389, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL AMOUNT " + dtv.Rows[0]["InvTotal"].ToString(), 470, 365, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Remarks :", 15, 349, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "E & O E", 15, 324, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name & Details", 15, 243, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Benificiary", 15, 220, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Payment Ref No", 350, 220, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name", 15, 205, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IFSC", 15, 190, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IFSC", 15, 175, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 145, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Prepared By", 120, 145, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for NERIDA SHIPPING PVT LTD", 450, 145, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMMUNICATION OFFICE ADDRESS :", 15, 105, 0);

                //horizontal line2 big
                cb.MoveTo(10, 776);
                cb.LineTo(660, 776);

                //horizontal line3 big
                cb.MoveTo(10, 680);
                cb.LineTo(660, 680);

                cb.MoveTo(10, 614);
                cb.LineTo(660, 614);

                cb.MoveTo(10, 596);
                cb.LineTo(660, 596);

                cb.MoveTo(10, 572);
                cb.LineTo(660, 572);

                cb.MoveTo(10, 530);
                cb.LineTo(660, 530);

                //Center line6 Small
                cb.MoveTo(160, 572);
                cb.LineTo(160, 401);
                cb.MoveTo(210, 572);
                cb.LineTo(210, 401);
                cb.MoveTo(280, 572);
                cb.LineTo(280, 401);
                cb.MoveTo(340, 572);
                cb.LineTo(340, 401);
                cb.MoveTo(380, 572);
                cb.LineTo(380, 359);

                cb.MoveTo(440, 572);
                cb.LineTo(440, 359);

                cb.MoveTo(440, 550);
                cb.LineTo(660, 550);

                //center small
                cb.MoveTo(460, 550);
                cb.LineTo(460, 401);

                cb.MoveTo(510, 550);
                cb.LineTo(510, 401);

                cb.MoveTo(535, 550);
                cb.LineTo(535, 401);

                cb.MoveTo(590, 550);
                cb.LineTo(590, 401);

                cb.MoveTo(615, 550);
                cb.LineTo(615, 401);

                cb.MoveTo(440, 380);
                cb.LineTo(660, 380);


                cb.MoveTo(10, 359);
                cb.LineTo(660, 359);

                //horizontal line6 small
                cb.MoveTo(10, 401);
                cb.LineTo(660, 401);

                cb.MoveTo(10, 339);
                cb.LineTo(660, 339);

                cb.MoveTo(10, 255);
                cb.LineTo(660, 255);
                cb.MoveTo(10, 235);
                cb.LineTo(660, 235);
                cb.MoveTo(10, 160);
                cb.LineTo(660, 160);


                cb.Stroke();
                //cb.BeginText();




                cb.EndText();

                writer.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=Invoices.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }
        }
        //public void BindFreightManifestPDF(string CusId, string idv)
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
        //    iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/logo.png"));
        //    png1.SetAbsolutePosition(40, 810);
        //    png1.ScalePercent(45f);
        //    doc.Add(png1);

        //    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddress.png"));
        //    png2.SetAbsolutePosition(400, 810);
        //    png2.ScalePercent(25f);
        //    doc.Add(png2);

        //    BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader2, 10);
        //    cb.SetColorFill(Color.BLUE);
        //    //center
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "TAX INVOICE", 160, 830, 0);
        //    cb.BeginText();
        //    //Border-Top//
        //    cb.MoveTo(10, 900);
        //    cb.LineTo(695, 900);

        //    //left//
        //    cb.MoveTo(10, 805);
        //    cb.LineTo(10, 700);
        //    //right//  
        //    //Top//
        //    cb.MoveTo(10, 805);
        //    cb.LineTo(660, 805);
        //    //Bottom//
        //    cb.MoveTo(10, 115);
        //    cb.LineTo(660, 115);
        //    //left//
        //    cb.MoveTo(10, 805);
        //    cb.LineTo(10, 115);
        //    //right//      
        //    cb.MoveTo(660, 805);
        //    cb.LineTo(660, 115);

        //    cb.SetFontAndSize(bfheader2, 9);
        //    cb.SetColorFill(Color.BLACK);


        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL TO", 15, 760, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 760, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "State Name", 15, 700, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 700, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GSTIN#:", 15, 688, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 688, 0);
        //    cb.EndText();


        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice No", 15, 795, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice Date", 200, 795, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Due Date", 345, 795, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL Number", 500, 795, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel/Voyage", 15, 670, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 170, 670, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 250, 670, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 345, 670, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POO", 470, 670, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POF", 580, 670, 0);
        //    cb.EndText();


        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage ETA", 15, 641, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage ETD", 170, 641, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Liner BL", 345, 641, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Volume", 470, 641, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gr.Wt (KGS/MT)", 560, 641, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Details :", 15, 602, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Services", 15, 545, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SAC", 170, 545, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Rate Per", 220, 545, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Quantity", 290, 545, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Ex.Rate", 350, 545, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Taxable", 390, 545, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Amount In    INR", 500, 555, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 445, 535, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SGST", 480, 535, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 515, 535, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CGST", 550, 535, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 595, 535, 0);
        //    cb.EndText();

        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGST", 630, 535, 0);
        //    cb.EndText();
        //    cb.BeginText();
        //    cb.SetFontAndSize(bfheader2, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORDS", 15, 389, 0);
        //    cb.EndText();

        //    cb.BeginText();

        //    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader3, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 745, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 736, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 745, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 736, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 100, 700, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 100, 688, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 700, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 688, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 783, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 200, 783, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 783, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 500, 783, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 656, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 170, 656, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 656, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 656, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 450, 656, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 580, 656, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 631, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 170, 631, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 250, 631, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 470, 631, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 570, 631, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 586, 0);

        //    //DataTable dtInvDtls = GetInvCusBillingdtls(dtv.Rows[0]["Id"].ToString());
        //    int RowGrd = 520;

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 170, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 220, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 290, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 350, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 390, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "18", 595, RowGrd, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0.00", 630, RowGrd, 0);
        //        RowGrd -= 10;

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IF ANY DISCREPENCY IN THE INVOICE BRING IT TO THE NOTICE WITHIN 7 DAYS FROM THE INVOICE DATE, OTHERWISE IT WILL BE PRESUMED THAT THE", 15, 309, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT REFLECTED ON THE BILL IS CORRECT AND HAVE BEEN VERIFIED AT YOUR END", 15, 300, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT MUST BE RECEIVED WITHIN THE AGREED CREDIT PERIOD,FAILING WHICH INTEREST @ 24% PER ANNUM WILL BE CHARGED ON OVERDUE INVOICES", 15, 291, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ALL OBJECTIONS/CLAIMS ARE SUBJECTED TO MUMBAI JURISDICTION", 15, 282, 0);


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NERIDA SHIPPING PVT LTD", 100, 220, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 550, 220, 0);


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ICICI BANK", 100, 205, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ICIC0001199", 100, 190, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 100, 175, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "119905000426", 550, 205, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Mumbai", 550, 190, 0);


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 130, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 120, 130, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorised Signatory", 450, 130, 0);

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "# 909, 1The Avenue', International Airport Road,Opp. The Leela, Andheri (East) Mumbai - 400 059 India.", 250, 105, 0);

        //    BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader5, 8);
        //    cb.SetColorFill(Color.BLACK);


        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 270, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Acc No", 350, 205, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Branch", 350, 190, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 350, 175, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 550, 170, 0);
        //    BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader6, 8);
        //    cb.SetColorFill(Color.BLACK);
        //    decimal TotalAmount = 0;

        //    //TotalAmount = decimal.Parse("");

        //    //decimal TotalW = Math.Round(Convert.ToDecimal(TotalAmount.ToString("#,#0.00")));
        //    //string Rupes = NumberConverWords.changeCurrencyToWords(TotalW.ToString());

        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 374, 0);

        //    BaseFont bfheader8 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    cb.SetFontAndSize(bfheader8, 8);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 390, 389, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 625, 389, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL AMOUNT ", 470, 365, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Remarks :", 15, 349, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "E & O E", 15, 324, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name & Details", 15, 243, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Benificiary", 15, 220, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Payment Ref No", 350, 220, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name", 15, 205, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IFSC", 15, 190, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 145, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Prepared By", 120, 145, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for NERIDA SHIPPING PVT LTD", 450, 145, 0);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMMUNICATION OFFICE ADDRESS :", 15, 105, 0);

        //    //horizontal line2 big
        //    cb.MoveTo(10, 776);
        //    cb.LineTo(660, 776);

        //    //horizontal line3 big
        //    cb.MoveTo(10, 680);
        //    cb.LineTo(660, 680);

        //    cb.MoveTo(10, 614);
        //    cb.LineTo(660, 614);

        //    cb.MoveTo(10, 596);
        //    cb.LineTo(660, 596);

        //    cb.MoveTo(10, 572);
        //    cb.LineTo(660, 572);

        //    cb.MoveTo(10, 530);
        //    cb.LineTo(660, 530);

        //    //Center line6 Small
        //    cb.MoveTo(160, 572);
        //    cb.LineTo(160, 401);
        //    cb.MoveTo(210, 572);
        //    cb.LineTo(210, 401);
        //    cb.MoveTo(280, 572);
        //    cb.LineTo(280, 401);
        //    cb.MoveTo(340, 572);
        //    cb.LineTo(340, 401);
        //    cb.MoveTo(380, 572);
        //    cb.LineTo(380, 359);

        //    cb.MoveTo(440, 572);
        //    cb.LineTo(440, 359);

        //    cb.MoveTo(440, 550);
        //    cb.LineTo(660, 550);

        //    //center small
        //    cb.MoveTo(460, 550);
        //    cb.LineTo(460, 401);

        //    cb.MoveTo(510, 550);
        //    cb.LineTo(510, 401);

        //    cb.MoveTo(535, 550);
        //    cb.LineTo(535, 401);

        //    cb.MoveTo(590, 550);
        //    cb.LineTo(590, 401);

        //    cb.MoveTo(615, 550);
        //    cb.LineTo(615, 401);

        //    cb.MoveTo(440, 380);
        //    cb.LineTo(660, 380);


        //    cb.MoveTo(10, 359);
        //    cb.LineTo(660, 359);

        //    //horizontal line6 small
        //    cb.MoveTo(10, 401);
        //    cb.LineTo(660, 401);

        //    cb.MoveTo(10, 339);
        //    cb.LineTo(660, 339);

        //    cb.MoveTo(10, 255);
        //    cb.LineTo(660, 255);
        //    cb.MoveTo(10, 235);
        //    cb.LineTo(660, 235);
        //    cb.MoveTo(10, 160);
        //    cb.LineTo(660, 160);


        //    cb.Stroke();
        //    cb.EndText();

        //    writer.CloseStream = false;
        //    doc.Close();
        //    Response.Buffer = true;
        //    Response.ContentType = "application/pdf";
        //    //Response.AddHeader("content-disposition", "attachment;filename=Invoices.pdf");
        //    Response.Cache.SetCacheability(HttpCacheability.NoCache);
        //    //Response.Write(doc);
        //    Response.End();
        //}


        public DataTable GetInvPDFValus(string InvID)
        {
            string _Query = "select Id, InvoiceNo, convert(varchar, invDate, 23) as InvDate,convert(varchar, InvDueDate, 23) as InvDueDate,PartyName,InvTotal,PartyTypes,PartyID, BLID,BranchID,TaxNo,Address,StateCode,InvTax,InvAmount,intBank,CurrencyID,(select top(1) GSTIN from NVO_CustomerMaster Inner Join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.ID where NVO_CustomerMaster.ID = PartyID) as GSTIN," +
              " (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_InvoiceCusBilling.AgentID ) as Agency," +
              " ( select top(1) UserName from NVO_UserDetails where ID = NVO_InvoiceCusBilling.UserID ) as CreatedBy from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID=" + InvID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvPDFDtlValues(string BLID)
        {
            string _Query = "Select B.ID,BookingNo,VesVoy,POL,POD,POO,FPOD,convert(varchar, (select top(1) ETA from NVO_VoyageDetails where ID = VesVoyID), 23) as ETA,convert(varchar, (select top(1) ETD from NVO_VoyageDetails where ID = VesVoyID), 23) as ETD,'20' + ' x ' + convert(varchar, CTQ20) + ',' + '40' + ' x ' + convert(varchar, CTQ40) as CntrCount,(select top(1) GrsWt from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) as GrsWt,(select top(1) CntrNo from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) +'/' + (select top(1) size from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) +'/' + ServiceType as CntrSizeService from NVO_Booking B INNER JOIN NVO_BOL ON NVO_BOL.BkgID = B.ID where B.ID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvCusBillingdtls(string InvID)
        {
            string _Query = "Select  Id,InvCusBillingID,NarrationDescription,RatePerUnit,Qty,ROE,LocalAmount,(select top(1) SACCODE from NVO_ChargeTB where ID = NarrationID) as SACCode from NVO_InvoiceCusBillingdtls where InvCusBillingID =" + InvID;
            return Manag.GetViewData(_Query, "");
        }
    }
}