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
using QRCoder;
using DataManager;
using System.ComponentModel;
using System.Configuration;
using System.Text.RegularExpressions;


namespace NVOCShipping.Controllers
{
    public class InvoicePDFController : Controller
    {

        DocumentManager Manag = new DocumentManager();
        NumberConverWords Convertwords = new NumberConverWords();
        // GET: InvoicePDF
        public ActionResult Index()
        {
            return View();
        }


        public FileResult InvoiceGNLPDF(string idv, string AgencyID, string CountryID)
        {
            string FileHidpath = "";
            string str = FileHidpath;
            string mime = "";
            MergeEx pdfmp = new MergeEx();
            DataTable dtv = GetInvPDFValus(idv);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 870);
                doc = new Document(rec);
                Paragraph para = new Paragraph();



                string pdfpath = Server.MapPath("./pdfpath/");

                pdfmp.SourceFolder = pdfpath;

                pdfmp.DestinationFile = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";
                FileHidpath = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";

                string _FileName = Session.SessionID.ToString() + idv + 1;
                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 830, 0);

                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                png2.SetAbsolutePosition(20, 790);
                png2.ScalePercent(55f);
                doc.Add(png2);

                cb.BeginText();
                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 15);
                cb.SetColorFill(Color.BLACK);

                cb.MoveTo(10, 850);
                cb.LineTo(655, 850);

                cb.MoveTo(655, 850);
                cb.LineTo(655, 30);

                cb.MoveTo(10, 850);
                cb.LineTo(10, 30);


                cb.MoveTo(10, 780);
                cb.LineTo(655, 780);


                cb.MoveTo(10, 760);
                cb.LineTo(655, 760);

                cb.MoveTo(130, 780);
                cb.LineTo(130, 740);


                cb.MoveTo(260, 780);
                cb.LineTo(260, 740);
                cb.MoveTo(420, 780);
                cb.LineTo(420, 740);
                cb.MoveTo(580, 780);
                cb.LineTo(580, 740);

                //POL


                cb.MoveTo(180, 670);
                cb.LineTo(180, 630);

             

                cb.MoveTo(320, 630);
                cb.LineTo(320, 670);

                cb.MoveTo(450, 630);
                cb.LineTo(450, 670);


                cb.MoveTo(570, 630);
                cb.LineTo(570, 670);




                //POL

                cb.MoveTo(10, 740);
                cb.LineTo(655, 740);

                cb.MoveTo(10, 720);
                cb.LineTo(655, 720);

                cb.MoveTo(10, 670);
                cb.LineTo(655, 670);



                cb.MoveTo(10, 650);
                cb.LineTo(655, 650);
                cb.MoveTo(10, 630);
                cb.LineTo(655, 630);

                cb.MoveTo(10, 560);
                cb.LineTo(655, 560);

                cb.MoveTo(10, 540);
                cb.LineTo(655, 540);





                cb.MoveTo(10, 420 - 180);
                cb.LineTo(655, 420 - 180);

                cb.MoveTo(50, 560);
                cb.LineTo(50, 420 - 180);


                cb.MoveTo(180, 560);
                cb.LineTo(180, 420 - 180);

                cb.MoveTo(260, 560);
                cb.LineTo(260, 420 - 180);

                cb.MoveTo(320, 560);
                cb.LineTo(320, 420 - 180);

                cb.MoveTo(420, 560);
                cb.LineTo(420, 380 - 180);

                cb.MoveTo(480, 560);
                cb.LineTo(480, 380 - 180);

                cb.MoveTo(580, 560);
                cb.LineTo(580, 400 - 180);

                cb.MoveTo(10, 400 - 180);
                cb.LineTo(655, 400 - 180);
                cb.MoveTo(10, 380 - 180);
                cb.LineTo(655, 380 - 180);
                cb.MoveTo(10, 360 - 180);
                cb.LineTo(655, 360 - 180);
                cb.MoveTo(10, 320 - 180);
                cb.LineTo(655, 320 - 180);


                cb.MoveTo(100, 300 - 180);
                cb.LineTo(100, 237 - 180);

                cb.MoveTo(320, 300 - 180);
                cb.LineTo(320, 237 - 180);

                cb.MoveTo(420, 300 - 180);
                cb.LineTo(420, 237 - 180);

                cb.MoveTo(10, 120);
                cb.LineTo(655, 120);

                cb.MoveTo(10, 100);
                cb.LineTo(655, 100);

                cb.MoveTo(10, 80);
                cb.LineTo(655, 80);

                cb.MoveTo(10, 60);
                cb.LineTo(655, 60);

                cb.MoveTo(10, 30);
                cb.LineTo(655, 30);


                //Remarks

                cb.MoveTo(320, 640);
                cb.LineTo(320, 500);

                if (AgencyID == "100")
                    AgencyID = "3";

                DataTable dtc = GetAgencyDetails(AgencyID);
                
                    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader3, 9);
                    cb.SetColorFill(Color.BLACK);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 180, 835, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TAX INVOICE", 480, 820, 0);

                    BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader4, 8);
                    cb.SetColorFill(Color.BLACK);
                    int AddRow = 820;
                    var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, LogoAddresss[a].ToString(), 180, AddRow, 0);
                    AddRow -= 10;
                }

                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No.3-12,MTBB2,JALAN BATU NILAM 16.", 180, 820, 0);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANDAR BUKIT TINGGI 2, 41200 KLANG,", 180, 810, 0);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SELANGOR D.E", 180, 800, 0);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Office: + 603-33193806 / +603-33195806 / +603-33195500", 180, 790, 0);
                

                BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader6, 8);
                cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOA DATE", 480, 800, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "INVOICE NO", 50, 770, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLNUMBER", 180, 770, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "INVOICE DATE", 330, 770, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DUE DATE", 480, 770, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "INVTYPE ", 600, 770, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL TO", 15, 730, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 80, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 240, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL/VOYAGE", 360, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA", 480, 660, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD", 590, 660, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER DETAILS", 15, 620, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REMARKS", 325, 620, 0);
                string[] splitremarks = Regex.Split(dtv.Rows[0]["Remarks"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(165));
                int ColumnRowsx = 610;
                string[] remarksplit;

                for (int x = 0; x < splitremarks.Length; x++)
                {
                    remarksplit = splitremarks[x].Split('\n');

                    for (int k = 0; k < remarksplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, remarksplit[k].ToString(), 330, ColumnRowsx, 0);
                        ColumnRowsx -= 9;
                       
                    }
                }


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "S.NO", 20, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHARGES DESCRIPTION", 70, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RATE PER UNIT", 190, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "QUANTITY", 270, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL AMOUNT", 340, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "EX.RATE", 440, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOCAL AMOUNT", 500, 550, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TAX", 600, 550, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORD", 15, 410 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL", 440, 410 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GRAND TOTAL", 422, 390 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "E & OE", 15, 370 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANK NAME & DETAILS", 15, 310 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BENEFICIARY", 15, 290 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANK NAME", 15, 270 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BANK ADDRESS", 15, 250 - 180, 0);

                
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PREPARED BY", 15, 230 - 180, 0);


                BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader7, 8);
                cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "29‐08‐2023", 560, 800, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["FinalInvoice"].ToString(), 50, 750, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 150, 750, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDate"].ToString(), 330, 750, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDueDate"].ToString(), 480, 750, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvTypesv"].ToString(), 600, 750, 0);


                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "MUTHU", 15, 710, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["PartyName"].ToString(), 15, 710, 0);
                cb.SetFontAndSize(bfheader3, 7);
                int ColumnRows = 695;
                int RowsColumn = 0;
                string[] ArrayAddress = Regex.Split(dtv.Rows[0]["Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                string[] Aaddsplit;

                for (int x = 0; x < ArrayAddress.Length; x++)
                {
                    Aaddsplit = ArrayAddress[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 15, ColumnRows, 0);
                        ColumnRows -= 9;
                        RowsColumn++;
                    }
                }


                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, dtv.Rows[0]["POL"].ToString(), 90, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, dtv.Rows[0]["POD"].ToString(), 250, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, dtv.Rows[0]["BLVesVoy"].ToString(), 390, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 470, 640, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETD"].ToString(), 580, 640, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "14", 280, 620, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 460, 620, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, 620, 0);




               // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RSTU20201910/20GP", 15, 580, 0);


                int ColumnRow = 15;
                int RowInx = 610;
                int Rx = 0;
                DataTable _dtx = GetInvPDFCntrValues(dtv.Rows[0]["BLID"].ToString());
                int CountCntr = 15;
                int TotalCountCntr = _dtx.Rows.Count;

                for (int z = 0; z < _dtx.Rows.Count; z++)
                {
                    Rx++;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[z]["CntrNos"].ToString(), ColumnRow, RowInx, 0);
                    ColumnRow += 80;
                    if (Rx > 4)
                    {
                        ColumnRow = 15;
                        RowInx -= 9;
                        Rx = 0;
                        if (CountCntr == (z + 1))
                        {
                            break;
                        }
                        continue;
                    }
                    if (CountCntr == (z + 1))
                    {
                        break;
                    }

                }


                decimal SGSTAmt = 0;
                decimal CGSTAmt = 0;
                decimal IGSTAmt = 0;
                DataTable dtInvDtls = GetInvCusBillingdtls(dtv.Rows[0]["Id"].ToString());
                int RowGrd = 520;
                int SRow = 1;
                decimal TotalValue = 0;
                for (int i = 0; i < dtInvDtls.Rows.Count; i++)
                {
                    
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, SRow.ToString(), 15, RowGrd, 0);
                    SRow++;
                    var splitDesc = SplitByLenght(dtInvDtls.Rows[i]["NarrationDescription"].ToString(), 33);
                    for (int k = 0; k < splitDesc.Length; k++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitDesc[k].ToString(), 55, RowGrd, 0);
                        if (splitDesc.Length >= 2)
                        {
                            if (k == 0)
                            {
                                RowGrd -= 10;
                            }
                        }
                    }
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["RatePerUnit"].ToString() + " " + dtInvDtls.Rows[i]["Currency"].ToString(), 240, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["Qty"].ToString() + " x " + dtInvDtls.Rows[i]["Size"].ToString(), 305, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["Amount"].ToString() + " " + dtInvDtls.Rows[i]["Currency"].ToString(), 400, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["ROE"].ToString(), 460, RowGrd, 0);

                    decimal Totalamt = (decimal.Parse(dtInvDtls.Rows[i]["LocalAmount"].ToString()));
                    TotalValue += Totalamt;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, Totalamt.ToString("#,#0.00") + "-" + dtv.Rows[0]["Currency"].ToString(), 570, RowGrd, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "0", 600, RowGrd, 0);
                    RowGrd -= 12;
                    if (i == 19)
                    {
                        break;
                    }

                    

                }

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1", 20, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OCEAN FREIGHT", 70, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "200 USD", 190, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " 2", 270, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "400 USD", 340, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "4.4", 440, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1760 MYR", 500, 530 - 180, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, 530 - 180, 0);


                decimal TotalW = Convert.ToDecimal(TotalValue.ToString("#,#0.00"));
                string Rupes = NumberConverWords.changeCurrencyToWords(TotalW.ToString());

                DataTable DtCu = GetCurrency(CountryID);
                if (DtCu.Rows.Count > 0)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Rupes.Replace("Rupees", dtv.Rows[0]["Currency"].ToString()).ToUpper(), 15, 390 - 180, 0);
                }
                else
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Rupes.ToUpper(), 15, 390 - 180, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, TotalValue.ToString("#,#0.00"), 500, 410 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, 410 - 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, TotalValue.ToString("#,#0.00") + " " + dtv.Rows[0]["Currency"].ToString(), 500, 390 - 180, 0);

                DataTable _dtb = GetInvPDFBankValues(dtv.Rows[0]["IntBank"].ToString());
                if (_dtb.Rows.Count > 0)
                {

                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ONE THOUSAND EIGHT HUNDRED FIFTY RINGET", 15, 390 - 180, 0);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1850", 500, 410 - 180, 0);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, 410 - 180, 0);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1850 MYR", 500, 390 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "1.A FEE OF RM50 WILL BE LEVIED ON ALL RETURNED CHEQUES", 15, 350 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "2.ALL LEGAL COST WILL BE ACCRUED AGAINST YOU IF ACTION IS NECESSARY.", 15, 335 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 105, 290 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["BankName"].ToString(), 105, 270 - 180, 0);
                    string[] ArrayBranchAddress = _dtb.Rows[0]["BranchAddress"].ToString().Split('\n');
                    int ColumnRowsBranch = 250;
                    for (int k = 0; k < ArrayBranchAddress.Length; k++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ArrayBranchAddress[k].ToString(), 105, ColumnRowsBranch - 180, 0);
                        ColumnRowsBranch -= 9;
                        
                    }

                   

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 180, 230 - 180, 0);
                  
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT REF.NO", 325, 290 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "INVOICE NO", 500, 290 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ACCOUNT NO", 325, 270 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SWIFT CODE", 325, 250 - 180, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["AccountNo"].ToString(), 500, 270 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["SwiftCode"].ToString(), 500, 250 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FOR "+ dtc.Rows[0]["AgencyName"].ToString(), 495, 230 - 180, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AUTHORISED SIGNATORY", 550, 213 - 180, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Prepared By", 40, 52, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CreatedBy"].ToString(), 35, 35, 0);
                    

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*** THIS IS AN ELECTRONIC INVOICE AND DOES NOT REQUIRE AUTHORIZATION ***", 200, 190 - 180, 0);
                    
                }

                cb.EndText();
                cb.Stroke();
                doc.Close();
                pdfmp.AddFile(_FileName + ".pdf");
                int newSheetNo = 1;
                DataTable dtz = GetInvDetentionValues(dtv.Rows[0]["Id"].ToString());
                if (dtz.Rows[0]["DDGID"].ToString() != "")
                {
                    if (dtz.Rows.Count > 0)
                    {

                        for (int a = 0; a < dtz.Rows.Count; a++)
                        {
                            if (dtz.Rows[a]["DDGID"].ToString() != "0")
                            {

                                string Filesv = "Attach" + idv;
                                string _AttFileName = Filesv;
                                
                                Document Attdocument = new Document(rec);
                                PdfWriter Attwriter = PdfWriter.GetInstance(Attdocument, new FileStream(pdfpath + (_AttFileName + newSheetNo) + ".pdf", FileMode.Create));
                                Attdocument.Open();
                                PdfContentByte Attcb = Attwriter.DirectContent;

                                Attcb.BeginText();
                                BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                                Attcb.SetFontAndSize(bfheader, 12);
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 830, 0);


                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL NUMBER", 25, 820, 0);
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BookingNo"].ToString().ToUpper(), 25, 800, 0);
                                //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Contract ID", 300, 820, 0);
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel/Voyage", 450, 820, 0);
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLVesVoy"].ToString(), 450, 800, 0);

                                int RowColumn = 730;
                                DataTable dt1 = GetDDGCntrExisting(dtz.Rows[a]["DDGID"].ToString());
                                for (int y = 0; y < dt1.Rows.Count; y++)
                                {

                                    Attcb.SetFontAndSize(bfheader, 12);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt1.Rows[y]["CntrNo"].ToString() + "-" + dt1.Rows[y]["CntrTypes"].ToString(), 25, RowColumn, 0);
                                    RowColumn -= 20;
                                    Attcb.SetFontAndSize(bfheader, 10);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Lower", 30, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Upper", 80, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "$", 120, RowColumn, 0);


                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "From", 170, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "To", 240, RowColumn, 0);

                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Days", 300, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "x", 345, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Rate", 400, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "=", 445, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Amount ($)", 500, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Local Amount", 580, RowColumn, 0);

                                    RowColumn -= 10;

                                    decimal PerCntrTotal = 0;
                                    decimal TotalCntrValues = 0;
                                    DataTable _dt2 = NewGetDisplayEstimateExisting(dtz.Rows[a]["DDGID"].ToString(), dt1.Rows[y]["CntrID"].ToString());
                                    for (int s = 0; s < _dt2.Rows.Count; s++)
                                    {
                                        Attcb.SetFontAndSize(bfheader, 9);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["LLimit"].ToString(), 30, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["ULimit"].ToString(), 80, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["Amount"].ToString(), 130, RowColumn, 0);

                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt2.Rows[s]["FDatev"].ToString(), 150, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt2.Rows[s]["TDatev"].ToString(), 220, RowColumn, 0);

                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["Days"].ToString(), 300, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "x", 345, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["Amount"].ToString(), 410, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "=", 450, RowColumn, 0);
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["Rate"].ToString(), 515, RowColumn, 0);
                                        PerCntrTotal += decimal.Parse(_dt2.Rows[s]["Rate"].ToString());
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dt2.Rows[s]["Total"].ToString(), 600, RowColumn, 0);
                                        TotalCntrValues += decimal.Parse(_dt2.Rows[s]["Total"].ToString());
                                        RowColumn -= 10;
                                    }
                                    RowColumn -= 10;
                                    Attcb.SetFontAndSize(bfheader, 10);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Per Cntr.Total", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, PerCntrTotal.ToString() + "($)", 525, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, TotalCntrValues.ToString(), 600, RowColumn, 0);

                                }

                                DataTable _dtT = GetImpDDGBiiledAmount(dtv.Rows[0]["BLID"].ToString(), dtz.Rows[a]["DDGID"].ToString());
                                if (_dtT.Rows.Count > 0)
                                {


                                    RowColumn -= 30;
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Total Amount for Per Containers", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dtT.Rows[0]["TotalAmount"].ToString(), 600, RowColumn, 0);
                                    RowColumn -= 15;
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Alredy Billed Amount/Per Container", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dtT.Rows[0]["AlreadyBilled"].ToString(), 600, RowColumn, 0);
                                    RowColumn -= 15;
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Percentage /Per Container", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "0.0", 600, RowColumn, 0);
                                    RowColumn -= 15;
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Waiver /Per Container", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, decimal.Parse(_dtT.Rows[0]["WavierAmt"].ToString()).ToString("#,#0.00"), 600, RowColumn, 0);
                                    RowColumn -= 15;
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Exchange Rate", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, _dtT.Rows[0]["ExRate"].ToString(), 600, RowColumn, 0);
                                    RowColumn -= 15;
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Estimated Total /Per Container", 450, RowColumn, 0);
                                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, (decimal.Parse(_dtT.Rows[0]["TotalAmount"].ToString()) - decimal.Parse(_dtT.Rows[0]["AlreadyBilled"].ToString()) - decimal.Parse(_dtT.Rows[0]["WavierAmt"].ToString())).ToString("#,#0.00") + "($)", 610, RowColumn, 0);

                                }



                                Attcb.SetFontAndSize(bfheader2, 9);
                                Attcb.SetColorFill(Color.BLACK);
                                Attcb.Stroke();
                                Attcb.EndText();
                                Attdocument.Close();
                                pdfmp.AddFile(_AttFileName + newSheetNo + ".pdf");
                                newSheetNo++;
                            }
                        }
                    }
                }
            }



            pdfmp.Execute();
            mime = MimeMapping.GetMimeMapping(FileHidpath);
            return File(FileHidpath, mime);
        }

    

        private static Byte[] BitmapToBytesCode(System.Drawing.Bitmap image)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return stream.ToArray();
            }
        }
        public DataTable GetNotesClauses()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=268";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetInvPDFValus(string InvID)
        {
            string _Query = " select NVO_InvoiceCusBilling.Id, InvoiceNo, convert(varchar, invDate, 103) as InvDate,convert(varchar, InvDueDate, 103) as InvDueDate,PartyName, " +
                            " InvTotal,PartyTypes,PartyID, BLID,BranchID,TaxNo,Address,StateCode,InvTax,InvAmount,intBank,CurrencyID, " +
                            " (select top(1) GSTIN from NVO_CustomerMaster " +
                            " Inner Join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.ID " +
                            " where NVO_CusBranchLocation.CID = PartyID) as GSTIN, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_InvoiceCusBilling.AgentID ) as Agency, " +
                            " (select  top(1) IRN from NVO_EInvoiceGeneration where InvID = NVO_InvoiceCusBilling.Id) as IRNNo, " +
                            " (select  top(1) SingnedQRCode from NVO_EInvoiceGeneration where InvID = NVO_InvoiceCusBilling.Id) as SingnedQRCode, " +
                            " intBank, " +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.Id = NVO_InvoiceCusBilling.CurrencyID) as Currency, " +
                            " CurrencyID,  (select top(1) UserName from NVO_UserDetails where ID = NVO_InvoiceCusBilling.UserID ) as CreatedBy, " +
                            " FinalInvoice,IsFinal,InvTypes,(select top(1) FinalInvoice from NVO_InvoiceCusBilling Inv " +
                            " where Inv.Id = NVO_InvoiceCusBilling.DrID) as DebitInvoice,NVO_InvoiceCusBilling.Remarks,TaxExemption, " +
                            " isnull(CreditNoteType, 0) CreditNoteType, " +
                            " BLNumber,(select top(1) PortName from NVO_PortMaster where NVO_PortMaster.ID = NVO_BOL.POLID)  as POL, " +
                            " (select top(1) PortName from NVO_PortMaster where NVO_PortMaster.ID = NVO_BOL.PODID)  as POD,BLVesVoy,case when NVO_InvoiceCusBilling.BLTypes=1 then 'Export' else 'Import' end as InvTypesv, " +
                            " case when NVO_InvoiceCusBilling.BLTypes = 2 then(select top(1) convert(varchar, ETA, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BKgId = NVO_BOL.BkgID) else (select  top(1) Convert(varchar, ETA, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) end as ETA,  " +
                            " case when NVO_InvoiceCusBilling.BLTypes = 2 then(select top(1) convert(varchar, ETD, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BKgId = NVO_BOL.BkgID) else (select   top(1) Convert(varchar, ETD, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) end as ETD " + 
                            " from NVO_InvoiceCusBilling " +
                            " inner join NVO_BOL on NVO_BOL.ID = NVO_InvoiceCusBilling.BLID " +
                            " where NVO_InvoiceCusBilling.ID = " + InvID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvPDFDtlValues(string BLID, string InvID)
        {
            //string _Query = " Select distinct B.ID,NVO_BOL.BLNumber as BookingNo, " +
            //                " case when NVO_InvoiceCusBilling.BLTypes = 2 then  (select Top(1)(select top(1)(select top(1) VesselName from NVO_VesselMaster where ID = NVO_Voyage.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = NVO_Voyage.ID) as VesVoy from NVO_Voyage  where NVO_Voyage.ID = NVO_VoyageOpenBLdtls.VoyageID) from NVO_VoyageOpenBLdtls where NVO_VoyageOpenBLdtls.BLID = NVO_BOL.ID) else VesVoy end as VesVoy," +
            //                " POL,(select top 1 PortName from NVO_PortMaster where ID =NVO_BOL.PODID) as POD,POO,FPOD,(select  top(1) Convert(varchar,ETA,105) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) as ETA,(select   top(1) Convert(varchar,ETD,105) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) as ETD,'20' + ' x ' + convert(varchar, CTQ20) + ',' + '40' + ' x ' + convert(varchar, CTQ40) as CntrCount,(select top(1) GrsWt from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) as GrsWt,(select top(1) CntrNo from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) +'/' + (select top(1) size from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) +'/' + ServiceType as CntrSizeService,(SELECT top 1 CompanyName  from NVO_NewCompnayDetails) AS CompanyName" +
            //                " from NVO_Booking B INNER JOIN NVO_BOL ON NVO_BOL.BkgID = B.ID  " +
            //                " inner join NVO_InvoiceCusBilling on NVO_InvoiceCusBilling.BLID=NVO_BOL.ID and NVO_InvoiceCusBilling.BkgId=B.ID " +
            //                " where NVO_BOL.ID=" + BLID + " and NVO_InvoiceCusBilling.Id = " + InvID;

            string _Query = " Select distinct B.ID,NVO_BOL.BLNumber as BookingNo, "+
                            "  case when NVO_InvoiceCusBilling.BLTypes = 2 then (select top(1)(select top(1)VesVoy from NVO_View_VoyageDetails  where NVO_View_VoyageDetails.ID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails where  NVO_BOLImpVoyageDetails.BLID=NVO_BOL.ID)  else VesVoy end as VesVoy, (select top 1 PortName from NVO_PortMaster where ID = NVO_BOL.POLID) as POL, " +
                            " (select top 1 PortName from NVO_PortMaster where ID = NVO_BOL.PODID) as POD,POO,  (select top 1 PortName from NVO_PortMaster where ID = NVO_BOL.FPODID) as FPOD, " +
                            " case when NVO_InvoiceCusBilling.BLTypes = 2 then (select top(1) convert(varchar, ETA, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BKgId=B.ID) else (select  top(1) Convert(varchar, ETA, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) end as ETA,  "+
                            " case when NVO_InvoiceCusBilling.BLTypes = 2 then(select top(1) convert(varchar, ETD, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BKgId = B.ID) else (select   top(1) Convert(varchar, ETD, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) end as ETD, " +
                            " '20' + ' x ' + convert(varchar, CTQ20) + ',' + '40' + ' x ' + convert(varchar, CTQ40) as CntrCount, " +
                            " (select top(1) GrsWt from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) as GrsWt1, "+
                            " (select top(1) GRWT from NVO_BLRelease where BLID = NVO_BOL.ID ) as GrsWt,"+
                            " (select top(1) CntrNo from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) +'/' + (select top(1) size from NVO_BOLCntrDetails " +
                            " where BLID = NVO_BOL.ID ) +'/' + ServiceType as CntrSizeService,(SELECT top 1 CompanyName from NVO_NewCompnayDetails) AS CompanyName "+
                            " from NVO_Booking B INNER JOIN NVO_BOL ON NVO_BOL.BkgID = B.ID "+
                            " inner join NVO_InvoiceCusBilling on NVO_InvoiceCusBilling.BLID = NVO_BOL.ID "+
                            " and NVO_InvoiceCusBilling.BkgId = B.ID  where NVO_BOL.ID=" + BLID + " and NVO_InvoiceCusBilling.Id = " + InvID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvPDFCntrValues(string BLID)
        {
            string _Query = " select NVO_Containers.CntrNo + ' / ' + size as CntrNos from NVO_BOLCntrDetails inner join NVO_Containers on NVO_Containers.ID = NVO_BOLCntrDetails.CntrID " +
                            " where NVO_BOLCntrDetails.BLID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvCusBillingdtls(string InvID)
        {
            string _Query = " Select  Id,InvCusBillingID,NarrationDescription,RatePerUnit, (select top(1) Size from NVO_tblCntrTypes where NVO_tblCntrTypes.ID = NVO_InvoiceCusBillingdtls.UnitID) as size, Qty,ROE,Amount,LocalAmount,(select top(1) SACCODE from NVO_ChargeTB where ID = NarrationID) as SACCode,(select top(1) CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.Id = CurrencyID) as Currency, " +
                            " isnull((select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID= NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID= 1 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as SGSTAmt, " +
                            " isnull((select Top(1) Tax_PCT from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 1 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as SGSTPec, " +
                            " isnull((select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 2 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as CGSTAmt, " +
                            " isnull((select Top(1) Tax_PCT from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 2 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as CGSTPec, " +
                            " isnull((select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 3 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as IGSTAmt, " +
                            " isnull((select Top(1) Tax_PCT from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 3 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as IGSTPec " +
                            " from NVO_InvoiceCusBillingdtls where InvCusBillingID =" + InvID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvPDFBankValues(string InvID)
        {
            string _Query = "select * from NVO_FinBankMaster where Id = " + InvID;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetInvDetentionValues(string InvID)
        {
            string _Query = "select distinct (select top(1) DDGID from NVO_BLCharges where Id = BLInvID) as DDGID from NVO_InvoiceCusBillingdtls where InvCusBillingID = " + InvID;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetImpDDGBiiledAmount(string BLID, string DDID)
        {
            //string _Query = " Select NVO_ImpBLDDGCharges.ID,ChargeID,CntrTypeID,ExRate, (select top(1) BLNumber from NVO_BOL where NVO_BOL.ID =NVO_ImpBLDDGCharges.BLID) as BLNumber,isnull((WaverTotalAmount / ExRate),0) as WavierAmt," +
            //                " (select top(1) BLVesVoy from NVO_BOL where NVO_BOL.ID =NVO_ImpBLDDGCharges.BLID) as BLVesVoy, ContID, (select sum(Rate) from NVO_ImpBLDDGChargedtls where  DDGID= NVO_ImpBLDDGCharges.ID) as " +
            //                " TotalAmount,(select isnull((select top(1) TotalAmount from Nvo_View_BLDDGTotalCalculationTwo Tow where One.BLID=Tow.BLID and Tow.RowNo < one.RowNo order by ID desc),0) " +
            //                " from Nvo_View_BLDDGTotalCalculation One where One.BLID=NVO_ImpBLDDGCharges.BLID  and One.ID = NVO_ImpBLDDGCharges.ID) as AlreadyBilled,((select isnull((select top(1) TotalAmount" +
            //                " from Nvo_View_BLDDGTotalCalculationTwo Tow where One.BLID=Tow.BLID " +
            //                " and Tow.RowNo < one.RowNo order by ID desc),0)  from Nvo_View_BLDDGTotalCalculation One " +
            //                " where One.BLID=NVO_ImpBLDDGCharges.BLID  and One.ID = NVO_ImpBLDDGCharges.ID)- isnull((select sum(Rate) from NVO_ImpBLDDGChargedtls where  DDGID= NVO_ImpBLDDGCharges.ID),0))  as BalanceAmt "+
            //                " from NVO_ImpBLDDGCharges WHERE NVO_ImpBLDDGCharges.ID =" + DDID + " AND NVO_ImpBLDDGCharges.BLID= " + BLID;

            //string _Query = " select ID,BLID,ContID,ExRate,BLVesVoy,BLNumber,WavierAmt,sum(TotalAmount) as TotalAmount, ChargeID,(select top(1) CntrTypeID from NVO_ImpBLDDGChargedtls where NVO_ImpBLDDGChargedtls.DDGID=v_NVO_BLDDGTotalCalculationNew.ID) as CntrTypeID,  " +
            //                " sum(isnull(AlreadyBilled, 0)) AlreadyBilled, (sum(TotalAmount) - sum(isnull(AlreadyBilled, 0))) BalanceAmt, ((sum(TotalAmount) - sum(isnull(AlreadyBilled, 0)))  / (select Count(CntrTypeID) from NVO_view_DDGQty  where NVO_view_DDGQty.DDGID = v_NVO_BLDDGTotalCalculationNew.ID and NVO_view_DDGQty.CntrTypeID =  (select top(1) CntrTypeID from NVO_ImpBLDDGChargedtls where NVO_ImpBLDDGChargedtls.DDGID=v_NVO_BLDDGTotalCalculationNew.ID))) as  BalanceAmtTotal  from v_NVO_BLDDGTotalCalculationNew where ID=" + DDID + " and BLID=" + BLID +
            //                " group by ID, BLID, ContID, ExRate, BLVesVoy, BLNumber, WavierAmt,ChargeID";

            //string _Query = " select ID,BLID,ContID,ExRate,BLVesVoy,BLNumber,WavierAmt,sum(TotalAmount) as TotalAmount, ChargeID,(select top(1) CntrTypeID from NVO_ImpBLDDGChargedtls where NVO_ImpBLDDGChargedtls.DDGID=v_NVO_BLDDGTotalCalculationNew.ID) as CntrTypeID,  " +
            //                "  isnull(AlreadyBilled,0) as AlreadyBilled,  (sum(TotalAmount) - isnull(AlreadyBilled, 0)) BalanceAmt,  (sum(TotalAmount) - (isnull(AlreadyBilled, 0))  / (select Count(CntrTypeID) from NVO_view_DDGQty  where NVO_view_DDGQty.DDGID = v_NVO_BLDDGTotalCalculationNew.ID and NVO_view_DDGQty.CntrTypeID = (select top(1) CntrTypeID from NVO_ImpBLDDGChargedtls where NVO_ImpBLDDGChargedtls.DDGID = v_NVO_BLDDGTotalCalculationNew.ID))) as BalanceAmtTotal  from v_NVO_BLDDGTotalCalculationNew where ID=" + DDID + " and BLID=
            //                " group by ID, BLID, ContID, ExRate, BLVesVoy, BLNumber, WavierAmt,ChargeID,AlreadyBilled";

            string _Query = " select  distinct DDGID as ID,  MainTable.BLID,MainTable.CntrID,    " +
                            " (select top(1) ContID from NVO_ImpBLDDGCharges where NVO_ImpBLDDGCharges.ID = MainTable.DDGID) as ContID, " +
                            " (select top(1) ChargeID from NVO_ImpBLDDGCharges where NVO_ImpBLDDGCharges.ID = MainTable.DDGID) as ChargeID,  " +
                            " (select top(1) CntrTypeID from NVO_ImpBLDDGChargedtls where NVO_ImpBLDDGChargedtls.DDGID = MainTable.DDGID) as CntrTypeID," +
                            " (select top(1) BLVesVoy from NVO_BOL where NVO_BOL.ID = MainTable.BLID) as BLVesVoy,             " +
                            " (select top(1) BLNumber from NVO_BOL where NVO_BOL.ID = MainTable.BLID) as BLNumber,            " +
                            " ((select top(1) isnull(WaverTotalAmount, 0) from NVO_ImpBLDDGCharges where ID = MainTable.DDGID) / MainTable.ExRate) as WavierAmt,ExRate,   " +
                            " sum(Rate) as TotalAmount,isnull((select top(1) Rate from v_NVO_BLDDGAlrderBilled subTable where subTable.DDGID < MainTable.DDGID " +
                            " and subTable.BLID = MainTable.BLID and subTable.CntrID = MainTable.CntrID order by DDGID desc),0) as AlreadyBilled, " +
                            " (sum(Rate) - isnull((select top(1) Rate from v_NVO_BLDDGAlrderBilled subTable where subTable.DDGID < MainTable.DDGID " +
                            " and subTable.BLID = MainTable.BLID and subTable.CntrID = MainTable.CntrID order by DDGID desc), 0)) as BalanceAmt, " +
                            " (sum(Rate) - isnull((select top(1) Rate from v_NVO_BLDDGAlrderBilled subTable where subTable.DDGID < MainTable.DDGID " +
                            " and subTable.BLID = MainTable.BLID and subTable.CntrID = MainTable.CntrID order by DDGID desc), 0)) as BalanceAmtTotal " +
                            " from NVO_ImpBLDDGChargedtls MainTable where BLID = " + BLID + " and DDGID = " + DDID +
                            " group by DDGID,BLID,CntrID,ExRate";

            return Manag.GetViewData(_Query, "");
        }


        //public void BindPDFOT(string idv, string AgencyID)
        //{
        //    DataTable dtv = GetInvPDFValus(idv);
        //    if (dtv.Rows.Count > 0)
        //    {
        //        Document doc = new Document();
        //        Rectangle rec = new Rectangle(670, 870);
        //        doc = new Document(rec);
        //        Paragraph para = new Paragraph();


        //        PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
        //        doc.Open();

        //        PdfContentByte cb = writer.DirectContent;
        //        cb.SetColorStroke(Color.BLACK);
        //        int _Xp = 10, _Yp = 785, YDiff = 10;

        //        BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader, 14);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 830, 0);
        //        //iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/logo.png"));
        //        //png1.SetAbsolutePosition(330, 810);
        //        //png1.ScalePercent(45f);
        //        //doc.Add(png1);
        //        DataTable dtc = GetAgencyDetails(AgencyID);
        //        if (dtc.Rows.Count > 0)
        //        {
        //            if (AgencyID == "13")
        //            {
        //                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaaddress1.png"));
        //                png2.SetAbsolutePosition(360, 810);
        //                png2.ScalePercent(18f);
        //                doc.Add(png2);
        //            }
        //            if (AgencyID == "14")
        //            {
        //                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddressmundra.png"));
        //                png2.SetAbsolutePosition(400, 810);
        //                png2.ScalePercent(25f);
        //                doc.Add(png2);
        //            }
        //            if (AgencyID == "15")
        //            {
        //                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddressdelhi.png"));
        //                png2.SetAbsolutePosition(400, 810);
        //                png2.ScalePercent(25f);
        //                doc.Add(png2);
        //            }
        //            if (AgencyID == "16")
        //            {
        //                iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaaddress1.png"));
        //                png2.SetAbsolutePosition(360, 810);
        //                png2.ScalePercent(18f);
        //                doc.Add(png2);
        //            }

        //        }


        //        //Top Header
        //        //cb.MoveTo(15, 835);
        //        //cb.LineTo(650, 835);
        //        cb.BeginText();
        //        BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader2, 15);
        //        cb.SetColorFill(Color.BLACK);
        //        //center
        //        //cb.MoveTo(345, 835);
        //        //cb.LineTo(345, 600);
        //        if (dtv.Rows[0]["InvTypes"].ToString() == "1")
        //        {

        //            if (dtv.Rows[0]["FinalInvoice"].ToString() != "")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "TAX INVOICE", 200, 830, 0);
        //            }
        //            else
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "PROFORMA INVOICE", 200, 830, 0);
        //            }
        //        }
        //        else
        //        {
        //            if (dtv.Rows[0]["FinalInvoice"].ToString() != "")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "CREDIT NOTE", 200, 830, 0);
        //            }
        //            else
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "PROFORMA CREDIT NOTE", 200, 830, 0);
        //            }
        //        }

        //        if (dtv.Rows[0]["FinalInvoice"].ToString() != "" && dtv.Rows[0]["SingnedQRCode"].ToString() != "")
        //        {
        //            QRCodeGenerator _qrCode = new QRCodeGenerator();
        //            QRCodeData _qrCodeData = _qrCode.CreateQrCode(dtv.Rows[0]["SingnedQRCode"].ToString(), QRCodeGenerator.ECCLevel.Q);
        //            QRCode qrCode = new QRCode(_qrCodeData);
        //            System.Drawing.Bitmap qrCodeImage = qrCode.GetGraphic(20);


        //            BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader, 14);
        //            iTextSharp.text.Image png11 = iTextSharp.text.Image.GetInstance(BitmapToBytesCode(qrCodeImage));
        //            png11.SetAbsolutePosition(30, 806);
        //            png11.ScalePercent(2f);
        //            doc.Add(png11);
        //        }
        //        cb.EndText();

        //        cb.BeginText();
        //        //Border-Top//
        //        cb.MoveTo(10, 900);
        //        cb.LineTo(695, 900);

        //        //left//
        //        cb.MoveTo(10, 805);
        //        cb.LineTo(10, 700);
        //        //right//  
        //        //Top//
        //        cb.MoveTo(10, 805);
        //        cb.LineTo(660, 805);
        //        //Bottom//
        //        //cb.MoveTo(10, 115);
        //        //cb.LineTo(660, 115);
        //        //left//
        //        cb.MoveTo(10, 805);
        //        cb.LineTo(10, 35);
        //        //right//      
        //        cb.MoveTo(660, 805);
        //        cb.LineTo(660, 35);

        //        //center
        //        //cb.MoveTo(330, 805);
        //        //cb.LineTo(330, 680);

        //        //cb.MoveTo(695, 935);
        //        //cb.LineTo(695, 842);
        //        cb.SetFontAndSize(bfheader2, 9);
        //        cb.SetColorFill(Color.BLACK);


        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL TO", 15, 760, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 760, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "State Name", 15, 700, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 700, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GSTIN#:", 15, 688, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 688, 0);
        //        cb.EndText();



        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice No", 15, 795, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice Date", 200, 795, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Due Date", 345, 795, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL Number", 500, 795, 0);
        //        cb.EndText();



        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IRN :", 15, 668, 0);
        //        cb.EndText();


        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Services", 15, 638, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SAC", 170, 638, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RatePer(Unit)", 220, 638, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Quantity", 290, 638, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Ex.Rate", 350, 638, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Taxable", 390, 638, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Amount In " + dtv.Rows[0]["Currency"].ToString(), 500, 648, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 445, 630, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SGST", 480, 630, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 515, 630, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CGST", 550, 630, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 595, 630, 0);
        //        cb.EndText();

        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGST", 630, 630, 0);
        //        cb.EndText();
        //        cb.BeginText();
        //        cb.SetFontAndSize(bfheader2, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORDS", 15, 300, 0);
        //        cb.EndText();




        //        cb.BeginText();
        //        //Center line Small
        //        //cb.MoveTo(470, 836);
        //        //cb.LineTo(470, 798);
        //        //Horizontal line Small
        //        //cb.MoveTo(330, 798);
        //        //cb.LineTo(470, 798);
        //        //horizontal line1 big
        //        //cb.MoveTo(10, 775);
        //        //cb.LineTo(660, 775);
        //        DataTable dtvs = GetInvPDFDtlValues(dtv.Rows[0]["BLID"].ToString(), idv);
        //        BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader3, 9);
        //        cb.SetColorFill(Color.BLACK);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["PartyName"].ToString(), 15, 745, 0);
        //        cb.SetFontAndSize(bfheader3, 7);
        //        int ColumnRows = 735;
        //        int RowsColumn = 0;
        //        string[] ArrayAddress = Regex.Split(dtv.Rows[0]["Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
        //        string[] Aaddsplit;

        //        for (int x = 0; x < ArrayAddress.Length; x++)
        //        {
        //            Aaddsplit = ArrayAddress[x].Split('\n');

        //            for (int k = 0; k < Aaddsplit.Length; k++)
        //            {

        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 15, ColumnRows, 0);
        //                ColumnRows -= 9;
        //                RowsColumn++;
        //            }
        //        }



        //        //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Address"].ToString(), 15, 736, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 745, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 736, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["StateCode"].ToString().ToUpper(), 100, 700, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["GSTIN"].ToString().ToUpper(), 100, 688, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 700, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 688, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["FinalInvoice"].ToString().ToUpper(), 15, 783, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDate"].ToString(), 200, 783, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDueDate"].ToString(), 345, 783, 0);


        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IRNNo"].ToString(), 45, 668, 0);


        //        int ColumnRow = 15;
        //        int RowInx = 585;
        //        int Rx = 0;


        //        decimal SGSTAmt = 0;
        //        decimal CGSTAmt = 0;
        //        decimal IGSTAmt = 0;
        //        DataTable dtInvDtls = GetInvCusBillingdtls(dtv.Rows[0]["Id"].ToString());
        //        int RowGrd = 608;
        //        for (int i = 0; i < dtInvDtls.Rows.Count; i++)
        //        {
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["NarrationDescription"].ToString(), 15, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["SACCode"].ToString(), 170, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["RatePerUnit"].ToString(), 257, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["Currency"].ToString(), 262, RowGrd, 0);


        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["Qty"].ToString() + " x " + dtInvDtls.Rows[i]["Size"].ToString(), 315, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["ROE"].ToString(), 375, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["LocalAmount"].ToString(), 430, RowGrd, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["SGSTPec"].ToString(), 458, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["SGSTAmt"].ToString(), 505, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["CGSTPec"].ToString(), 530, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["CGSTAmt"].ToString(), 585, RowGrd, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["IGSTPec"].ToString(), 610, RowGrd, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["IGSTAmt"].ToString(), 655, RowGrd, 0);
        //            SGSTAmt += decimal.Parse(dtInvDtls.Rows[i]["SGSTAmt"].ToString());
        //            CGSTAmt += decimal.Parse(dtInvDtls.Rows[i]["CGSTAmt"].ToString());
        //            IGSTAmt += decimal.Parse(dtInvDtls.Rows[i]["IGSTAmt"].ToString());

        //            RowGrd -= 10;

        //        }
        //        // int RowGrdv = 309;

        //        int RowGrdv = 190;

        //        DataTable _dtn = GetNotesClauses();
        //        for (int d = 0; d < _dtn.Rows.Count; d++)
        //        {
        //            string[] ArrayNotes = Regex.Split(_dtn.Rows[d]["Notes"].ToString().Trim(), char.ConvertFromUtf32(13));
        //            string[] Aaddsplitv;
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, (d + 1).ToString(), 20, RowGrdv, 0);
        //            for (int x = 0; x < ArrayNotes.Length; x++)
        //            {
        //                Aaddsplitv = ArrayNotes[x].Split('\n');
        //                for (int k = 0; k < Aaddsplitv.Length; k++)
        //                {
        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplitv[k].ToString(), 30, RowGrdv, 0);
        //                    RowGrdv -= 9;

        //                }
        //            }
        //        }










        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 130, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CreatedBy"].ToString(), 15, 40, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorised Signatory", 530, 40, 0);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "# 909, 1The Avenue', International Airport Road,Opp. The Leela, Andheri (East) Mumbai - 400 059 India.", 175, 20, 0);

        //        BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader5, 8);
        //        cb.SetColorFill(Color.BLACK);


        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 270, 0);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 350, 175, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 550, 170, 0);
        //        BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader6, 8);
        //        cb.SetColorFill(Color.BLACK);
        //        decimal TotalAmount = 0;

        //        TotalAmount = decimal.Parse(dtv.Rows[0]["InvTotal"].ToString());

        //        decimal TotalW = Convert.ToDecimal(TotalAmount.ToString("#,#0.00"));


        //        string Rupes = NumberConverWords.changeCurrencyToWords(TotalW.ToString());

        //        string Ruppev = "";
        //        if ( dtv.Rows[0]["CurrencyID"].ToString() == "146")
        //        {
        //            Ruppev = Rupes.Replace("Rupees", "DOLLAR");
        //        }
        //        else
        //        {
        //            Ruppev = Rupes;
        //        }

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Ruppev.ToUpper(), 15, 280, 0);

        //        BaseFont bfheader8 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //        cb.SetFontAndSize(bfheader8, 8);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvAmount"].ToString(), 390, 280, 0);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, SGSTAmt.ToString(), 470, 295, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, CGSTAmt.ToString(), 550, 295, 0);
        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, IGSTAmt.ToString(), 623, 295, 0);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL AMOUNT   " + dtv.Rows[0]["InvTotal"].ToString(), 470, 278, 0);

        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Remarks :", 15, 260, 0);
        //        var Remarks = SplitByLenght(dtv.Rows[0]["Remarks"].ToString(), 120);
        //        int RemarksL = 242;
        //        for (int s = 0; s < Remarks.Length; s++)
        //        {
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Remarks[s].ToString(), 15, RemarksL, 0);
        //            RemarksL-=10;
        //        }
        //        if (dtv.Rows[0]["InvTypes"].ToString() == "2")
        //        {
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Original Dr.Note No:   " + dtv.Rows[0]["DebitInvoice"].ToString(), 45, 250, 0);
        //        }


        //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "E & O E", 15, 200, 0);

        //        DataTable _dtb = GetInvPDFBankValues(dtv.Rows[0]["IntBank"].ToString());
        //        if (_dtb.Rows.Count > 0)
        //        {
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name & Details", 15, 135, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Benificiary", 15, 115, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NERIDA SHIPPING PVT LTD", 100, 115, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name", 15, 100, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["BankName"].ToString(), 100, 100, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IFSC", 15, 85, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["IFSCCode"].ToString(), 100, 85, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Payment Ref No", 350, 115, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["FinalInvoice"].ToString(), 550, 115, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Acc No", 350, 100, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["AccountNo"].ToString(), 550, 100, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Branch", 350, 85, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["BranchName"].ToString(), 550, 85, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 145, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Prepared By", 15, 55, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for NERIDA SHIPPING PVT LTD", 530, 55, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMMUNICATION OFFICE ADDRESS :", 15, 20, 0);
        //        }
        //        //horizontal line2 big
        //        cb.MoveTo(10, 776);
        //        cb.LineTo(660, 776);

        //        //horizontal line3 big
        //        //cb.MoveTo(10, 680);
        //        //cb.LineTo(660, 680);

        //        cb.MoveTo(10, 678);
        //        cb.LineTo(660, 678);

        //        //cb.MoveTo(10, 648);
        //        //cb.LineTo(660, 648);

        //        cb.MoveTo(10, 658);
        //        cb.LineTo(660, 658);

        //        //cb.MoveTo(10, 565);
        //        //cb.LineTo(660, 565);

        //        cb.MoveTo(10, 618);
        //        cb.LineTo(660, 618);

        //        //Center line6 Small
        //        cb.MoveTo(160, 658);
        //        cb.LineTo(160, 310);

        //        cb.MoveTo(210, 658);
        //        cb.LineTo(210, 310);

        //        cb.MoveTo(280, 658);
        //        cb.LineTo(280, 310);

        //        cb.MoveTo(340, 658);
        //        cb.LineTo(340, 310);

        //        cb.MoveTo(380, 658);
        //        cb.LineTo(380, 270);

        //        cb.MoveTo(440, 658);
        //        cb.LineTo(440, 270);

        //        cb.MoveTo(440, 640);
        //        cb.LineTo(660, 640);

        //        //center small
        //        cb.MoveTo(460, 640);
        //        cb.LineTo(460, 310);



        //        cb.MoveTo(510, 640);
        //        cb.LineTo(510, 290);
        //        //Muthu



        //        cb.MoveTo(535, 640);
        //        cb.LineTo(535, 310);

        //        cb.MoveTo(590, 640);
        //        cb.LineTo(590, 290);

        //        cb.MoveTo(615, 640);
        //        cb.LineTo(615, 310);

        //        cb.MoveTo(440, 290);
        //        cb.LineTo(660, 290);

        //        //horizontal line6 small
        //        cb.MoveTo(10, 310);
        //        cb.LineTo(660, 310);


        //        cb.MoveTo(10, 270);
        //        cb.LineTo(660, 270);

        //        cb.MoveTo(10, 210);
        //        cb.LineTo(660, 210);

        //        cb.MoveTo(10, 150);
        //        cb.LineTo(660, 150);


        //        cb.MoveTo(10, 70);
        //        cb.LineTo(660, 70);


        //        cb.MoveTo(10, 35);
        //        cb.LineTo(660, 35);


        //        cb.Stroke();
        //        //cb.BeginText();




        //        cb.EndText();

        //        writer.CloseStream = false;
        //        doc.Close();
        //        Response.Buffer = true;
        //        Response.ContentType = "application/pdf";
        //        //Response.AddHeader("content-disposition", "attachment;filename=Invoices.pdf");
        //        Response.Cache.SetCacheability(HttpCacheability.NoCache);
        //        //Response.Write(doc);
        //        Response.End();
        //    }
        //}


        private string[] SplitByLenght(string Values, int split)
        {

            List<string> list = new List<string>();
            int SplitTheLoop = Values.Length / split;
            for (int i = 0; i < SplitTheLoop; i++)
                list.Add(Values.Substring(i * split, split));
            if (SplitTheLoop * split != Values.Length)
                list.Add(Values.Substring(SplitTheLoop * split));

            return list.ToArray();
        }


        public DataTable GetCurrency(string CountryID)
        {
            string _Query = "select * from NVO_CurrencyMaster where CountryID = " + CountryID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetDDGCntrExisting(string Id)
        {
            string _Query = " select distinct CntrID,sum(Rate) as USDAmount , sum(Total) as Total," +
                            " (select top(1) CntrNo from NVO_Containers where ID = NVO_ImpBLDDGChargedtls.CntrID) as CntrNo, " +
                            " (select top(1) Size from NVO_tblCntrTypes where ID = NVO_ImpBLDDGChargedtls.CntrTypeID) as CntrTypes " +
                            " from NVO_ImpBLDDGChargedtls where DDGID = " + Id + "  group by CntrID,CntrTypeID";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable NewGetDisplayEstimateExisting(string ID, string CntrID)
        {
            string _Query = " select convert(varchar,LLimitDt, 103) as FDatev,convert(varchar,ULimitDt, 103) as TDatev,LLimit,ULimit,Rate,CurrID,Days,Amount,Total,NVO_ImpBLDDGCharges.ExRate,TotalAmount" +
                            " from NVO_ImpBLDDGChargedtls " +
                            " inner join NVO_ImpBLDDGCharges on NVO_ImpBLDDGCharges.ID=NVO_ImpBLDDGChargedtls.DDGID" +
                            " where DDGID = " + ID + " and NVO_ImpBLDDGChargedtls.CntrID=" + CntrID;
            return Manag.GetViewData(_Query, "");
        }

    }
}