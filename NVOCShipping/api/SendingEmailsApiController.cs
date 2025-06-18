using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using DataManager;
using DataTier;
using System.Net.Mail;
using System.Text;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;



namespace NVOCShipping.api
{

    public class SendingEmailsApiController : ApiController
    {

        // GET api/<controller>

        [ActionName("EmailsendingRateSheet")]
        public List<MySendingEmailAlert> EmailsendingRateSheet(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.RateSheetSendingEmail(Data);
            return st;
        }



        [ActionName("EmailsendingBooking")]
        public List<MySendingEmailAlert> EmailsendingBooking(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.BookingSendingEmail(Data);
            return st;
        }

        [ActionName("SendAutoEmails")]
        public List<MySendingEmailAlert> SendAutoEmails(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.EmailData(Data);
            return st;
        }

        [ActionName("SendAutoEmailsCANTo")]
        public List<MyImpCAN> SendAutoEmailsCANTo(MyImpCAN Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MyImpCAN> st = AccMange.EmailDataTo(Data);
            return st;
        }



        [ActionName("CROSendAutoEmails")]
        public List<MySendingEmailAlert> CROSendAutoEmails(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.CROSendingEmail(Data);
            return st;
        }

        [ActionName("CROSendAutoEmailsDepo")]
        public List<MySendingEmailAlert> CROSendAutoEmailsDepo(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.CROSendingEmailDepo(Data);
            return st;
        }

        [ActionName("EmailsendingCAN")]
        public List<MySendingEmailAlert> EmailsendingCAN(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.CANSendingEmail(Data);
            return st;
        }

        [ActionName("BookingSaveConfirmEmail")]
        public List<MySendingEmailAlert> BookingSaveConfirmEmail(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.EmailAlertBookingSaveConfirm(Data);
            return st;
        }

        [ActionName("EmailAlertBookingSaveConfirmPdf")]
        public List<MySendingEmailAlert> EmailAlertBookingSaveConfirmPdf(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.EmailAlertBookingSaveConfirmPdf(Data);
            return st;
        }


        [ActionName("SendAutoEmailsForCros")]
        public List<MySendingEmailAlert> SendAutoEmailsForCros(MySendingEmailAlert Data)
        {

            SendEmailManagerLocalValues AccMange = new SendEmailManagerLocalValues();
            List<MySendingEmailAlert> st = AccMange.GetSendAutoEmailsForCros(Data);
            return st;
        }
    }

    public class SendEmailManagerLocalValues
    {
        DocumentManager Manag = new DocumentManager();


        public List<MySendingEmailAlert> RateSheetSendingEmail(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            string strHTML = "";

            try
            {
                DataTable dtv = GetRRPDFValus(Data.ID.ToString());
                if (dtv.Rows.Count > 0)
                {

                    Document doc = new Document();
                    Rectangle rec = new Rectangle(670, 900);
                    doc = new Document(rec);
                    Paragraph para = new Paragraph();
                    MemoryStream memoryStream = new MemoryStream();
                    PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                    doc.Open();
                    // PdfWriter pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
                    //// PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
                    //// pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create));
                    // doc.Open();

                    #region Header LOGO COMPANY NAME
                    //-------------HEADER-------------------//

                    iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                    tbllogo.Width = 100;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    //tbllogo.Cellpadding = 1;
                    tbllogo.BorderWidth = 0;
                    Cell cell = new Cell();
                    //cell.Width = 10;

                    var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));

                    img.ScaleAbsolute(100f, 45f);
                    img.Alignment = Element.ALIGN_LEFT;
                    cell = new Cell(img);
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    cell.Width = 100;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tbllogo.AddCell(cell);

                    ///--SPACE--//
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    //cell.Colspan = 3;
                    tbllogo.AddCell(cell);
                    ///----/////
                    ///

                    DataTable dtc = GetCompnayDetails();
                    if (dtc.Rows.Count > 0)
                    {
                        cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);
                    }


                    //cell = new Cell(new Phrase("Global Network Lines Lines Pte.Ltd", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //tbllogo.Alignment = Element.ALIGN_LEFT;
                    //tbllogo.AddCell(cell);

                    cell = new Cell(new Phrase("RATE REQUEST ", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //cell.Colspan = 3;
                    tbllogo.AddCell(cell);


                    var Addresssa = Regex.Split(dtc.Rows[0]["CompanyAddress"].ToString(), "\r\n|\r|\n");
                    for (int a = 0; a < Addresssa.Length; a++)
                    {
                        cell = new Cell(new Phrase(Addresssa[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        cell.Colspan = 2;
                        tbllogo.AddCell(cell);
                    }

                    //cell = new Cell(new Phrase("33 UBI AVENUE 3, #08-68, VERTEX TOWER A, SINGAPORE 408868 ", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //tbllogo.Alignment = Element.ALIGN_LEFT;
                    //cell.Colspan = 2;
                    //tbllogo.AddCell(cell);

                    //cell = new Cell(new Phrase(DS.Tables[0].Rows[0]["Location"].ToString() + " - " + DS.Tables[0].Rows[0]["Pincode"].ToString() + " Tel # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Phone1"].ToString() + "   Fax # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Fax"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 6;
                    // tbllogo.AddCell(cell);

                    doc.Add(tbllogo);

                    para = new Paragraph("");
                    doc.Add(para);

                    para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                    para.Alignment = Element.ALIGN_RIGHT;
                    doc.Add(para);

                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                    doc.Add(Tblspace2);

                    //------------------------------------------------------------------------
                    #endregion

                    #region Booking Party and Ratesheet details
                    //-------------------Bookingparty and Ratesheet details-----------
                    PdfContentByte content = pdfWriter.DirectContent;
                    PdfPTable mtable = new PdfPTable(2);
                    mtable.WidthPercentage = 100;
                    mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                    PdfPTable Tbl1 = new PdfPTable(1);
                    Tbl1.WidthPercentage = 50;
                    PdfPCell cell1 = new PdfPCell(new Phrase("Booking Party", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell1.Colspan = 6;
                    cell1.HorizontalAlignment = 1;
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell1.BorderWidth = 0;
                    cell1.FixedHeight = 23f;
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.Colspan = 1;
                    Tbl1.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["Customer"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);

                    var Addresssb = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                    for (int a = 0; a < Addresssb.Length; a++)
                    {
                        cell1 = new PdfPCell(new Phrase(Addresssb[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.BorderWidth = 0;
                        Tbl1.AddCell(cell1);
                    }

                    // cell1 = new PdfPCell(new Phrase("MUMBAI CITY, MAHARASHTRA, 400092", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));

                    // cell1.BorderWidth = 0;
                    //Tbl1.AddCell(cell1);
                    mtable.AddCell(Tbl1);



                    Tbl1 = new PdfPTable(2);
                    Tbl1.WidthPercentage = 50;
                    Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                    cell1 = new PdfPCell(new Phrase("RATE REQUEST N0", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["RatesheetNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("SHIPMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ShipmentTypes"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("SALES", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["SalesPerson"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("VALID TILL", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ValidDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase("STATUS", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["Status"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 1;
                    cell1.FixedHeight = 25f;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tbl1.AddCell(cell1);

                    mtable.AddCell(Tbl1);
                    doc.Add(mtable);

                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspaces2 = new iTextSharp.text.Table(1);
                    doc.Add(Tblspaces2);

                    //------------------------------------------------------------------------
                    #endregion

                    #region Location POL POD POO
                    // /----------------------- LocTable-----------------------///

                    PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                    TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                    TblLocs.WidthPercentage = 100;

                    cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POOLs"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblLocs.AddCell(cell1);

                    doc.Add(TblLocs);



                    #endregion

                    #region Container And Free DayDetails
                    //----------------- Container And Free DayDetails--------------//

                    iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                    Tbl3.Width = 100;
                    Tbl3.Alignment = Element.ALIGN_LEFT;
                    Tbl3.Cellpadding = 0;
                    Tbl3.BorderWidth = 0;

                    //Sub Heading
                    cell = new Cell(new Phrase("Container And Free Days Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    Tbl3.AddCell(cell);
                    doc.Add(Tbl3);

                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspaces1 = new iTextSharp.text.Table(1);
                    doc.Add(Tblspaces1);

                    //------------------------------------------------------------------------

                    PdfPTable TblCntrDtls = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                    TblCntrDtls.HorizontalAlignment = Element.ALIGN_LEFT;
                    TblCntrDtls.WidthPercentage = 100;

                    cell1 = new PdfPCell(new Phrase("Container Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Cargo", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase("DG Class", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("VGM", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    DataTable _dtCnt = GetRRPDFCntrTypesValus(Data.ID.ToString());
                    for (int i = 0; i < _dtCnt.Rows.Count; i++)
                    {
                        cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["CntrTypes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrDtls.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrDtls.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["cargo"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrDtls.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["DGClass"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrDtls.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["VGM"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrDtls.AddCell(cell1);

                    }

                    doc.Add(TblCntrDtls);


                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspaces = new iTextSharp.text.Table(1);
                    doc.Add(Tblspaces);

                    //------------------------------------------------------------------------


                    PdfPTable Tblfreedtls = new PdfPTable(new float[] { 2, 2, 2 });
                    Tblfreedtls.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tblfreedtls.WidthPercentage = 100;

                    cell1 = new PdfPCell(new Phrase("Freedays Mode", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblfreedtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Export Freedays", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblfreedtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("Import Freedays", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblfreedtls.AddCell(cell1);

                    DataTable _dtFreeday = GetRRFreedaysDtls(Data.ID.ToString());
                    for (int i = 0; i < _dtFreeday.Rows.Count; i++)
                    {
                        cell1 = new PdfPCell(new Phrase(_dtFreeday.Rows[i]["Mode"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tblfreedtls.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtFreeday.Rows[i]["ExpFreeDays"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tblfreedtls.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtFreeday.Rows[i]["ImpFreeDays"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tblfreedtls.AddCell(cell1);


                    }

                    doc.Add(Tblfreedtls);
                    // }
                    #endregion





                    #region Special Instruction and OTHER  REMARKS

                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspaces8 = new iTextSharp.text.Table(1);
                    doc.Add(Tblspaces8);

                    //----------------------------------------------------------------------

                    //-------------------Bookingparty and Ratesheet details-----------

                    PdfPTable mtable1 = new PdfPTable(2);
                    mtable1.WidthPercentage = 100;
                    mtable1.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                    PdfPTable TblSplInsOthers = new PdfPTable(1);
                    TblSplInsOthers.WidthPercentage = 50;
                    cell1 = new PdfPCell(new Phrase("Special Instruction – Slot Details & Services :", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    cell1.Colspan = 1;
                    TblSplInsOthers.AddCell(cell1);



                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["Remarks"].ToString() + " \n \n\n\n\n\n", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblSplInsOthers.AddCell(cell1);
                    mtable1.AddCell(TblSplInsOthers);



                    TblSplInsOthers = new PdfPTable(1);
                    TblSplInsOthers.WidthPercentage = 50;
                    TblSplInsOthers.HorizontalAlignment = Element.ALIGN_RIGHT;

                    cell1 = new PdfPCell(new Phrase("Other Remarks :", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    cell1.Colspan = 1;
                    TblSplInsOthers.AddCell(cell1);


                    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["OtherRemarks"].ToString() + " \n\n\n\n\n\n", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblSplInsOthers.AddCell(cell1);
                    mtable1.AddCell(TblSplInsOthers);
                    doc.Add(mtable1);


                    #endregion

                    #region Terms & Condition


                    iTextSharp.text.Table Tbl11 = new iTextSharp.text.Table(1);
                    Tbl11.Width = 100;
                    Tbl11.Alignment = Element.ALIGN_LEFT;
                    Tbl11.Cellpadding = 1;
                    Tbl11.BorderWidth = 0;


                    cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    Tbl11.AddCell(cell);



                    cell = new Cell(new Phrase(" * Rate details as per annexure sheet ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl11.AddCell(cell);

                    cell = new Cell(new Phrase(" * Rate was agreed & accepted to respective market price ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl11.AddCell(cell);

                    cell = new Cell(new Phrase(" * Tariff break up attached below in annexure sheet ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl11.AddCell(cell);

                    cell = new Cell(new Phrase(" \n\n\n\n\n\n\n", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl11.AddCell(cell);


                    doc.Add(Tbl11);
                    #endregion



                    #region second page

                    //#region logo
                    ////----------------------------DYNAMIC-LETTER-HEAD-ADDRESS---------------------------
                    //iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(8);
                    //tbllogo1.Width = 100;
                    //tbllogo1.Alignment = Element.ALIGN_LEFT;
                    //tbllogo1.Cellpadding = 1;
                    //tbllogo1.BorderWidth = 0;

                    //var img1 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/OCEANUS LOGO.png"));
                    ////var img1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/OCEANUS LOGO.png"));
                    //img1.Alignment = Element.ALIGN_CENTER;
                    //cell = new Cell(img1);
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 2;
                    //cell.Rowspan = 3;
                    //cell.Width = 20;
                    //tbllogo1.AddCell(cell);


                    //cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //tbllogo1.Alignment = Element.ALIGN_LEFT;
                    ////cell.Colspan = 2;
                    //tbllogo1.AddCell(cell);

                    //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    ////cell.Colspan = 3;
                    //tbllogo1.AddCell(cell);


                    //var Addresssaa = Regex.Split(dtc.Rows[0]["CompanyAddress"].ToString(), "\r\n|\r|\n");
                    //for (int a = 0; a < Addresssaa.Length; a++)
                    //{
                    //    cell = new Cell(new Phrase(Addresssaa[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    //    cell.BorderWidth = 0;
                    //    tbllogo1.AddCell(cell);

                    //}
                    //doc.Add(tbllogo1);

                    //iTextSharp.text.Table Tblline1 = new iTextSharp.text.Table(1);
                    //Tblline1.Width = 100;
                    //Tblline1.DefaultCellBorder = Rectangle.NO_BORDER;
                    //Tblline1.Border = Rectangle.NO_BORDER;
                    //Tblline1.Cellpadding = 1;
                    //Tblline1.BorderWidth = 0.5F;

                    //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                    //cell.BorderWidthTop = 0;
                    //cell.BorderWidthRight = 0;
                    //cell.BorderWidthLeft = 0;
                    //cell.BorderWidthBottom = 0;
                    //cell.BackgroundColor = new Color(98, 141, 214);
                    //Tblline1.AddCell(cell);
                    //doc.Add(Tblline1);

                    #endregion

                    #region Header LOGO COMPANY NAME
                    //-------------HEADER-------------------//

                    iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(2);
                    tbllogo1.Width = 100;
                    tbllogo1.Alignment = Element.ALIGN_LEFT;
                    //tbllogo.Cellpadding = 1;
                    tbllogo1.BorderWidth = 0;
                    Cell cell2 = new Cell();
                    //cell.Width = 10;

                    var img1 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/gnlLOGO.png"));
                    img1.ScaleAbsolute(100f, 45f);
                    img1.Alignment = Element.ALIGN_LEFT;
                    cell2 = new Cell(img);
                    cell2.BorderWidth = 0;
                    cell2.Colspan = 1;
                    cell2.Width = 100;
                    cell2.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tbllogo1.AddCell(cell2);

                    ///--SPACE--//
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo1.Alignment = Element.ALIGN_LEFT;
                    //cell.Colspan = 3;
                    tbllogo1.AddCell(cell);
                    ///----/////


                    cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));


                    cell.BorderWidth = 0;
                    tbllogo1.Alignment = Element.ALIGN_LEFT;
                    //cell.Colspan = 2;
                    tbllogo1.AddCell(cell);

                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    //cell.Colspan = 3;
                    tbllogo1.AddCell(cell);


                    var Addresssaa = Regex.Split(dtc.Rows[0]["CompanyAddress"].ToString(), "\r\n|\r|\n");
                    for (int a = 0; a < Addresssaa.Length; a++)
                    {
                        cell = new Cell(new Phrase(Addresssaa[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo1.AddCell(cell);
                    }

                    //cell = new Cell(new Phrase(DS.Tables[0].Rows[0]["Location"].ToString() + " - " + DS.Tables[0].Rows[0]["Pincode"].ToString() + " Tel # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Phone1"].ToString() + "   Fax # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Fax"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 6;
                    // tbllogo.AddCell(cell);

                    doc.Add(tbllogo1);

                    para = new Paragraph("");
                    doc.Add(para);

                    para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                    para.Alignment = Element.ALIGN_RIGHT;
                    doc.Add(para);



                    //------------------------------------------------------------------------
                    #endregion


                    #region header

                    iTextSharp.text.Table tblHead = new iTextSharp.text.Table(1);
                    tblHead.Width = 100;
                    tblHead.Alignment = Element.ALIGN_LEFT;
                    tblHead.Cellpadding = 1;
                    tblHead.BorderWidth = 0;

                    cell = new Cell(new Phrase("RATE REQUEST- TARIFF ANNEXURE", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    //cell.Colspan = 3;
                    tblHead.AddCell(cell);
                    doc.Add(tblHead);

                    iTextSharp.text.Table tblHead1 = new iTextSharp.text.Table(4);
                    tblHead1.Width = 100;
                    tblHead1.Alignment = Element.ALIGN_LEFT;
                    tblHead1.BorderWidth = 0;

                    cell = new Cell(new Phrase("RATE REQUEST NO :", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.BorderWidth = 0;
                    tblHead1.AddCell(cell);

                    cell = new Cell(new Phrase(dtv.Rows[0]["RatesheetNo"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.BorderWidth = 0;
                    tblHead1.AddCell(cell);

                    cell = new Cell(new Phrase("STATUS  :", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell.BorderWidth = 0;
                    tblHead1.AddCell(cell);

                    cell = new Cell(new Phrase(dtv.Rows[0]["Status"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.BorderWidth = 0;
                    tblHead1.AddCell(cell);
                    doc.Add(tblHead1);

                    #endregion

                    #region Frieght Charges


                    iTextSharp.text.Table TblFrieght = new iTextSharp.text.Table(1);
                    TblFrieght.Width = 100;
                    TblFrieght.Alignment = Element.ALIGN_LEFT;
                    TblFrieght.Cellpadding = 0;
                    TblFrieght.BorderWidth = 0;

                    cell = new Cell(new Phrase("FREIGHT CHARGES", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    TblFrieght.AddCell(cell);
                    doc.Add(TblFrieght);

                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspace3 = new iTextSharp.text.Table(1);
                    doc.Add(Tblspace3);

                    //------------------------------------------------------------------------;


                    PdfPTable TblFrieght1 = new PdfPTable(new float[] { 1, 1, 2, 1, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f });
                    TblFrieght1.HorizontalAlignment = Element.ALIGN_LEFT;
                    TblFrieght1.WidthPercentage = 100;

                    cell1 = new PdfPCell(new Phrase("S.NO", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("EQUIP TYPE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("CHARGES", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("CUR", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("COLLECTION MODE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("REQUESTED RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("MANIFEST RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("CUSTOMER RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("RATE DIFFERENCE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblFrieght1.AddCell(cell1);

                    DataTable _dtfrt = GetRRTariffCharges(Data.ID.ToString(), "135");
                    for (int i = 0; i < _dtfrt.Rows.Count; i++)
                    {

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["SNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["CntrSize"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["ChgDesc"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["CurrencyCode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["PaymentMode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["ReqRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["ManifRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["CustomerRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtfrt.Rows[i]["RateDiff"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblFrieght1.AddCell(cell1);

                    }

                    doc.Add(TblFrieght1);

                    #endregion

                    #region THC Charges

                    iTextSharp.text.Table TblTHC = new iTextSharp.text.Table(1);
                    TblTHC.Width = 100;
                    TblTHC.Alignment = Element.ALIGN_LEFT;
                    TblTHC.Cellpadding = 0;
                    TblTHC.BorderWidth = 0;

                    cell = new Cell(new Phrase("TERMINAL HANDLING CHARGES ", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    TblTHC.AddCell(cell);
                    doc.Add(TblTHC);

                    //----------SPACE----------------------------------
                    iTextSharp.text.Table Tblspace4 = new iTextSharp.text.Table(1);
                    doc.Add(Tblspace4);

                    //------------------------------------------------------------------------;

                    PdfPTable TblTHC1 = new PdfPTable(new float[] { 1, 1, 2, 1, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f });
                    TblTHC1.HorizontalAlignment = Element.ALIGN_LEFT;
                    TblTHC1.WidthPercentage = 100;

                    cell1 = new PdfPCell(new Phrase("S.NO", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("EQUIP TYPE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("CHARGES", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("CUR", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("COLLECTION MODE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("REQUESTED RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("MANIFEST RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("CUSTOMER RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("RATE DIFFERENCE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    TblTHC1.AddCell(cell1);

                    DataTable _dtThc = GetRRTariffCharges(Data.ID.ToString(), "136");
                    for (int i = 0; i < _dtThc.Rows.Count; i++)
                    {

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["SNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["CntrSize"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["ChgDesc"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["CurrencyCode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["PaymentMode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["ReqRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["ManifRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["CustomerRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtThc.Rows[i]["RateDiff"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblTHC1.AddCell(cell1);

                    }

                    doc.Add(TblTHC1);


                    #endregion

                    #region HAULAGE Charges
                    DataTable _dtHAUL = GetRRTariffCharges(Data.ID.ToString(), "137");

                    if (_dtHAUL.Rows.Count == 0)
                    {
                        iTextSharp.text.Table TblHaulNil = new iTextSharp.text.Table(1);
                        TblHaulNil.Width = 100;
                        TblHaulNil.Alignment = Element.ALIGN_LEFT;
                        TblHaulNil.Cellpadding = 0;
                        TblHaulNil.BorderWidth = 0;

                        cell = new Cell(new Phrase("HAULAGE CHARGES - NIL", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        TblHaulNil.AddCell(cell);
                        doc.Add(TblHaulNil);

                    }
                    else
                    {
                        iTextSharp.text.Table TblHaul = new iTextSharp.text.Table(1);
                        TblHaul.Width = 100;
                        TblHaul.Alignment = Element.ALIGN_LEFT;
                        TblHaul.Cellpadding = 0;
                        TblHaul.BorderWidth = 0;

                        cell = new Cell(new Phrase("HAULAGE CHARGES", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        TblHaul.AddCell(cell);
                        doc.Add(TblHaul);

                        //----------SPACE----------------------------------
                        iTextSharp.text.Table Tblspace5 = new iTextSharp.text.Table(1);
                        doc.Add(Tblspace5);

                        //------------------------------------------------------------------------;

                        PdfPTable TblHaul1 = new PdfPTable(new float[] { 1, 1, 2, 1, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f });
                        TblHaul1.HorizontalAlignment = Element.ALIGN_LEFT;
                        TblHaul1.WidthPercentage = 100;

                        cell1 = new PdfPCell(new Phrase("S.NO", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("EQUIP TYPE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CHARGES", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CUR", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("COLLECTION MODE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("REQUESTED RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("MANIFEST RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CUSTOMER RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("RATE DIFFERENCE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblHaul1.AddCell(cell1);


                        for (int i = 0; i < _dtHAUL.Rows.Count; i++)
                        {

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["SNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["CntrSize"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["ChgDesc"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["CurrencyCode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["PaymentMode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["ReqRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["ManifRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["CustomerRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtHAUL.Rows[i]["RateDiff"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblHaul1.AddCell(cell1);

                        }

                        doc.Add(TblHaul1);

                    }
                    #endregion

                    #region Local Charges Orgin

                    DataTable _dtLCO = GetRRLocalTariffCharges(Data.ID.ToString(), "138", "18");
                    if (_dtLCO.Rows.Count == 0)
                    {
                        iTextSharp.text.Table TblLCONil = new iTextSharp.text.Table(1);
                        TblLCONil.Width = 100;
                        TblLCONil.Alignment = Element.ALIGN_LEFT;
                        TblLCONil.Cellpadding = 0;
                        TblLCONil.BorderWidth = 0;

                        cell = new Cell(new Phrase("LOCAL CHARGES (ORIGIN) - NIL", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        TblLCONil.AddCell(cell);
                        doc.Add(TblLCONil);
                    }
                    else
                    {
                        iTextSharp.text.Table TblLCO = new iTextSharp.text.Table(1);
                        TblLCO.Width = 100;
                        TblLCO.Alignment = Element.ALIGN_LEFT;
                        TblLCO.Cellpadding = 0;
                        TblLCO.BorderWidth = 0;

                        cell = new Cell(new Phrase("LOCAL CHARGES - ORIGIN", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        TblLCO.AddCell(cell);
                        doc.Add(TblLCO);

                        //----------SPACE----------------------------------
                        iTextSharp.text.Table Tblspace6 = new iTextSharp.text.Table(1);
                        doc.Add(Tblspace6);

                        //------------------------------------------------------------------------;

                        PdfPTable TblLCO1 = new PdfPTable(new float[] { 1, 1, 2, 1, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f });
                        TblLCO1.HorizontalAlignment = Element.ALIGN_LEFT;
                        TblLCO1.WidthPercentage = 100;

                        cell1 = new PdfPCell(new Phrase("S.NO", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("EQUIP TYPE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CHARGES", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CUR", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("COLLECTION MODE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("REQUESTED RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("MANIFEST RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CUSTOMER RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("RATE DIFFERENCE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCO1.AddCell(cell1);



                        for (int i = 0; i < _dtLCO.Rows.Count; i++)
                        {

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["SNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["CntrSize"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["ChgDesc"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["CurrencyCode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["PaymentMode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["ReqRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["ManifRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["CustomerRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCO.Rows[i]["RateDiff"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCO1.AddCell(cell1);

                        }

                        doc.Add(TblLCO1);

                    }
                    //}


                    #endregion

                    #region Local Charges Destination

                    DataTable _dtLCDest = GetRRLocalTariffCharges(Data.ID.ToString(), "138", "19");
                    if (_dtLCDest.Rows.Count == 0)
                    {
                        iTextSharp.text.Table TblLCDestNil = new iTextSharp.text.Table(1);
                        TblLCDestNil.Width = 100;
                        TblLCDestNil.Alignment = Element.ALIGN_LEFT;
                        TblLCDestNil.Cellpadding = 0;
                        TblLCDestNil.BorderWidth = 0;

                        cell = new Cell(new Phrase("LOCAL CHARGES (DESTINATION) - NIL", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        TblLCDestNil.AddCell(cell);
                        doc.Add(TblLCDestNil);
                    }
                    else
                    {
                        iTextSharp.text.Table TblLCDest = new iTextSharp.text.Table(1);
                        TblLCDest.Width = 100;
                        TblLCDest.Alignment = Element.ALIGN_LEFT;
                        TblLCDest.Cellpadding = 0;
                        TblLCDest.BorderWidth = 0;

                        cell = new Cell(new Phrase("LOCAL CHARGES - DESTINATION", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        TblLCDest.AddCell(cell);
                        doc.Add(TblLCDest);

                        //----------SPACE----------------------------------
                        iTextSharp.text.Table Tblspace6 = new iTextSharp.text.Table(1);
                        doc.Add(Tblspace6);

                        //------------------------------------------------------------------------;

                        PdfPTable TblLCDest1 = new PdfPTable(new float[] { 1, 1, 2, 1, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f });
                        TblLCDest1.HorizontalAlignment = Element.ALIGN_LEFT;
                        TblLCDest1.WidthPercentage = 100;

                        cell1 = new PdfPCell(new Phrase("S.NO", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("EQUIP TYPE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CHARGES", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CUR", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("COLLECTION MODE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("REQUESTED RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("MANIFEST RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("CUSTOMER RATE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("RATE DIFFERENCE", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        TblLCDest1.AddCell(cell1);



                        for (int i = 0; i < _dtLCDest.Rows.Count; i++)
                        {

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["SNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["CntrSize"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["ChgDesc"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["CurrencyCode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["PaymentMode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["ReqRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["ManifRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["CustomerRate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtLCDest.Rows[i]["RateDiff"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                            TblLCDest1.AddCell(cell1);

                        }

                        doc.Add(TblLCDest1);

                    }
                    //}


                    #endregion

                    //#endregion



                    #region FOOTER
                    ///---------FOOTER----------------//
                    iTextSharp.text.Table Tbl12 = new iTextSharp.text.Table(4);
                    Tbl12.Width = 100;
                    Tbl12.Alignment = Element.ALIGN_LEFT;
                    Tbl12.Cellpadding = 1;
                    Tbl12.BorderWidth = 1;

                    cell = new Cell(new Phrase("Created By : " + dtv.Rows[0]["CreatedBy"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tbl12.AddCell(cell);

                    cell = new Cell(new Phrase("Created On :  " + dtv.Rows[0]["CreatedOn"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tbl12.AddCell(cell);

                    cell = new Cell(new Phrase("Approved By :  " + dtv.Rows[0]["ApprovedBy"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tbl12.AddCell(cell);

                    cell = new Cell(new Phrase("Approved On : " + dtv.Rows[0]["DtAppr"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tbl12.AddCell(cell);
                    doc.Add(Tbl12);
                    #endregion


                    pdfWriter.CloseStream = false;
                    doc.Close();

                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();
                    DataTable dtE = GetEmailsedingRRView(Data.ID.ToString());
                    if (dtE.Rows.Count > 0)
                    {
                        string HTMLstr = "";

                        HTMLstr = "<table border='1' cellpadding='0' cellspacing='0' width='50%' style='font-family:Arial; border: 1px solid #2196f3;'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td style='background-color:#2196f3;color:#fff;border-right:0px solid #303297;border-left:0px solid #303297;border-top:0px solid #303297;border-bottom:0px solid #303297'> " +
                                  " <table cellpadding='0' cellspacing='0' width='100%'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td style='font-family:Arial;font-size:14px;font-weight:bold;text-align:left;color:#fff;padding-top:2px;padding-left:3px;padding-bottom:2px;padding-right:3px;'>RR NUMBER : <a href='https:/RRPDF/" + dtE.Rows[0]["RatesheetNo"].ToString() + ".pdf' style='color:#fff;'>" + dtE.Rows[0]["RatesheetNo"].ToString() + "</a></td> " +
                                  " <td style='font-family:Arial;font-size:14px;font-weight:bold;padding-top:2px;padding-left:3px;padding-bottom:2px;padding-right:3px;text-align:right;'> STATUS : " + dtE.Rows[0]["Status"].ToString() + "</td> " +
                                  " </tr> " +
                                  " </tbody> " +
                                  " </table> " +
                                  " </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td style='border:none;'> " +
                                  " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px; font-size:14px;'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td colspan='2'> " +
                                  " <p style='margin-top: 10px;margin-bottom:0px;margin-left:10px;padding-bottom:0;'>Hi,</p><p style='padding-top:0;margin-top:0;'>This is the system generated message to notify the rate request details</p> " +
                                  " </td> " +
                                  " </tr> " +
                                  " </tbody> " +
                                  " </table> " +
                                  " </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td style ='border:none;'> " +
                                  " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Rate Request Details</td> " +
                                  " </tr> " +
                                  " </tbody> " +
                                  " </table> " +
                                  " </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td style='border:none;'> " +
                                  " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px; padding-bottom:8px; font-size:14px;'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td colspan='2'> " +
                                  " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'> Customer </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Customer"].ToString() + "</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'> Loading </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'> " + dtE.Rows[0]["Loading"].ToString() + " </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'>T/S Port</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["TSPort"].ToString() + "</td> " +
                                  " </tr> " +
                                   //" <tr> " +
                                   // " <td colspan='2' style='font-weight:bold;'>Detention Freedays</td> " +
                                   //" </tr> " +
                                   //" <tr> " +
                                   //" <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Freedays"].ToString() + " Days</td> " +
                                   //" </tr> " +
                                   // " <tr> " +
                                   // " <td colspan='2' style='font-weight:bold;'>Demmurage Freedays</td> " +
                                   //" </tr> " +
                                   //" <tr> " +
                                   //" <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Freedays"].ToString() + " Days</td> " +
                                   //" </tr> " +

                                   " <tr> " +
                                   " <td colspan='2' style='font-weight:bold;'>Collection Mode</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["PaymentMode"].ToString() + "</td> " +
                                  " </tr>" +
                                   " <tr> " +
                                   " <td colspan='2' style='font-weight:bold;'>Cargo</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Cargo"].ToString() + "</td> " +
                                  " </tr>" +

                                  "<tr>" +
                                   " <td colspan='2' style='font-weight:bold;'>" + dtE.Rows[0]["Status"].ToString() + " REASON:</td> " +
                                  " </tr>";

                        if (dtE.Rows[0]["RSStatus"].ToString() == "2")
                        {
                            HTMLstr += "<tr>" +
                                      "<td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["RJRemarks"].ToString() + "</td> " +
                                      " </tr>";
                        }
                        else if (dtE.Rows[0]["RSStatus"].ToString() == "3")
                        {
                            HTMLstr += "<tr>" +
                                      "<td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["RERemarks"].ToString() + "</td> " +
                                      " </tr>";
                        }
                        else if (dtE.Rows[0]["RSStatus"].ToString() == "4")
                        {
                            HTMLstr += "<tr>" +
                                      "<td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["APRemarks"].ToString() + "</td> " +
                                      " </tr>";
                        }

                        HTMLstr += "</tbody> " +
                                  " </table> " +
                                  " </td> " +
                                  " <td colspan='2'> " +
                                  " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'> Sales Person </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["SalesPerson"].ToString() + "</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'> Discharge </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'> " + dtE.Rows[0]["Discharge"].ToString() + " </td> " +
                                  " </tr> " +

                                   " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'> Final POD </td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'> " + dtE.Rows[0]["FPOD"].ToString() + " </td> " +
                                  " </tr> " +

                                  " <tr> " +
                                  " <td colspan='2' style='font-weight:bold;'>Commodity</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>GENERAL CARGO</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                   " <td colspan='2' style='font-weight:bold;'>Validity Till</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["ValidDate"].ToString() + "</td> " +
                                  " </tr> " +
                                   " <tr> " +
                                   " <td colspan='2' style='font-weight:bold;'>Remarks</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                  " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Remarks"].ToString() + "</td> " +
                                  " </tr> " +
                                  " </tbody> " +
                                  " </table> " +
                                  " </td> " +
                                  " </tr> " +
                                  " </tbody> " +
                                  " </table> " +
                                  " </td> " +
                                  " </tr> " +

                                   " <tr> " +
                                  " <td style ='border:none;'> " +
                                  " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                                  " <tbody> " +
                                  " <tr> " +
                                  " <td colspan='3' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Free days</td> " +
                                  " </tr> " +
                                  " <tr> " +
                                   " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Mode</td>" +
                                   " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Export Freedays</td>" +
                                   " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Import Freedays</td>" +
                                  " </tr>";

                        DataTable _dtFreeday1 = GetRRFreedaysDtls(Data.ID.ToString());
                        for (int i = 0; i < _dtFreeday1.Rows.Count; i++)
                        {

                            HTMLstr += " <tr> " +
                                   " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtFreeday1.Rows[i]["Mode"].ToString() + "</td>" +
                                   " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtFreeday1.Rows[i]["ExpFreeDays"].ToString() + "</td>" +
                                   " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtFreeday1.Rows[i]["ImpFreeDays"].ToString() + "</td>" +
                                   " </tr>";
                        }

                        HTMLstr += " </tbody> " +
                           " </table> " +
                           " </td> " +
                           " </tr> " +
                           " <tr> " +
                           " <td style ='border:none;'> " +
                           " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                           " <tbody> " +
                           " <tr> " +
                           " <td colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Equipment Type & Rate</td> " +
                           " </tr> " +
                           " </tbody> " +
                           " </table> " +
                           " </td> " +
                           " </tr> " +
                           " <tr> " +
                           " <td> " +
                           " <table border='1' cellpadding='0' cellspacing='0'> " +
                           " <tr> " +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Size  Type</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Commodity</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>FRT</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>BAF</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>DGE</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>ECRS</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>CAF</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>EWRS</td>" +
                           " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>LSS</td>" +
                           " </tr>";
                        DataTable _dtR = GetRRNotificationChargewiseRates(Data.ID.ToString());
                        for (int x = 0; x < _dtR.Rows.Count; x++)
                        {
                            HTMLstr += "<tr>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["Size"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["Commodity"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["FRT"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["BAF"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["DGS"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["ECRS"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["CAF"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["EWRS"].ToString() + "</td>" +
                                " <td style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dtR.Rows[x]["LSS"].ToString() + "</td>" +
                            " </tr>";
                        }

                        DataTable _dthC = GetRateSheetNotificationTHCandIHCValues(Data.ID.ToString());
                        HTMLstr += " <tr> " +
                                    "<td  style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Size  Type</td> " +
                                    "<td  colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>LOADING PORT THC</td> " +
                                    "<td  colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>DEST THC</td> " +
                                    "<td  colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>EXPORT IHC</td> " +
                                    "<td colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>IMPORT IHC</td> " +
                                    "</tr>";

                        for (int y = 0; y < _dthC.Rows.Count; y++)
                        {
                            HTMLstr += " <tr> " +
                                   "<td  style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dthC.Rows[y]["Size"].ToString() + "</td> " +
                                   "<td  colspan='2' style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dthC.Rows[y]["LoadTHCRate"].ToString() + "</td> " +
                                   "<td  colspan='2' style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dthC.Rows[y]["DestTHCRate"].ToString() + "</td> " +
                                   "<td  colspan='2' style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dthC.Rows[y]["ExportIHCRate"].ToString() + "</td> " +
                                   "<td colspan='2' style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>" + _dthC.Rows[y]["ImportIHCRate"].ToString() + "</td> " +
                                   "</tr>";
                        }

                        HTMLstr += "  <tr> " +
                                    " <td colspan='5' style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Created By: " + dtE.Rows[0]["CreatedBy"].ToString() + "</td> " +
                                    " <td colspan='5' style='padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'>Created on: " + dtE.Rows[0]["CreateDate"].ToString() + "</td> " +
                                    " </tr>";

                        HTMLstr += "</table>" +
                                   " </td> " +
                                   " </tr> " +
                                   " </tbody> " +
                                   " </table>";

                        DataTable _dtCom = GetCompnayDetails();
                        MailMessage EmailObject = new MailMessage();
                        EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                        //DataTable dtAuto = GetCustomerEmailsending(Data.AgentID);
                        DataTable dtAuto = GetCustomerEmailsending(dtE.Rows[0]["AgentId"].ToString());
                        if (dtAuto.Rows.Count > 0)
                        {
                            var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                            for (int y = 0; y < EmailID.Length; y++)
                            {
                                if (EmailID[y].ToString() != "")
                                {
                                    EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                                }
                            }
                            EmailObject.To.Add(new MailAddress(dtE.Rows[0]["UserEmailID"].ToString()));
                            EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"));
                            EmailObject.Body = HTMLstr;
                            EmailObject.IsBodyHtml = true;
                            EmailObject.Priority = MailPriority.Normal;
                            EmailObject.Subject = "Rate Request: " + dtE.Rows[0]["RatesheetNo"].ToString() + " - " + dtE.Rows[0]["Status"].ToString();
                            EmailObject.Priority = MailPriority.Normal;
                            SmtpClient SMTPServer = new SmtpClient();
                            SMTPServer.UseDefaultCredentials = true;
                            //SMTPServer.Credentials = new NetworkCredential("noreply@blue-wave.app", "Automail@123");
                            SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());
                            SMTPServer.Host = "smtp.office365.com";
                            // SMTPServer.Timeout = 2000000;
                            SMTPServer.ServicePoint.MaxIdleTime = 1;
                            SMTPServer.Port = 587;
                            SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                            SMTPServer.EnableSsl = true;
                            SMTPServer.Send(EmailObject);

                            ViewList.Add(new MySendingEmailAlert
                            {
                                AlertMessage = "Email sent successfully"

                            });
                        }
                        else
                        {
                            ViewList.Add(new MySendingEmailAlert
                            {
                                AlertMessage = "email id not a configuration agancy master  kindly add and  retry"

                            });
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ViewList.Add(new MySendingEmailAlert
                {
                    AlertMessage = ex.Message.ToString()

                });
            }


            return ViewList;
        }
        public List<MySendingEmailAlert> BookingSendingEmail(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            string strHTML = "";

            MemoryStream memoryStream = new MemoryStream();
            DataTable dtv = GetBkgPDFValus(Data.BkgID);
            //if (dtv.Rows.Count > 0)
            //{
            //    Document doc = new Document();
            //    Rectangle rec = new Rectangle(670, 900);
            //    doc = new Document(rec);
            //    Paragraph para = new Paragraph();


            //    PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
            //    doc.Open();

            //    #region Header LOGO COMPANY NAME
            //    //-------------HEADER-------------------//

            //    iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
            //    tbllogo.Width = 100;
            //    //tbllogo.Alignment = Element.ALIGN_LEFT;
            //    //tbllogo.Cellpadding = 1;
            //    tbllogo.BorderWidth = 0;
            //    Cell cell = new Cell();
            //    cell.Width = 10;


            //    //var img = iTextSharp.text.Image.GetInstance(Path.Combine(HttpListenerContext.Current.Server.MapPath( "~/assets/img/logo.png"));

            //    DataTable dtc = GetAgencyDetailssending(Data.AgentID);
            //    if(dtc.Rows.Count >0)
            //    {
            //        var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
            //        img.Alignment = Element.ALIGN_LEFT;
            //        img.ScaleAbsolute(45f, 45f);
            //        cell = new Cell(img);
            //        cell.BorderWidth = 0;
            //        cell.Colspan = 1;
            //        cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            //        tbllogo.AddCell(cell);
            //    }


            //    ///--SPACE--//
            //    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    tbllogo.Alignment = Element.ALIGN_LEFT;
            //    //cell.Colspan = 3;
            //    tbllogo.AddCell(cell);

            //    ///--SPACE--//

            //    cell = new Cell(new Phrase(dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    tbllogo.Alignment = Element.ALIGN_LEFT;
            //    //cell.Colspan = 2;
            //    tbllogo.AddCell(cell);

            //    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    tbllogo.Alignment = Element.ALIGN_LEFT;
            //    //cell.Colspan = 3;
            //    tbllogo.AddCell(cell);
            //    ///----/////
            //    cell = new Cell(new Phrase("Agent of " + dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    tbllogo.Alignment = Element.ALIGN_LEFT;
            //    //cell.Colspan = 2;
            //    tbllogo.AddCell(cell);

            //    cell = new Cell(new Phrase("Booking Confirmation", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            //    //cell.Colspan = 3;
            //    tbllogo.AddCell(cell);

            //    var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString() + "\r\n" + "GST: " + dtc.Rows[0]["TaxGSTNo"].ToString(), "\r\n|\r|\n");
            //    for (int a = 0; a < LogoAddresss.Length; a++)
            //    {
            //        cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
            //        cell.BorderWidth = 0;
            //        tbllogo.Alignment = Element.ALIGN_LEFT;
            //        tbllogo.AddCell(cell);
            //    }

            //    doc.Add(tbllogo);

            //    para = new Paragraph("");
            //    doc.Add(para);

            //    para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
            //    para.Alignment = Element.ALIGN_RIGHT;
            //    doc.Add(para);

            //    //----------SPACE----------------------------------
            //    iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
            //    doc.Add(Tblspace2);

            //    //------------------------------------------------------------------------

            //    #endregion

            //    #region Bookingparty and Ratesheet details
            //    //-------------------Bookingparty and Ratesheet details-----------
            //    PdfContentByte content = pdfWriter.DirectContent;
            //    PdfPTable mtable = new PdfPTable(2);
            //    mtable.WidthPercentage = 100;
            //    mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


            //    PdfPTable Tbl1 = new PdfPTable(1);
            //    Tbl1.WidthPercentage = 50;
            //    PdfPCell cell1 = new PdfPCell(new Phrase("Booking Party", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
            //    cell1.Colspan = 6;
            //    cell1.HorizontalAlignment = 1;
            //    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            //    cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
            //    cell1.BorderWidth = 0;
            //    cell1.FixedHeight = 23f;
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.Colspan = 1;
            //    Tbl1.AddCell(cell1);


            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell1.BorderWidth = 0;
            //    Tbl1.AddCell(cell1);

            //    var Addresss = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
            //    for (int a = 0; a < Addresss.Length; a++)
            //    {
            //        cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //        cell1.BorderWidth = 0;
            //        Tbl1.AddCell(cell1);
            //    }
            //    mtable.AddCell(Tbl1);



            //    Tbl1 = new PdfPTable(2);
            //    Tbl1.WidthPercentage = 50;
            //    Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


            //    cell1 = new PdfPCell(new Phrase("BOOKING N0", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);


            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BookingNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("SHIPMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ShipmentType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("SALES", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["SalesPerson"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);


            //    cell1 = new PdfPCell(new Phrase("SERVICE TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BorderWidth = 1;
            //    cell1.FixedHeight = 25f;
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    Tbl1.AddCell(cell1);

            //    mtable.AddCell(Tbl1);
            //    doc.Add(mtable);

            //    #endregion

            //    #region Location POL POD POO

            //    PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
            //    TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
            //    TblLocs.SpacingBefore = 10;
            //    TblLocs.WidthPercentage = 100;

            //    cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);


            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POO"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);


            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblLocs.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
            //    TblLocs.AddCell(cell1);

            //    doc.Add(TblLocs);



            //    #endregion

            //    #region Booking Details
            //    //----------------- Booking Details--------------//

            //    iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
            //    Tbl3.Width = 100;
            //    Tbl3.Alignment = Element.ALIGN_LEFT;
            //    Tbl3.Cellpadding = 0;
            //    Tbl3.BorderWidth = 0;

            //    //Sub Heading
            //    cell = new Cell(new Phrase("Booking Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

            //    cell.BorderWidth = 0;
            //    cell.Colspan = 1;
            //    Tbl3.AddCell(cell);
            //    doc.Add(Tbl3);

            //    #region CntrValues

            //    PdfPTable TblCntrVal = new PdfPTable(new float[] { 2, 2, 2 });
            //    TblCntrVal.HorizontalAlignment = Element.ALIGN_LEFT;
            //    TblCntrVal.SpacingBefore = 10;
            //    TblCntrVal.WidthPercentage = 100;

            //    cell1 = new PdfPCell(new Phrase("Size", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblCntrVal.AddCell(cell1);

            //    cell1 = new PdfPCell(new Phrase("Quantity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblCntrVal.AddCell(cell1);


            //    cell1 = new PdfPCell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //    cell1.BackgroundColor = new Color(152, 178, 209);
            //    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //    TblCntrVal.AddCell(cell1);

            //    DataTable dtCT = GetBkgCntrValus(Data.BkgID);
            //    if (dtCT.Rows.Count > 0)
            //    {
            //        for (int i = 0; i < dtCT.Rows.Count; i++)
            //        {

            //            cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //            TblCntrVal.AddCell(cell1);

            //            cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //            TblCntrVal.AddCell(cell1);

            //            cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //            TblCntrVal.AddCell(cell1);
            //        }
            //        doc.Add(TblCntrVal);
            //    }

            //    #endregion

            //    iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
            //    Tbl5.Width = 100;
            //    Tbl5.Alignment = Element.ALIGN_LEFT;
            //    Tbl5.Cellpadding = 1;
            //    Tbl5.BorderWidth = 0;

            //    ////Caption
            //    //cell = new Cell(new Phrase("Volume", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    //cell.BorderWidth = 0;
            //    //cell.Colspan = 2;
            //    //Tbl5.AddCell(cell);
            //    ////Value
            //    //DataTable dtCT = GetBkgCntrValus(Data.BkgID);
            //    //if (dtCT.Rows.Count > 0)
            //    //{
            //    //    for (int i = 0; i < dtCT.Rows.Count; i++)
            //    //    {
            //    //        cell = new Cell(new Phrase(": " + dtCT.Rows[i]["Size"].ToString() + " * " + dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    //        cell.BorderWidth = 0;
            //    //        cell.Colspan = 3;
            //    //        Tbl5.AddCell(cell);
            //    //    }
            //    //}

            //    ////Caption
            //    //cell = new Cell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    //cell.BorderWidth = 0;
            //    //cell.Colspan = 3;
            //    //Tbl5.AddCell(cell);
            //    ////Value
            //    //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CommodityType"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    //cell.BorderWidth = 0;
            //    //cell.Colspan = 3;
            //    //Tbl5.AddCell(cell);


            //    //Caption
            //    cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 2;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 4;
            //    Tbl5.AddCell(cell);

            //    //Caption
            //    cell = new Cell(new Phrase("ETA / ETD", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);

            //    //Value
            //    cell = new Cell(new Phrase(" :  " + dtv.Rows[0]["ETDDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);


            //    //Caption
            //    cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 2;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CutDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 4;
            //    Tbl5.AddCell(cell);


            //    //Caption
            //    cell = new Cell(new Phrase("Box Operator Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + "ONL", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);

            //    //Caption
            //    cell = new Cell(new Phrase("Loading Terminal", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 2;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + " GTI", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 4;
            //    Tbl5.AddCell(cell);

            //    //Caption
            //    cell = new Cell(new Phrase("SCN No", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["SCNNo"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);

            //    //Caption
            //    cell = new Cell(new Phrase("Carrier", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 2;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CarrierName"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 4;
            //    Tbl5.AddCell(cell);

            //    //Caption
            //    cell = new Cell(new Phrase("Shipper", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);

            //    //Caption
            //    cell = new Cell(new Phrase("Pick Up Depot", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 2;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 4;
            //    Tbl5.AddCell(cell);

            //    cell = new Cell(new Phrase("Port Net Reference", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);
            //    //Value
            //    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PortNtRef"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 3;
            //    Tbl5.AddCell(cell);

            //    doc.Add(Tbl5);


            //    #endregion

            //    #region Terms & Condition

            //    iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
            //    Tbl7.Width = 100;
            //    Tbl7.Alignment = Element.ALIGN_LEFT;
            //    Tbl7.Cellpadding = 1;
            //    Tbl7.BorderWidth = 0;

            //    cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 1;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * Kindly advise us at least 5 days before vessel’s arrival if you are unable to load on the said vessel", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * Kindly provide us the Shipping Instruction 2 days before prior vessel ETD", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * Bill of Lading Release - Cash against document. THE ABOVE MENTIONED CLOSING TIME IS SUBJECT TO CHANGE DEPENDING ON ACTUAL VESSEL ARRIVAL AT PORT OF LOADING", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * Detention charges as per standard detention slab to be applied after the normal drop cost free days and refer container ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * Either Line /Agent will not bear any storage charges. ", new Font(Font.HELVETICA, 11, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * ETA of vessels are subject to change without prior notice. Please consistently check with the port system for updated information. ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * All dangerous and fumigation cargo must be declared at the time of the booking. ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * Pre inspection certificate for all scrap cargos at the time of BL submission for all India sectors. ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    cell = new Cell(new Phrase(" * MSDG certificate required for all types of DG cargo ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    Tbl7.AddCell(cell);

            //    doc.Add(Tbl7);

            //    #endregion

            //    #region FOOTER
            //    ///---------FOOTER----------------//
            //    ///
            //    //Sub Heading
            //    iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
            //    Tbl8.Width = 100;
            //    Tbl8.Alignment = Element.ALIGN_LEFT;
            //    Tbl8.Cellpadding = 0;
            //    Tbl8.BorderWidth = 0;


            //    cell = new Cell(new Phrase("Thank you very much on your booking confirmation with us. & Looking forward to your valuable support for future bookings.", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));

            //    cell.BorderWidth = 0;
            //    cell.Colspan = 1;
            //    Tbl8.AddCell(cell);
            //    doc.Add(Tbl8);

            //    iTextSharp.text.Table Tbl9 = new iTextSharp.text.Table(1);
            //    Tbl9.Width = 100;
            //    Tbl9.Alignment = Element.ALIGN_LEFT;
            //    Tbl9.Cellpadding = 1;
            //    Tbl9.BorderWidth = 0;

            //    cell = new Cell(new Phrase("Best regards,", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 1;
            //    Tbl9.AddCell(cell);

            //    cell = new Cell(new Phrase("Customer service team", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    cell.Colspan = 1;
            //    Tbl9.AddCell(cell);
            //    doc.Add(Tbl9);



            //    #endregion

            //    pdfWriter.CloseStream = false;
            //    doc.Close();


            //}

            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                //PdfWriter pdfWriter1 = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/BKGPDF\\" + dtv.Rows[0]["BookingNo"].ToString() + ".pdf"), FileMode.Create));
                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);

                doc.Open();

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;

                DataTable dtc = GetAgencyDetails(Data.AgentID);
                if (dtc.Rows.Count > 0)
                {


                    var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                    img.Alignment = Element.ALIGN_LEFT;
                    img.ScaleAbsolute(150f, 80f);
                    cell = new Cell(img);
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tbllogo.AddCell(cell);
                }


                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--////

                cell = new Cell(new Phrase(dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);



                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);
                ///----/////

                //DataTable dta = GetCompanyDetails();
                //if (dta.Rows.Count > 0)
                //{
                //    if (dta.Rows[0]["CompanyID"].ToString() != "")
                //    {
                //        cell = new Cell(new Phrase("Agent Of " + dta.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //        cell.BorderWidth = 0;
                //        tbllogo.Alignment = Element.ALIGN_LEFT;
                //        //cell.Colspan = 2;
                //        tbllogo.AddCell(cell);
                //    }
                //}
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("Booking Confirmation", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString() + "\r\n" + "GST: " + dtc.Rows[0]["TaxGSTNo"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                }

                doc.Add(tbllogo);

                //para = new Paragraph("");
                //doc.Add(para);

                para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace2);

                //------------------------------------------------------------------------

                #endregion

                #region Bookingparty and Ratesheet details
                //-------------------Bookingparty and Ratesheet details-----------
                PdfContentByte content = pdfWriter.DirectContent;
                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Booking Party", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }
                mtable.AddCell(Tbl1);



                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("BOOKING N0", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BookingNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SHIPMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ShipmentType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SALES", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["SalesPerson"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("SERVICE TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region Location POL POD POO

                PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                TblLocs.SpacingBefore = 10;
                TblLocs.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POO"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                TblLocs.AddCell(cell1);

                doc.Add(TblLocs);



                #endregion


                #region Booking Details
                //----------------- Booking Details--------------//

                iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                Tbl3.Width = 100;
                Tbl3.Alignment = Element.ALIGN_LEFT;
                Tbl3.Cellpadding = 0;
                Tbl3.BorderWidth = 0;

                //Sub Heading
                cell = new Cell(new Phrase("Booking Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl3.AddCell(cell);
                doc.Add(Tbl3);

                #region CntrValues

                PdfPTable TblCntrVal = new PdfPTable(new float[] { 2, 2, 2 });
                TblCntrVal.HorizontalAlignment = Element.ALIGN_LEFT;
                TblCntrVal.SpacingBefore = 10;
                TblCntrVal.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Size", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Quantity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);

                DataTable dtCT = GetBkgCntrValus(Data.BkgID);
                if (dtCT.Rows.Count > 0)
                {
                    for (int i = 0; i < dtCT.Rows.Count; i++)
                    {

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);
                    }
                    doc.Add(TblCntrVal);
                }

                #endregion


                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 1;
                Tbl5.BorderWidth = 0;

                //Caption
                //cell = new Cell(new Phrase("Volume", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 2;
                //Tbl5.AddCell(cell);
                ////Value

                //DataTable dtCT = GetBkgCntrValus(BkgID);
                //if (dtCT.Rows.Count > 0)
                //{
                //    for (int i = 0; i < dtCT.Rows.Count; i++)
                //    {
                //        cell = new Cell(new Phrase(" : " + dtCT.Rows[i]["Size"].ToString() + " * " + dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //    cell.BorderWidth = 0;
                //    cell.Colspan = 2;
                //    Tbl5.AddCell(cell);
                //    }
                //}


                ////Caption
                //cell = new Cell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 3;
                //Tbl5.AddCell(cell);
                ////Value
                //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CommodityType"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 3;
                //Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("ETA / ETD", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dtv.Rows[0]["ETADate"].ToString() + "/ " + dtv.Rows[0]["ETDDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                if (Data.AgentID == "3")
                {
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["ClosingTime"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }
                else
                {
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CutDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }


                //Caption
                cell = new Cell(new Phrase("Next Port ETA", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["NextPortETA"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Loading Terminal", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Terminal"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Box Operator Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                if (Data.AgentID == "3")
                {
                    cell = new Cell(new Phrase(" : " + "BWS", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                }
                else
                {
                    cell = new Cell(new Phrase(" : " + "BWS", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                }

                if (Data.AgentID != "3")
                {
                    //Caption
                    cell = new Cell(new Phrase("Carrier", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CarrierName"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }
                else
                {  //Caption
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);

                }

                //Caption
                cell = new Cell(new Phrase("Shipper", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("VESSEL CLOSING TIME", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["ClosingTime"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("Port Net Reference", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PortNtRef"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("SCN No", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["SCNNo"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("BS CODE", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["BSCODE"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("Vessel ID", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesselIDValue"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                if (Data.AgentID != "3")
                {
                    cell = new Cell(new Phrase("Pick Up Depot", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);

                }
                doc.Add(Tbl5);


                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 1;
                Tbl7.BorderWidth = 0;


                cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);

                DataTable _dtv = GetNotesClausesBooking();
                if (_dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < _dtv.Rows.Count; i++)
                    {
                        cell = new Cell(new Phrase("*" + _dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        Tbl7.AddCell(cell);

                    }
                }



                doc.Add(Tbl7);

                #endregion

                #region FOOTER
                ///---------FOOTER----------------//
                ///
                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("Thank you very much on your booking confirmation with us. & Looking forward to your valuable support for future bookings.", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl8.AddCell(cell);
                doc.Add(Tbl8);

                iTextSharp.text.Table Tbl9 = new iTextSharp.text.Table(1);
                Tbl9.Width = 100;
                Tbl9.Alignment = Element.ALIGN_LEFT;
                Tbl9.Cellpadding = 1;
                Tbl9.BorderWidth = 0;

                cell = new Cell(new Phrase("Best regards,", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);

                cell = new Cell(new Phrase("Customer service team", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);
                doc.Add(Tbl9);



                #endregion

                pdfWriter.CloseStream = false;
                doc.Close();

            }


            DataTable _dtCom = GetCompnayDetails();
            DataTable _dtAgc = GetAgencyDetailssending(Data.AgentID);
            if (_dtAgc.Rows.Count > 0)
            {

                DataTable _dtBkg = GetBkgDetailssending(Data.BkgID);
                strHTML = "<table border='1' cellpadding='0' cellspacing='0' width='80%' style='font-family:Arial; border:1px solid #2196f3;'>" +
                                 "<tbody>" +
                                 "<tr> " +
                                 "<td style='border:none!important;'>" +
                                 "<table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:15px;'>" +
                                 "<tr><td style='background-image: url(http://" + _dtCom.Rows[0]["Website"].ToString() + "/emailimg/booking1.jpg); background-position: center; background-size: cover; height: 200px;'> " +
                                 "<table cellpadding='0' cellspacing='0' width='65%' style='margin-top:-79px; padding-left:30px;'><tr><td style='font-weight:bold; font-size:16px; padding-top:7px;'> " + _dtAgc.Rows[0]["AgencyName"].ToString() + " </td></tr>" +
                                 "<tr><td style='font-size: 13px; padding-top:7px;'>Agent Of " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr><tr><td style='font-size: 13px;'> " + _dtAgc.Rows[0]["Address"].ToString() + ",<br/> GST:  " + _dtAgc.Rows[0]["TaxGSTNo"].ToString() + " </td></tr></table></tr> " +
                                 "</table>" +
                                 "</td></tr><tr><td style='border:none!important;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:50px;'>" +
                                 "<tr><td style='padding-left:3px; font-weight:bold;'>Hi,</td> </tr></table></td></tr><tr><td style='border:none!important;'>" +
                                 "<table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'><tr><td style='padding-left: 3px; font-weight: bold; text-align: center; color:#3eb348;'> " +
                                 "Your Booking has been Confirmed!</td></tr><tr><td style='text-align:center;'><img src='http:http:/" + _dtCom.Rows[0]["Website"].ToString() + "/assets/img/emailimg/statusimg.jpg' style='margin-top:15px;' class='img-responsive' /> " +
                                 "</td></tr><tr><td><span style='padding-left: 170px; font-size:12px;color:#ccc;'>PENDING</span><span style='float:right; padding-right: 160px; color: #0ebc67; font-weight: bold; font-size: 12px;'>CONFIRMED</span>" +
                                 "</td></tr></table></td></tr><tr><td style='border:none!important;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'><tr> " +
                                 "<td style='padding-left: 3px; font-weight: bold; text-align: center; color:#3eb348; font-size:15px;'>Booking No : " + _dtBkg.Rows[0]["BookingNo"].ToString() + "</td> </tr></table></td></tr>" +
                                 "<tr><td style='border:none!important;'><table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:50px;padding-right:50px;padding-top:20px;padding-bottom:20px; font-size: 14px; background-color:#cccccc1f'>" +
                                 "<tbody><tr><td colspan='2'><table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:94px;'><tbody><tr>" +
                                 "<td colspan='2' style='font-weight:bold;'>Port Of Loading</td></tr><tr><td colspan='2' style='padding-bottom:8px;'>" + _dtBkg.Rows[0]["POL"].ToString() + "</td>" +
                                 "</tr><tr><td colspan='2' style='font-weight:bold;'>Volume</td></tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtBkg.Rows[0]["CntrTypes"].ToString() + "</td></tr><tr>" +
                                 "<td colspan='2' style='font-weight:bold;'>Vessel & Voyage</td></tr><tr><td colspan='2' style='padding-bottom:8px;'>" + _dtBkg.Rows[0]["vesvoy"].ToString() + "</td></tr></tbody>" +
                                 "</table></td><td colspan='2'><table border='0' cellpadding='0' cellspacing='0' width= '100%'><tbody><tr><td colspan='2' style='font-weight:bold;'> Port Of Discharge</td>" +
                                 "</tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtBkg.Rows[0]["POD"].ToString() + "</td></tr><tr><td colspan='2' style='font-weight:bold;'> Commodity </td>" +
                                 "</tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtBkg.Rows[0]["Commodity"].ToString() + "  </td></tr><tr><td colspan='2' style='font-weight:bold;'> Cut - Off Date</td></tr> " +
                                 "<tr> <td colspan='2' style='padding-bottom:8px;'>  " + _dtBkg.Rows[0]["CutDate"].ToString() + " </td></tr></tbody></table></td></tr></tbody></table></td> </tr><tr> " +
                                 "<td style='border:none!important;'><table cellpadding='0' cellspacing='0' width= '100%' style='margin-top:20px;'><tr><td style='padding-left:3px; font-weight:600;font-size:13px;'>" +
                                 "Thank you very much on your booking confirmation with us.Looking forward to your valuable support for future bookings</td></tr></table></td></tr>" +
                                 "<tr><td style='border:none!important;'><table cellpadding='0' cellspacing='0' width= '100%' style='margin-top:20px;'><tr><td style='padding-left:3px; font-weight:600;font-size:13px;'>" +
                                 "Best regards,</td></tr><tr><td style='padding-left:3px; font-weight:600;font-size:13px;'>Customer Service Team</td></tr></table> </td></tr></tbody> </table>";
            }

            MailMessage EmailObject = new MailMessage();
            EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());

            var EmailIDTo = Data.To.Split(',');
            for (int y = 0; y < EmailIDTo.Length; y++)
            {
                if (EmailIDTo[y].ToString() != "")
                {
                    EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                }
            }

            var EmailIDCC = Data.CC.Split(',');
            for (int y = 0; y < EmailIDCC.Length; y++)
            {
                if (EmailIDCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDCC[y].ToString()));
                }
            }

            var EmailIDBCC = Data.BCC.Split(',');
            for (int y = 0; y < EmailIDBCC.Length; y++)
            {
                if (EmailIDBCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDBCC[y].ToString()));
                }
            }

            byte[] bytes = memoryStream.ToArray();
            memoryStream.Close();
            EmailObject.Bcc.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
            EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "BookingPDF.pdf"));
            EmailObject.Body = strHTML;
            EmailObject.IsBodyHtml = true;
            EmailObject.Priority = MailPriority.Normal;
            EmailObject.Subject = "BOOKING:";
            EmailObject.Priority = MailPriority.Normal;
            SmtpClient SMTPServer = new SmtpClient();
            SMTPServer.UseDefaultCredentials = true;
            SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());
            SMTPServer.Host = "smtp.office365.com";
            SMTPServer.ServicePoint.MaxIdleTime = 1;
            SMTPServer.Port = 587;
            SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
            SMTPServer.EnableSsl = true;
            SMTPServer.Send(EmailObject);

            ViewList.Add(new MySendingEmailAlert
            {
                AlertMessage = "Email sent successfully"

            });

            return ViewList;
        }


        public List<MySendingEmailAlert> CROSendingEmail(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            string strHTML = "";
            DataTable _dtCom = GetCompnayDetails();
            MemoryStream memoryStream = new MemoryStream();

            DataTable dt = GetCROPDFValus(Data.CRId);
            if (dt.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                #region First Page

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;


                DataTable dtc = GetAgencyDetails(Data.AgentID);
                if (dtc.Rows.Count > 0)
                {
                    if (dtc.Rows[0]["LogoPath"].ToString() != "")
                    {

                        var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        img.ScaleAbsolute(150f, 80f);
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        tbllogo.AddCell(cell);




                    }
                    else
                    {
                        var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        img.ScaleAbsolute(150f, 80f);
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        tbllogo.AddCell(cell);
                    }
                }

                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--//

                cell = new Cell(new Phrase(dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);
                ///----/////

                DataTable dta = GetCompanyDetails();
                if (dta.Rows.Count > 0)
                {
                    if (dta.Rows[0]["CompanyID"].ToString() != "")
                    {

                        cell = new Cell(new Phrase("Agent Of " + dta.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        //cell.Colspan = 2;
                        tbllogo.AddCell(cell);


                    }
                }



                cell = new Cell(new Phrase("Container Release Order", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);


                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString() + "\r\n" + "GST: " + dtc.Rows[0]["TaxGSTNo"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);

                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                }

                doc.Add(tbllogo);

                para = new Paragraph("");
                doc.Add(para);

                para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace2);

                //------------------------------------------------------------------------

                #endregion

                #region Release to and Container Release Order details
                //-------------------Bookingparty and Ratesheet details-----------
                PdfContentByte content = pdfWriter.DirectContent;
                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Release To", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["BkgParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss = Regex.Split(dt.Rows[0]["Address"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }
                mtable.AddCell(Tbl1);



                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("RELEASE ORDER NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["ReleaseOrderNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("RELEASE DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["Date"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("BOOKING NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["BookingNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SERVICE TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("VALID TILL", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["ValidTill"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region Location POL POD POO

                // /----------------------- LocTable-----------------------///

                PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                TblLocs.SpacingBefore = 10;
                TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                TblLocs.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["POO"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                TblLocs.AddCell(cell1);

                doc.Add(TblLocs);



                #endregion

                #region Release Order Details
                //----------------- Release Order Details--------------//

                PdfPTable Tbl3 = new PdfPTable(1);
                Tbl3.WidthPercentage = 100;
                Tbl3.SpacingBefore = 10;
                Tbl3.SpacingAfter = 0;
                Tbl3.HorizontalAlignment = Element.ALIGN_LEFT;

                //Sub Heading
                cell1 = new PdfPCell(new Phrase("Release Order Details", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                Tbl3.AddCell(cell1);
                doc.Add(Tbl3);

                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(13);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 1;
                Tbl5.BorderWidth = 0;

                //Caption
                cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("ETA", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["ETADate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase(" ETD", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["ETDDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["CUTDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("Line Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["Linecode"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                doc.Add(Tbl5);


                #endregion

                #region Container Type & Release Quantity

                iTextSharp.text.Table TblCntrTypes = new iTextSharp.text.Table(4);
                TblCntrTypes.Width = 100;
                TblCntrTypes.Alignment = Element.ALIGN_LEFT;
                TblCntrTypes.Cellpadding = 1;
                TblCntrTypes.BorderWidth = 1;

                cell = new Cell(new Phrase("Container Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);

                cell = new Cell(new Phrase("Release Quantity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);


                DataTable dtCroDtls = GetCRODetailsPDFValues(Data.CRId);
                for (int i = 0; i < dtCroDtls.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrTypes.AddCell(cell);

                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["ReqQty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrTypes.AddCell(cell);

                }
                doc.Add(TblCntrTypes);

                #endregion

                #region Shipper,Surveyor Details,Pick Up Depo ,Remarks 

                iTextSharp.text.Table TblShipper = new iTextSharp.text.Table(2);
                TblShipper.Width = 100;
                TblShipper.Alignment = Element.ALIGN_LEFT;
                TblShipper.Cellpadding = 0;
                TblShipper.BorderWidth = 0;

                ////---------SHIPPER-----------///
                cell = new Cell(new Phrase("Shipper", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);


                cell = new Cell(new Phrase(" :   " + dt.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);

                ////---------Surveyor Details -----------///

                cell = new Cell(new Phrase("Surveyor Details", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["SurveyorName"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(dt.Rows[0]["SurveyorAddress"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                ////---------Pick Up Depo -----------///

                cell = new Cell(new Phrase("Pick Up Depo", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(dt.Rows[0]["DepotAddress"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                //var DepotAddress = Regex.Split(dt.Rows[0]["DepotAddress"].ToString(), "\r\n|\r|\n");
                //for (int a = 0; a < DepotAddress.Length; a++)
                //{
                //    cell = new Cell(new Phrase(DepotAddress[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                //    cell.BorderWidth = 0;
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.Rowspan = 2;
                //    TblShipper.AddCell(cell);
                //}

                ////---------Remarks -----------///

                cell = new Cell(new Phrase("Remarks", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);

                cell = new Cell(new Phrase(" :    " + dt.Rows[0]["Remarks"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);

                doc.Add(TblShipper);

                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 0;
                Tbl7.BorderWidth = 0;

                cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);


                //cell = new Cell(new Phrase(" * Ensure the empty container received from our yard is in clean and sound condition. Costs for any subsequent rejection will be to your account.", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" * Any loss or damage to the container while in custody of shipper, transporter, forwarder shal be fully identified for repair / replacement / reimbursement as notified by owner / hirer. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" *  loading list needs to be sent to the shipping line by 48 hour prior to vessel cut off. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" * For Non-availability of Containers kindly contact our Operations Incharge - Mr.Nivrutti - (M) +91-90223 45131 will be on your account.", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" *FORM 13 to be collected & Shipping Bill to be handed over to our surveyor as mentioned in above. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //doc.Add(Tbl7);

                DataTable _dtv = GetNotesClauses();
                if (_dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < _dtv.Rows.Count; i++)
                    {
                        cell = new Cell(new Phrase(_dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        Tbl7.AddCell(cell);
                    }

                    cell = new Cell(new Phrase("" + "\n \n", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Colspan = 1;
                    Tbl7.AddCell(cell);
                }

                //cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                //cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_CENTER;
                //cell.Colspan = 1;
                //Tbl7.AddCell(cell);

                doc.Add(Tbl7);
                #endregion

                #region FOOTER
                ///---------FOOTER----------------//
                ///
                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_CENTER;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.RED)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.Colspan = 1;
                Tbl8.AddCell(cell);
                doc.Add(Tbl8);



                #endregion

                #endregion




                pdfWriter.CloseStream = false;
                doc.Close();

            }




            DataTable _dtAgc = GetAgencyDetailssending(Data.AgentID);
            if (_dtAgc.Rows.Count > 0)
            {

                DataTable _dtCRO = GetCROSending(Data.CRId);
                strHTML = " <table border='1' cellpadding='0' cellspacing='0' width='100%' style='font-family:Arial; border:1px solid #2196f3; margin:0 auto;'><tbody><tr>" +
                           " <td style='border:none!important;'>" +
                           " <table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:15px;'>" +
                           " <tr><td style='background-image: url(http://" + _dtCom.Rows[0]["Website"].ToString() + "/assets/img/emailimg/booking1.jpg); background-position: center; background-size: cover; height: 200px;'> " +
                           " <table cellpadding='0' cellspacing='0' width='65%' style='margin-top:-79px; padding-left:30px;'><tr><td style='font-weight:bold; font-size:16px; padding-top:7px;'> " + _dtAgc.Rows[0]["AgencyName"].ToString() + " </td></tr>" +
                           " <tr><td style='font-size: 13px; padding-top:7px;'>Agent Of " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr><tr><td style='font-size: 13px;'> " + _dtAgc.Rows[0]["Address"].ToString() + ",<br/> GST:  " + _dtAgc.Rows[0]["TaxGSTNo"].ToString() + " </td></tr></table></tr> " +
                           "</table></td></tr><tr><td style='border:none !important; padding-left: 30px;'>" +
                           "<table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'><tr><td style='padding-left:3px; font-weight:bold;'>" +
                           " Hi,</td></tr> </table></td></tr><tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'>" +
                           " <tr><td style='padding-left:3px; font-weight:bold; font-size:14px;'>Release To : " + _dtCRO.Rows[0]["BkgParty"].ToString() + "</td></tr></table></td></tr> " +
                           " <tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'> " +
                           " <tr><td style='padding-left: 3px; font-weight: bold; color:#3eb348; font-size:14px; margin-top:20px;'> Booking No : " + _dtCRO.Rows[0]["BookingNo"].ToString() + "</td></tr></table>" +
                           " </td></tr><tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'>" +
                           " <tr><td style='padding-left:3px; font-weight: bold; color:#3eb348; font-size:14px;'>Release Order No : " + _dtCRO.Rows[0]["ReleaseOrderNo"].ToString() + "</td></tr></table></td></tr><tr> " +
                           " <td style='border:none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'><tr> " +
                           " <td style='padding-left: 3px; font-weight: bold; color:red; font-size:14px;'>Validity Till : " + _dtCRO.Rows[0]["ValidTill"].ToString() + "</td></tr></table></td></tr> " +
                           " <tr><td style='border:none!important;'><table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:30px;padding-right:50px;padding-top:20px;padding-bottom:20px; font-size: 14px; background-color:#cccccc1f'>" +
                           " <tbody><tr><td colspan='2'><table border='0' cellpadding='0' cellspacing='0' width='100%'><tbody><tr><td colspan='2' style='font-weight:bold;'>Port Of Loading</td>" +
                           " </tr><tr><td colspan='2' style='padding-bottom:8px;'>" + _dtCRO.Rows[0]["POL"].ToString() + "</td></tr><tr><td colspan='2' style='font-weight:bold;'> Volume</td></tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtCRO.Rows[0]["CntrTypes"].ToString() + "</td>" +
                           " </tr><tr><td colspan='2' style='font-weight:bold;'>Vessel & Voyage</td></tr><tr><td colspan='2' style='padding-bottom:8px;'>CAPE FAWLEY</td></tr></tbody></table></td>" +
                           " <td colspan='2'><table border= '0' cellpadding= '0' cellspacing= '0' width= '100%'><tbody><tr><td colspan='2' style='font-weight:bold;'> Port Of Discharge</td>" +
                           " </tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtCRO.Rows[0]["POD"].ToString() + "</td></tr><tr><td colspan='2' style='font-weight:bold;'> Commodity </td></tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtCRO.Rows[0]["Commodity"].ToString() + " </td>" +
                           " </tr><tr><td colspan='2' style='font-weight:bold;' >Cut - Off Date</td></tr><tr><td colspan='2' style='padding-bottom:8px;' > " + _dtCRO.Rows[0]["CutDate"].ToString() + "</td></tr>" +
                           " </tbody></table> </td></tr> <tr> <td style='border:none!important;'> <table cellpadding='0' cellspacing='0' width='100%' style='margin-top:20px;'><tr><td style='font-weight:bold;font-size:13px;'>" +
                           " Sureveyor </td><td style='font-weight:bold;font-size:13px;'>:</td><td style='font-weight:bold;font-size:13px;'>" + _dtCRO.Rows[0]["SurveyorName"].ToString() + " </td></tr></table></td> </tr><tr>" +
                           " <td style='border:none!important;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-top:20px;'><tr><td style='adding-left:3px; font-weight:600;font-size:13px; width:100px;'> " +
                           " Pick Up Depo</td><td>:</td><td style='padding-left:10px; font-weight:600;font-size:13px;'><table cellpadding='0' cellspacing='0' width='100%'><tr><td> " + _dtCRO.Rows[0]["PickUpDepot"].ToString() + " " +
                           " </td></tr><tr><td style='font-weight:normal;'>" + _dtCRO.Rows[0]["DepotAddress"].ToString() + " </td></tr></table></td></tr></table>" +
                           " </td></tr></tbody></table></td></tr><tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width= '100%' style='margin-top:20px;'>" +
                           " <tr><td style='padding-left:3px; font-weight:600;font-size:13px;'> Best regards,</td></tr><tr><td style='padding-left:3px; font-weight:600;font-size:13px;'>Admin</td>" +
                           " </tr></table></td></tr> </tbody></table>";
            }

            MailMessage EmailObject = new MailMessage();
            EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString()); ;

            var EmailIDTo = Data.To.Split(',');
            for (int y = 0; y < EmailIDTo.Length; y++)
            {
                if (EmailIDTo[y].ToString() != "")
                {
                    EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                }
            }

            var EmailIDCC = Data.CC.Split(',');
            for (int y = 0; y < EmailIDCC.Length; y++)
            {
                if (EmailIDCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDCC[y].ToString()));
                }
            }

            var EmailIDBCC = Data.BCC.Split(',');
            for (int y = 0; y < EmailIDBCC.Length; y++)
            {
                if (EmailIDBCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDBCC[y].ToString()));
                }
            }

            byte[] bytes = memoryStream.ToArray();
            memoryStream.Close();

            //EmailObject.To.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
            EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), dt.Rows[0]["ReleaseOrderNo"].ToString() + ".pdf"));
            EmailObject.Body = strHTML;
            EmailObject.IsBodyHtml = true;
            EmailObject.Priority = MailPriority.Normal;
            EmailObject.Subject = "CROPDF:"+ dt.Rows[0]["ReleaseOrderNo"].ToString();

            EmailObject.Priority = MailPriority.Normal;
            SmtpClient SMTPServer = new SmtpClient();
            SMTPServer.UseDefaultCredentials = true;
            SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());

            SMTPServer.Host = "smtp.office365.com";
            SMTPServer.ServicePoint.MaxIdleTime = 1;
            SMTPServer.Port = 587;
            SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
            SMTPServer.EnableSsl = true;
            SMTPServer.Send(EmailObject);

            ViewList.Add(new MySendingEmailAlert
            {
                AlertMessage = "Email sent successfully"

            });

            return ViewList;
        }

        public List<MySendingEmailAlert> CROSendingEmailDepo(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            string strHTML = "";
            DataTable _dtCom = GetCompnayDetails();
            MemoryStream memoryStream = new MemoryStream();

            DataTable dt = GetCROPDFValus(Data.CRId);
            if (dt.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                #region First Page

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;


                DataTable dtc = GetAgencyDetails(Data.AgentID);
                if (dtc.Rows.Count > 0)
                {
                    if (dtc.Rows[0]["LogoPath"].ToString() != "")
                    {

                        var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        img.ScaleAbsolute(150f, 80f);
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        tbllogo.AddCell(cell);




                    }
                    else
                    {
                        var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        img.ScaleAbsolute(150f, 80f);
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        tbllogo.AddCell(cell);
                    }
                }

                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--//

                cell = new Cell(new Phrase(dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);
                ///----/////

                DataTable dta = GetCompanyDetails();
                if (dta.Rows.Count > 0)
                {
                    if (dta.Rows[0]["CompanyID"].ToString() != "")
                    {

                        cell = new Cell(new Phrase("Agent Of " + dta.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        //cell.Colspan = 2;
                        tbllogo.AddCell(cell);


                    }
                }



                cell = new Cell(new Phrase("Container Release Order", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);


                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString() + "\r\n" + "GST: " + dtc.Rows[0]["TaxGSTNo"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);

                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                }

                doc.Add(tbllogo);

                para = new Paragraph("");
                doc.Add(para);

                para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace2);

                //------------------------------------------------------------------------

                #endregion

                #region Release to and Container Release Order details
                //-------------------Bookingparty and Ratesheet details-----------
                PdfContentByte content = pdfWriter.DirectContent;
                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Release To", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["BkgParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss = Regex.Split(dt.Rows[0]["Address"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }
                mtable.AddCell(Tbl1);



                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("RELEASE ORDER NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["ReleaseOrderNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("RELEASE DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["Date"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("BOOKING NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["BookingNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SERVICE TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("VALID TILL", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["ValidTill"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region Location POL POD POO

                // /----------------------- LocTable-----------------------///

                PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                TblLocs.SpacingBefore = 10;
                TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                TblLocs.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["POO"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dt.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                TblLocs.AddCell(cell1);

                doc.Add(TblLocs);



                #endregion

                #region Release Order Details
                //----------------- Release Order Details--------------//

                PdfPTable Tbl3 = new PdfPTable(1);
                Tbl3.WidthPercentage = 100;
                Tbl3.SpacingBefore = 10;
                Tbl3.SpacingAfter = 0;
                Tbl3.HorizontalAlignment = Element.ALIGN_LEFT;

                //Sub Heading
                cell1 = new PdfPCell(new Phrase("Release Order Details", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                Tbl3.AddCell(cell1);
                doc.Add(Tbl3);

                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(13);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 1;
                Tbl5.BorderWidth = 0;

                //Caption
                cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("ETA", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["ETADate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase(" ETD", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["ETDDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["CUTDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("Line Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["Linecode"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                doc.Add(Tbl5);


                #endregion

                #region Container Type & Release Quantity

                iTextSharp.text.Table TblCntrTypes = new iTextSharp.text.Table(4);
                TblCntrTypes.Width = 100;
                TblCntrTypes.Alignment = Element.ALIGN_LEFT;
                TblCntrTypes.Cellpadding = 1;
                TblCntrTypes.BorderWidth = 1;

                cell = new Cell(new Phrase("Container Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);

                cell = new Cell(new Phrase("Release Quantity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);


                DataTable dtCroDtls = GetCRODetailsPDFValues(Data.CRId);
                for (int i = 0; i < dtCroDtls.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrTypes.AddCell(cell);

                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["ReqQty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrTypes.AddCell(cell);

                }
                doc.Add(TblCntrTypes);

                #endregion

                #region Shipper,Surveyor Details,Pick Up Depo ,Remarks 

                iTextSharp.text.Table TblShipper = new iTextSharp.text.Table(2);
                TblShipper.Width = 100;
                TblShipper.Alignment = Element.ALIGN_LEFT;
                TblShipper.Cellpadding = 0;
                TblShipper.BorderWidth = 0;

                ////---------SHIPPER-----------///
                cell = new Cell(new Phrase("Shipper", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);


                cell = new Cell(new Phrase(" :   " + dt.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);

                ////---------Surveyor Details -----------///

                cell = new Cell(new Phrase("Surveyor Details", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["SurveyorName"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(dt.Rows[0]["SurveyorAddress"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                ////---------Pick Up Depo -----------///

                cell = new Cell(new Phrase("Pick Up Depo", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblShipper.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(dt.Rows[0]["DepotAddress"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblShipper.AddCell(cell);

                //var DepotAddress = Regex.Split(dt.Rows[0]["DepotAddress"].ToString(), "\r\n|\r|\n");
                //for (int a = 0; a < DepotAddress.Length; a++)
                //{
                //    cell = new Cell(new Phrase(DepotAddress[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                //    cell.BorderWidth = 0;
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.Rowspan = 2;
                //    TblShipper.AddCell(cell);
                //}

                ////---------Remarks -----------///

                cell = new Cell(new Phrase("Remarks", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);

                cell = new Cell(new Phrase(" :    " + dt.Rows[0]["Remarks"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);

                doc.Add(TblShipper);

                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 0;
                Tbl7.BorderWidth = 0;

                cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);


                //cell = new Cell(new Phrase(" * Ensure the empty container received from our yard is in clean and sound condition. Costs for any subsequent rejection will be to your account.", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" * Any loss or damage to the container while in custody of shipper, transporter, forwarder shal be fully identified for repair / replacement / reimbursement as notified by owner / hirer. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" *  loading list needs to be sent to the shipping line by 48 hour prior to vessel cut off. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" * For Non-availability of Containers kindly contact our Operations Incharge - Mr.Nivrutti - (M) +91-90223 45131 will be on your account.", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //cell = new Cell(new Phrase(" *FORM 13 to be collected & Shipping Bill to be handed over to our surveyor as mentioned in above. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //Tbl7.AddCell(cell);

                //doc.Add(Tbl7);

                DataTable _dtv = GetNotesClauses();
                if (_dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < _dtv.Rows.Count; i++)
                    {
                        cell = new Cell(new Phrase(_dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        Tbl7.AddCell(cell);
                    }

                    cell = new Cell(new Phrase("" + "\n \n", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Colspan = 1;
                    Tbl7.AddCell(cell);
                }

                //cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                //cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_CENTER;
                //cell.Colspan = 1;
                //Tbl7.AddCell(cell);

                doc.Add(Tbl7);
                #endregion

                #region FOOTER
                ///---------FOOTER----------------//
                ///
                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_CENTER;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.RED)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.Colspan = 1;
                Tbl8.AddCell(cell);
                doc.Add(Tbl8);



                #endregion

                #endregion




                pdfWriter.CloseStream = false;
                doc.Close();

            }




            DataTable _dtAgc = GetAgencyDetailssending(Data.AgentID);
            if (_dtAgc.Rows.Count > 0)
            {

                DataTable _dtCRO = GetCROSending(Data.CRId);
                strHTML = " <table border='1' cellpadding='0' cellspacing='0' width='100%' style='font-family:Arial; border:1px solid #2196f3; margin:0 auto;'><tbody><tr>" +
                           " <td style='border:none!important;'>" +
                           " <table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:15px;'>" +
                           " <tr><td style='background-image: url(http://" + _dtCom.Rows[0]["Website"].ToString() + "/assets/img/emailimg/booking1.jpg); background-position: center; background-size: cover; height: 200px;'> " +
                           " <table cellpadding='0' cellspacing='0' width='65%' style='margin-top:-79px; padding-left:30px;'><tr><td style='font-weight:bold; font-size:16px; padding-top:7px;'> " + _dtAgc.Rows[0]["AgencyName"].ToString() + " </td></tr>" +
                           " <tr><td style='font-size: 13px; padding-top:7px;'>Agent Of " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr><tr><td style='font-size: 13px;'> " + _dtAgc.Rows[0]["Address"].ToString() + ",<br/> GST:  " + _dtAgc.Rows[0]["TaxGSTNo"].ToString() + " </td></tr></table></tr> " +
                           "</table></td></tr><tr><td style='border:none !important; padding-left: 30px;'>" +
                           "<table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'><tr><td style='padding-left:3px; font-weight:bold;'>" +
                           " Hi,</td></tr> </table></td></tr><tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'>" +
                           " <tr><td style='padding-left:3px; font-weight:bold; font-size:14px;'>Release To : " + _dtCRO.Rows[0]["BkgParty"].ToString() + "</td></tr></table></td></tr> " +
                           " <tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'> " +
                           " <tr><td style='padding-left: 3px; font-weight: bold; color:#3eb348; font-size:14px; margin-top:20px;'> Booking No : " + _dtCRO.Rows[0]["BookingNo"].ToString() + "</td></tr></table>" +
                           " </td></tr><tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'>" +
                           " <tr><td style='padding-left:3px; font-weight: bold; color:#3eb348; font-size:14px;'>Release Order No : " + _dtCRO.Rows[0]["ReleaseOrderNo"].ToString() + "</td></tr></table></td></tr><tr> " +
                           " <td style='border:none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:20px;'><tr> " +
                           " <td style='padding-left: 3px; font-weight: bold; color:red; font-size:14px;'>Validity Till : " + _dtCRO.Rows[0]["ValidTill"].ToString() + "</td></tr></table></td></tr> " +
                           " <tr><td style='border:none!important;'><table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:30px;padding-right:50px;padding-top:20px;padding-bottom:20px; font-size: 14px; background-color:#cccccc1f'>" +
                           " <tbody><tr><td colspan='2'><table border='0' cellpadding='0' cellspacing='0' width='100%'><tbody><tr><td colspan='2' style='font-weight:bold;'>Port Of Loading</td>" +
                           " </tr><tr><td colspan='2' style='padding-bottom:8px;'>" + _dtCRO.Rows[0]["POL"].ToString() + "</td></tr><tr><td colspan='2' style='font-weight:bold;'> Volume</td></tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtCRO.Rows[0]["CntrTypes"].ToString() + "</td>" +
                           " </tr><tr><td colspan='2' style='font-weight:bold;'>Vessel & Voyage</td></tr><tr><td colspan='2' style='padding-bottom:8px;'>CAPE FAWLEY</td></tr></tbody></table></td>" +
                           " <td colspan='2'><table border= '0' cellpadding= '0' cellspacing= '0' width= '100%'><tbody><tr><td colspan='2' style='font-weight:bold;'> Port Of Discharge</td>" +
                           " </tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtCRO.Rows[0]["POD"].ToString() + "</td></tr><tr><td colspan='2' style='font-weight:bold;'> Commodity </td></tr><tr><td colspan='2' style='padding-bottom:8px;'> " + _dtCRO.Rows[0]["Commodity"].ToString() + " </td>" +
                           " </tr><tr><td colspan='2' style='font-weight:bold;' >Cut - Off Date</td></tr><tr><td colspan='2' style='padding-bottom:8px;' > " + _dtCRO.Rows[0]["CutDate"].ToString() + "</td></tr>" +
                           " </tbody></table> </td></tr> <tr> <td style='border:none!important;'> <table cellpadding='0' cellspacing='0' width='100%' style='margin-top:20px;'><tr><td style='font-weight:bold;font-size:13px;'>" +
                           " Sureveyor </td><td style='font-weight:bold;font-size:13px;'>:</td><td style='font-weight:bold;font-size:13px;'>" + _dtCRO.Rows[0]["SurveyorName"].ToString() + " </td></tr></table></td> </tr><tr>" +
                           " <td style='border:none!important;'><table cellpadding='0' cellspacing='0' width='100%' style='margin-top:20px;'><tr><td style='adding-left:3px; font-weight:600;font-size:13px; width:100px;'> " +
                           " Pick Up Depo</td><td>:</td><td style='padding-left:10px; font-weight:600;font-size:13px;'><table cellpadding='0' cellspacing='0' width='100%'><tr><td> " + _dtCRO.Rows[0]["PickUpDepot"].ToString() + " " +
                           " </td></tr><tr><td style='font-weight:normal;'>" + _dtCRO.Rows[0]["DepotAddress"].ToString() + " </td></tr></table></td></tr></table>" +
                           " </td></tr></tbody></table></td></tr><tr><td style='border: none !important; padding-left: 30px;'><table cellpadding='0' cellspacing='0' width= '100%' style='margin-top:20px;'>" +
                           " <tr><td style='padding-left:3px; font-weight:600;font-size:13px;'> Best regards,</td></tr><tr><td style='padding-left:3px; font-weight:600;font-size:13px;'>Admin</td>" +
                           " </tr></table></td></tr> </tbody></table>";
            }

            MailMessage EmailObject = new MailMessage();
            EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString()); ;

            var EmailIDTo = Data.To.Split(',');
            for (int y = 0; y < EmailIDTo.Length; y++)
            {
                if (EmailIDTo[y].ToString() != "")
                {
                    EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                }
            }

            var EmailIDCC = Data.CC.Split(',');
            for (int y = 0; y < EmailIDCC.Length; y++)
            {
                if (EmailIDCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDCC[y].ToString()));
                }
            }

            var EmailIDBCC = Data.BCC.Split(',');
            for (int y = 0; y < EmailIDBCC.Length; y++)
            {
                if (EmailIDBCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDBCC[y].ToString()));
                }
            }

            byte[] bytes = memoryStream.ToArray();
            memoryStream.Close();

            //EmailObject.To.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
            EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), dt.Rows[0]["ReleaseOrderNo"].ToString() + ".pdf"));
            EmailObject.Body = strHTML;
            EmailObject.IsBodyHtml = true;
            EmailObject.Priority = MailPriority.Normal;
            EmailObject.Subject = "CROPDF:" + dt.Rows[0]["ReleaseOrderNo"].ToString();

            EmailObject.Priority = MailPriority.Normal;
            SmtpClient SMTPServer = new SmtpClient();
            SMTPServer.UseDefaultCredentials = true;
            SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());

            SMTPServer.Host = "smtp.office365.com";
            SMTPServer.ServicePoint.MaxIdleTime = 1;
            SMTPServer.Port = 587;
            SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
            SMTPServer.EnableSsl = true;
            SMTPServer.Send(EmailObject);

            ViewList.Add(new MySendingEmailAlert
            {
                AlertMessage = "Email sent successfully"

            });

            return ViewList;
        }



        public List<MySendingEmailAlert> CANSendingEmail(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            string strHTML = "";
            DataTable _dtCom = GetCompnayDetails();
            MemoryStream memoryStream = new MemoryStream();

            Document doc = new Document();
            //Rectangle rec = new Rectangle(670, 900);
            Rectangle rec = new Rectangle(700, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
            doc.Open();

            PdfContentByte cb = writer.DirectContent;
            cb.SetColorStroke(Color.BLACK);
            int _Xp = 10, _Yp = 785, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

            //Document doc = new Document();
            //Rectangle rec = new Rectangle(670, 900);
            //doc = new Document(rec);
            //Paragraph para = new Paragraph();


            //PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
            //doc.Open();

            //PdfContentByte cb = writer.DirectContent;
            //cb.SetColorStroke(Color.BLACK);
            //int _Xp = 10, _Yp = 785, YDiff = 10;

            //BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            //cb.SetFontAndSize(bfheader, 14);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

            DataTable dtc = GetAgencyDetails(Data.AgentID);
            if (dtc.Rows.Count > 0)
            {


                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/logoBWS.png"));
                //png1.SetAbsolutePosition(300, 840);
                //png1.ScalePercent(80f);
                //doc.Add(png1);
                png1.SetAbsolutePosition(25, 835);     //logo fixed location
                png1.ScalePercent(10f);
                doc.Add(png1);


            }


            cb.BeginText();
            //BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            //cb.SetFontAndSize(bfheader2, 15);
            //cb.SetColorFill(Color.BLACK);


            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 13);
            cb.SetColorFill(Color.BLACK);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtCom.Rows[0]["EmailHeader"].ToString(), 180, 835, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NOTICE OF ARRIVAL WITH INVOICE", 450, 820, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CARGO ARRIVAL NOTICE", 250, 860, 0);//right

            //DataTable dt = GetCanInvoice(Data.BLID.ToString());
            //if (dt.Rows.Count > 0)
            //{
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BLNumber"].ToString(), 50, 750, 0);
            //    // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "20‐08‐2023", 210, 750, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesVoy"].ToString(), 290, 750, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["Terminal"].ToString(), 450, 750, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["ETA"].ToString(), 600, 750, 0);



            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["CONSIGNEE"].ToString(), 15, 710, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["CANReleaseAddress"].ToString(), 15, 690, 0);




            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["POL"].ToString(), 40, 640, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["POD"].ToString(), 190, 640, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BSCode"].ToString(), 360, 640, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["SCNCode"].ToString(), 430, 640, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesselID1"].ToString(), 530, 640, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["ImpFreeDays"].ToString(), 280, 620, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 460, 620, 0);
            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 600, 620, 0);

            //}
            //int IndexRows = 580;
            //int ColumnIndex = 15;
            //DataTable dtcntr = GetCanContainerPrint(Data.BLID.ToString());
            //for (int i = 0; i < dtcntr.Rows.Count; i++)
            //{

            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows[i]["CntrDtls"].ToString() + "", ColumnIndex, IndexRows, 0);
            //    ColumnIndex += 140;
            //}
            ////IndexRows -= 12;

            ////cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RSTU20201910/20GP", 15, 580, 0);




            //cb.EndText();

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


            BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader4, 13);
            cb.SetColorFill(Color.BLACK);


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

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMS & CONDITIONS:", 25, 375, 0);//right


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorized Signatory ", 25, 75, 0);//

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for BLUE WAVE SHIPPING (M) SDN BHD (1024530U) ", 25, 60, 0);//

            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorized Signature BLUE WAVE SHIPPING (M) SDN BHD (1024530U)", 25, 70, 0);//right

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

            DataTable _dtx = GetCanPrint(Data.BLID.ToString());
            if (_dtx.Rows.Count > 0)
            {

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["Grwt"].ToString() + " KGS", 590, 605, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtx.Rows[0]["Cbm"].ToString(), 25, 575, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "20 X " + _dtx.Rows[0]["GP20"].ToString() + "    40  X  " + _dtx.Rows[0]["GP40"].ToString() + "", 590, 650, 0);
            }

            DataTable dt = GetCanInvoice(Data.BLID.ToString());
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

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["BSCode"].ToString(), 25, 650, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["SCNCode"].ToString(), 25, 605, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dt.Rows[0]["VesselID"].ToString(), 200, 605, 0);

            }
            cb.SetFontAndSize(bfheader1, 9);
            cb.SetColorFill(Color.BLACK);
            int ColumnRows = 760; int RowsColumn = 0;
            RowsColumn = 0;
            string[] ArrayAddress1 = Regex.Split(dt.Rows[0]["CONSIGNEE"].ToString().Trim().ToUpper() + "\r" + dt.Rows[0]["CANReleaseAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
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

            //int IndexRows = 540;
            //int ColumnIndex = 25;
            //int PageWidth = 595;
            //int ColumnWidth = 140;
            //int MaxColumnsPerRow = (PageWidth - ColumnIndex) / ColumnWidth;

            int IndexRows = 540;
            int ColumnIndex = 25;
            int PageWidth = 595;
            int ColumnWidth = 140;
            int MaxColumnsPerRow = (PageWidth - ColumnIndex) / ColumnWidth;
            int RowHeight = 15;
            int ColumnSpacing = 160;

            DataTable dtcntr = GetCanContainerPrint(Data.BLID.ToString());

            for (int i = 0; i < dtcntr.Rows.Count; i++)
            {


                if ((i % MaxColumnsPerRow) == 0 && i != 0)
                {
                    ColumnIndex = 25;
                    IndexRows -= RowHeight;
                }


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtcntr.Rows[i]["CntrDtls"].ToString(), ColumnIndex, IndexRows, 0);
                ColumnIndex += ColumnSpacing;
            }
            cb.Stroke();
            cb.EndText();
            writer.CloseStream = false;
            doc.Close();




            #region  #Email
            var PODV1 = dt.Rows[0]["POD"].ToString().Split('-');
            var POLV1 = dt.Rows[0]["POL"].ToString().Split('-');
            strHTML = "<table border='0' cellpadding='0' cellspacing='0' width='80%' style='font-family:Arial;'>" +
                                 "<tbody>" +
                                 "<tr>" +
                                  "<td style='border:none!important;'>" +
                                 "<table cellpadding='0' cellspacing='0' width='100%' style='margin-bottom:15px;'>" +
                                 "<tr><td style='padding-bottom:17px;'>Dear Valued Customer,</td></tr>" +
                                 "<tr><td style='padding-bottom:17px;'>This is auto-generated message to notify the <b>CARGO ARRIVAL NOTICE</b> for <b>BL#</b> " + dt.Rows[0]["BLNumber"].ToString() + "</td></tr>" +
                                 "<tr style='padding-bottom:20px;'>Attachment for your kind reference</tr>" +
                                 "<tr><td>" +
                                 "<table style='width:100%;' border='0'>" +
                                 "<tr>" +
                                 "<td style='color:#535353;font-weight:600; padding-bottom:16px;'>Port Of Loading</td><td style='padding-bottom:16px;'>" + POLV1[0].ToString() + "</td><td style='color:#535353;font-weight:600;padding-bottom:16px;'>Port Of Discharge</td><td style='padding-bottom:16px;'>" + PODV1[0].ToString() + "</td>" +
                                 "</tr>" +
                                  "<tr>" +
                                 "<td style='color:#535353;font-weight:600;padding-bottom:16px;'>Vessel/Voyage</td><td style='padding-bottom:16px;'>" + dt.Rows[0]["VesVoy"].ToString() + "</td><td style='color:#535353;font-weight:600;padding-bottom:16px;'>ETA</td><td style='padding-bottom:16px;'>" + dt.Rows[0]["ETA"].ToString() + "</td>" +
                                 "</tr>" +
                                 "</table>" +
                                 "</td><tr>" +
                                 "</table>" +
                                 "</td>" +
                                 "<tr>" +
                                 "<tr><td style='font-weight:bold;padding-bottom:16px;'>Note: This is an notification email & please do not respond/reply to this mail</tr>" +
                                 "<tr><td style='font-weight:bold;padding-bottom:10px;'></tr>";
            strHTML += "<tr><td style='font-weight:bold;padding-bottom:20px;'></tr>";
            strHTML += " <tr><td style='padding-bottom:10px;'>Regards,</tr>" +
                                 "<tr><td style='font-weight:bold;padding-bottom:15px;'>ADMIN</tr>" +
                                 "<tr><td style='font-weight:bold;padding-bottom:20px;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + "</tr>" +
                                 "<tr><td><img src='http://" + _dtCom.Rows[0]["Website"].ToString() + "/assets/img/gnllogo.png' style='margin-top:15px; width:100px;' class='img-responsive' /></td></tr>" +
                                 "<tbody>" +
                                 "</table>";



            MailMessage EmailObject = new MailMessage();
            EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());

            var EmailIDTo = Data.To.Split(',');
            for (int y = 0; y < EmailIDTo.Length; y++)
            {
                if (EmailIDTo[y].ToString() != "")
                {
                    EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                }
            }

            var EmailIDCC = Data.CC.Split(',');
            for (int y = 0; y < EmailIDCC.Length; y++)
            {
                if (EmailIDCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDCC[y].ToString()));
                }
            }

            var EmailIDBCC = Data.BCC.Split(',');
            for (int y = 0; y < EmailIDBCC.Length; y++)
            {
                if (EmailIDBCC[y].ToString() != "")
                {
                    EmailObject.CC.Add(new MailAddress(EmailIDBCC[y].ToString()));
                }
            }

            byte[] bytes = memoryStream.ToArray();
            memoryStream.Close();

            //EmailObject.To.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
            EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "CANPDF.pdf"));
            EmailObject.Body = strHTML;
            EmailObject.IsBodyHtml = true;
            EmailObject.Priority = MailPriority.Normal;
            EmailObject.Subject = "CAN: BLNo: " + dt.Rows[0]["BLNumber"].ToString() + "----" + "Vessel/Voyage: " + dt.Rows[0]["VesVoy"].ToString() + "----" + "ETA: " + dt.Rows[0]["ETA"].ToString();
            EmailObject.Priority = MailPriority.Normal;
            SmtpClient SMTPServer = new SmtpClient();
            SMTPServer.UseDefaultCredentials = true;
            SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());

            SMTPServer.Host = "smtp.office365.com";
            SMTPServer.ServicePoint.MaxIdleTime = 1;
            SMTPServer.Port = 587;
            SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
            SMTPServer.EnableSsl = true;
            SMTPServer.Send(EmailObject);

            #endregion 
            ViewList.Add(new MySendingEmailAlert
            {
                AlertMessage = "Email sent successfully"

            });

            return ViewList;
        }


        public DataTable GetNotesClausesBooking()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=266";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCanInvoice(string id)
        {
            //string _Query = " select BLNumber, POL, POD, VesVoy,(select top(1)(select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID = NVO_BOLImpVoyageDetails.VesVoyID) " +
            //                " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as VesVoy,BLVesVoyID, " +
            //                " (select(select top(1) VesselCallSign from NVO_VesselMaster where NVO_VesselMaster.ID = NVO_Voyage.VesselID) " +
            //                " from NVO_Voyage where NVO_Voyage.ID = NVO_BOL.BLVesVoyID) as VesselID, " +
            //                " (select top(1) convert(VARCHAR, ETA, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as ETA, " +
            //                " ImpFreeDays, " +
            //                " (select top(1) (select top(1) (select top(1)TerminalName from  NVO_TerminalMaster where NVO_TerminalMaster.ID=NVO_VoyageRoute.TerminalID)  " +
            //                " from NVO_VoyageRoute  where NVO_VoyageRoute.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails  where " +
            //                " NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID)  as Terminal, " +
            //                " (select top(1) Address from  NVO_CusBranchLocation where CID=NVO_ImpCAN.ReleaseTo) as CANReleaseAddress,  " +
            //                " (select Top(1) CustomerName from NVO_view_CustomerDetails where NVO_view_CustomerDetails.CID = NVO_ImpCAN.ReleaseTo) as CONSIGNEE, " +
            //                " (select top(1)(select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 283 and NVO_VoyageNotesDtls.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) " +
            //                " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as BSCode, " +
            //                 " (select top(1)(select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 282 and NVO_VoyageNotesDtls.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) " +
            //                " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as SCNCode, " +
            //                " (select (select top(1) (select top(1) VesselID from NVO_VesselMaster where NVO_VesselMaster.ID=NVO_Voyage.VesselID) from NVO_Voyage where NVO_Voyage.ID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails " +
            //                " where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) VesselID1 " +
            //                " from NVO_BOL " +
            //                " inner join NVO_Booking on NVO_Booking.ID = NVO_BOL.BkgID " +
            //                " inner join NVO_ImpCAN on NVO_ImpCAN.BLID = NVO_BOL.ID " +
            //                " where NVO_BOL.Id=" + id;

            string _Query = " select BLNumber, RRID, convert(varchar,getdate(), 103) as NOADate,ImpFreeDays, POL, POD, VesVoy,(select top(1)(select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID = NVO_BOLImpVoyageDetails.VesVoyID) " +
                           " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = NVO_BOL.ID) as VesVoy,BLVesVoyID, " +
                           " (select(select top(1) VesselCallSign from NVO_VesselMaster where NVO_VesselMaster.ID = NVO_Voyage.VesselID) " +
                           " from NVO_Voyage where NVO_Voyage.ID = NVO_BOL.BLVesVoyID) as VesselID, " +
                           " (select top(1) convert(varchar, ETA, 103) from NVO_VoyageRoute where VoyageID = NVO_BOL.BLVesVoyID order by RID Desc) as ETA, " +
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
                           " (select(select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_BOLCustomerDetails.PartID) from NVO_BOLCustomerDetails where PartyTypeID = 1 and BLID =NVO_BOL.ID) as Shipper, " +
                           " (select(select top(1) Address from  NVO_CusBranchLocation where CID = NVO_BOLCustomerDetails.PartID) from NVO_BOLCustomerDetails where PartyTypeID = 1 and BLID =NVO_BOL.ID) as ShipperAddress,NVO_ImpCAN.HBLNo " +
                           " from NVO_BOL " +
                           " inner join NVO_Booking on NVO_Booking.ID = NVO_BOL.BkgID " +
                           " left outer join NVO_ImpCAN on NVO_ImpCAN.BLID = NVO_BOL.ID " +
                           " where NVO_BOL.Id=" + id;
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
               " (select top(1) Rate from NVO_view_DailyExRate where AgencyID= " + AgencyID + " and FCurrencyID= NVO_BLCharges.CurrencyID) as ExRate " +
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
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }


        public List<MySendingEmailAlert> EmailData(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            DataTable dt = GetEmailsending(Data);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ViewList.Add(new MySendingEmailAlert
                {
                    EmailIDs = dt.Rows[i]["EmailID"].ToString(),

                });
            }
            return ViewList;
        }

        public List<MyImpCAN> EmailDataTo(MyImpCAN Data)
        {
            List<MyImpCAN> ViewList = new List<MyImpCAN>();
            DataTable dt = GetCANEmailToValues(Data);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ViewList.Add(new MyImpCAN
                {
                    Email = dt.Rows[i]["email"].ToString(),

                });
            }
            return ViewList;
        }

        public DataTable GetCANEmailToValues(MyImpCAN Data)
        {
            string _Query = "select (select  EmailID from NVO_CusBranchLocation where NVO_CusBranchLocation.CID = NVO_ImpCAN.ReleaseTo)as email, *  from NVO_ImpCAN where ReleaseTo=" + Data.ReleaseTo;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRateSheetNotificationTHCandIHCValues(string RID)
        {
            string _Query = "select R.ID,TC.Size,LoadPortTHC.LoadTHCRate, LoadPortTHC.CurrencyCode As Curr1, ExportIHC.CurrencyCode As Curr2,ImportIHC.CurrencyCode As Curr3,DestTHC.CurrencyCode As Curr4,DestTHC.DestTHCRate,ExportIHC.ExportIHCRate," +

              " ImportIHC.ImportIHCRate from  NVO_Ratesheet R Inner join NVO_RatesheetCntrTypes RCT ON RCT.RRID = R.ID " +

              " Inner join NVO_tblCntrTypes TC ON TC.ID = RCT.CntrTypeID OUTER APPLY(Select  isnull((ManifRate),0) as LoadTHCRate,  CM.CurrencyCode  from NVO_RatesheetCharges RSC Inner Join NVO_CurrencyMaster CM on CM.ID = RSC.CurrencyID  where RSC.ChargeCodeID = 4 and R.ID = RSC.RRID AND RSC.TariffTypeID = 136 and ChargeTypeID = 1 and RSC.CntrType = RCT.CntrTypeID ) LoadPortTHC " +

              "  OUTER APPLY(Select   (ManifRate) as DestTHCRate,CM.CurrencyCode from NVO_RatesheetCharges RSC  Inner Join NVO_CurrencyMaster CM on CM.ID = RSC.CurrencyID where RSC.ChargeCodeID = 9 and R.ID = RSC.RRID AND RSC.TariffTypeID = 136 and ChargeTypeID = 1 and RSC.CntrType = RCT.CntrTypeID) DestTHC  " +

             "    OUTER APPLY(Select ( ManifRate) As ExportIHCRate,CM.CurrencyCode from NVO_RatesheetCharges RSC  Inner Join NVO_CurrencyMaster CM on CM.ID = RSC.CurrencyID   where R.ID = RSC.RRID  AND RSC.TariffTypeID = 137 and RSC.PaymentModeID = 18 and ChargeTypeID = 1 and RSC.CntrType = RCT.CntrTypeID) ExportIHC  " +

             "  OUTER APPLY(Select  ( ManifRate ) As ImportIHCRate,CM.CurrencyCode from NVO_RatesheetCharges RSC  Inner Join NVO_CurrencyMaster CM on CM.ID = RSC.CurrencyID where R.ID = RSC.RRID AND RSC.TariffTypeID = 137 and RSC.PaymentModeID = 19 and ChargeTypeID = 1 and RSC.CntrType = RCT.CntrTypeID) ImportIHC WHERE R.ID = " + RID;

            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetEmailsending(MySendingEmailAlert Data)
        {
            string _Query = "select EmailID from NVO_AgencyEmailDtls where AlertTypeID = " + Data.EmailTypes + " and AgencyID=" + Data.AgentID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetSendAutoEmailsForCrosending(MySendingEmailAlert Data)
        {
            string _Query = " Select(SELECT TOP(1) STUFF((SELECT CAST(',' AS VARCHAR(MAX)) + bb.EmailID "+

                " FROM NVO_CusEmailAlerts  bb INNER JOIN NVO_CusBranchLocation cb on cb.CID = NVO_CROMaster.CusID  " +
                " WHERE bb.CustomerID = cb.CustomerID  AND bb.AlertType=69 FOR XML PATH(''), TYPE).value('.', 'VARCHAR(MAX)'), 1, 1, '')) " +
				" AS CustomerEmail,(SELECT TOP(1) STUFF((SELECT CAST(',' AS VARCHAR(MAX)) + bb.EmailID "+
                " FROM NVO_DepotMaster  bb   WHERE bb.ID = NVO_CROMaster.PickDepoID FOR XML PATH(''), " +
                " TYPE).value('.', 'VARCHAR(MAX)'), 1, 1, '')) AS DepoEmail from NVO_CROMaster "+
                " where NVO_CROMaster.ID= " + Data.ID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCustomerEmailsending(string AgentID)
        {
            string _Query = "select EmailID from NVO_AgencyEmailDtls where AlertTypeID = 210 and AgencyID=" + AgentID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetAgencyDetailssending(string AgentID)
        {
            string _Query = "select * from  NVO_AgencyMaster where Id=" + AgentID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetNotesClauses()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=265";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBkgDetailssending(string BkgID)
        {
            string _Query = " select ID, BookingNo, POL, POD,vesvoy, " +
                            " (select top(1) CntrNo from v_MultipleBookingCntrTypes where BKgID = NVO_Booking.ID) as CntrTypes, " +
                            " (select top(1) Commodity from v_MultipleBookingCntrTypes where BKgID = NVO_Booking.ID) as Commodity, " +
                            " (select top(1) convert(varchar, CutDate, 103) from NVO_CROMaster where BkgID = NVO_Booking.ID) as CutDate " +
                            " from NVO_Booking where Id=" + BkgID;
            return Manag.GetViewData(_Query, "");
        }

        //public DataTable GetBkgPDFValus(string BkgId)
        //{
        //    string _Query = " select BookingNo, convert(varchar, BkgDate, 106) as BkgDate,RRID,RRNo,SlotRefNo,BkgPartyID,BkgParty,(select top(1) Address from NVO_CusBranchLocation where CustomerID = BkgPartyID) as CustomerAddress,ShipmentTypeID,ShipmentType,POOID,POO,POLID,	" +
        //                    " POL,FPODID,FPOD,ServiceTypeID,ServiceType,CommodityTypeID,CommodityType,SalesPersonID,SalesPerson,CarrierID,Carrier,VesVoyID,VesVoy,	 " +
        //                    " ShipperID,Shipper,PickUpDepotID,PickUpDepot,ValidTill,PortNtRef,NVO_Booking.Remarks,AgentID,UserID,NVO_Booking.CurrentDate,PODID,POD,PreparedBYID,PreparedBY,CTQ20, " +
        //                    " CTQ40,TSPORT,TSPORTID,DestinationAgent,DestinationAgentID, " +
        //                    " convert(varchar,(select top(1) ETADate from NVO_CROMaster where BkgID = NVO_Booking.id), 103) as ETDDate,convert(varchar, (select top(1) CutDate from NVO_CROMaster where BkgID = NVO_Booking.id), 103) as CutDate,'' as SCNNo," +
        //                    " (select top(1) CustomerName from NVO_CustomerMaster where ID = NVO_Booking.SlotOperatorID) as CarrierName  " +
        //                    " from NVO_Booking " +
        //                    " where NVO_Booking.ID=" + BkgId;
        //    return Manag.GetViewData(_Query, "");
        //}


        public DataTable GetCompanyDetails()
        {

            string _Query = "Select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBkgPDFValus(string BkgId)
        {
            string _Query = " select BookingNo, convert(varchar, BkgDate, 106) as BkgDate,RRID,RRNo,SlotRefNo,BkgPartyID,BkgParty,(select top(1) Address from NVO_CusBranchLocation where CID = BkgPartyID) as CustomerAddress,ShipmentTypeID,ShipmentType,POOID,POO,POLID,	" +
                            " POL,FPODID,FPOD,ServiceTypeID,ServiceType,CommodityTypeID,CommodityType,SalesPersonID,SalesPerson,CarrierID,Carrier,VesVoyID,VesVoy,	 " +
                            " ShipperID,Shipper,PickUpDepotID,PickUpDepot,ValidTill,PortNtRef,NVO_Booking.Remarks,AgentID,UserID,NVO_Booking.CurrentDate,PODID,POD,PreparedBYID,PreparedBY,CTQ20, " +
                            " CTQ40,TSPORT,TSPORTID,DestinationAgent,DestinationAgentID, " +
                            "  convert(varchar,((select top(1) ETA from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID asc)), 103) as ETADate, " +
                            " convert(varchar, ((select top(1) ETD from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID asc)), 103) as ETDDate, " +
                                                        //"  convert(varchar,((select top(1) ETA from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID DESC)), 103) as NextPortETA, " +
                            " case when ( select count(vr.RID) from NVO_VoyageRoute vr where vr.VoyageID = NVO_Booking.VesVoyID) >2 then " +
                            " convert(varchar, ((select top(1) ETA from NVO_VoyageRoute vr inner join NVO_PortMaster pm on pm.MainPortID = vr.PortID " +
                            " where VoyageID = NVO_Booking.VesVoyID and pm.id = NVO_Booking.PODID)), 103) else convert(varchar,((select top(1) ETA from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID DESC)), 103) end as NextPortETA," +
                            " convert(varchar, (select top(1) CutDate from NVO_CROMaster where BkgID = NVO_Booking.id), 103) as CutDate, " +
                            " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_Booking.SlotOperatorID) as CarrierName,  " +
                            " (select top(1)(select top(1) TerminalName from  NVO_TerminalMaster where NVO_TerminalMaster.ID = TerminalID) from  NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID asc) as Terminal, " +
                            " isnull((select top 1 Notes from NVO_VoyageNotesDtls inner join NVO_GeneralMaster GM on GM.ID =  NVO_VoyageNotesDtls.NotesTypeID where VoyageID = NVO_Booking.VesVoyID AND GM.GeneralName = 'SCN No'),'') AS SCNNo, " +
                           " isnull((select top 1 Notes from NVO_VoyageNotesDtls inner join NVO_GeneralMaster GM on GM.ID =  NVO_VoyageNotesDtls.NotesTypeID where VoyageID = NVO_Booking.VesVoyID  AND GM.GeneralName = 'BS CODE'),'') AS BSCODE, " +
                           " isnull((select top 1 Notes from NVO_VoyageNotesDtls inner join NVO_GeneralMaster GM on GM.ID =  NVO_VoyageNotesDtls.NotesTypeID where VoyageID = NVO_Booking.VesVoyID  AND GM.GeneralName = 'VESSEL CLOSING TIME'),'') AS ClosingTime , " +
                           " isnull((select top 1 VM.VesselID from NVO_VesselMaster VM inner  join NVO_Voyage on NVO_Voyage.VesselID = VM.ID   where NVO_Voyage.ID = NVO_Booking.VesVoyID ),'') AS VesselIDValue " +
                            " from NVO_Booking " +
                            " where NVO_Booking.ID=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCROPDFValus(string CROId)
        {
            string _Query = " select BookingNo, BkgParty,convert(varchar, Date, 103) as Date,(select top(1) Address from NVO_CusBranchLocation where CustomerID = BkgPartyID) as Address,ServiceType, Linecode," +
                            " convert(varchar, NVO_CROMaster.ValidTill, 103) as ValidTill,ReleaseOrderNo,POO,POL,POD,FPOD,NVO_Booking.VesVoy,TsPort," +
                            " convert(varchar, ETADate, 103) as ETADate,convert(varchar, ETDDate, 103) as ETDDate,convert(varchar, CutDate, 103) as CUTDate,Shipper,(select top(1) CustomerName from NVO_CustomerMaster where ID = NVO_CROMaster.Surveyor) as SurveyorName," +
                            " (select top(1) Address from NVO_CusBranchLocation where CustomerID = NVO_CROMaster.Surveyor) as SurveyorAddress," +
                            " (select top(1) DepName from NVO_DepotMaster where ID = NVO_CROMaster.PickDepoID) as PickUpDepot," +
                            " (select top(1) DepAddress from NVO_DepotMaster where ID = NVO_CROMaster.PickDepoID) as DepotAddress,NVO_CROMaster.Remarks " +
                            " from NVO_Booking inner join NVO_CROMaster on NVO_CROMaster.BkgID = NVO_Booking.ID where NVO_CROMaster.Id = " + CROId;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCRODetailsPDFValues(string CROId)
        {
            string _Query = "select CROID,Size,ReqQty from NVO_CRODETAILS CRD INNER JOIN NVO_tblCntrTypes CT ON CT.ID = CRD.CntrTypeID where CROID = " + CROId;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetCROSending(string CROId)
        {
            string _Query = " select NVO_Booking.ID, BookingNo,BkgParty, vesvoy,POL, POD,ReleaseOrderNo, (select top(1) CntrNo from v_MultipleBookingCntrTypes where BKgID = NVO_Booking.ID) as CntrTypes, " +
                            " (select top(1) Commodity from v_MultipleBookingCntrTypes where BKgID = NVO_Booking.ID) as Commodity, (select top(1) convert(varchar, CutDate, 103) from NVO_CROMaster where BkgID = NVO_Booking.ID) as CutDate, " +
                            " convert(varchar, NVO_CROMaster.ValidTill, 103) as ValidTill,(select top(1) CustomerName from NVO_CustomerMaster where ID = NVO_CROMaster.Surveyor) as SurveyorName, " +
                            " (select top(1) DepName from NVO_DepotMaster where ID = NVO_CROMaster.PickDepoID) as PickUpDepot, " +
                            " (select top(1) DepAddress from NVO_DepotMaster where ID = NVO_CROMaster.PickDepoID) as DepotAddress " +
                            " from NVO_Booking inner join NVO_CROMaster on NVO_CROMaster.BkgID = NVO_Booking.ID " +
                            " where NVO_CROMaster.Id=" + CROId;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRPDFValus(string RID)
        {
            string _Query = " select NVO_Ratesheet.Id,RatesheetNo,Remarks,OtherRemarks,convert(varchar, ValidTill, 106) as ValidDate,convert(varchar, Date, 106) as CreatedOn,convert(varchar, DtApproved, 106) as DtAppr,RebateAmt,IsRebate,Case When IsRebate=1 then 'POL' when IsRebate=2 then 'POD' else 'N/A' End as Rebatedtls,(select top(1) Status from NVO_RRStatusMaster where Id = NVO_Ratesheet.RSStatus) as Status,  " +
                             " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = BookingPartyID) as Customer,case when ShipmentID= 1  then 'EXPORT' ELSE 'IMPORT' end as ShipmentTypes," +
                             "  (select top(1) Address from NVO_CusBranchLocation where CustomerID = BookingPartyID) as CustomerAddress,(select top(1) UserName from NVO_UserDetails where ID = SalesPersonID) as SalesPerson, " +
                                " (select top(1) UserName from NVO_UserDetails where ID = ApprovedBy) as ApprovedBy, " +
                                "(select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.UserID) as CreatedBy, " +
                               " (select top(1) PortCode from NVO_PortMaster where ID = PortOfLoading) as POLCode,  " +
                             " (select top(1) PortCode from NVO_PortMaster where ID = PlaceofDischargeId) as PODCode,  " +
                             " (select top(1) CityName from NVO_CityMaster where ID = PortOfOrgin) as POOLS,  " +
                             " (select top(1) PortName from NVO_PortMaster where ID = PortOfLoading) as POL,  " +
                             " (select top(1) PortName from NVO_PortMaster where ID = PlaceofDischargeId) as POD,  " +
                             " (select top(1) CityName from NVO_CityMaster where ID = FinalPODId) as FPOD,  " +
                             " (select top(1) PortName from NVO_PortMaster where ID = TranshimentPortID) as TSPort, '14' as Freedays,Remarks, " +
                             " (select top(1)(select top(1) GeneralName from NVO_GeneralMaster where ID = NVO_RatesheetMRG.FreightTerms) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as PePaid, " +
                             " (select top(1)(select top(1) Size from NVO_tblCntrTypes where SizeId = 1 and ID = NVO_RatesheetMRG.CntrTypes) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as Cntr20, " +
                             " (select sum(QuotedAmount) from NVO_RatesheetMRG where CntrTypes = 1 and RRID = NVO_Ratesheet.Id) as MrgRate20, " +
                             " (select top(1)(select top(1) Size from NVO_tblCntrTypes where SizeId = 2 and ID = NVO_RatesheetMRG.CntrTypes) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as Cntr40, " +
                             " (select sum(QuotedAmount) from NVO_RatesheetMRG where CntrTypes = 2 and RRID = NVO_Ratesheet.Id) as MrgRate40, " +
                             " (select sum(SlotAmt) from NVO_RatesheetSLOTDtls where CntrTypes = 1 and Commodity = 3 and RRID = NVO_Ratesheet.Id) as Slot20, " +
                             " (select sum(SlotAmt) from NVO_RatesheetSLOTDtls where CntrTypes = 2 and Commodity = 4 and RRID = NVO_Ratesheet.Id) as Slot40, " +
                             "  (select top(1) (select top(1) CustomerName from NVO_CustomerMaster where ID =slotOperator) from NVO_RatesheetSLOT where NVO_RatesheetSLOT.RRID = NVO_Ratesheet.Id)  as SlotOpt " +
                             " from NVO_Ratesheet where Id = " + RID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetRRPDFCntrTypesValus(string RID)
        {
            string _Query = "select (select top(1) Size from NVO_tblCntrTypes where ID = NVO_RatesheetCntrTypes.CntrTypeID) as CntrTypes, " +
                            " (select top(1) GeneralName from NVO_GeneralMaster where Id = CommodityTypeID) as Commodity, ExFreeday, ImFreeday, VGM,DGClass,RR_Cargo as Cargo " +
                            " from NVO_RatesheetCntrTypes where RRId = " + RID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRFreedaysDtls(string RID)
        {
            string _Query = "select ID,RRID,case when ModeID = 1 then 'COMBINED' WHEN ModeID = 2 THEN 'DETENTION' WHEN ModeID = 3 THEN 'DEMURRAGE' END AS Mode,ExpFreeDays,ImpFreeDays from NVO_Ratesheetmode where RRID = " + RID + " ";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRTariffCharges(string RID, string TariffTypeID)
        {
            string _Query = "Select   ROW_NUMBER() Over (Order by  R.TariffID) As [SNo], R.TariffID as  RCID,CT.Size as CntrSize,CTB.ChgDesc,CM.CurrencyCode,GM.GeneralName As PaymentMode,ReqRate,ManifRate,CustomerRate,RateDiff from NVO_RatesheetCharges R " +
             " Inner Join NVO_tblCntrTypes CT ON CT.ID = R.CntrType Inner Join NVO_ChargeTB CTB ON CTB.ID = R.ChargeCodeID " +
         " Inner Join NVO_CurrencyMaster CM ON CM.ID = R.CurrencyID Inner Join NVO_GeneralMaster GM ON GM.ID = R.PaymentModeID AND GM.SeqNo = 9 where RRId = " + RID + " and TariffTypeID = " + TariffTypeID + "  and ChargeTypeID=1 ";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetRRLocalTariffCharges(string RID, string TariffTypeID, string PaymentMode)
        {
            string _Query = "Select   ROW_NUMBER() Over (Order by  R.TariffID) As [SNo], R.TariffID as  RCID,CT.Size as CntrSize,CTB.ChgDesc,CM.CurrencyCode,GM.GeneralName As PaymentMode,ReqRate,ManifRate,CustomerRate,RateDiff from NVO_RatesheetCharges R " +
             " Inner Join NVO_tblCntrTypes CT ON CT.ID = R.CntrType Inner Join NVO_ChargeTB CTB ON CTB.ID = R.ChargeCodeID " +
         " Inner Join NVO_CurrencyMaster CM ON CM.ID = R.CurrencyID Inner Join NVO_GeneralMaster GM ON GM.ID = R.PaymentModeID AND GM.SeqNo = 9 where RRId = " + RID + " and TariffTypeID = " + TariffTypeID + "  and PaymentModeID = " + PaymentMode + "  and ChargeTypeID=1 ";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetEmailsedingRRView(string RRID)
        {
            string _Query = " select NVO_Ratesheet.Id,RatesheetNo,convert(varchar, ValidTill, 106) as ValidDate,(select top(1) Status from NVO_RRStatusMaster where Id = NVO_Ratesheet.RSStatus) as Status,RSStatus,APRemarks,RERemarks,RJRemarks,  " +
                            " (select top(1) CustomerName from NVO_view_CustomerDetails where NVO_view_CustomerDetails.CID = BookingPartyID) as Customer, " +
                            " (select top(1) UserName from NVO_UserDetails where ID = SalesPersonID) as SalesPerson,(select top(1) UserName from NVO_UserDetails where NVO_UserDetails.ID = NVO_Ratesheet.UserID) as CreatedBy,convert(varchar, Date, 103) as CreateDate, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = PortOfLoading) as Loading, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = PlaceofDischargeId) as Discharge, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = TranshimentPortID) as TSPort," +
                            " (select top(1) CityName from NVO_CityMaster where ID = FinalPODId) as FPOD, " +
                            " (select top(1) ExpFreeDays from  NVO_RatesheetMode WHERE RRID =NVO_Ratesheet.Id ) as Freedays,Remarks, " +
                            " (select top(1)(select top(1) GeneralName from NVO_GeneralMaster where ID = NVO_RatesheetMRG.FreightTerms) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as PePaid, " +
                            " (select top(1)(select top(1) Size from NVO_tblCntrTypes where SizeId = 1 and ID = NVO_RatesheetMRG.CntrTypes) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as Cntr20, " +
                            " (select sum(QuotedAmount) from NVO_RatesheetMRG where CntrTypes = 1 and RRID = NVO_Ratesheet.Id) as MrgRate20, " +
                            " (select top(1)(select top(1) Size from NVO_tblCntrTypes where SizeId = 2 and ID = NVO_RatesheetMRG.CntrTypes) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as Cntr40, " +
                            " (select sum(QuotedAmount) from NVO_RatesheetMRG where CntrTypes = 2 and RRID = NVO_Ratesheet.Id) as MrgRate40, " +
                            " (select sum(SlotAmt) from NVO_RatesheetSLOTDtls where CntrTypes = 1 and Commodity = 3 and RRID = NVO_Ratesheet.Id) as Slot20, " +
                            " (select sum(SlotAmt) from NVO_RatesheetSLOTDtls where CntrTypes = 2 and Commodity = 4 and RRID = NVO_Ratesheet.Id) as Slot40, " +
                            " (select top(1) (select top(1) CustomerName from NVO_CustomerMaster where ID =slotOperator) from NVO_RatesheetSLOT where NVO_RatesheetSLOT.RRID = NVO_Ratesheet.Id)  as SlotOpt,(select top(1) EmailID from NVO_UserDetails where  NVO_UserDetails.ID = NVO_Ratesheet.UserID) as UserEmailID, " +
                            " (select top(1) case when PaymentModeID=18  then 'PREPAID' ELSE 'COLLECT' END from NVO_RatesheetCharges where TariffTypeID= 135 and RRID= NVO_Ratesheet.Id) AS PaymentMode,AgentId,"+
                            " (select top(1) RR_Cargo from NVO_RatesheetCntrTypes where RRID =NVO_Ratesheet.Id) as Cargo "+
                            " from NVO_Ratesheet where Id = " + RRID;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetEmailsedingRRFreedaysView(string RRID)
        {
            string _Query = " Select  " +
                            " (select top(1) description from NVO_RatesheetModeItems where Id = NVO_RatesheetMode.ModeID) as items, " +
                            " ExpFreeDays, ImpFreeDays from NVO_RatesheetMode  where RRID = " + RRID;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetCanPrintNew(string id)
        {
            string _Query = " SELECT ReleaseTo, " +
                            " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_ImpCAN.ReleaseTo) as Release, " +
                            " (select top(1) Address from  NVO_CusBranchLocation where CID = NVO_ImpCAN.ReleaseTo) as CANReleaseAddress, " +
                            " (select top(1) CustomerName from NVO_CustomerMaster where ID = NVO_BOL.CFS) AS CFSName, " +
                            " (select TOP(1) Address from NVO_CusBranchLocation  where NVO_CusBranchLocation.CustomerID = CFS) as CFSAddress, " +
                            " BLNumber,(select top(1)(select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID =NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID =NVO_BOL.ID) as VesVoy, CONVERT(VARCHAR, NVO_ImpCAN.CETA, 103) AS ETADate,"+
                             " (select top(1) PortName from NVO_PortMaster where ID = NVO_BOL.POLID) as POL,(select top(1) PortName from NVO_PortMaster where ID = NVO_BOL.PODID) as POD, NVO_ImpCAN.LineNumber, " +
                            " (select top(1)(select top(1) PkgCode from NVO_CargoPkgMaster where ID = PakgType) from NVO_BOLCntrDetails " +
                            " where NVO_BOLCntrDetails.BLID = NVO_BOL.ID) as Pakage,    " +
                            " (select sum(NoOfPkg) from NVO_BOLCntrDetails where BLID = NVO_BOL.ID) as NoofPak,IGMNo,case when  IGMDate = '01/01/1900' then '' else  convert(varchar,IGMDate, 103) end  as IGMDate, " +
                            " (select sum(GrsWt) from NVO_BOLCntrDetails where BLID = NVO_BOL.ID) as Grwt,                       " +
                            " (select count(CntrID) from NVO_BOLCntrDetails where BLID = NVO_BOL.ID and Size in('20GP', '20FR', '20OT', '20RF', '20HC', '20TK')) as GP20,  " +
                            " (select count(CntrID) from NVO_BOLCntrDetails where BLID = NVO_BOL.ID and Size in ('40GP', '40HC', '40FR', '40OT', '40RF', '40TK')) as GP40 " +
                            " FROM NVO_Booking " +
                            " INNER JOIN NVO_BOL ON NVO_BOL.BkgID = NVO_Booking.ID " +
                            " INNER JOIN NVO_ImpCAN on NVO_ImpCAN.BLID = NVO_BOL.ID " +
                            " WHERE  BLID = " + id;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCanPrint(string id)
        {
            string _Query = "select * from Nvo_V_ImportCanPrint where Id =" + id;

            //string _Query = "select * from V_NVOImpPrintValue where Id =" + id;
            return Manag.GetViewData(_Query, "");
        }
   
        public DataTable GetCanContainerPrint(string id)
        {
            string _Query = " Select(select top(1) CntrNo from NVO_Containers where Id = NVO_BOLCntrDetails.CntrID) + '/ ' + size + '/ ' + SealNo  as CntrDtls from NVO_BOLCntrDetails where BLID =" + id;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRNotificationChargewiseRates(string RID)
        {

            string _Query = "select Id,RatesheetNo,(select top(1) Size + '-' + Type from NVO_tblCntrTypes where ID = CntrTypeID) AS Size,(select top(1) GeneralName from NVO_GeneralMaster where ID = CommodityTypeID) AS Commodity," +

                " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 1),0) as FRT, " +

                " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID  where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 1),'') as FRTCurr," +

              " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 22),0) as BAF, " +

             " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 22),'') as BAFCurr, " +

              " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 27),0) as DGS, " +

             " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 27),0) as DGSCurr," +

             " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 15),0) as ECRS, " +

             " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 15),'') as ECRSCurr ," +

              " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 46),0) as CAF, " +

              " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 46),'') as CAFCurr," +

               " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 35),0) as EWRS, " +
               " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID = CurrencyID where RRID = NVO_Ratesheet.ID and TariffTypeID = 135 and CntrType = NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid = 35),0) as EWRSCurr," +

               " isnull((select top(1) ManifRate from NVO_RatesheetCharges where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 23),0) as LSS," +

               " isnull((select top(1) CurrencyCode from NVO_RatesheetCharges Inner Join NVO_CurrencyMaster CM on CM.ID =CurrencyID where RRID= NVO_Ratesheet.ID and TariffTypeID= 135 and CntrType= NVO_RatesheetCntrTypes.CntrTypeID and chargecodeid= 23),'') as LSSCurr " +
               "  from NVO_Ratesheet inner join NVO_RatesheetCntrTypes on NVO_RatesheetCntrTypes.RRID=NVO_Ratesheet.ID where ID =" + RID;

            return Manag.GetViewData(_Query, "");
        }

        public List<MySendingEmailAlert> GetSendAutoEmailsForCros(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            DataTable dt = GetSendAutoEmailsForCrosending(Data);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ViewList.Add(new MySendingEmailAlert
                {
                    EmailIDs = dt.Rows[i]["CustomerEmail"].ToString(),
                    EmailTypesCC = dt.Rows[i]["DepoEmail"].ToString(),
                });
            }
            return ViewList;
        }
        public List<MySendingEmailAlert> EmailAlertBookingSaveConfirm(MySendingEmailAlert Data)
        {
            MemoryStream memoryStream = new MemoryStream();
            DataTable dtv = GetBkgPDFValus(Data.BkgID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                //#region Header LOGO COMPANY NAME
                ////-------------HEADER-------------------//

                //iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                //tbllogo.Width = 100;
                ////tbllogo.Alignment = Element.ALIGN_LEFT;
                ////tbllogo.Cellpadding = 1;
                //tbllogo.BorderWidth = 0;
                //Cell cell = new Cell();
                //cell.Width = 10;


                ////var img = iTextSharp.text.Image.GetInstance(Path.Combine(HttpListenerContext.Current.Server.MapPath( "~/assets/img/logo.png"));

                //var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));


                //img.Alignment = Element.ALIGN_LEFT;
                //img.ScaleAbsolute(45f, 45f);
                //cell = new Cell(img);
                //cell.BorderWidth = 0;
                //cell.Colspan = 1;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                //tbllogo.AddCell(cell);

                /////--SPACE--//
                //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 3;
                //tbllogo.AddCell(cell);

                /////--SPACE--//

                //cell = new Cell(new Phrase(" Shipping Pvt Ltd", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 2;
                //tbllogo.AddCell(cell);

                //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 3;
                //tbllogo.AddCell(cell);
                /////----/////
                //cell = new Cell(new Phrase("Agent of  Container Lines Pte.Ltd ", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 2;
                //tbllogo.AddCell(cell);

                //cell = new Cell(new Phrase("Booking Confirmation", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                ////cell.Colspan = 3;
                //tbllogo.AddCell(cell);

                //var LogoAddresss = Regex.Split("B 504, , Swapn Nagari, Mulund WestMumbai - 400 080 Maharashtra, India. , ", "\r\n|\r|\n");
                //for (int a = 0; a < LogoAddresss.Length; a++)
                //{
                //    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //    cell.BorderWidth = 0;
                //    tbllogo.Alignment = Element.ALIGN_LEFT;
                //    tbllogo.AddCell(cell);
                //}

                //doc.Add(tbllogo);

                //para = new Paragraph("");
                //doc.Add(para);

                //para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                //para.Alignment = Element.ALIGN_RIGHT;
                //doc.Add(para);

                ////----------SPACE----------------------------------
                //iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                //doc.Add(Tblspace2);

                ////------------------------------------------------------------------------

                //#endregion


                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;

                DataTable dtc = GetAgencyDetails(Data.AgentID);
                if (dtc.Rows.Count > 0)
                {


                    var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                    //var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                    img.Alignment = Element.ALIGN_LEFT;
                    img.ScaleAbsolute(150f, 80f);
                    cell = new Cell(img);
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tbllogo.AddCell(cell);
                }


                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--////

                cell = new Cell(new Phrase(dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);



                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);
                ///----/////

                DataTable dta = GetCompanyDetails();
                //if (dta.Rows.Count > 0)
                //{
                //    if (dta.Rows[0]["CompanyID"].ToString() != "")
                //    {
                //        cell = new Cell(new Phrase("Agent Of " + dta.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //        cell.BorderWidth = 0;
                //        tbllogo.Alignment = Element.ALIGN_LEFT;
                //        //cell.Colspan = 2;
                //        tbllogo.AddCell(cell);
                //    }
                //}
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("Booking Confirmation", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString() + "\r\n" + "GST: " + dtc.Rows[0]["TaxGSTNo"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                }

                doc.Add(tbllogo);

                //para = new Paragraph("");
                //doc.Add(para);

                para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace2);

                //------------------------------------------------------------------------

                #endregion
                #region Bookingparty and Ratesheet details
                //-------------------Bookingparty and Ratesheet details-----------
                PdfContentByte content = pdfWriter.DirectContent;
                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Booking Party", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }
                mtable.AddCell(Tbl1);



                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("BOOKING N0", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BookingNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SHIPMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ShipmentType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SALES", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["SalesPerson"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("SERVICE TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region Location POL POD POO

                PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                TblLocs.SpacingBefore = 10;
                TblLocs.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POO"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                TblLocs.AddCell(cell1);

                doc.Add(TblLocs);



                #endregion

                #region Booking Details
                //----------------- Booking Details--------------//

                iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                Tbl3.Width = 100;
                Tbl3.Alignment = Element.ALIGN_LEFT;
                Tbl3.Cellpadding = 0;
                Tbl3.BorderWidth = 0;

                //Sub Heading
                cell = new Cell(new Phrase("Booking Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl3.AddCell(cell);
                doc.Add(Tbl3);

                #region CntrValues

                PdfPTable TblCntrVal = new PdfPTable(new float[] { 2, 2, 2 });
                TblCntrVal.HorizontalAlignment = Element.ALIGN_LEFT;
                TblCntrVal.SpacingBefore = 10;
                TblCntrVal.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Size", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Quantity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);

                DataTable dtCT = GetBkgCntrValus(Data.BkgID);
                if (dtCT.Rows.Count > 0)
                {
                    for (int i = 0; i < dtCT.Rows.Count; i++)
                    {

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);
                    }
                    doc.Add(TblCntrVal);
                }

                #endregion


                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 1;
                Tbl5.BorderWidth = 0;

                //Caption
                //cell = new Cell(new Phrase("Volume", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 2;
                //Tbl5.AddCell(cell);
                ////Value

                //DataTable dtCT = GetBkgCntrValus(BkgID);
                //if (dtCT.Rows.Count > 0)
                //{
                //    for (int i = 0; i < dtCT.Rows.Count; i++)
                //    {
                //        cell = new Cell(new Phrase(" : " + dtCT.Rows[i]["Size"].ToString() + " * " + dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //    cell.BorderWidth = 0;
                //    cell.Colspan = 2;
                //    Tbl5.AddCell(cell);
                //    }
                //}


                ////Caption
                //cell = new Cell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 3;
                //Tbl5.AddCell(cell);
                ////Value
                //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CommodityType"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 3;
                //Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("ETA / ETD", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dtv.Rows[0]["ETADate"].ToString() + "/ " + dtv.Rows[0]["ETDDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                if (Data.AgentID == "3")
                {
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["ClosingTime"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }
                else
                {
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CutDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }


                //Caption
                cell = new Cell(new Phrase("Next Port ETA", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["NextPortETA"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Loading Terminal", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Terminal"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Box Operator Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                if (Data.AgentID == "3")
                {
                    cell = new Cell(new Phrase(" : " + "BWS", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                }
                else
                {
                    cell = new Cell(new Phrase(" : " + "BWS", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                }

                if (Data.AgentID != "3")
                {
                    //Caption
                    cell = new Cell(new Phrase("Carrier", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CarrierName"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }
                else
                {  //Caption
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);

                }

                //Caption
                cell = new Cell(new Phrase("Shipper", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("VESSEL CLOSING TIME", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["ClosingTime"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("Port Net Reference", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PortNtRef"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("SCN No", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["SCNNo"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("BS CODE", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["BSCODE"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("Vessel ID", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesselIDValue"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                if (Data.AgentID != "3")
                {
                    cell = new Cell(new Phrase("Pick Up Depot", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);

                }
                doc.Add(Tbl5);


                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 1;
                Tbl7.BorderWidth = 0;


                cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);

                DataTable _dtv = GetNotesClausesBooking();
                if (_dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < _dtv.Rows.Count; i++)
                    {
                        cell = new Cell(new Phrase("*" + _dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        Tbl7.AddCell(cell);

                    }
                }



                doc.Add(Tbl7);

                #endregion

                #region FOOTER
                ///---------FOOTER----------------//
                ///
                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("Thank you very much on your booking confirmation with us. & Looking forward to your valuable support for future bookings.", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl8.AddCell(cell);
                doc.Add(Tbl8);

                iTextSharp.text.Table Tbl9 = new iTextSharp.text.Table(1);
                Tbl9.Width = 100;
                Tbl9.Alignment = Element.ALIGN_LEFT;
                Tbl9.Cellpadding = 1;
                Tbl9.BorderWidth = 0;

                cell = new Cell(new Phrase("Best regards,", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);

                cell = new Cell(new Phrase("Customer service team", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);
                doc.Add(Tbl9);



                #endregion

                pdfWriter.CloseStream = false;
                doc.Close();


            }
            byte[] bytes = memoryStream.ToArray();
            memoryStream.Close();
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();
            DataTable _dtx = GetBookingDetailsEmail(Data.BkgID);
            if (_dtx.Rows.Count > 0)
            {
                DataTable _dtCom = GetCompnayDetails();
                DateTime dtDate = Convert.ToDateTime(System.DateTime.Now.Date.ToShortDateString());
                var creation_date = String.Format("{0:dd-MMM-yyyy}", dtDate);
                string strSubHeader = "<td style='font-family:Arial;font-weight: bold;font-size:14px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;'>";
                string strSubHeader1 = "<td colspan='4' style='font-family:Arial;font-weight: bold;font-size:15px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;text-align:center;'>";
                string sHtml = "";
                sHtml = "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:17px; font-weight:bold; text-decoration:underline; color:#FF6600; '>System Generated Message:</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Hi,</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from "+ _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Booking for bllow details</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                sHtml += "<tr>";
                sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >Booking Generation</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + " Online Booking</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader + " Booking Party</td>";
                sHtml += strSubHeader + " Booking No</td>";
                sHtml += strSubHeader + " Booking Date</td>";
                sHtml += strSubHeader + " Rate Request No</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader + _dtx.Rows[0]["BkgParty"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["BookingNo"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["Datev"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["RRNo"].ToString() + "</td>";
                sHtml += "</tr>";
                sHtml += "</table>";

                sHtml += "</table>";
                sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>"+ _dtCom.Rows[0]["EmailHeader"].ToString() + ".</td></tr>";

                sHtml += "</table>";

               // DataTable _dtCom = GetCompnayDetails();
                if (_dtCom.Rows.Count > 0)
                {
                    MailMessage EmailObject = new MailMessage();
                    DataTable dtAuto1 = GetBookingEmailsending(Data.AgentID);
                    if (dtAuto1.Rows.Count > 0)
                    {
                        var EmailID = dtAuto1.Rows[0]["EmailID"].ToString().Split(',');
                        for (int y = 0; y < EmailID.Length; y++)
                        {
                            if (EmailID[y].ToString() != "")
                            {
                                EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                            }
                        }
                        EmailObject.To.Add(new MailAddress(_dtx.Rows[0]["UserEmailID"].ToString()));
                    }
                       
                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    DataTable dtAuto = GetBookingEmailsendingParty(Data.BkgID);
                    if (dtAuto.Rows.Count > 0)
                    {
                        if (dtAuto.Rows[0]["EmailID"].ToString() != "null")
                        {
                            var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                            for (int y = 0; y < EmailID.Length; y++)
                            {
                                if (EmailID[y].ToString() != "")
                                {
                                    EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                                }
                            }
                        }
                        else
                        {
                            ViewList.Add(new MySendingEmailAlert
                            {
                                AlertMessage = "Party EmailId not updated please update"

                            });
                            return ViewList;
                        }
                        //var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                        //for (int y = 0; y < EmailID.Length; y++)
                        //{
                        //    if (EmailID[y].ToString() != "")
                        //    {
                        //        EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                        //    }
                        //}
                        EmailObject.To.Add(new MailAddress(_dtx.Rows[0]["UserEmailID"].ToString()));
                        EmailObject.Bcc.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
                        EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "BookingPDF.pdf"));
                        EmailObject.Body = sHtml;
                        EmailObject.IsBodyHtml = true;
                        EmailObject.Priority = MailPriority.Normal;
                        EmailObject.Subject = "Booking No: " + _dtx.Rows[0]["BookingNo"].ToString();
                        EmailObject.Priority = MailPriority.Normal;
                        SmtpClient SMTPServer = new SmtpClient();
                        SMTPServer.UseDefaultCredentials = true;
                        SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());
                        SMTPServer.Host = "smtp.office365.com";
                        SMTPServer.ServicePoint.MaxIdleTime = 1;
                        SMTPServer.Port = 587;
                        SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                        SMTPServer.EnableSsl = true;
                        SMTPServer.Send(EmailObject);


                    }

                }


            }
            ViewList.Add(new MySendingEmailAlert
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;


        }

        public List<MySendingEmailAlert> EmailAlertBookingSaveConfirmPdf(MySendingEmailAlert Data)
        {
            List<MySendingEmailAlert> ViewList = new List<MySendingEmailAlert>();

            MemoryStream memoryStream = new MemoryStream();
            DataTable dtv = GetBkgPDFValus(Data.BkgID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                //#region Header LOGO COMPANY NAME
                ////-------------HEADER-------------------//

                //iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                //tbllogo.Width = 100;
                ////tbllogo.Alignment = Element.ALIGN_LEFT;
                ////tbllogo.Cellpadding = 1;
                //tbllogo.BorderWidth = 0;
                //Cell cell = new Cell();
                //cell.Width = 10;


                ////var img = iTextSharp.text.Image.GetInstance(Path.Combine(HttpListenerContext.Current.Server.MapPath( "~/assets/img/logo.png"));

                //var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));


                //img.Alignment = Element.ALIGN_LEFT;
                //img.ScaleAbsolute(45f, 45f);
                //cell = new Cell(img);
                //cell.BorderWidth = 0;
                //cell.Colspan = 1;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                //tbllogo.AddCell(cell);

                /////--SPACE--//
                //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 3;
                //tbllogo.AddCell(cell);

                /////--SPACE--//

                //cell = new Cell(new Phrase(" Shipping Pvt Ltd", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 2;
                //tbllogo.AddCell(cell);

                //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 3;
                //tbllogo.AddCell(cell);
                /////----/////
                //cell = new Cell(new Phrase("Agent of  Container Lines Pte.Ltd ", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 2;
                //tbllogo.AddCell(cell);

                //cell = new Cell(new Phrase("Booking Confirmation", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                ////cell.Colspan = 3;
                //tbllogo.AddCell(cell);

                //var LogoAddresss = Regex.Split("B 504, , Swapn Nagari, Mulund WestMumbai - 400 080 Maharashtra, India. , ", "\r\n|\r|\n");
                //for (int a = 0; a < LogoAddresss.Length; a++)
                //{
                //    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                //    cell.BorderWidth = 0;
                //    tbllogo.Alignment = Element.ALIGN_LEFT;
                //    tbllogo.AddCell(cell);
                //}

                //doc.Add(tbllogo);

                //para = new Paragraph("");
                //doc.Add(para);

                //para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                //para.Alignment = Element.ALIGN_RIGHT;
                //doc.Add(para);

                ////----------SPACE----------------------------------
                //iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                //doc.Add(Tblspace2);

                ////------------------------------------------------------------------------

                //#endregion


                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;

                DataTable dtc = GetAgencyDetails(Data.AgentID);
                if (dtc.Rows.Count > 0)
                {


                    var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                    //var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                    img.Alignment = Element.ALIGN_LEFT;
                    img.ScaleAbsolute(150f, 80f);
                    cell = new Cell(img);
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tbllogo.AddCell(cell);
                }


                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--////

                cell = new Cell(new Phrase(dtc.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);



                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);
                ///----/////

                DataTable dta = GetCompanyDetails();
                //if (dta.Rows.Count > 0)
                //{
                //    if (dta.Rows[0]["CompanyID"].ToString() != "")
                //    {
                //        cell = new Cell(new Phrase("Agent Of " + dta.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //        cell.BorderWidth = 0;
                //        tbllogo.Alignment = Element.ALIGN_LEFT;
                //        //cell.Colspan = 2;
                //        tbllogo.AddCell(cell);
                //    }
                //}
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("Booking Confirmation", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString() + "\r\n" + "GST: " + dtc.Rows[0]["TaxGSTNo"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cell = new Cell(new Phrase(LogoAddresss[a].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);
                }

                doc.Add(tbllogo);

                //para = new Paragraph("");
                //doc.Add(para);

                para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace2);

                //------------------------------------------------------------------------

                #endregion
                #region Bookingparty and Ratesheet details
                //-------------------Bookingparty and Ratesheet details-----------
                PdfContentByte content = pdfWriter.DirectContent;
                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Booking Party", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }
                mtable.AddCell(Tbl1);



                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("BOOKING N0", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BookingNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SHIPMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ShipmentType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SALES", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["SalesPerson"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["BkgDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("SERVICE TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region Location POL POD POO

                PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2, 2, 2 });
                TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                TblLocs.SpacingBefore = 10;
                TblLocs.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POO"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                TblLocs.AddCell(cell1);

                doc.Add(TblLocs);



                #endregion

                #region Booking Details
                //----------------- Booking Details--------------//

                iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                Tbl3.Width = 100;
                Tbl3.Alignment = Element.ALIGN_LEFT;
                Tbl3.Cellpadding = 0;
                Tbl3.BorderWidth = 0;

                //Sub Heading
                cell = new Cell(new Phrase("Booking Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl3.AddCell(cell);
                doc.Add(Tbl3);

                #region CntrValues

                PdfPTable TblCntrVal = new PdfPTable(new float[] { 2, 2, 2 });
                TblCntrVal.HorizontalAlignment = Element.ALIGN_LEFT;
                TblCntrVal.SpacingBefore = 10;
                TblCntrVal.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Size", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Quantity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrVal.AddCell(cell1);

                DataTable dtCT = GetBkgCntrValus(Data.BkgID);
                if (dtCT.Rows.Count > 0)
                {
                    for (int i = 0; i < dtCT.Rows.Count; i++)
                    {

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(dtCT.Rows[i]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblCntrVal.AddCell(cell1);
                    }
                    doc.Add(TblCntrVal);
                }

                #endregion


                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 1;
                Tbl5.BorderWidth = 0;

                //Caption
                //cell = new Cell(new Phrase("Volume", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 2;
                //Tbl5.AddCell(cell);
                ////Value

                //DataTable dtCT = GetBkgCntrValus(BkgID);
                //if (dtCT.Rows.Count > 0)
                //{
                //    for (int i = 0; i < dtCT.Rows.Count; i++)
                //    {
                //        cell = new Cell(new Phrase(" : " + dtCT.Rows[i]["Size"].ToString() + " * " + dtCT.Rows[i]["Qty"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //    cell.BorderWidth = 0;
                //    cell.Colspan = 2;
                //    Tbl5.AddCell(cell);
                //    }
                //}


                ////Caption
                //cell = new Cell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 3;
                //Tbl5.AddCell(cell);
                ////Value
                //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CommodityType"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 3;
                //Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("ETA / ETD", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dtv.Rows[0]["ETADate"].ToString() + "/ " + dtv.Rows[0]["ETDDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                if (Data.AgentID == "3")
                {
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["ClosingTime"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }
                else
                {
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CutDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }


                //Caption
                cell = new Cell(new Phrase("Next Port ETA", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["NextPortETA"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Loading Terminal", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Terminal"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("Box Operator Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                if (Data.AgentID == "3")
                {
                    cell = new Cell(new Phrase(" : " + "BWS", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                }
                else
                {
                    cell = new Cell(new Phrase(" : " + "BWS", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                }

                if (Data.AgentID != "3")
                {
                    //Caption
                    cell = new Cell(new Phrase("Carrier", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["CarrierName"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);
                }
                else
                {  //Caption
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 4;
                    Tbl5.AddCell(cell);

                }

                //Caption
                cell = new Cell(new Phrase("Shipper", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("VESSEL CLOSING TIME", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["ClosingTime"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("Port Net Reference", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PortNtRef"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("SCN No", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["SCNNo"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("BS CODE", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["BSCODE"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("Vessel ID", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["VesselIDValue"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                if (Data.AgentID != "3")
                {
                    cell = new Cell(new Phrase("Pick Up Depot", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);
                    //Value
                    cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                    Tbl5.AddCell(cell);

                }
                doc.Add(Tbl5);


                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 1;
                Tbl7.BorderWidth = 0;


                cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);

                DataTable _dtv = GetNotesClausesBooking();
                if (_dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < _dtv.Rows.Count; i++)
                    {
                        cell = new Cell(new Phrase("*" + _dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        Tbl7.AddCell(cell);

                    }
                }



                doc.Add(Tbl7);

                #endregion

                #region FOOTER
                ///---------FOOTER----------------//
                ///
                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("Thank you very much on your booking confirmation with us. & Looking forward to your valuable support for future bookings.", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl8.AddCell(cell);
                doc.Add(Tbl8);

                iTextSharp.text.Table Tbl9 = new iTextSharp.text.Table(1);
                Tbl9.Width = 100;
                Tbl9.Alignment = Element.ALIGN_LEFT;
                Tbl9.Cellpadding = 1;
                Tbl9.BorderWidth = 0;

                cell = new Cell(new Phrase("Best regards,", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);

                cell = new Cell(new Phrase("Customer service team", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);
                doc.Add(Tbl9);



                #endregion

                pdfWriter.CloseStream = false;
                doc.Close();


            }
            DataTable _dtx = GetBookingDetailsEmail(Data.BkgID);
            if (_dtx.Rows.Count > 0)
            {
                DataTable _dtCom = GetCompnayDetails();
                DateTime dtDate = Convert.ToDateTime(System.DateTime.Now.Date.ToShortDateString());
                var creation_date = String.Format("{0:dd-MMM-yyyy}", dtDate);
                string strSubHeader = "<td style='font-family:Arial;font-weight: bold;font-size:14px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;'>";
                string strSubHeader1 = "<td colspan='4' style='font-family:Arial;font-weight: bold;font-size:15px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;text-align:center;'>";
                string sHtml = "";
                sHtml = "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:17px; font-weight:bold; text-decoration:underline; color:#FF6600; '>System Generated Message:</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Hi,</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from "+ _dtCom.Rows[0]["EmailHeader"].ToString()+"</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Booking for bllow details</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                sHtml += "<tr>";
                sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >Booking Confirmation</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + "  Online Booking</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader + " Booking Party</td>";
                sHtml += strSubHeader + " Booking No</td>";
                sHtml += strSubHeader + " Booking Date</td>";
                sHtml += strSubHeader + " Rate Request No</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader + _dtx.Rows[0]["BkgParty"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["BookingNo"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["Datev"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["RRNo"].ToString() + "</td>";
                sHtml += "</tr>";
                sHtml += "</table>";

                sHtml += "</table>";
                sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>"+ _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";

                sHtml += "</table>";

                //DataTable _dtCom = GetCompnayDetails();
                if (_dtCom.Rows.Count > 0)
                {
                    MailMessage EmailObject = new MailMessage();
                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    DataTable dtAuto = GetBookingEmailsendingParty(Data.BkgID);
                    if (dtAuto.Rows.Count > 0)
                    {
                        if (dtAuto.Rows[0]["EmailID"].ToString() != "null")
                        {
                            var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                            for (int y = 0; y < EmailID.Length; y++)
                            {
                                if (EmailID[y].ToString() != "")
                                {
                                    EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                                }
                            }
                        }
                        else
                        {
                            ViewList.Add(new MySendingEmailAlert
                            {
                                AlertMessage = "Party EmailId not updated please update"

                            });
                            return ViewList;
                        }
                        byte[] bytes = memoryStream.ToArray();
                        memoryStream.Close();
                        EmailObject.To.Add(new MailAddress(_dtx.Rows[0]["UserEmailID"].ToString()));
                        EmailObject.Bcc.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
                        //EmailObject.Bcc.Add(new MailAddress("venkat@neridashipping.com"));
                        //EmailObject.Bcc.Add(new MailAddress("ganesh@rmjtech.in"));
                        EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "BookingPDF.pdf"));
                        EmailObject.Body = sHtml;
                        EmailObject.IsBodyHtml = true;
                        EmailObject.Priority = MailPriority.Normal;
                        EmailObject.Subject = "Booking No: " + _dtx.Rows[0]["BookingNo"].ToString();
                        EmailObject.Priority = MailPriority.Normal;
                        SmtpClient SMTPServer = new SmtpClient();
                        SMTPServer.UseDefaultCredentials = true;
                        SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());
                        SMTPServer.Host = "smtp.office365.com";
                        SMTPServer.ServicePoint.MaxIdleTime = 1;
                        SMTPServer.Port = 25;
                        SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                        SMTPServer.EnableSsl = true;
                        SMTPServer.Send(EmailObject);

                    }

                }


            }
            ViewList.Add(new MySendingEmailAlert
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;


        }
        public DataTable GetBkgCntrValus(string BkgId)
        {
            string _Query = " select BKgID,NVO_tblCntrTypes.Size,Qty,GeneralName as Commodity from NVO_BookingCntrTypes inner join NVO_tblCntrTypes on NVO_tblCntrTypes.ID = NVO_BookingCntrTypes.CntrTypes inner join NVO_GeneralMaster on NVO_GeneralMaster.ID = NVO_BookingCntrTypes.CommodityType where NVO_BookingCntrTypes.BkgID=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetBookingEmailsending(string AgentID)
        {
            string _Query = "select EmailID from NVO_AgencyEmailDtls where AlertTypeID = 211 and AgencyID=" + AgentID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBookingDetailsEmail(string BkgId)
        {
            string _Query = " select BookingNo,RRNo,BkgParty,Convert(varchar,BkgDate, 106) as Datev, " +
                            " (select top(1) EmailID from NVO_UserDetails where Id = NVO_Booking.UserID) as UserEmailID from NVO_Booking where Id=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }
     
        public DataTable GetCompnayDetails()
        {
            string _Query = "select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBookingEmailsendingParty(string BkgId)
        {
            string _Query = " select ID,(select top(1) EmailID from NVO_CusEmailAlerts where AlertType = 70 and CustomerID = NVO_Booking.BkgPartyID) as EmailID " +
                            " from NVO_Booking where Id=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }
    }
}