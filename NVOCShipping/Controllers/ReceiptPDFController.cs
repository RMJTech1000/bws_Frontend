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
    public class ReceiptPDFController : Controller
    {
        AccountMaster Manag = new AccountMaster();

        // GET: ReceiptPDF
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ReceiptPDF(string idv, string AgencyID)
        {
            CreateReceiptPDF(idv, AgencyID);
            return View();

        }
        public ActionResult PaymentPDF(string idv, string AgencyID)
        {
            CreatePaymentPDF(idv, AgencyID);
            return View();

        }
        public void CreateReceiptPDF(string idv, string AgencyID)
        {

            // DataTable dtv = GetBkgPDFValus(BkgID);
            //if (dtv.Rows.Count > 0)
            //{
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            //PdfWriter pdfWriter1 = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/BKGPDF\\" + dtv.Rows[0]["BookingNo"].ToString() + ".pdf"), FileMode.Create));
            PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();

            #region Header LOGO COMPANY NAME
            //-------------HEADER-------------------//


            iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(1);
            tbllogo.Width = 50;
            tbllogo.Alignment = Element.ALIGN_LEFT;
            //tbllogo.Cellpadding = 1;
            tbllogo.BorderWidth = 0;
            Cell cell = new Cell();
            cell.Width = 50;

             DataTable dtc = GetAgencyDetails(AgencyID);
            //if (dtc.Rows.Count > 0)
            //{
            //    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
            //    img.Alignment = Element.ALIGN_LEFT;
            //    cell = new Cell(img);
            //    cell.BorderWidth = 0;
            //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //    tbllogo.AddCell(cell);


            //}

            var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
            img.Alignment = Element.ALIGN_LEFT;
            cell = new Cell(img);
            cell.BorderWidth = 0;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbllogo.AddCell(cell);
            doc.Add(tbllogo);

            ///--SPACE--//

            iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(2);
            tbllogo1.Width = 100;
            tbllogo1.Alignment = Element.ALIGN_LEFT;
            tbllogo1.BorderWidth = 0;

            cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLUE)));
            cell.BorderWidth = 0;
            tbllogo1.Alignment = Element.ALIGN_LEFT;
            tbllogo1.AddCell(cell);

            cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLUE)));
            cell.BorderWidth = 0;
            tbllogo1.Alignment = Element.ALIGN_LEFT;
            tbllogo1.AddCell(cell);

  

            doc.Add(tbllogo1);

            para = new Paragraph("");
            doc.Add(para);

            para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLUE));
            para.Alignment = Element.ALIGN_RIGHT;
            doc.Add(para);

            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace2);

            //------------------------------------------------------------------------
            #endregion

            #region Customer and Receipt details
            //-------------------Bookingparty and Ratesheet details-----------
            PdfContentByte content = pdfWriter.DirectContent;
            PdfPTable mtable = new PdfPTable(2);
            mtable.WidthPercentage = 100;
            mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

            DataTable _dtv = GetReceiptDtls(idv);
            if (_dtv.Rows.Count > 0)
            {
                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Customer Name", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["CustomerName"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss1 = Regex.Split(_dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss1.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss1[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }

                mtable.AddCell(Tbl1);


                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("RECEIPT NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["ReceiptNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("RECEIPT DATE ", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["ReceiptDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("BL NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["BLNumber"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("RECEIPT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["ReceiptTypeV"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region NARRATION DETAILS

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace3 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace3);

                //------------------------------------------------------------------------

                PdfPTable Tbl2 = new PdfPTable(1);
                Tbl2.WidthPercentage = 100;
                Tbl2.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                cell1 = new PdfPCell(new Phrase("NARRATION", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 12;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl2.AddCell(cell1);
                doc.Add(Tbl2);

                iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                Tbl3.Width = 100;
                Tbl3.DefaultCell.Border = 0;
                Tbl3.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl3.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase(_dtv.Rows[0]["Remarks"].ToString() , new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl3.AddCell(cell);

                doc.Add(Tbl3);


                iTextSharp.text.Table Tblline = new iTextSharp.text.Table(1);
                Tblline.Width = 100;
                Tblline.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblline.Border = Rectangle.NO_BORDER;
                Tblline.Cellpadding = 1;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                cell.BackgroundColor = new Color(98, 141, 214);
                Tblline.AddCell(cell);
                doc.Add(Tblline);

                #endregion

                #region Cash/Cheque Details 

                //Sub Heading
                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(1);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 0;
                Tbl5.BorderWidth = 0;

                cell = new Cell(new Phrase("Cash/Cheque Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl5.AddCell(cell);
                doc.Add(Tbl5);

                iTextSharp.text.Table TblReceiptDtls = new iTextSharp.text.Table(8);
                TblReceiptDtls.Width = 100;
                TblReceiptDtls.Alignment = Element.ALIGN_LEFT;
                TblReceiptDtls.Cellpadding = 1;
                TblReceiptDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Mode Of Payment", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Bank Name", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Cheque No/UTR No", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Clearance Date", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);


                cell = new Cell(new Phrase("Currency", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Collection Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Local Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);
                //DataTable _dtColl = GetReceiptCollDetails(idv);

                //for (int i = 0; i < _dtColl.Rows.Count; i++)
                //{
                    cell = new Cell(new Phrase(_dtv.Rows[0]["PaymentMade"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan =2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["BankName"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["Reference"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["PaymentDate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["Currency"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.BorderWidth = 0.5f;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["Amount"].ToString(), new Font(Font.HELVETICA,8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["LocalAmount"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    doc.Add(TblReceiptDtls);
                //}


                #endregion

                #region Invoice Details 

                //Sub Heading
                iTextSharp.text.Table Tbl6 = new iTextSharp.text.Table(1);
                Tbl6.Width = 100;
                Tbl6.Alignment = Element.ALIGN_LEFT;
                Tbl6.Cellpadding = 0;
                Tbl6.BorderWidth = 0;


                DataTable _dtn = GetReceiptInvoiceDtls(idv);

                cell = new Cell(new Phrase("Invoice Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl6.AddCell(cell);
                doc.Add(Tbl6);


                iTextSharp.text.Table TblInvoiceDtls = new iTextSharp.text.Table(8);
                TblInvoiceDtls.Width = 100;
                TblInvoiceDtls.Alignment = Element.ALIGN_LEFT;
                TblInvoiceDtls.Cellpadding = 1;
                TblInvoiceDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Document Number", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Doc Date", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
               // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Currency", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Invoice Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Total Received", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
               // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                //cell = new Cell(new Phrase("Due Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                //cell.BackgroundColor = new Color(98, 141, 214);
                //cell.BorderWidth = 0.5f;
                ////cell.Colspan = 1;
                //cell.HorizontalAlignment = Element.ALIGN_CENTER;
                //TblInvoiceDtls.AddCell(cell);

            
                cell = new Cell(new Phrase("TDS Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
               // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("TDS Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);



                for (int i = 0; i < _dtn.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(_dtn.Rows[i]["InvoiceNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtn.Rows[i]["InvoiceDate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["Currency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtn.Rows[i]["InvoiceAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtn.Rows[i]["ReceivedAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                decimal InvAmt = 0;
                decimal RecvdAmt = 0;
                decimal DueAmt = 0;

                InvAmt = decimal.Parse(_dtn.Rows[i]["InvoiceAmt"].ToString());
                RecvdAmt = decimal.Parse(_dtn.Rows[i]["ReceivedAmt"].ToString());
                DueAmt = (InvAmt - RecvdAmt);

               // cell = new Cell(new Phrase(DueAmt.ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
               // cell.BorderWidth = 0.5f;
               //// cell.Colspan = 1;
               // cell.HorizontalAlignment = Element.ALIGN_CENTER;
               // TblInvoiceDtls.AddCell(cell);


                cell = new Cell(new Phrase(_dtn.Rows[i]["TDSType"].ToString(), new Font(Font.HELVETICA,8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
               // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);
                if(_dtn.Rows[i]["TDSAmt"].ToString() == "0.00")
                    {
                        cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0.5f;
                        //cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblInvoiceDtls.AddCell(cell);
                    }
                else
                    {
                        cell = new Cell(new Phrase(_dtn.Rows[i]["TDSAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0.5f;
                        //cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblInvoiceDtls.AddCell(cell);
                    }
               
               }
                doc.Add(TblInvoiceDtls);

                if(_dtv.Rows[0]["RoundOffType"].ToString() != "")
                {

             
                iTextSharp.text.Table TblTransTypeDtls = new iTextSharp.text.Table(4);
                TblTransTypeDtls.Width = 100;
                TblTransTypeDtls.Alignment = Element.ALIGN_LEFT;
                TblTransTypeDtls.Cellpadding = 1;
                TblTransTypeDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Transaction Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblTransTypeDtls.AddCell(cell);

                cell = new Cell(new Phrase("Currency", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblTransTypeDtls.AddCell(cell);

                cell = new Cell(new Phrase("Local Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblTransTypeDtls.AddCell(cell);


                cell = new Cell(new Phrase(_dtv.Rows[0]["RoundOffType"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblTransTypeDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["RFCurrency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblTransTypeDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["ExcessLocalAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblTransTypeDtls.AddCell(cell);

              


                doc.Add(TblTransTypeDtls);

                }

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.DefaultCell.Border = 0;
                Tbl7.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl7.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase(" \n \n \n \n \n \n \n \n \n \n \n \n ", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                Tbl7.AddCell(cell);
                doc.Add(Tbl7);

                #endregion

                #region Footer

                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(2);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("Receipt Prepared By :" + _dtv.Rows[0]["CreatedBy"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);



                cell = new Cell(new Phrase(" Prepared On  : " + _dtv.Rows[0]["ReceiptDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl8.AddCell(cell);

                doc.Add(Tbl8);


                iTextSharp.text.Table Tblline1 = new iTextSharp.text.Table(1);
                Tblline1.Width = 100;
                Tblline1.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblline1.Border = Rectangle.NO_BORDER;
                Tblline1.Cellpadding = 1;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                cell.BackgroundColor = new Color(98, 141, 214);
                Tblline1.AddCell(cell);
                doc.Add(Tblline1);


                iTextSharp.text.Table Tblfoot = new iTextSharp.text.Table(1);
                Tblfoot.Width = 100;
                Tblfoot.Alignment = Element.ALIGN_CENTER;
                Tblfoot.DefaultCell.Border = 0;
                Tblfoot.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblfoot.Border = Rectangle.NO_BORDER;



                cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 9, Font.BOLD, Color.RED)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblfoot.AddCell(cell);
                doc.Add(Tblfoot);

                #endregion



                pdfWriter.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=Receipt.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }

        }

        public void CreatePaymentPDF(string idv, string AgencyID)
        {

            // DataTable dtv = GetBkgPDFValus(BkgID);
            //if (dtv.Rows.Count > 0)
            //{
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            //PdfWriter pdfWriter1 = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/BKGPDF\\" + dtv.Rows[0]["BookingNo"].ToString() + ".pdf"), FileMode.Create));
            PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();

            #region Header LOGO COMPANY NAME
            //-------------HEADER-------------------//


            iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(1);
            tbllogo.Width = 50;
            tbllogo.Alignment = Element.ALIGN_LEFT;
            //tbllogo.Cellpadding = 1;
            tbllogo.BorderWidth = 0;
            Cell cell = new Cell();
            cell.Width = 50;

            DataTable dtc = GetAgencyDetails(AgencyID);
            if (dtc.Rows.Count > 0)
            {
                //if (AgencyID == "13")
                //{
                //    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaaddress1.png"));
                //    img.Alignment = Element.ALIGN_LEFT;
                //    cell = new Cell(img);
                //    cell.BorderWidth = 0;
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    tbllogo.AddCell(cell);


                //}
                //if (AgencyID == "14")
                //{
                //    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddressmundra.png"));
                //    img.Alignment = Element.ALIGN_LEFT;
                //    cell = new Cell(img);
                //    cell.BorderWidth = 0;
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    tbllogo.AddCell(cell);

                //}
                //if (AgencyID == "15")
                //{
                //    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/invaddressdelhi.png"));
                //    img.Alignment = Element.ALIGN_LEFT;
                //    cell = new Cell(img);
                //    cell.BorderWidth = 0;
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    tbllogo.AddCell(cell);

                //}
                //if (AgencyID == "16")
                //{

                //    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaaddress1.png"));
                //    img.Alignment = Element.ALIGN_LEFT;
                //    cell = new Cell(img);
                //    cell.BorderWidth = 0;
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    tbllogo.AddCell(cell);


                //}

                var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                img.Alignment = Element.ALIGN_LEFT;
                img.ScaleAbsolute(150f, 80f);
                cell = new Cell(img);
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                tbllogo.AddCell(cell);

            }



            doc.Add(tbllogo);

            ///--SPACE--//

            iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(2);
            tbllogo1.Width = 100;
            tbllogo1.Alignment = Element.ALIGN_LEFT;
            tbllogo1.BorderWidth = 0;

            cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLUE)));
            cell.BorderWidth = 0;
            tbllogo1.Alignment = Element.ALIGN_LEFT;
            tbllogo1.AddCell(cell);

            cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLUE)));
            cell.BorderWidth = 0;
            tbllogo1.Alignment = Element.ALIGN_LEFT;
            tbllogo1.AddCell(cell);

            //cell = new Cell(new Phrase("Nerida Shipping Pvt Ltd", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
            //cell.BorderWidth = 0;
            //tbllogo1.Alignment = Element.ALIGN_LEFT;
            //tbllogo1.AddCell(cell);

            //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
            //cell.BorderWidth = 0;
            //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            //tbllogo1.AddCell(cell);
            //DataTable dta = GetCompanyDetails();
            //if (dta.Rows.Count > 0)
            //{
            //    cell = new Cell(new Phrase(dta.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    tbllogo1.Alignment = Element.ALIGN_LEFT;
            //    cell.Colspan = 2;
            //    tbllogo1.AddCell(cell);
            //}





            //var Addresss = Regex.Split("B 504, Willows Twin Tower CHS LTD, Swapn Nagari, Mulund WestMumbai - 400 080 Maharashtra, India. , GST: 27AAGCN5719A1ZN", "\r\n|\r|\n");
            //for (int a = 0; a < Addresss.Length; a++)
            //{
            //    cell = new Cell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
            //    cell.BorderWidth = 0;
            //    tbllogo1.AddCell(cell);
            //}
            //cell = new Cell(new Phrase("Receipt", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
            //cell.BorderWidth = 0;
            //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            //tbllogo1.AddCell(cell);

            //cell = new Cell(new Phrase(DS.Tables[0].Rows[0]["Location"].ToString() + " - " + DS.Tables[0].Rows[0]["Pincode"].ToString() + " Tel # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Phone1"].ToString() + "   Fax # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Fax"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD)));
            //cell.BorderWidth = 0;
            //cell.Colspan = 6;
            // tbllogo.AddCell(cell);

            doc.Add(tbllogo1);

            para = new Paragraph("");
            doc.Add(para);

            para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLUE));
            para.Alignment = Element.ALIGN_RIGHT;
            doc.Add(para);

            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace2);

            //------------------------------------------------------------------------
            #endregion

            #region Customer and Receipt details
            //-------------------Bookingparty and Ratesheet details-----------
            PdfContentByte content = pdfWriter.DirectContent;
            PdfPTable mtable = new PdfPTable(2);
            mtable.WidthPercentage = 100;
            mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

            DataTable _dtv = GetPaymentDtls(idv);
            if (_dtv.Rows.Count > 0)
            {
                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Customer Name", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["CustomerName"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss1 = Regex.Split(_dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss1.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss1[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl1.AddCell(cell1);
                }

                mtable.AddCell(Tbl1);


                Tbl1 = new PdfPTable(2);
                Tbl1.WidthPercentage = 50;
                Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("PAYMENT NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["ReceiptNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("PAYMENT DATE ", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["ReceiptDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("BL NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["BLNumber"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("PAYMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(_dtv.Rows[0]["ReceiptTypeV"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region NARRATION DETAILS

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace3 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace3);

                //------------------------------------------------------------------------

                PdfPTable Tbl2 = new PdfPTable(1);
                Tbl2.WidthPercentage = 100;
                Tbl2.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                cell1 = new PdfPCell(new Phrase("NARRATION", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 12;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl2.AddCell(cell1);
                doc.Add(Tbl2);

                iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                Tbl3.Width = 100;
                Tbl3.DefaultCell.Border = 0;
                Tbl3.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl3.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase(_dtv.Rows[0]["Remarks"].ToString() , new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl3.AddCell(cell);

                doc.Add(Tbl3);


                iTextSharp.text.Table Tblline = new iTextSharp.text.Table(1);
                Tblline.Width = 100;
                Tblline.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblline.Border = Rectangle.NO_BORDER;
                Tblline.Cellpadding = 1;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                cell.BackgroundColor = new Color(98, 141, 214);
                Tblline.AddCell(cell);
                doc.Add(Tblline);

                #endregion

                #region Cash/Cheque Details 

                //Sub Heading
                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(1);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 0;
                Tbl5.BorderWidth = 0;

                cell = new Cell(new Phrase("Cash/Cheque Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl5.AddCell(cell);
                doc.Add(Tbl5);

                iTextSharp.text.Table TblReceiptDtls = new iTextSharp.text.Table(8);
                TblReceiptDtls.Width = 100;
                TblReceiptDtls.Alignment = Element.ALIGN_LEFT;
                TblReceiptDtls.Cellpadding = 1;
                TblReceiptDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Mode Of Payment", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Bank Name", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("ChequeNo/UTR No", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Clearance Date", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);


                cell = new Cell(new Phrase("Currency", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Collection Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Local Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);
                //DataTable _dtColl = GetReceiptCollDetails(idv);

                //for (int i = 0; i < _dtColl.Rows.Count; i++)
                //{
                cell = new Cell(new Phrase(_dtv.Rows[0]["PaymentMade"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["BankName"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["Reference"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["PaymentDate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["Currency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 1;
                cell.BorderWidth = 0.5f;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["Amount"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase(_dtv.Rows[0]["LocalAmount"].ToString(), new Font(Font.HELVETICA,8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                doc.Add(TblReceiptDtls);
                //}


                #endregion

                #region Invoice Details 

                //Sub Heading
                iTextSharp.text.Table Tbl6 = new iTextSharp.text.Table(1);
                Tbl6.Width = 100;
                Tbl6.Alignment = Element.ALIGN_LEFT;
                Tbl6.Cellpadding = 0;
                Tbl6.BorderWidth = 0;


                DataTable _dtn = GetPaymentInvoiceDtls(idv);

                cell = new Cell(new Phrase("Invoice Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl6.AddCell(cell);
                doc.Add(Tbl6);


                iTextSharp.text.Table TblInvoiceDtls = new iTextSharp.text.Table(9);
                TblInvoiceDtls.Width = 100;
                TblInvoiceDtls.Alignment = Element.ALIGN_LEFT;
                TblInvoiceDtls.Cellpadding = 1;
                TblInvoiceDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Document Number", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Doc Date", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Currency", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Invoice Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Total Received", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Due Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);


                cell = new Cell(new Phrase("TDS Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                // cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("TDS Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);



                for (int i = 0; i < _dtn.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(_dtn.Rows[i]["InvoiceNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtn.Rows[i]["InvoiceDate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtn.Rows[i]["Currency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtn.Rows[i]["InvoiceAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtn.Rows[i]["ReceivedAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    decimal InvAmt = 0;
                    decimal RecvdAmt = 0;
                    decimal DueAmt = 0;

                    InvAmt = decimal.Parse(_dtn.Rows[i]["InvoiceAmt"].ToString());
                    RecvdAmt = decimal.Parse(_dtn.Rows[i]["ReceivedAmt"].ToString());
                    DueAmt = (InvAmt - RecvdAmt);

                    cell = new Cell(new Phrase(DueAmt.ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    // cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);


                    cell = new Cell(new Phrase(_dtn.Rows[i]["TDSType"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    // cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);
                    if (_dtn.Rows[i]["TDSAmt"].ToString() == "0.00")
                    {
                        cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0.5f;
                        //cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblInvoiceDtls.AddCell(cell);
                    }
                    else
                    {
                        cell = new Cell(new Phrase(_dtn.Rows[i]["TDSAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0.5f;
                        //cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblInvoiceDtls.AddCell(cell);
                    }

                }
                doc.Add(TblInvoiceDtls);

                if (_dtv.Rows[0]["RoundOffType"].ToString() != "")
                {


                    iTextSharp.text.Table TblTransTypeDtls = new iTextSharp.text.Table(4);
                    TblTransTypeDtls.Width = 100;
                    TblTransTypeDtls.Alignment = Element.ALIGN_LEFT;
                    TblTransTypeDtls.Cellpadding = 1;
                    TblTransTypeDtls.BorderWidth = 0.5f;

                    cell = new Cell(new Phrase("Transaction Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                    cell.BackgroundColor = new Color(98, 141, 214);
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblTransTypeDtls.AddCell(cell);

                    cell = new Cell(new Phrase("Currency", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                    cell.BackgroundColor = new Color(98, 141, 214);
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblTransTypeDtls.AddCell(cell);

                    cell = new Cell(new Phrase("Local Amount", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                    cell.BackgroundColor = new Color(98, 141, 214);
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblTransTypeDtls.AddCell(cell);


                    cell = new Cell(new Phrase(_dtv.Rows[0]["RoundOffType"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblTransTypeDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["RFCurrency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblTransTypeDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtv.Rows[0]["ExcessLocalAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblTransTypeDtls.AddCell(cell);




                    doc.Add(TblTransTypeDtls);

                }

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.DefaultCell.Border = 0;
                Tbl7.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl7.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase(" \n \n \n \n \n \n \n \n \n \n \n \n ", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                Tbl7.AddCell(cell);
                doc.Add(Tbl7);

                #endregion

                #region Footer

                //Sub Heading
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(2);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.Cellpadding = 0;
                Tbl8.BorderWidth = 0;


                cell = new Cell(new Phrase("Receipt Prepared By :" + _dtv.Rows[0]["CreatedBy"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);



                cell = new Cell(new Phrase(" Prepared On  : " + _dtv.Rows[0]["ReceiptDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl8.AddCell(cell);

                doc.Add(Tbl8);


                iTextSharp.text.Table Tblline1 = new iTextSharp.text.Table(1);
                Tblline1.Width = 100;
                Tblline1.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblline1.Border = Rectangle.NO_BORDER;
                Tblline1.Cellpadding = 1;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                cell.BackgroundColor = new Color(98, 141, 214);
                Tblline1.AddCell(cell);
                doc.Add(Tblline1);


                iTextSharp.text.Table Tblfoot = new iTextSharp.text.Table(1);
                Tblfoot.Width = 100;
                Tblfoot.Alignment = Element.ALIGN_CENTER;
                Tblfoot.DefaultCell.Border = 0;
                Tblfoot.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblfoot.Border = Rectangle.NO_BORDER;



                cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 9, Font.BOLD, Color.RED)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblfoot.AddCell(cell);
                doc.Add(Tblfoot);

                #endregion



                pdfWriter.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=Receipt.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }

        }
        public DataTable GetReceiptDtls(string idv)
        {
            string _Query = "select Distinct ID,ReceiptNo,convert(varchar,DtReceipt,103)as ReceiptDate, " +
                            " case when ReceiptTypes = 193 then 'On Account' else case when ReceiptTypes = 194 then 'Bill Payment' else " +
                            " case when ReceiptTypes = 195 then 'Un Deposit Account' end end end as ReceiptTypeV, " +
                            "  (select top 1  upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where NVO_CusBranchLocation.CID =NVO_Receipts.PartyID ) as CustomerName, " +
                            " (select Address from NVO_CusBranchLocation where NVO_CusBranchLocation.CID = NVO_Receipts.PartyID)as CustomerAddress, " +
                            " (select GeneralName From NVO_GeneralMaster where NVO_GeneralMaster.ID = NVO_Receipts.PaymentTypes AND NVO_GeneralMaster.SeqNo=59 )as PaymentMade,Reference,(select Top(1) BankName from NVO_FinBankMaster where NVO_FinBankMaster.ID = NVO_Receipts.Bank)as BankName, " +
                            " convert(varchar, PaymentDate, 103) as PaymentDate, " +
                            " (select CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_Receipts.Currency)as Currency, " +
                            " Amount,LocalAmount,Remarks, " +
                            " (select top(1)BLNumber from NVO_v_ReceiptPrintBLNumberFinal WHERE ReceiptID =NVO_Receipts.ID) as BLNumber,(select top 1 UserName from NVO_UserDetails where ID=NVO_Receipts.UserID) as  Createdby ,convert(varchar, CreatedOn, 103) as CreatedOn, " +
                             " (select top(1) GeneralName from NVO_GeneralMaster where Id = RoundOffTypeID) as RoundOffType, ExcessLocalAmt,(select CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_Receipts.RoundOffCurrency)as RFCurrency " +
                            " from NVO_Receipts where ID = " + idv;
            return Manag.GetViewData(_Query, "");


        }
        //public DataTable GetReceiptCollDetails(string idv)
        //{
        //    string _Query = "select ReceiptID,PaymentType,UTRReference,(select Top(1) BankName from NVO_FinBankMaster where NVO_FinBankMaster.ID = NVO_Receiptdtls.BankID)as BankName, " +
        //                    " convert(varchar, ClearanceDate, 103) as PaymentDate, " +
        //                    " (select CurrencyName from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_Receiptdtls.CurrencyID) as Currency, " +
        //                    " CollectionAmt,LocalAmount " +
        //                    " from NVO_Receiptdtls where ReceiptID = " + idv;
        //    return Manag.GetViewData(_Query, "");


        //}
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCompanyDetails()
        {

            string _Query = "Select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetReceiptInvoiceDtls(string idv)
        {
            string _Query = "select NVO_ReceiptBL.ReceiptID, (select Top(1) FinalInvoice from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId) as InvoiceNo, " +
                            " (select Top(1) Convert(varchar, InvDate, 103) from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId) as InvoiceDate, " +
                            " (select Top(1) InvTotal from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId and NVO_InvoiceCusBilling.InvTypes = 1) - isnull((select sum(InvTotal)  from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.InvTypes = 2 and NVO_InvoiceCusBilling.DRID = NVO_ReceiptBL.InvCusBillingId ),0) as InvoiceAmt, " +
                            " Amount as ReceivedAmt, " +
                            " (select Top(1)(select CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_InvoiceCusBilling.CurrencyID) " +
                            " from NVO_InvoiceCusBilling)as Currency, " +
                            " (select Top(1) TDSAmt from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId) as TDSAmt,  isnull((select TOP 1 (GLCode + '-' + GLDesc) from NVO_GLMaster WHERE ID = TDS),'') as TDStype ,* " +
                            "  from NVO_ReceiptBL where ReceiptID = " + idv;
            return Manag.GetViewData(_Query, "");


        }

        public DataTable GetPaymentDtls(string idv)
        {
            string _Query = "select Distinct ID,ReceiptNo,convert(varchar,DtReceipt,103)as ReceiptDate, " +
                            " case when ReceiptTypes = 193 then 'On Account' else case when ReceiptTypes = 194 then 'Bill Payment' else " +
                            " case when ReceiptTypes = 195 then 'Un Deposit Account' end end end as ReceiptTypeV, " +
                            "  (select top 1  upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where NVO_CusBranchLocation.CID =NVO_Payments.PartyID ) as CustomerName, " +
                            " (select Address from NVO_CusBranchLocation where NVO_CusBranchLocation.CID = NVO_Payments.PartyID)as CustomerAddress, " +
                            " (select GeneralName From NVO_GeneralMaster where NVO_GeneralMaster.ID = NVO_Payments.PaymentTypes AND NVO_GeneralMaster.SeqNo=59 ) as PaymentMade,Reference,(select Top(1) BankName from NVO_FinBankMaster where NVO_FinBankMaster.ID = NVO_Payments.Bank)as BankName, " +
                            " convert(varchar, PaymentDate, 103) as PaymentDate, " +
                            " (select CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_Payments.Currency)as Currency, " +
                            " Amount,LocalAmount,Remarks, " +
                            " (select top(1)BLNumber from NVO_v_ReceiptPrintBLNumberFinal WHERE ReceiptID =NVO_Payments.ID) as BLNumber,(select top 1 UserName from NVO_UserDetails where ID=NVO_Payments.UserID) as  Createdby ,convert(varchar, CreatedOn, 103) as CreatedOn, " +
                            " (select top(1) GeneralName from NVO_GeneralMaster where Id = RoundOffTypeID) as RoundOffType, ExcessLocalAmt "+
                            " from NVO_Payments where ID = " + idv;
            return Manag.GetViewData(_Query, "");


        }

        public DataTable GetPaymentInvoiceDtls(string idv)
        {
            string _Query = "select NVO_PaymentBL.ReceiptID, (select Top(1) FinalInvoice from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_PaymentBL.InvCusBillingId) as InvoiceNo, " +
                            " (select Top(1) Convert(varchar, InvDate, 103) from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_PaymentBL.InvCusBillingId) as InvoiceDate, " +
                            " (select Top(1) InvTotal from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_PaymentBL.InvCusBillingId) as InvoiceAmt, " +
                            " (select Top(1) Amount from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_PaymentBL.InvCusBillingId) as ReceivedAmt, " +
                            " (select Top(1)(select CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_InvoiceCusBilling.CurrencyID) " +
                            " from NVO_InvoiceCusBilling)as Currency, " +
                            " (select Top(1) TDSAmt from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_PaymentBL.InvCusBillingId) as TDSAmt, " +
                            " isnull((select TOP 1(GLCode + '-' + GLDesc) from NVO_GLMaster WHERE ID = TDS),'') as TDStype ," +
                            "  * from NVO_PaymentBL where ReceiptID = " + idv;
            return Manag.GetViewData(_Query, "");


        }

    }
}
