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
    public class ReceivableMatchingController : Controller
    {
        AccountMaster Manag = new AccountMaster();
        // GET: ReceivableMatching
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ReceivableMatching(string RecID, string AgencyID)
        {
            CreateRecepMatchingPDF(RecID, AgencyID);
            return View();

        }
        public void CreateRecepMatchingPDF(string RecID, string AgencyID)
        {

            DataTable dtv = GetReceiptMatchDtls(RecID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                //PdfWriter pdfWriter1 = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/BKGPDF\\" + dtv.Rows[0]["BookingNo"].ToString() + ".pdf"), FileMode.Create));
                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//


                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
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
                    cell = new Cell(img);
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell);

                }

                cell = new Cell(new Phrase("RECEIVABLE MATCHING", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
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


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Party Name", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.FixedHeight = 23f;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss1 = Regex.Split(dtv.Rows[0]["StrAccount"].ToString(), "\r\n|\r|\n");
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


                cell1 = new PdfPCell(new Phrase("MATCHING NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["Reference"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("MATCHING DATE ", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["VoucherDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion



                #region Cash/Cheque Details 

                //Sub Heading
                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(1);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 0;
                Tbl5.BorderWidth = 0;

                cell = new Cell(new Phrase("DEBIT TRANSACTION:", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl5.AddCell(cell);
                doc.Add(Tbl5);

                iTextSharp.text.Table TblReceiptDtls = new iTextSharp.text.Table(5);
                TblReceiptDtls.Width = 100;
                TblReceiptDtls.Alignment = Element.ALIGN_LEFT;
                TblReceiptDtls.Cellpadding = 1;
                TblReceiptDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Voucher Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);

                cell = new Cell(new Phrase("Document Name", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
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

                cell = new Cell(new Phrase("Matched Amt", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblReceiptDtls.AddCell(cell);



                DataTable _dtColl = GetReceiptDRAmountDetails(RecID);

                for (int i = 0; i < _dtColl.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(_dtColl.Rows[i]["Invoice"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtColl.Rows[i]["VoucherNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtColl.Rows[i]["Currency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtColl.Rows[i]["AdjustmentAmt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0.5f;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblReceiptDtls.AddCell(cell);
                   
                }
                doc.Add(TblReceiptDtls);

                #endregion

                #region Invoice Details 

                //Sub Heading
                iTextSharp.text.Table Tbl6 = new iTextSharp.text.Table(1);
                Tbl6.Width = 100;
                Tbl6.Alignment = Element.ALIGN_LEFT;
                Tbl6.Cellpadding = 0;
                Tbl6.BorderWidth = 0;




                cell = new Cell(new Phrase("CREDIT TRANSACTION:", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl6.AddCell(cell);
                doc.Add(Tbl6);


                iTextSharp.text.Table TblInvoiceDtls = new iTextSharp.text.Table(5);
                TblInvoiceDtls.Width = 100;
                TblInvoiceDtls.Alignment = Element.ALIGN_LEFT;
                TblInvoiceDtls.Cellpadding = 1;
                TblInvoiceDtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("Voucher Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                cell = new Cell(new Phrase("Document Name", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
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

                cell = new Cell(new Phrase("Matched Amt", new Font(Font.HELVETICA, 9, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 0.5f;
                //cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblInvoiceDtls.AddCell(cell);

                DataTable _dtCr = GetReceiptCRAmountDetails(RecID);
                for (int j = 0; j < _dtCr.Rows.Count; j++)
                {
                    cell = new Cell(new Phrase(_dtCr.Rows[j]["Invoice"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));

                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtCr.Rows[j]["VoucherNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));

                    cell.BorderWidth = 0.5f;
                    // cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtCr.Rows[j]["Currency"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));

                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtCr.Rows[j]["AdjAmount"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));

                    cell.BorderWidth = 0.5f;
                    //cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblInvoiceDtls.AddCell(cell);

                }

                doc.Add(TblInvoiceDtls);





                #endregion
                //Sub Heading
                iTextSharp.text.Table Tbl10 = new iTextSharp.text.Table(2);
                Tbl10.Width = 100;
                Tbl10.Alignment = Element.ALIGN_LEFT;
                Tbl10.Cellpadding = 0;
                Tbl10.BorderWidth = 0;

                cell = new Cell(new Phrase("SUMMARY:", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl10.AddCell(cell);
                doc.Add(Tbl10);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace5 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace5);

                //------------------------------------------------------------------------
                PdfPTable sumtable = new PdfPTable(2);
                sumtable.WidthPercentage = 100;
                sumtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                PdfPTable Tbl11 = new PdfPTable(2);
                Tbl11.WidthPercentage = 100;
                //Tbl11.HorizontalAlignment = Element.ALIGN_RIGHT;


                cell1 = new PdfPCell(new Phrase("TOTAL DR", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl11.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TotalDR"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl11.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("TOTAL CR", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl11.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["TotalCR"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl11.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("DIFFERENCE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl11.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(((decimal.Parse(dtv.Rows[0]["TotalDR"].ToString())) - (decimal.Parse(dtv.Rows[0]["TotalCR"].ToString()))).ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl11.AddCell(cell1);
                doc.Add(Tbl11);
                sumtable.AddCell(Tbl11);
                #region Invoice Details 



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


                cell = new Cell(new Phrase("Matched By :" + dtv.Rows[0]["UserName"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);



                cell = new Cell(new Phrase(" Prepared On  : " + dtv.Rows[0]["PreparedOn"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
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
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetReceiptMatchDtls(string idv)
        {
            string _Query = "  select Reference,CONVERT(VARCHAR, VoucherDate, 103) as VoucherDate,StrAccount,TotalDR,TotalCR,(select top(1) UserName from NVO_UserDetails where NVO_UserDetails.ID = NVO_financeInvoiceSetOff.UserID) as UserName, convert(varchar, PreparedOn, 103) as PreparedOn from NVO_financeInvoiceSetOff where ID=" + idv;
            return Manag.GetViewData(_Query, "");


        }

        public DataTable GetReceiptDRAmountDetails(string RecID)
        {
            string _Query = " select 'Invoice' as Invoice,VoucherNo,(select top(1)(select top(1) CurrencyCode from NVO_CurrencyMaster where ID = NVO_InvoiceCusBilling.CurrencyID) " +
                            " from NVO_InvoiceCusBilling where Id = NVO_financeInvoiceSetOffDR.VoucherID) as Currency,AdjustmentAmt " +
                            " from NVO_financeInvoiceSetOffDR where InvId =" + RecID;
            return Manag.GetViewData(_Query, "");


        }

        public DataTable GetReceiptCRAmountDetails(string RecID)
        {
            string _Query = " select 'Invoice' as Invoice,VoucherNo,(select top(1)(select top(1) CurrencyCode from NVO_CurrencyMaster where ID = NVO_InvoiceCusBilling.CurrencyID) " +
                            " from NVO_InvoiceCusBilling where Id = NVO_financeInvoiceSetOffCR.VoucherID) as Currency,AdjAmount " +
                            " from NVO_financeInvoiceSetOffCR where InvId =" + RecID;
            return Manag.GetViewData(_Query, "");


        }



    }
}