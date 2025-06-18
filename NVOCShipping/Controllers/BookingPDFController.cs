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
    public class BookingPDFController : Controller
    {
        // GET: BookingPDF
        DocumentManager Manag = new DocumentManager();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult BookingPDF(string id, string agencyid, string GeoLocID)
        {
            if (agencyid == "")
            {
                CreatePDF(id, agencyid, GeoLocID);
                return View();
            }
            else
            {
                CreatePDF(id, agencyid, GeoLocID);
                return View();

            }
        }
        public void CreatePDF(string BkgID, string AgencyID, string GeoLocID)
        {
            DataTable dtv = GetBkgPDFValus(BkgID);
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
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;

                DataTable dtc = GetAgencyDetails(AgencyID);
                if (dtc.Rows.Count > 0)
                {


                    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
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

                DataTable dtCT = GetBkgCntrValus(BkgID);
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
                if (AgencyID == "3")
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
                if (AgencyID == "3")
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

                if (AgencyID != "3")
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

                if (GeoLocID == "1" || GeoLocID == "2" || GeoLocID == "3")
                {
                    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.WHITE)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 24;
                    Tbl5.AddCell(cell);
                    ////Value
                    //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PortNtRef"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 3;
                    //Tbl5.AddCell(cell);

                    //cell = new Cell(new Phrase("SCN No", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 2;
                    //Tbl5.AddCell(cell);
                    ////Value
                    //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["SCNNo"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 4;
                    //Tbl5.AddCell(cell);

                    //cell = new Cell(new Phrase("BS CODE", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 3;
                    //Tbl5.AddCell(cell);
                    ////Value
                    //cell = new Cell(new Phrase(" : " + dtv.Rows[0]["BSCODE"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell.BorderWidth = 0;
                    //cell.Colspan = 3;
                    //Tbl5.AddCell(cell);
                }
                else
                {
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
                }

              

                //if (AgencyID != "3")
                //{
                   
                //}
                cell = new Cell(new Phrase("Pick Up Depot", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dtv.Rows[0]["PickUpDepot"].ToString() + " \n " + dtv.Rows[0]["DepAddress"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


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
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=BookingConfirmation.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();

            }

        }
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
                            " case when ( select count(vr.RID) from NVO_VoyageRoute vr where vr.VoyageID = NVO_Booking.VesVoyID) >2 then "+
                            " convert(varchar, ((select top(1) ETA from NVO_VoyageRoute vr inner join NVO_PortMaster pm on pm.MainPortID = vr.PortID "+
                            " where VoyageID = NVO_Booking.VesVoyID and pm.id = NVO_Booking.PODID)), 103) else convert(varchar,((select top(1) ETA from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID DESC)), 103) end as NextPortETA,"+
                            " convert(varchar, (select top(1) CutDate from NVO_CROMaster where BkgID = NVO_Booking.id), 103) as CutDate, " +
                            " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_Booking.SlotOperatorID) as CarrierName,  " +
                            " (select top(1)(select top(1) TerminalName from  NVO_TerminalMaster where NVO_TerminalMaster.ID = TerminalID) from  NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID order by RID asc) as Terminal, " +
                            " isnull((select top 1 Notes from NVO_VoyageNotesDtls inner join NVO_GeneralMaster GM on GM.ID =  NVO_VoyageNotesDtls.NotesTypeID where VoyageID = NVO_Booking.VesVoyID AND GM.GeneralName = 'SCN No'),'') AS SCNNo, " +
                           " isnull((select top 1 Notes from NVO_VoyageNotesDtls inner join NVO_GeneralMaster GM on GM.ID =  NVO_VoyageNotesDtls.NotesTypeID where VoyageID = NVO_Booking.VesVoyID  AND GM.GeneralName = 'BS CODE'),'') AS BSCODE, " +
                           " isnull((select top 1 Notes from NVO_VoyageNotesDtls inner join NVO_GeneralMaster GM on GM.ID =  NVO_VoyageNotesDtls.NotesTypeID where VoyageID = NVO_Booking.VesVoyID  AND GM.GeneralName = 'VESSEL CLOSING TIME'),'') AS ClosingTime , " +
                           " isnull((select top 1 VM.VesselID from NVO_VesselMaster VM inner  join NVO_Voyage on NVO_Voyage.VesselID = VM.ID   where NVO_Voyage.ID = NVO_Booking.VesVoyID ),'') AS VesselIDValue, " +
                           "  (select top(1) DepAddress from NVO_DepotMaster  where NVO_DepotMaster.Id = NVO_Booking.PickUpDepotID)  as DepAddress "+
                            " from NVO_Booking " +
                            " where NVO_Booking.ID=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBkgCntrValus(string BkgId)
        {
            string _Query = " select BKgID,NVO_tblCntrTypes.Size,Qty,GeneralName as Commodity from NVO_BookingCntrTypes inner join NVO_tblCntrTypes on NVO_tblCntrTypes.ID = NVO_BookingCntrTypes.CntrTypes inner join NVO_GeneralMaster on NVO_GeneralMaster.ID = NVO_BookingCntrTypes.CommodityType where NVO_BookingCntrTypes.BkgID=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }
        public ActionResult CROPDF(string id, string agencyid)
        {

            //if (agencyid == "3")
            //{
                CreatePortKlangCROPDF(id, agencyid);
                return View();
            //}
           // //else
           // //{
           // CreateCROPDF(id, agencyid);
           //     return View();

           //// }
        }

        public void CreateCROPDF(string CRId, string AgencyID)
        {

            DataTable dt = GetCROPDFValus(CRId);
            if (dt.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
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


                DataTable dtc = GetAgencyDetails(AgencyID);
                if (dtc.Rows.Count > 0)
                {
                   
                }
                
                var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                img.Alignment = Element.ALIGN_LEFT;
                img.ScaleAbsolute(150f, 80f);
                cell = new Cell(img);
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                tbllogo.AddCell(cell);
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

                cell = new Cell(new Phrase("Container Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);

                cell = new Cell(new Phrase("Release Quantity", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);


                DataTable dtCroDtls = GetCRODetailsPDFValues(CRId);
                for (int i = 0; i < dtCroDtls.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrTypes.AddCell(cell);

                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["ReqQty"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
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
                cell = new Cell(new Phrase("Shipper :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);


                cell = new Cell(new Phrase(dt.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);
                doc.Add(TblShipper);

                ////---------Surveyor Details -----------///


                iTextSharp.text.Table TblSurveyor = new iTextSharp.text.Table(2);
                TblSurveyor.Width = 100;
                TblSurveyor.Alignment = Element.ALIGN_LEFT;
                TblSurveyor.Cellpadding = 0;
                TblSurveyor.BorderWidth = 0;

                cell = new Cell(new Phrase("Surveyor Details :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblSurveyor.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["SurveyorName"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblSurveyor.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblSurveyor.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(dt.Rows[0]["SurveyorAddress"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblSurveyor.AddCell(cell);

                doc.Add(TblSurveyor);

                ////---------Pick Up Depo -----------///

                iTextSharp.text.Table TblPickUpdepo = new iTextSharp.text.Table(2);
                TblPickUpdepo.Width = 100;
                TblPickUpdepo.Alignment = Element.ALIGN_LEFT;
                TblPickUpdepo.Cellpadding = 0;
                TblPickUpdepo.BorderWidth = 0;
                cell = new Cell(new Phrase("Pick Up Depo :", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblPickUpdepo.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblPickUpdepo.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblPickUpdepo.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(" " + dt.Rows[0]["DepotAddress"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblPickUpdepo.AddCell(cell);
                doc.Add(TblPickUpdepo);

                ////---------Remarks -----------///
                iTextSharp.text.Table TblRemarks = new iTextSharp.text.Table(2);
                TblRemarks.Width = 100;
                TblRemarks.Alignment = Element.ALIGN_LEFT;
                TblRemarks.Cellpadding = 0;
                TblRemarks.BorderWidth = 0;



                cell = new Cell(new Phrase("Remarks :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblRemarks.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["Remarks"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblRemarks.AddCell(cell);

                doc.Add(TblRemarks);

                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 0;
                Tbl7.BorderWidth = 0;
                cell = new Cell(new Phrase("Terms & Condition :" + "\n \n", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);
                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--//
                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace5 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace5);

                if (AgencyID != "14")
                {

                    DataTable _dtv = GetNotesClauses();
                    if (_dtv.Rows.Count > 0)
                    {
                        for (int i = 0; i < _dtv.Rows.Count; i++)
                        {
                            cell = new Cell(new Phrase(_dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                            cell.BorderWidth = 0;
                            Tbl7.AddCell(cell);
                        }
                    }

                    cell = new Cell(new Phrase("" + "\n \n", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Colspan = 1;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Colspan = 1;
                    Tbl7.AddCell(cell);

                    doc.Add(Tbl7);
                }
                else
                {

                    cell = new Cell(new Phrase("1.Ensure the empty container received from our yard is in clean and sound condition. Costs for any subsequent Rejection will be to your account and Line will not take responsibility for the same.", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("2.Liner free detention is 10 Days post the same Detention Charges will be applicable as per Company policy,", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("3.Any loss or damage to the container while in custody of shipper, transporter and forwarder shall be fully identified for repair/replacement/ reimbursement as notified by Owner / Hirer.", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("4.Loading list needs to be sent to the shipping line 24 hours prior to vessel cut off.", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("5.For Non-availability of Containers kindly contact our Operations Incharge - Mr.Durgesh Tandon, Mobile No: +91-97249 61403, Mr.Rahul, Mobile No: +91-90336 73828 will be on your account.", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("6.FORM 13 to be collected & Shipping Bill to be handed over to our surveyor.", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("7.For Surveyor details, Please refer CRO" + "\n \n", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    Tbl7.AddCell(cell);

                    cell = new Cell(new Phrase("****This is system generated file, doesn’t require any seal/stamp*****", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Colspan = 1;
                    Tbl7.AddCell(cell);

                    doc.Add(Tbl7);
                }



                #endregion

                //#region FOOTER
                /////---------FOOTER----------------//
                /////
                ////Sub Heading
                //iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                //Tbl8.Width = 100;
                //Tbl8.Alignment = Element.ALIGN_CENTER;
                //Tbl8.Cellpadding = 0;
                //Tbl8.BorderWidth = 0;


                //doc.Add(Tbl8);



                //#endregion

                #endregion




                pdfWriter.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=RRQuotation.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                Response.End();
            }
        }

        public void CreatePortKlangCROPDF(string CRId, string AgencyID)
        {

            DataTable dt = GetCROPDFValus(CRId);
            if (dt.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
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


                DataTable dtc = GetAgencyDetails(AgencyID);
                if (dtc.Rows.Count > 0)
                {
                    if (dtc.Rows[0]["LogoPath"].ToString() != "")
                    {

                        var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
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
                        var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
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

                cell = new Cell(new Phrase("Container Type", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);

                cell = new Cell(new Phrase("Release Quantity", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrTypes.AddCell(cell);


                DataTable dtCroDtls = GetCRODetailsPDFValues(CRId);
                for (int i = 0; i < dtCroDtls.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrTypes.AddCell(cell);

                    cell = new Cell(new Phrase(dtCroDtls.Rows[i]["ReqQty"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
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
                cell = new Cell(new Phrase("Shipper :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);


                cell = new Cell(new Phrase(dt.Rows[0]["Shipper"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblShipper.AddCell(cell);
                doc.Add(TblShipper);

                ////---------cargo,haulier,operator -----------///


                iTextSharp.text.Table Tblcargo = new iTextSharp.text.Table(2);
                Tblcargo.Width = 100;
                Tblcargo.Alignment = Element.ALIGN_LEFT;
                Tblcargo.Cellpadding = 0;
                Tblcargo.BorderWidth = 0;

                cell = new Cell(new Phrase("Cargo :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblcargo.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["Cargo"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblcargo.AddCell(cell);

                doc.Add(Tblcargo);

                iTextSharp.text.Table Tblhaulier = new iTextSharp.text.Table(2);
                Tblhaulier.Width = 100;
                Tblhaulier.Alignment = Element.ALIGN_LEFT;
                Tblhaulier.Cellpadding = 0;
                Tblhaulier.BorderWidth = 0;

                cell = new Cell(new Phrase("Haulier :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblhaulier.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["Haulier"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblhaulier.AddCell(cell);

                doc.Add(Tblhaulier);


                iTextSharp.text.Table Tbloperator = new iTextSharp.text.Table(2);
                Tbloperator.Width = 100;
                Tbloperator.Alignment = Element.ALIGN_LEFT;
                Tbloperator.Cellpadding = 0;
                Tbloperator.BorderWidth = 0;

                cell = new Cell(new Phrase("Operator :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbloperator.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["Operator"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbloperator.AddCell(cell);

                doc.Add(Tbloperator);

                ////---------Pick Up Depo -----------///

                iTextSharp.text.Table TblPickUpdepo = new iTextSharp.text.Table(2);
                TblPickUpdepo.Width = 100;
                TblPickUpdepo.Alignment = Element.ALIGN_LEFT;
                TblPickUpdepo.Cellpadding = 0;
                TblPickUpdepo.BorderWidth = 0;
                cell = new Cell(new Phrase("Pick Up Depo :", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblPickUpdepo.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["PickUpDepot"].ToString(), new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblPickUpdepo.AddCell(cell);

                //blank//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                // cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblPickUpdepo.AddCell(cell);

                //address//
                cell = new Cell(new Phrase(" " + dt.Rows[0]["DepotAddress"].ToString() + "\n \n \n", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.GRAY)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Rowspan = 2;
                TblPickUpdepo.AddCell(cell);
                doc.Add(TblPickUpdepo);

                ////---------Remarks -----------///
                iTextSharp.text.Table TblRemarks = new iTextSharp.text.Table(2);
                TblRemarks.Width = 100;
                TblRemarks.Alignment = Element.ALIGN_LEFT;
                TblRemarks.Cellpadding = 0;
                TblRemarks.BorderWidth = 0;



                cell = new Cell(new Phrase("Remarks :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblRemarks.AddCell(cell);

                cell = new Cell(new Phrase(dt.Rows[0]["Remarks"].ToString() + "\n \n", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 1;
                TblRemarks.AddCell(cell);

                doc.Add(TblRemarks);

                #endregion

                #region Terms & Condition

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 0;
                Tbl7.BorderWidth = 0;
                cell = new Cell(new Phrase("Terms & Condition :" + "\n \n", new Font(Font.HELVETICA, 11, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);
                ///--SPACE--//
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                ///--SPACE--//
                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace5 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace5);



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
                

                cell = new Cell(new Phrase("*********This is system generated file, doesn’t require any seal/stamp************", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.RED)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.Colspan = 1;
                Tbl7.AddCell(cell);

                doc.Add(Tbl7);



                #endregion

                //#region FOOTER
                /////---------FOOTER----------------//
                /////
                ////Sub Heading
                //iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(1);
                //Tbl8.Width = 100;
                //Tbl8.Alignment = Element.ALIGN_CENTER;
                //Tbl8.Cellpadding = 0;
                //Tbl8.BorderWidth = 0;


                //doc.Add(Tbl8);



                //#endregion

                #endregion




                pdfWriter.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                //Response.AddHeader("content-disposition", "attachment;filename=RRQuotation.pdf");
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
        public DataTable GetCROPDFValus(string CROId)
        {
            string _Query = " select BookingNo, BkgParty, convert(varchar, Date, 103) as Date,(select top(1) Address from NVO_CusBranchLocation where CID = BkgPartyID) as Address,ServiceType, Linecode," +
                            " convert(varchar, NVO_CROMaster.ValidTill, 103) as ValidTill,ReleaseOrderNo,POO,POL,POD,FPOD,NVO_Booking.VesVoy,TsPort," +
                            " convert(varchar, ETADate, 103) as ETADate,convert(varchar, ETDDate, 103) as ETDDate,FORMAT(CutDate,'dd/MM/yyyy hh:mm:ss tt') as CUTDate,Shipper,(select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_CROMaster.Surveyor) as SurveyorName," +
                            " (select top(1) Address from NVO_CusBranchLocation where CID = NVO_CROMaster.Surveyor) as SurveyorAddress," +
                            " (select top(1) DepName from NVO_DepotMaster where ID = NVO_CROMaster.PickDepoID) as PickUpDepot," +
                            " (select top(1) DepAddress from NVO_DepotMaster where ID = NVO_CROMaster.PickDepoID) as DepotAddress,NVO_CROMaster.Remarks,NVO_CROMaster.Cargo,NVO_CROMaster.Haulier,NVO_CROMaster.Operator " +
                            " from NVO_Booking inner join NVO_CROMaster on NVO_CROMaster.BkgID = NVO_Booking.ID where NVO_CROMaster.Id = " + CROId;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCRODetailsPDFValues(string CROId)
        {
            string _Query = "select CROID,Size,ReqQty from NVO_CRODETAILS CRD INNER JOIN NVO_tblCntrTypes CT ON CT.ID = CRD.CntrTypeID where CROID = " + CROId;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetNotesClauses()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=265";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetNotesClausesBooking()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=266";
            return Manag.GetViewData(_Query, "");
        }
    }
}