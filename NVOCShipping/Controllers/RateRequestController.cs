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
using System.Net.Mail;
using System.Net;


namespace NVOCShipping.Controllers
{

    public class RateRequestController : Controller
    {
        SalesManager Manag = new SalesManager();
        // GET: RateRequest
        public ActionResult Index()
        {
            return View();
        }

  
        public ActionResult RateRequestPdfSample(string id)
        {
           /// CreatePDF(id);
            //CreatePDFMODIFIED(id);
            return View();

        }
        public void CreatePDFold(string ID)
        {

            DataTable dtv = GetRRPDFValus(ID);
            if (dtv.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
                pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create));
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
                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);
                ///----/////
                DataTable dtc = GetCompanyDetails();
                if (dtc.Rows.Count > 0)
                {
                    cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLUE)));
                }
                
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("RATE REQUEST ", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase(dtc.Rows[0]["CompanyAddress"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                cell.Colspan = 2;
                tbllogo.AddCell(cell);

                //cell = new Cell(new Phrase(DS.Tables[0].Rows[0]["Location"].ToString() + " - " + DS.Tables[0].Rows[0]["Pincode"].ToString() + " Tel # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Phone1"].ToString() + "   Fax # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Fax"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 6;
                // tbllogo.AddCell(cell);

                doc.Add(tbllogo);

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

                #region Bookingparty and Ratesheet details
                //-------------------Bookingparty and Ratesheet details-----------
                PdfContentByte content = pdfWriter.DirectContent;
                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 50;
                PdfPCell cell1 = new PdfPCell(new Phrase("Booking Party", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLUE)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["Customer"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                cell1.BorderWidth = 0;
                Tbl1.AddCell(cell1);

                var Addresss = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
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


                cell1 = new PdfPCell(new Phrase("RATE REQUEST N0", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["RatesheetNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SHIPMENT TYPE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ShipmentTypes"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SALES", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["SalesPerson"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("VALID TILL ", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["ValidDate"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("COLLECTION MODE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["PePaid"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell1.BorderWidth = 1;
                //cell1.Colspan = 1;
                //cell1.Rowspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl1.AddCell(cell1);

                mtable.AddCell(Tbl1);
                doc.Add(mtable);

                #endregion

                #region Location POL POD POO
                // /----------------------- LocTable-----------------------///

                iTextSharp.text.Table TblLocs = new iTextSharp.text.Table(6);
                TblLocs.Width = 100;
                TblLocs.Alignment = Element.ALIGN_LEFT;
                TblLocs.Cellpadding = 1;
                TblLocs.BorderWidth = 1;

                cell = new Cell(new Phrase("Place Of Origin", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase("Place Of Final Destination", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase("Transhipment Port", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                //for (int i = 0; i <= Rates_DS.Tables[0].Rows.Count - 1; i++)
                //{
                cell = new Cell(new Phrase(dtv.Rows[0]["POOLs"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["FPOD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["TSPort"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell);
                // }
                doc.Add(TblLocs);
                // }

                #endregion

                #region Container And Free DayDetails
                //----------------- Container And Free DayDetails--------------//

                iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                Tbl3.Width = 100;
                Tbl3.Alignment = Element.ALIGN_LEFT;
                Tbl3.Cellpadding = 0;
                Tbl3.BorderWidth = 0;

                //Sub Heading
                cell = new Cell(new Phrase("Container And Free Days Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLUE)));

                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl3.AddCell(cell);
                doc.Add(Tbl3);

                iTextSharp.text.Table TblCntrDtls = new iTextSharp.text.Table(5);
                TblCntrDtls.Width = 100;
                TblCntrDtls.Alignment = Element.ALIGN_LEFT;
                TblCntrDtls.Cellpadding = 1;
                TblCntrDtls.BorderWidth = 1;

                cell = new Cell(new Phrase("Container Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrDtls.AddCell(cell);

                cell = new Cell(new Phrase("Commodity", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrDtls.AddCell(cell);

                cell = new Cell(new Phrase("Freedays @ POL", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrDtls.AddCell(cell);

                cell = new Cell(new Phrase("Freedays @ POL", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BackgroundColor = new Color(152, 178, 209);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrDtls.AddCell(cell);


                DataTable _dtCnt = GetRRPDFCntrTypesValus(ID);
                for (int i = 0; i < _dtCnt.Rows.Count; i++)
                {
                    cell = new Cell(new Phrase(_dtCnt.Rows[i]["CntrTypes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 2;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtCnt.Rows[i]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtCnt.Rows[i]["ExFreeday"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell);

                    cell = new Cell(new Phrase(_dtCnt.Rows[i]["ImFreeday"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell.BorderWidth = 1;
                    cell.Colspan = 1;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell);
                }

                doc.Add(TblCntrDtls);
                // }
                #endregion

                #region MRG & Slot Details-
                //----------------- MRG & Slot Details--------------//



                iTextSharp.text.Table Tblline = new iTextSharp.text.Table(1);
                Tblline.Width = 100;
                Tblline.Border = 0;
                Tblline.Cellpadding = 2;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 1;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                Tblline.AddCell(cell);
                doc.Add(Tblline);


                //Sub Heading
                iTextSharp.text.Table Tbl4 = new iTextSharp.text.Table(1);
                Tbl4.Width = 100;
                Tbl4.Alignment = Element.ALIGN_LEFT;
                Tbl4.Cellpadding = 0;
                Tbl4.BorderWidth = 0;

                cell = new Cell(new Phrase("MRG & Slot Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl4.AddCell(cell);
                doc.Add(Tbl4);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace3 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace3);
                //------------------------------------------------------------------------

                PdfPTable mtable1 = new PdfPTable(2);
                mtable1.WidthPercentage = 100;
                mtable1.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable TblMRGDtls = new PdfPTable(2);
                TblMRGDtls.WidthPercentage = 50;
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                //cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                TblMRGDtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                // cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblMRGDtls.AddCell(cell1);

                var img1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/Images/Location.png"));
                img1.Alignment = Element.ALIGN_LEFT;
                cell1 = new PdfPCell(img1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                TblMRGDtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["POLCode"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                // cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                TblMRGDtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["PODCode"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                // cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                TblMRGDtls.AddCell(cell1);

                mtable1.AddCell(TblMRGDtls);


                TblMRGDtls = new PdfPTable(3);
                TblMRGDtls.WidthPercentage = 50;
                TblMRGDtls.HorizontalAlignment = Element.ALIGN_RIGHT;

                cell1 = new PdfPCell(new Phrase("Size Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                //cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblMRGDtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("OCEAN FRT", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                // cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblMRGDtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("SLOT", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                //cell.BackgroundColor = new Color(152, 178, 209);
                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblMRGDtls.AddCell(cell1);


                DataTable _dtSL = GetRRPDFMRGSLOTValus(ID);

                for (int i = 0; i < _dtSL.Rows.Count; i++)
                {
                    cell1 = new PdfPCell(new Phrase(_dtSL.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell1.BorderWidth = 0;
                    cell1.Colspan = 1;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblMRGDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtSL.Rows[i]["MrgAmt"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell1.BorderWidth = 0;
                    cell1.Colspan = 1;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblMRGDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtSL.Rows[i]["SlotAmt"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));
                    cell1.BorderWidth = 0;
                    cell1.Colspan = 1;
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblMRGDtls.AddCell(cell1);
                }
                mtable1.AddCell(TblMRGDtls);

                doc.Add(mtable1);


                iTextSharp.text.Table Tblslot = new iTextSharp.text.Table(10);
                Tblslot.Width = 100;
                Tblslot.Alignment = Element.ALIGN_RIGHT;
                Tblslot.Cellpadding = 2;
                Tblslot.BorderWidth = 0;


                cell = new Cell(new Phrase("SLOT Operator", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.Colspan = 2;
                cell.BorderWidth = 0;
                Tblslot.AddCell(cell);


                //cell = new Cell(new Phrase(" : " + dt.Rows[0]["CANNo"].ToString().ToUpper(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                cell = new Cell(new Phrase(":" + dtv.Rows[0]["SlotOpt"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tblslot.AddCell(cell);
                doc.Add(Tblslot);

                #endregion

                #region Tariff Details

                iTextSharp.text.Table Tblline1 = new iTextSharp.text.Table(1);
                Tblline1.Width = 100;
                Tblline1.Border = 0;
                Tblline1.Cellpadding = 2;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 1;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                Tblline1.AddCell(cell);
                doc.Add(Tblline1);


                //Sub Heading
                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(1);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 0;
                Tbl5.BorderWidth = 0;

                cell = new Cell(new Phrase("Tariff Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl5.AddCell(cell);
                doc.Add(Tbl5);

                iTextSharp.text.Table TblTariffDtls = new iTextSharp.text.Table(10);
                TblTariffDtls.Width = 100;
                TblTariffDtls.Alignment = Element.ALIGN_RIGHT;
                TblTariffDtls.Cellpadding = 2;
                TblTariffDtls.BorderWidth = 0;

                DataTable _dtT = GetRRPDFTariffValus(ID, "1");

                cell = new Cell(new Phrase("Export Local Charges", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.Colspan = 2;
                cell.BorderWidth = 0;
                TblTariffDtls.AddCell(cell);
                if (_dtT.Rows.Count > 0)
                {
                    //cell = new Cell(new Phrase(" : " + dt.Rows[0]["CANNo"].ToString().ToUpper(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                    cell = new Cell(new Phrase(":" + _dtT.Rows[0]["Curr"].ToString() + " " + _dtT.Rows[0]["Amt"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 3;
                }
                TblTariffDtls.AddCell(cell);


                cell = new Cell(new Phrase("Import Local Charges", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.Colspan = 2;
                cell.BorderWidth = 0;
                TblTariffDtls.AddCell(cell);
                DataTable _dtT1 = GetRRPDFTariffValus(ID, "2");
                if (_dtT1.Rows.Count > 0)
                {
                    //cell = new Cell(new Phrase(" : " + dt.Rows[0]["CANDate"].ToString().ToUpper(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                    cell = new Cell(new Phrase(":" + _dtT1.Rows[0]["Curr"].ToString() + " " + _dtT1.Rows[0]["Amt"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLUE)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 2;
                    TblTariffDtls.AddCell(cell);
                }
                doc.Add(TblTariffDtls);


                //Sub Heading
                iTextSharp.text.Table Tbl6 = new iTextSharp.text.Table(1);
                Tbl6.Width = 100;
                Tbl6.Alignment = Element.ALIGN_LEFT;
                Tbl6.Cellpadding = 0;
                Tbl6.BorderWidth = 0;
                #endregion

                #region Terms & Condition
                cell = new Cell(new Phrase("Terms & Condition :", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLUE)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl6.AddCell(cell);
                doc.Add(Tbl6);

                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.Cellpadding = 1;
                Tbl7.BorderWidth = 0;

                cell = new Cell(new Phrase(" * Rate was agreed & accepted to respective market price ", new Font(Font.HELVETICA, 11, Font.NORMAL, Color.BLUE)));
                cell.BorderWidth = 0;
                Tbl7.AddCell(cell);
                doc.Add(Tbl7);
                #endregion

                #region FOOTER
                ///---------FOOTER----------------//
                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(4);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.Cellpadding = 1;
                Tbl8.BorderWidth = 1;

                cell = new Cell(new Phrase("Created By : Venkat", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);

                cell = new Cell(new Phrase("Created On : 21/04/2021", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);

                cell = new Cell(new Phrase("Approved By : Shyam", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);

                cell = new Cell(new Phrase("Approved On :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLUE)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);
                doc.Add(Tbl8);
                #endregion

                //string str = " RRPDF/" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf";
                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "_POPUP_Window", "<script>window.open('" + str + "','new','left=10,top=10,width=1200,height=600,scrollbars=yes')</script>", false);
                pdfWriter.CloseStream = false;
                doc.Close();
                Response.Buffer = true;
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "attachment;filename=RRQuotation.pdf");
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.Write(doc);
                Response.End();
            }

        }


        public void RRPDFSave(string ID, string AgentID)
        {

            DataTable dtv = GetRRPDFValus(ID);
            if (dtv.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
               // PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
               // pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create));
                doc.Open();

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                //cell.Width = 10;

                var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                img.Alignment = Element.ALIGN_LEFT;
                img.ScaleAbsolute(150f, 80f);
                img.Alignment = Element.ALIGN_LEFT;
                cell = new Cell(img);
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.Width = 20;
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
                DataTable dtc = GetCompanyDetails();
                if (dtc.Rows.Count > 0)
                {
                    cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                }
                
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("RATE REQUEST ", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase(dtc.Rows[0]["CompanyAddress"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                cell.Colspan = 2;
                tbllogo.AddCell(cell);

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

                var Addresss = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
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

                PdfPTable TblCntrDtls = new PdfPTable(new float[] { 2, 2, 2, 2, });
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

                cell1 = new PdfPCell(new Phrase("DG Class", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrDtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("VGM", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblCntrDtls.AddCell(cell1);

                DataTable _dtCnt = GetRRPDFCntrTypesValus(ID);
                for (int i = 0; i < _dtCnt.Rows.Count; i++)
                {
                    cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["CntrTypes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
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

                DataTable _dtFreeday = GetRRFreedaysDtls(ID);
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

                #region second page

                #region logo
                //----------------------------DYNAMIC-LETTER-HEAD-ADDRESS---------------------------
                iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(8);
                tbllogo1.Width = 100;
                tbllogo1.Alignment = Element.ALIGN_LEFT;
                tbllogo1.Cellpadding = 1;
                tbllogo1.BorderWidth = 0;

                var img1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                img1.Alignment = Element.ALIGN_CENTER;
                cell = new Cell(img1);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                cell.Rowspan = 3;
                cell.Width = 20;
                tbllogo1.AddCell(cell);


                cell = new Cell(new Phrase("GNL Lines Pte.Ltd", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 6;
                tbllogo1.AddCell(cell);


                cell = new Cell(new Phrase("33 UBI AVENUE 3, #08-68, VERTEX TOWER A," + "\n" + "SINGAPORE 408868 ", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                cell.Colspan = 2;
                tbllogo1.AddCell(cell);
                doc.Add(tbllogo1);

                iTextSharp.text.Table Tblline1 = new iTextSharp.text.Table(1);
                Tblline1.Width = 100;
                Tblline1.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblline1.Border = Rectangle.NO_BORDER;
                Tblline1.Cellpadding = 1;
                Tblline1.BorderWidth = 0.5F;

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthLeft = 0;
                cell.BorderWidthBottom = 0;
                cell.BackgroundColor = new Color(98, 141, 214);
                Tblline1.AddCell(cell);
                doc.Add(Tblline1);

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

                DataTable _dtfrt = GetRRTariffCharges(ID, "135");
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

                DataTable _dtThc = GetRRTariffCharges(ID, "136");
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
                DataTable _dtHAUL = GetRRTariffCharges(ID, "137");

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

                DataTable _dtLCO = GetRRLocalTariffCharges(ID, "138", "18");
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

                DataTable _dtLCDest = GetRRLocalTariffCharges(ID, "138", "19");
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

                #endregion



                //string str = "RRPDF/" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf";
                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "_POPUP_Window", "<script>window.open('" + str + "','new','left=10,top=10,width=1200,height=600,scrollbars=yes')</script>", false);
               // pdfWriter.CloseStream = false;
                doc.Close();
               // doc.Dispose();
                //Response.Buffer = true;
                //Response.ContentType = "application/pdf";
                ////Response.AddHeader("content-disposition", "attachment;filename=RRQuotation.pdf");
                //Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Response.Write(doc);
                //Response.End();
               


                DataTable dtE = GetEmailsedingRRView(ID);
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
                              " <td style='font-family:Arial;font-size:14px;font-weight:bold;text-align:left;color:#fff;padding-top:2px;padding-left:3px;padding-bottom:2px;padding-right:3px;'>RR NUMBER : <a href='https://nerida.rmjtech.in/RRPDF/" + dtE.Rows[0]["RatesheetNo"].ToString() + ".pdf' style='color:#fff;'>" + dtE.Rows[0]["RatesheetNo"].ToString() + "</a></td> " +
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
                              " <tr> " +
                               " <td colspan='2' style='font-weight:bold;'>Freedays</td> " +
                              " </tr> " +
                              " <tr> " +
                              " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Freedays"].ToString() + " Days</td> " +
                              " </tr> " +
                               " <tr> " +
                               " <td colspan='2' style='font-weight:bold;'>Collection Mode</td> " +
                              " </tr> " +
                              " <tr> " +
                              " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["PePaid"].ToString() + "</td> " +
                              " </tr> " +
                              " </tbody> " +
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
                              " <td colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Rate</td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                               " <tr> " +
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>20GP - " + dtE.Rows[0]["Cntr20"].ToString() + " USD</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                               " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>40HC - " + dtE.Rows[0]["Cntr40"].ToString() + " USD</td>" +
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
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='background-color:#2196f3;color:#fff; padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Slot Cost </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                               " <tr> " +
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>20GP - " + dtE.Rows[0]["Slot20"].ToString() + " USD</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                               " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>40HC - " + dtE.Rows[0]["Slot40"].ToString() + " USD</td>" +
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
                                " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='color:#000; padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Slot Operator </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                               " <tr> " +
                                " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>" + dtE.Rows[0]["SlotOpt"].ToString() + "</td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                                " <tr> " +
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;padding-top:6px;border-top:1px solid #2196f3;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>Created By: Venkat</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                               " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding:left:64px!important;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>Created On : 10/05/2021</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table>";
                    DataTable _dtCom = GetCompnayDetails();
                    MailMessage EmailObject = new MailMessage();
                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    DataTable dtAuto = GetCustomerEmailsending(AgentID);
                    var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                    for (int y = 0; y < EmailID.Length; y++)
                    {
                        if (EmailID[y].ToString() != "")
                        {
                            EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                        }
                    }
                    EmailObject.Attachments.Add(new Attachment(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf")));
                    EmailObject.Body = HTMLstr;
                    EmailObject.IsBodyHtml = true;
                    EmailObject.Priority = MailPriority.Normal;
                    EmailObject.Subject = "Rate Request: " + dtE.Rows[0]["RatesheetNo"].ToString() + " - APPROVAL PENDING";
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


        public DataTable GetCompnayDetails()
        {
            string _Query = "select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCustomerEmailsending(string AgentID)
        {
            string _Query = "select EmailID from NVO_AgencyEmailDtls where AlertTypeID = 1 and AgencyID=" + AgentID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetEmailsedingRRView(string RRID)
        {
            string _Query = " select NVO_Ratesheet.Id,RatesheetNo,convert(varchar, ValidTill, 106) as ValidDate,(select top(1) Status from NVO_RRStatusMaster where Id = NVO_Ratesheet.RSStatus) as Status,  " +
                            " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = BookingPartyID) as Customer, " +
                            " (select top(1) UserName from NVO_UserDetails where ID = SalesPersonID) as SalesPerson, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = PortOfLoading) as Loading, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = PlaceofDischargeId) as Discharge, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = TranshimentPortID) as TSPort, '14' as Freedays,Remarks, " +
                            " (select top(1)(select top(1) GeneralName from NVO_GeneralMaster where ID = NVO_RatesheetMRG.FreightTerms) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as PePaid, " +
                            " (select top(1)(select top(1) Size from NVO_tblCntrTypes where SizeId = 1 and ID = NVO_RatesheetMRG.CntrTypes) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as Cntr20, " +
                            " (select sum(QuotedAmount) from NVO_RatesheetMRG where CntrTypes = 1 and RRID = NVO_Ratesheet.Id) as MrgRate20, " +
                            " (select top(1)(select top(1) Size from NVO_tblCntrTypes where SizeId = 2 and ID = NVO_RatesheetMRG.CntrTypes) from NVO_RatesheetMRG where RRID = NVO_Ratesheet.Id) as Cntr40, " +
                            " (select sum(QuotedAmount) from NVO_RatesheetMRG where CntrTypes = 2 and RRID = NVO_Ratesheet.Id) as MrgRate40, " +
                            " (select sum(SlotAmt) from NVO_RatesheetSLOTDtls where CntrTypes = 1 and Commodity = 3 and RRID = NVO_Ratesheet.Id) as Slot20, " +
                            " (select sum(SlotAmt) from NVO_RatesheetSLOTDtls where CntrTypes = 2 and Commodity = 4 and RRID = NVO_Ratesheet.Id) as Slot40, " +
                            " (select top(1) (select top(1) CustomerName from NVO_view_CustomerDetails where CID =slotOperator) from NVO_RatesheetSLOT where NVO_RatesheetSLOT.RRID = NVO_Ratesheet.Id)  as SlotOpt " +
                            " from NVO_Ratesheet where Id = " + RRID;
            return Manag.GetViewData(_Query, "");
        }
        public void RRPDF(string ID)
        {

            DataTable dtv = GetRRPDFValus(ID);
            if (dtv.Rows.Count > 0)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
               // pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create));
                doc.Open();

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                tbllogo.Width = 100;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                //cell.Width = 10;

                var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                img.Alignment = Element.ALIGN_LEFT;
                img.ScaleAbsolute(150f, 80f);
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
                DataTable dtc = GetCompanyDetails();
                if (dtc.Rows.Count > 0)
                {
                    cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                }


                cell.BorderWidth = 0;
                tbllogo.Alignment = Element.ALIGN_LEFT;
                //cell.Colspan = 2;
                tbllogo.AddCell(cell);

                cell = new Cell(new Phrase("RATE REQUEST ", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.Colspan = 3;
                tbllogo.AddCell(cell);


                var Addresss = Regex.Split(dtc.Rows[0]["CompanyAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell = new Cell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    tbllogo.AddCell(cell);
                }

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

                var Addresss11 = Regex.Split(dtv.Rows[0]["CustomerAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss11.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss11[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
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

                PdfPTable TblCntrDtls = new PdfPTable(new float[] { 2, 2, 2, 2,2 });
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

                DataTable _dtCnt = GetRRPDFCntrTypesValus(ID);
                for (int i = 0; i < _dtCnt.Rows.Count; i++)
                {
                    cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["CntrTypes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["Commodity"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblCntrDtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtCnt.Rows[0]["Cargo"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
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

                DataTable _dtFreeday = GetRRFreedaysDtls(ID);
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

                //#region REBATE Details

                //iTextSharp.text.Table Tbl4 = new iTextSharp.text.Table(1);
                //Tbl4.Width = 100;
                //Tbl4.Alignment = Element.ALIGN_LEFT;
                //Tbl4.Cellpadding = 0;
                //Tbl4.BorderWidth = 0;

                ////Sub Heading
                //cell = new Cell(new Phrase("Rebate Details :", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 1;
                //Tbl4.AddCell(cell);
                //doc.Add(Tbl4);



                //iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(4);
                //Tbl5.Width = 60;
                //Tbl5.Alignment = Element.ALIGN_LEFT;
                //Tbl5.BorderWidth = 0;
                //if (dtv.Rows[0]["IsRebate"].ToString() == "1")
                //{
                //    cell = new Cell(new Phrase(dtv.Rows[0]["RebateDtls"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.BorderWidth = 0;
                //    Tbl5.AddCell(cell);

                //    cell = new Cell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.BorderWidth = 0;
                //    Tbl5.AddCell(cell);
                //}
                //else if (dtv.Rows[0]["IsRebate"].ToString() == "2")
                //{
                //    cell = new Cell(new Phrase(dtv.Rows[0]["RebateDtls"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.BorderWidth = 0;
                //    Tbl5.AddCell(cell);

                //    cell = new Cell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.BorderWidth = 0;
                //    Tbl5.AddCell(cell);
                //}
                //else
                //{
                //    cell = new Cell(new Phrase(dtv.Rows[0]["RebateDtls"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.BorderWidth = 0;
                //    Tbl5.AddCell(cell);

                //    cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //    cell.BorderWidth = 0;
                //    Tbl5.AddCell(cell);
                //}
                //cell = new Cell(new Phrase("Rebate Amount  :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.BorderWidth = 0;
                //Tbl5.AddCell(cell);

                //cell = new Cell(new Phrase(dtv.Rows[0]["RebateAmt"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.BorderWidth = 0;
                //Tbl5.AddCell(cell);
                //doc.Add(Tbl5);

                //#endregion

                //Sub Heading




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


                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspaces10 = new iTextSharp.text.Table(1);
                doc.Add(Tblspaces10);

                //------------------------------------------------------------------------

                #region SlotOp


                PdfPTable Tblslotdtls = new PdfPTable(new float[] { 5, 1, 1,1 });
                Tblslotdtls.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblslotdtls.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("Slot Operator", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblslotdtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Slot Terms", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblslotdtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Slot for 20'", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblslotdtls.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("Slot for 40'", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblslotdtls.AddCell(cell1);

                DataTable _dtSlot = GetRRSlotDtls(ID);
                for (int i = 0; i < _dtSlot.Rows.Count; i++)
                {
                    cell1 = new PdfPCell(new Phrase(_dtSlot.Rows[i]["SlotOperator"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblslotdtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtSlot.Rows[i]["SlotTerms"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblslotdtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtSlot.Rows[i]["SlotAmt20"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblslotdtls.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(_dtSlot.Rows[i]["SlotAmt40"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    Tblslotdtls.AddCell(cell1);


                }

                doc.Add(Tblslotdtls);


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


                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(2);
                tbllogo1.Width = 100;
                tbllogo1.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo1.BorderWidth = 0;
                Cell cell2 = new Cell();
                //cell.Width = 10;

                var img1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
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


                //#region logo
                ////----------------------------DYNAMIC-LETTER-HEAD-ADDRESS---------------------------
                //iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(8);
                //tbllogo1.Width = 100;
                //tbllogo1.Alignment = Element.ALIGN_LEFT;
                //tbllogo1.Cellpadding = 1;
                //tbllogo1.BorderWidth = 0;

                //var img1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/OCEANUS LOGO.png"));
                //img1.Alignment = Element.ALIGN_CENTER;
                //cell = new Cell(img1);
                //cell.BorderWidth = 0;
                //cell.Colspan = 2;
                //cell.Rowspan = 3;
                //cell.Width = 20;
                //tbllogo1.AddCell(cell);


                //DataTable dtc1 = GetCompanyDetails();
                //if (dtc1.Rows.Count > 0)
                //{
                //    cell = new Cell(new Phrase(dtc1.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                //    cell.Colspan = 6;
                //}



                //cell.BorderWidth = 0;
                //tbllogo1.Alignment = Element.ALIGN_LEFT;
                ////cell.Colspan = 2;
                //tbllogo1.AddCell(cell);




                //var Addresss2 = Regex.Split(dtc1.Rows[0]["CompanyAddress"].ToString(), "\r\n|\r|\n");
                //for (int a = 0; a < Addresss2.Length; a++)
                //{
                //    cell = new Cell(new Phrase(Addresss2[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                //    cell.BorderWidth = 0;
                //    tbllogo1.Alignment = Element.ALIGN_LEFT;
                //    cell.Colspan = 2;
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

                //#endregion

                //#region header

                //iTextSharp.text.Table tblHead = new iTextSharp.text.Table(1);
                //tblHead.Width = 100;
                //tblHead.Alignment = Element.ALIGN_LEFT;
                //tblHead.Cellpadding = 1;
                //tblHead.BorderWidth = 0;

                //cell = new Cell(new Phrase("RATE REQUEST- TARIFF ANNEXURE", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                //cell.BorderWidth = 0;
                //cell.HorizontalAlignment = Element.ALIGN_CENTER;
                ////cell.Colspan = 3;
                //tblHead.AddCell(cell);
                //doc.Add(tblHead);

                //iTextSharp.text.Table tblHead1 = new iTextSharp.text.Table(4);
                //tblHead1.Width = 100;
                //tblHead1.Alignment = Element.ALIGN_LEFT;
                //tblHead1.BorderWidth = 0;

                //cell = new Cell(new Phrase("RATE REQUEST NO :", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.BorderWidth = 0;
                //tblHead1.AddCell(cell);

                //cell = new Cell(new Phrase(dtv.Rows[0]["RatesheetNo"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.BorderWidth = 0;
                //tblHead1.AddCell(cell);

                //cell = new Cell(new Phrase("STATUS  :", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                //cell.BorderWidth = 0;
                //tblHead1.AddCell(cell);

                //cell = new Cell(new Phrase(dtv.Rows[0]["Status"].ToString(), new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.BorderWidth = 0;
                //tblHead1.AddCell(cell);
                //doc.Add(tblHead1);

                //#endregion

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

                DataTable _dtfrt = GetRRTariffCharges(ID, "135");
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

                DataTable _dtThc = GetRRTariffCharges(ID, "136");
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
                DataTable _dtHAUL = GetRRTariffCharges(ID, "137");

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

                DataTable _dtLCO = GetRRLocalTariffCharges(ID, "138", "18");
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

                DataTable _dtLCDest = GetRRLocalTariffCharges(ID, "138", "19");
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

                #endregion
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


                //string str = "RRPDF/" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf";
                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "_POPUP_Window", "<script>window.open('" + str + "','new','left=10,top=10,width=1200,height=600,scrollbars=yes')</script>", false);
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

        public DataTable GetRRPDFValus(string RID)
        {
            string _Query = " select NVO_Ratesheet.Id,RatesheetNo,Remarks,OtherRemarks,convert(varchar, ValidTill, 106) as ValidDate,convert(varchar, Date, 106) as CreatedOn,convert(varchar, DtApproved, 106) as DtAppr,RebateAmt,IsRebate,Case When IsRebate=1 then 'POL' when IsRebate=2 then 'POD' else 'N/A' End as Rebatedtls,(select top(1) Status from NVO_RRStatusMaster where Id = NVO_Ratesheet.RSStatus) as Status,  " +
                             " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = BookingPartyID) as Customer,case when ShipmentID= 1  then 'EXPORT' ELSE 'IMPORT' end as ShipmentTypes," +
                             "  (select top(1) Address from NVO_CusBranchLocation where CID = BookingPartyID) as CustomerAddress,(select top(1) UserName from NVO_UserDetails where ID = SalesPersonID) as SalesPerson, " +
                                " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.ApprovedBy) as ApprovedBy, " +
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
                             "  (select top(1) (select top(1) CustomerName from NVO_view_CustomerDetails where CID =slotOperator) from NVO_RatesheetSLOT where NVO_RatesheetSLOT.RRID = NVO_Ratesheet.Id)  as SlotOpt " +
                             " from NVO_Ratesheet where Id = " + RID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRPDFCntrTypesValus(string RID)
        {
            string _Query = "select (select top(1) Size from NVO_tblCntrTypes where ID = NVO_RatesheetCntrTypes.CntrTypeID) as CntrTypes, " +
                            " (select top(1) GeneralName from NVO_GeneralMaster where Id = CommodityTypeID) as Commodity, ExFreeday, ImFreeday, VGM,DGClass, RR_Cargo as Cargo " +
                            " from NVO_RatesheetCntrTypes where RRId = " + RID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRPDFMRGSLOTValus(string RID)
        {
            string _Query = " select Size, sum(QuotedAmount) as MrgAmt, sum(slotAmt) as SlotAmt from NVO_tblCntrTypes " +
                            " inner join NVO_RatesheetMRG on NVO_RatesheetMRG.CntrTypes = NVO_tblCntrTypes.ID " +
                            " inner join NVO_RatesheetSLOTDtls on NVO_RatesheetSLOTDtls.CntrTypes = NVO_tblCntrTypes.ID " +
                            " where NVO_RatesheetMRG.RRID = " + RID +
                            " group by Size";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRPDFTariffValus(string RID, string Types)
        {
            string _Query = " select  (select top(1) CurrencyCode from NVO_CurrencyMaster where ID = CurrencyID) as Curr,  " +
                            " sum(Amt) as Amt from NVO_RatesheetRevRate where RRID = " + RID + " and ShipmentType = " + Types + " group by CurrencyID";
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

        public DataTable GetRRFreedaysDtls(string RID)
        {
            string _Query = "select ID,RRID,case when ModeID = 1 then 'COMBINED' WHEN ModeID = 2 THEN 'DETENTION' WHEN ModeID = 3 THEN 'DEMURRAGE' END AS Mode,ExpFreeDays,ImpFreeDays from NVO_Ratesheetmode where RRID = " + RID + " ";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetRRSlotDtls(string RID)
        {
            string _Query = " select (select top(1) CustomerName from  NVO_view_CustomerDetails where  CID= SlotOperatorID) as SlotOperator, " +
                            " SlotOperatorID,SlotAmt20,SlotAmt40,SlotTermID,(select top(1) Description from NVO_tblDLValues where ID = SlotTermID) as SlotTerms " +
                            " from NVO_RatesheetSlotCharges where RRID = " + RID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCompanyDetails()
        {

            string _Query = "Select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }

        public void RRApprovalStatusEmailsending(string ID, string AgentID)
        {

          
                DataTable dtE = GetEmailsedingRRView(ID);
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
                              " <td style='font-family:Arial;font-size:14px;font-weight:bold;text-align:left;color:#fff;padding-top:2px;padding-left:3px;padding-bottom:2px;padding-right:3px;'>RR NUMBER : <a href='https://nerida.rmjtech.in/RRPDF/" + dtE.Rows[0]["RatesheetNo"].ToString() + ".pdf' style='color:#fff;'>" + dtE.Rows[0]["RatesheetNo"].ToString() + "</a></td> " +
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
                              " <tr> " +
                               " <td colspan='2' style='font-weight:bold;'>Freedays</td> " +
                              " </tr> " +
                              " <tr> " +
                              " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["Freedays"].ToString() + " Days</td> " +
                              " </tr> " +
                               " <tr> " +
                               " <td colspan='2' style='font-weight:bold;'>Collection Mode</td> " +
                              " </tr> " +
                              " <tr> " +
                              " <td colspan='2' style='padding-bottom:8px;'>" + dtE.Rows[0]["PePaid"].ToString() + "</td> " +
                              " </tr> " +
                              " </tbody> " +
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
                              " <td colspan='2' style='background-color:#2196f3;color:#fff;padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Rate</td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                               " <tr> " +
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>20GP - " + dtE.Rows[0]["Cntr20"].ToString() + " USD</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                               " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>40HC - " + dtE.Rows[0]["Cntr40"].ToString() + " USD</td>" +
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
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='background-color:#2196f3;color:#fff; padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Slot Cost </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                               " <tr> " +
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>20GP - " + dtE.Rows[0]["Slot20"].ToString() + " USD</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                               " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>40HC - " + dtE.Rows[0]["Slot40"].ToString() + " USD</td>" +
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
                                " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='color:#000; padding-left:4px;padding-top:4px;padding-bottom:4px;font-size:14px;'> Slot Operator </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                               " <tr> " +
                                " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>" + dtE.Rows[0]["SlotOpt"].ToString() + "</td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                                " <tr> " +
                              " <td style='border:none;'> " +
                              " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding-left:15px;padding-right:15px;padding-bottom:8px;padding-top:6px;border-top:1px solid #2196f3;font-size:14px;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>Created By: Venkat</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                               " <td colspan='2'>" +
                               " <table border='0' cellpadding='0' cellspacing='0' width='100%' style='padding:left:64px!important;'> " +
                              " <tbody> " +
                              " <tr> " +
                              " <td colspan='2' style='font-weight:bold;'>Created On : 10/05/2021</td>" +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table> " +
                              " </td> " +
                              " </tr> " +
                              " </tbody> " +
                              " </table>";

                DataTable _dtCom = GetCompnayDetails();
                MailMessage EmailObject = new MailMessage();
                EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                DataTable dtAuto = GetCustomerEmailsending(AgentID);
                    var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                    for (int y = 0; y < EmailID.Length; y++)
                    {
                        if (EmailID[y].ToString() != "")
                        {
                            EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                        }
                    }
                    EmailObject.Attachments.Add(new Attachment(Server.MapPath("~/RRPDF\\" + dtE.Rows[0]["RatesheetNo"].ToString() + ".pdf")));
                    EmailObject.Body = HTMLstr;
                    EmailObject.IsBodyHtml = true;
                    EmailObject.Priority = MailPriority.Normal;
                    EmailObject.Subject = "Rate Request: " + dtE.Rows[0]["RatesheetNo"].ToString() + " - APPROVAL " + dtE.Rows[0]["Status"].ToString() + "";
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
}