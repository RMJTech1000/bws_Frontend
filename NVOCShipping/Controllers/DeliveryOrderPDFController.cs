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
    public class DeliveryOrderPDFController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: DeliveryOrderPDF
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult DeliveryOrderPDF(string id, string DoID, string AgencyID,string LocID)
        {
            CreateDOPDF(id, DoID, AgencyID,LocID);
            return View();

        }

        public void CreateDOPDF(string id, string DoID, string AgencyID, string LocID)
        {
            DataTable dtv = GetDOPrint(id, DoID);
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();


                //PdfWriter pdfWriter1 = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/BKGPDF\\" + dtv.Rows[0]["BookingNo"].ToString() + ".pdf"), FileMode.Create));
                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
                doc.Open();
                PdfContentByte content = pdfWriter.DirectContent;

                #region Header LOGO COMPANY NAME
                //-------------HEADER-------------------//

                iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(4);
                tbllogo.Width = 60;
              
                //tbllogo.Alignment = Element.ALIGN_LEFT;
                //tbllogo.Cellpadding = 1;
                tbllogo.BorderWidth = 0;
                Cell cell = new Cell();
                cell.Width = 10;


                var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/gnllogo.png"));
                img.Alignment = Element.ALIGN_LEFT;
                img.ScaleAbsolute(80f, 80f);
                cell = new Cell(img);
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                tbllogo.AddCell(cell);

                DataTable dtx = GetAgencyAddress(AgencyID);
                cell = new Cell(new Phrase(dtx.Rows[0]["AgencyName"].ToString()  + " " + dtx.Rows[0]["Address"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
               
                cell.Colspan = 3;
                cell.BorderWidth = 0;
                tbllogo.AddCell(cell);

                doc.Add(tbllogo);


                para = new Paragraph("");
                doc.Add(para);

                para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLUE));
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                ////----------SPACE----------------------------------
                //iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                //doc.Add(Tblspace2);

                ////------------------------------------------------------------------------
                #endregion

                PdfPTable Tbl1 = new PdfPTable(1);
                Tbl1.WidthPercentage = 103;
                //Tbl1.DefaultRowspan = 2;

                PdfPCell cell1 = new PdfPCell(new Phrase("DELIVERY ORDER", new Font(Font.HELVETICA, 12, Font.BOLD, Color.WHITE)));
                cell1.Colspan = 12;
                //cell.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                cell1.VerticalAlignment = Element.ALIGN_CENTER;
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 25f;
                cell1.BackgroundColor = new Color(98, 141, 214);
                cell1.Colspan = 1;
                Tbl1.AddCell(cell1);
                doc.Add(Tbl1);

                #region 1st page border

                Rectangle rectangle1 = new Rectangle(doc.PageSize);
                rectangle1.Left += doc.LeftMargin - 9;
                rectangle1.Right -= doc.RightMargin - 9;
                rectangle1.Top -= doc.TopMargin + 113;
                rectangle1.Bottom += doc.BottomMargin + 4;
                content.SetColorStroke(Color.BLACK);
                content.Rectangle(rectangle1.Left, rectangle1.Bottom, rectangle1.Width, rectangle1.Height);
                content.Stroke();

                #endregion

                #region DO DETAILS, DATE, MBL NUMBER BL DATE,LINE NUMBER
                // /----------------------- LocTable-----------------------///

                iTextSharp.text.Table TblDOdtls = new iTextSharp.text.Table(7);
                TblDOdtls.Width = 103;
                TblDOdtls.Alignment = Element.ALIGN_CENTER;
                // TblDOdtls.DefaultCellBorder = Rectangle.NO_BORDER;
                // TblDOdtls.Border = Rectangle.NO_BORDER;
                TblDOdtls.DefaultCell.Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                TblDOdtls.Cellpadding = 1;
                TblDOdtls.BorderWidth = 0.5f;

                cell = new Cell(new Phrase("DO NUMBER", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                //cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.BorderWidthLeft = -8;
                cell.HorizontalAlignment = 1;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase("DO DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase("MBL NUMBER", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);

                cell.BorderWidth = 1;
                cell.Colspan = 2;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase("BL DATE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase("LINE NUMBER", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BackgroundColor = new Color(98, 141, 214);
                // cell.BorderWidth = 1;
                cell.Colspan = 1;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                /////values//////
                cell = new Cell(new Phrase(dtv.Rows[0]["DONo"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                //cell.BorderWidth = 1;
                cell.Colspan = 2;

                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["IssueDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;

                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["BLNumber"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 1;
                cell.Colspan = 2;

                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["BLDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 1;
                cell.Colspan = 1;

                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["LineNumber"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 1;
                //cell.Colspan = 1;

                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                TblDOdtls.AddCell(cell);

                doc.Add(TblDOdtls);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace3 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace3);

                //------------------------------------------------------------------------
                #endregion

                #region To and CHA details

                PdfPTable mtable = new PdfPTable(2);
                mtable.WidthPercentage = 100;
                mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                PdfPTable Tbl2 = new PdfPTable(1);
                Tbl2.WidthPercentage = 50;
                cell1 = new PdfPCell(new Phrase("To", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.Left = cell1.EffectivePaddingLeft + 200;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                //cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl2.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("THE MANAGER - CFS ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl2.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["CFSName"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl2.AddCell(cell1);

                var Addresss = Regex.Split(dtv.Rows[0]["CFSAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    Tbl2.AddCell(cell1);
                }

                //Tbl2.AddCell(cell1);
                mtable.AddCell(Tbl2);


                Tbl2 = new PdfPTable(1);
                Tbl2.WidthPercentage = 50;
                Tbl2.HorizontalAlignment = Element.ALIGN_RIGHT;

                cell1 = new PdfPCell(new Phrase("CHA", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                // cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl2.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["CHA"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl2.AddCell(cell1);

                var Addresss1 = Regex.Split(dtv.Rows[0]["CHAAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss1.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss1[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tbl2.AddCell(cell1);
                }

                // Tbl2.AddCell(cell1);

                mtable.AddCell(Tbl2);
                doc.Add(mtable);



                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace4 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace4);

                //------------------------------------------------------------------------
                #endregion

                #region Subject( Dear sir details)


                PdfPTable mtable1 = new PdfPTable(2);
                mtable1.WidthPercentage = 100;
                mtable1.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                mtable1.DefaultCell.Padding = 4;

                PdfPTable Tbl3 = new PdfPTable(1);
                Tbl3.WidthPercentage = 50;
                cell1 = new PdfPCell(new Phrase("Dear Sir,", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                //cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl3.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase("We shall be glad, if you will deliver the goods to M/S.", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl3.AddCell(cell1);

                //--blank---//
                cell1 = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl3.AddCell(cell1);
                mtable1.AddCell(Tbl3);


                Tbl3 = new PdfPTable(1);
                Tbl3.WidthPercentage = 50;
                Tbl3.HorizontalAlignment = Element.ALIGN_RIGHT;

                ///--No value for this balnk space---////
                cell1 = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell1.Colspan = 6;
                cell1.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell1.BorderWidth = 0;
                // cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.Colspan = 1;
                Tbl3.AddCell(cell1);


                cell1 = new PdfPCell(new Phrase(dtv.Rows[0]["DOParty"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell1.BorderWidth = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl3.AddCell(cell1);

                var Addresss2 = Regex.Split(dtv.Rows[0]["DOPartyAddress"].ToString(), "\r\n|\r|\n");

                for (int a = 0; a < Addresss2.Length; a++)
                {
                    cell1 = new PdfPCell(new Phrase(Addresss2[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell1.BorderWidth = 0;
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    Tbl3.AddCell(cell1);
                }


                mtable1.AddCell(Tbl3);
                doc.Add(mtable1);

                ////----------SPACE----------------------------------
                //iTextSharp.text.Table Tblspace5 = new iTextSharp.text.Table(1);
                //doc.Add(Tblspace5);

                ////------------------------------------------------------------------------


                iTextSharp.text.Table TblPara = new iTextSharp.text.Table(1);
                TblPara.Width = 60;
                TblPara.Alignment = Element.ALIGN_LEFT;
                TblPara.DefaultCell.Border = 0;
                TblPara.DefaultCellBorder = Rectangle.NO_BORDER;
                TblPara.Border = Rectangle.NO_BORDER;
                //TblPara.DefaultCell.UseBorderPadding;

                cell = new Cell(new Phrase("It is required to take proper receipt for the same", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                TblPara.AddCell(cell);
                doc.Add(TblPara);


                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace8 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace8);

                //------------------------------------------------------------------------
                #endregion

                #region TRANSPORTATION DETAILS

                PdfPTable Tbl4 = new PdfPTable(1);
                Tbl4.WidthPercentage = 103;
                //Tbl1.DefaultRowspan = 2;

                cell1 = new PdfPCell(new Phrase("TRANSPORTATION DETAILS", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell1.Colspan = 12;
                //cell.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_LEFT;
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 20f;
                cell1.BackgroundColor = new Color(98, 141, 214);
                cell1.Colspan = 1;
                Tbl4.AddCell(cell1);
                doc.Add(Tbl4);


                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(4);
                Tbl5.Width = 100;
                Tbl5.DefaultCell.Border = 0;
                Tbl5.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl5.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase("From", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("To", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("Vessel/VoyNo", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase(" ETA", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl5.AddCell(cell);


                //DataTable _dtSL = GetRRPDFMRGSLOTValus(ID);

                //for (int i = 0; i < _dtSL.Rows.Count; i++)
                //{
                cell = new Cell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["ETA"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tbl5.AddCell(cell);


                ///}

                doc.Add(Tbl5);



                iTextSharp.text.Table Tbl6 = new iTextSharp.text.Table(5);
                Tbl6.Width = 100;
                Tbl6.DefaultCell.Border = 0;
                Tbl6.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl6.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase("Volume", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase("Movement Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase("Item No", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase("IGM Number", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase("IGM Date", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);


                //DataTable _dtSL = GetRRPDFMRGSLOTValus(ID);

                //for (int i = 0; i < _dtSL.Rows.Count; i++)
                //{

                cell = new Cell(new Phrase("20 X " + dtv.Rows[0]["GP20"].ToString() + " 40 X " + dtv.Rows[0]["GP40"].ToString() + "", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["ServiceType"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["LineNumber"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["IGMNo"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["IGMDate"].ToString() + "\n \n", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);



                cell = new Cell(new Phrase("HBLNo", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase("HBLDate", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(".", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);


                //
                cell = new Cell(new Phrase(dtv.Rows[0]["HBLNo"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["HBLDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);

                cell = new Cell(new Phrase(".", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl6.AddCell(cell);



                ///}

                doc.Add(Tbl6);


                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace2);

                //------------------------------------------------------------------------

                #endregion

                #region Container Details

                PdfPTable Tbl51 = new PdfPTable(1);
                Tbl51.WidthPercentage = 103;
                //Tbl1.DefaultRowspan = 2;

                cell1 = new PdfPCell(new Phrase("CONTAINER DETAILS", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell1.Colspan = 12;
                //cell.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_LEFT;
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 20f;
                cell1.BackgroundColor = new Color(98, 141, 214);
                cell1.Colspan = 1;
                Tbl51.AddCell(cell1);
                doc.Add(Tbl51);

                PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2 });
                TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                TblLocs.SpacingBefore = 10;
                TblLocs.WidthPercentage = 100;

                cell1 = new PdfPCell(new Phrase("CONTAINER NUMBER/SIZE TYPE/COMMODITY", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                cell1 = new PdfPCell(new Phrase("DO VALIDITY UPTO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                //cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                DataTable dts = GetDoContDetails(id, DoID);

                for (int i = 0; i < dts.Rows.Count; i++)
                {
                    cell1 = new PdfPCell(new Phrase(dts.Rows[i]["CntrDetails"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(dts.Rows[i]["ValidityDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                    //cell1.BackgroundColor = new Color(152, 178, 209);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    TblLocs.AddCell(cell1);

                }

                doc.Add(TblLocs);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace7 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace7);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace9 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace9);

                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace10 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace10);

                #endregion

                #region second page

                PdfPTable Tbl52 = new PdfPTable(1);
                Tbl52.WidthPercentage = 103;
                //Tbl1.DefaultRowspan = 2;

                cell1 = new PdfPCell(new Phrase("SURVEYOUR", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
                cell1.Colspan = 12;
                //cell.HorizontalAlignment = 1;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.VerticalAlignment = Element.ALIGN_LEFT;
                cell1.BorderWidth = 1;
                cell1.FixedHeight = 20f;
                cell1.BackgroundColor = new Color(98, 141, 214);
                cell1.Colspan = 1;
                Tbl52.AddCell(cell1);
                doc.Add(Tbl52);


                iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                Tbl7.Width = 100;
                Tbl7.Alignment = Element.ALIGN_LEFT;
                Tbl7.DefaultCell.Border = 0;
                Tbl7.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl7.Border = Rectangle.NO_BORDER;




                cell = new Cell(new Phrase(dtv.Rows[0]["Surveyor"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl7.AddCell(cell);

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl7.AddCell(cell);

                var Addresss3 = Regex.Split(dtv.Rows[0]["SurveyorAddress"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < Addresss3.Length; a++)
                {
                    cell = new Cell(new Phrase(Addresss3[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Colspan = 2;
                    Tbl7.AddCell(cell);
                }


                doc.Add(Tbl7);


                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspace6 = new iTextSharp.text.Table(1);
                doc.Add(Tblspace6);

                //------------------------------------------------------------------------


                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspac13 = new iTextSharp.text.Table(1);
                doc.Add(Tblspac13);

                //------------------------------------------------------------------------


                //----------SPACE----------------------------------
                iTextSharp.text.Table Tblspac14 = new iTextSharp.text.Table(1);
                doc.Add(Tblspac14);

                //------------------------------------------------------------------------



                iTextSharp.text.Table Tbl8 = new iTextSharp.text.Table(2);
                Tbl8.Width = 100;
                Tbl8.Alignment = Element.ALIGN_LEFT;
                Tbl8.DefaultCell.Border = 0;
                Tbl8.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl8.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase("Delivery Order Remarks  : ", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl8.AddCell(cell);

                cell = new Cell(new Phrase(dtv.Rows[0]["Remarks"].ToString() + " \n \n \n \n \n \n \n \n \n \n \n  ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                Tbl8.AddCell(cell);

                doc.Add(Tbl8);



                #region Clauses

                //Sub Heading
                iTextSharp.text.Table Tbl9 = new iTextSharp.text.Table(1);
                Tbl9.Width = 100;
                Tbl9.Alignment = Element.ALIGN_LEFT;
                Tbl9.Cellpadding = 0;
                Tbl9.BorderWidth = 0;


                cell = new Cell(new Phrase("Clauses :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl9.AddCell(cell);
                doc.Add(Tbl9);

                iTextSharp.text.Table Tbl10 = new iTextSharp.text.Table(1);
                Tbl10.Width = 70;
                Tbl10.Alignment = Element.ALIGN_LEFT;
                Tbl10.Cellpadding = 1;
                Tbl10.BorderWidth = 0;

                DataTable _dtv = GetNotesClauses();
                if (_dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < _dtv.Rows.Count; i++)
                    {

                        cell = new Cell(new Phrase(_dtv.Rows[i]["Notes"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        Tbl10.AddCell(cell);

                    }
                }

                doc.Add(Tbl10);

                iTextSharp.text.Table Tbl11 = new iTextSharp.text.Table(2);
                Tbl11.Width = 100;
                Tbl11.Alignment = Element.ALIGN_LEFT;
                Tbl11.DefaultCell.Border = 0;
                Tbl11.DefaultCellBorder = Rectangle.NO_BORDER;
                Tbl11.Border = Rectangle.NO_BORDER;

                cell = new Cell(new Phrase("Please note that these Delivery Order is valid up to: ", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl11.AddCell(cell);

                cell = new Cell(new Phrase(dts.Rows[0]["ValidityDate"].ToString() + " \n \n \n \n \n \n \n \n \n \n \n \n", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tbl11.AddCell(cell);

                doc.Add(Tbl11);


                #endregion

                #region Signature and img

                iTextSharp.text.Table Tblsign = new iTextSharp.text.Table(1);
                Tblsign.Width = 100;
                Tblsign.Alignment = Element.ALIGN_LEFT;
                Tblsign.DefaultCell.Border = 0;
                Tblsign.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblsign.Border = Rectangle.NO_BORDER;



                cell = new Cell(new Phrase("FOR BLUEWAVE SHIPPING & LOGISTIC PVT LTD ", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblsign.AddCell(cell);



               

                cell = new Cell(new Phrase("As Agent", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblsign.AddCell(cell);

                cell = new Cell(new Phrase("(BLUE WAVE SHIPPING & LOGISTIC PTE LTD)", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                Tblsign.AddCell(cell);
                doc.Add(Tblsign);

                iTextSharp.text.Table Tblfoot = new iTextSharp.text.Table(1);
                Tblfoot.Width = 100;
                Tblfoot.Alignment = Element.ALIGN_CENTER;
                Tblfoot.DefaultCell.Border = 0;
                Tblfoot.DefaultCellBorder = Rectangle.NO_BORDER;
                Tblfoot.Border = Rectangle.NO_BORDER;



                cell = new Cell(new Phrase("Note : This is computer generated document Hence No signature required", new Font(Font.HELVETICA, 8, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                Tblfoot.AddCell(cell);
                doc.Add(Tblfoot);
                #endregion



                #endregion

                #region second page border
                // Add border to page

                Rectangle rectangle = new Rectangle(doc.PageSize);
                rectangle.Left += doc.LeftMargin - 8;
                rectangle.Right -= doc.RightMargin;
                rectangle.Top -= doc.TopMargin;
                rectangle.Bottom += doc.BottomMargin;
                content.SetColorStroke(Color.BLACK);
                content.Rectangle(rectangle.Left, rectangle.Bottom, rectangle.Width, rectangle.Height);
                content.Stroke();
                #endregion


                pdfWriter.CloseStream = false;
                doc.Close();
            }
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            //Response.AddHeader("content-disposition", "attachment;filename=DeliveryOrder.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();

        }
        public DataTable GetNotesClauses()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=267";
            return Manag.GetViewData(_Query, "");
        }
        public ActionResult EmptyReturnPDF(string id, string bkgid, string DoId, string AgencyID, string LocID)
        {
            CreateEmptyReturnPDF(id, bkgid, DoId, AgencyID, LocID);
            return View();

        }

        public void CreateEmptyReturnPDF(string id, string bkgid, string DoId, string AgencyID, string LocID)
        {
            DataTable dtv = GetEmptyReturnPrint(id, bkgid, DoId, AgencyID);
            //    DataTable dtv = GetBkgPDFValus(BkgID);
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

            iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(4);
            tbllogo.Width = 60;

            //tbllogo.Alignment = Element.ALIGN_LEFT;
            //tbllogo.Cellpadding = 1;
            tbllogo.BorderWidth = 0;
            Cell cell = new Cell();
            cell.Width = 10;


            var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/gnllogo.png"));
            img.Alignment = Element.ALIGN_LEFT;
            img.ScaleAbsolute(80f, 80f);
            cell = new Cell(img);
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            tbllogo.AddCell(cell);


            DataTable dtx = GetAgencyAddress(AgencyID);
            cell = new Cell(new Phrase(dtx.Rows[0]["AgencyName"].ToString()  + "  " + dtx.Rows[0]["Address"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.Colspan = 3;
            cell.BorderWidth = 0;
            tbllogo.AddCell(cell);

            doc.Add(tbllogo);


            para = new Paragraph("");
            doc.Add(para);

            para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLUE));
            para.Alignment = Element.ALIGN_RIGHT;
            doc.Add(para);

            ////----------SPACE----------------------------------
            //iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
            //doc.Add(Tblspace2);

            ////------------------------------------------------------------------------
            #endregion

            PdfPTable Tbl1 = new PdfPTable(1);
            Tbl1.WidthPercentage = 103;
            //Tbl1.DefaultRowspan = 2;

            PdfPCell cell1 = new PdfPCell(new Phrase("EMPTY RETURN", new Font(Font.HELVETICA, 12, Font.BOLD, Color.WHITE)));
            cell1.Colspan = 12;
            //cell.HorizontalAlignment = 1;
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.VerticalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidth = 1;
            cell1.FixedHeight = 25f;
            cell1.BackgroundColor = new Color(98, 141, 214);
            cell1.Colspan = 1;
            Tbl1.AddCell(cell1);
            doc.Add(Tbl1);


            // Add border to page
            PdfContentByte content = pdfWriter.DirectContent;
            Rectangle rectangle = new Rectangle(doc.PageSize);
            rectangle.Left += doc.LeftMargin - 9;
            rectangle.Right -= doc.RightMargin - 9;
            rectangle.Top -= doc.TopMargin + 113;
            rectangle.Bottom += doc.BottomMargin;
            content.SetColorStroke(Color.BLACK);
            content.Rectangle(rectangle.Left, rectangle.Bottom, rectangle.Width, rectangle.Height);
            content.Stroke();

            #region DO DETAILS, DATE, MBL NUMBER BL DATE,LINE NUMBER
            // /----------------------- LocTable-----------------------///

            iTextSharp.text.Table TblDOdtls = new iTextSharp.text.Table(6);
            TblDOdtls.Width = 103;
            TblDOdtls.Alignment = Element.ALIGN_CENTER;
            TblDOdtls.Cellpadding = 1;
            // TblDOdtls.DefaultCellBorder = Rectangle.NO_BORDER;
            // TblDOdtls.Border = Rectangle.NO_BORDER;
            TblDOdtls.DefaultCell.Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
            TblDOdtls.BorderWidth = 1;

            cell = new Cell(new Phrase("DO NUMBER :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell.BackgroundColor = new Color(98, 141, 214);
            //cell.BorderWidth = 1;
            cell.Colspan = 2;
            cell.BorderWidthLeft = -8;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase("DO DATE :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase("MBL NUMBER :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell.BackgroundColor = new Color(98, 141, 214);

            cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase("BL DATE :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase("LINE NUMBER :", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell.BackgroundColor = new Color(98, 141, 214);
            // cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            /////values//////
            cell = new Cell(new Phrase(dtv.Rows[0]["DONo"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.WHITE)));
            // cell.BorderWidth = 1;
            cell.Colspan = 2;
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["IssueDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.WHITE)));
            cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["BLNumber"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.WHITE)));
            cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["BLDate"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.WHITE)));
            cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["LineNumber"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.WHITE)));
            // cell.BorderWidth = 1;
            cell.Colspan = 1;
            cell.BackgroundColor = new Color(98, 141, 214);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            TblDOdtls.AddCell(cell);

            doc.Add(TblDOdtls);

            #endregion

            #region To details

            iTextSharp.text.Table TblTO = new iTextSharp.text.Table(2);
            TblTO.Width = 60;
            TblTO.Alignment = Element.ALIGN_LEFT;
            TblTO.DefaultCell.Border = 0;
            TblTO.DefaultCellBorder = Rectangle.NO_BORDER;
            TblTO.Border = Rectangle.NO_BORDER;

            cell = new Cell(new Phrase("To : " + dtv.Rows[0]["ReturnDepo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.BorderWidthLeft = -8f;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            TblTO.AddCell(cell);

            //cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
            //cell.BorderWidth = 0;
            //cell.Colspan = 1;
            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //TblTO.AddCell(cell);

            var Addresss3 = Regex.Split(dtv.Rows[0]["DepoAddress"].ToString().Trim(), "\r\n|\r|\n");
            for (int a = 0; a < Addresss3.Length; a++)
            {
                cell = new Cell(new Phrase(Addresss3[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.BorderWidthLeft = -8;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 2;
                TblTO.AddCell(cell);
            }


            doc.Add(TblTO);



            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace4 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace4);

            //------------------------------------------------------------------------
            #endregion

            #region Shipment DETAILS
            PdfPTable Tbl4 = new PdfPTable(1);
            Tbl4.WidthPercentage = 103;
            //Tbl1.DefaultRowspan = 2;

            cell1 = new PdfPCell(new Phrase("SHIPMENT DETAILS", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell1.Colspan = 12;
            //cell.HorizontalAlignment = 1;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.VerticalAlignment = Element.ALIGN_LEFT;
            cell1.BorderWidth = 1;
            cell1.BorderWidthLeft = -8;
            cell1.FixedHeight = 20f;
            cell1.BackgroundColor = new Color(98, 141, 214);
            cell1.Colspan = 1;
            Tbl4.AddCell(cell1);
            doc.Add(Tbl4);


            iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(8);
            Tbl5.Width = 100;
            Tbl5.DefaultCell.Border = 0;
            Tbl5.DefaultCellBorder = Rectangle.NO_BORDER;
            Tbl5.Border = Rectangle.NO_BORDER;

            cell = new Cell(new Phrase("From", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase("To", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase("Vessel/VoyNo", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase(" ETA", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase("Volume", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);
            //DataTable _dtSL = GetRRPDFMRGSLOTValus(ID);

            //for (int i = 0; i < _dtSL.Rows.Count; i++)
            //{
            cell = new Cell(new Phrase(dtv.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 2;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["ETA1"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);

            cell = new Cell(new Phrase("20 X " + dtv.Rows[0]["GP20"].ToString() + " 40 X " + dtv.Rows[0]["GP40"].ToString() + "", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl5.AddCell(cell);
            ///}
            ///
            //




            ///}

            doc.Add(Tbl5);


            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace2);

            //------------------------------------------------------------------------





            iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
            Tbl7.Width = 100;
            Tbl7.Alignment = Element.ALIGN_LEFT;
            Tbl7.DefaultCell.Border = 0;
            Tbl7.DefaultCellBorder = Rectangle.NO_BORDER;
            Tbl7.Border = Rectangle.NO_BORDER;

            cell = new Cell(new Phrase("Please accept below mentioned empty container(s) at your yard on or before the valid date of return mentioined against below container from ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.BorderWidthLeft = -8;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            Tbl7.AddCell(cell);




            doc.Add(Tbl7);


            #endregion

            #region Consignee details

            iTextSharp.text.Table TblConsignee = new iTextSharp.text.Table(1);
            TblConsignee.Width = 90;
            TblConsignee.Alignment = Element.ALIGN_LEFT;
            TblConsignee.DefaultCell.Border = 0;
            TblConsignee.DefaultCellBorder = Rectangle.NO_BORDER;
            TblConsignee.Border = Rectangle.NO_BORDER;

            cell = new Cell(new Phrase("CONSIGNEE", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.BorderWidthLeft = -8;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            TblConsignee.AddCell(cell);

            cell = new Cell(new Phrase(dtv.Rows[0]["Consignee"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            TblConsignee.AddCell(cell);

            var Addresss4 = Regex.Split(dtv.Rows[0]["ConsigneeAddress"].ToString(), "\r\n|\r|\n");
            for (int a = 0; a < Addresss4.Length; a++)
            {
                cell = new Cell(new Phrase(Addresss4[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.BorderWidthLeft = -8;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 2;
                TblConsignee.AddCell(cell);
            }


            doc.Add(TblConsignee);



            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace5 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace5);

            //------------------------------------------------------------------------
            #endregion



            #region Container Details

            PdfPTable Tbl51 = new PdfPTable(1);
            Tbl51.WidthPercentage = 103;
            //Tbl1.DefaultRowspan = 2;

            cell1 = new PdfPCell(new Phrase("CONTAINER DETAILS", new Font(Font.HELVETICA, 10, Font.BOLD, Color.WHITE)));
            cell1.Colspan = 12;
            //cell.HorizontalAlignment = 1;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.VerticalAlignment = Element.ALIGN_LEFT;
            cell1.BorderWidth = 1;
            cell1.FixedHeight = 20f;
            cell1.BackgroundColor = new Color(98, 141, 214);
            cell1.Colspan = 1;
            Tbl51.AddCell(cell1);
            doc.Add(Tbl51);

            PdfPTable TblLocs = new PdfPTable(new float[] { 2 });
            TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
            TblLocs.SpacingBefore = 10;
            TblLocs.WidthPercentage = 100;

            cell1 = new PdfPCell(new Phrase("CONTAINER NUMBER/SIZE TYPE/COMMODITY", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            //cell1.BackgroundColor = new Color(152, 178, 209);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            TblLocs.AddCell(cell1);

            //cell1 = new PdfPCell(new Phrase("DO VALIDITY UPTO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            ////cell1.BackgroundColor = new Color(152, 178, 209);
            //cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            //TblLocs.AddCell(cell1);

            DataTable dts = GetDoContDetails(id, DoId);

            for (int i = 0; i < dts.Rows.Count; i++)
            {
                cell1 = new PdfPCell(new Phrase(dts.Rows[i]["CntrDetails"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                //cell1.BackgroundColor = new Color(152, 178, 209);
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                TblLocs.AddCell(cell1);

                //cell1 = new PdfPCell(new Phrase(dts.Rows[i]["ValidityDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                ////cell1.BackgroundColor = new Color(152, 178, 209);
                //cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                //TblLocs.AddCell(cell1);

            }

            doc.Add(TblLocs);

            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace7 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace7);

            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace9 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace9);

            //----------SPACE----------------------------------
            iTextSharp.text.Table Tblspace10 = new iTextSharp.text.Table(1);
            doc.Add(Tblspace10);

            PdfPTable Tbl9 = new PdfPTable(2);
            Tbl9.WidthPercentage = 103;
            //Tbl1.DefaultRowspan = 2;

            cell1 = new PdfPCell(new Phrase("CONTAINERS ARE TO BE RETURNED ON OR BEFORE:", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell1.Colspan = 12;
            //cell.HorizontalAlignment = 1;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.VerticalAlignment = Element.ALIGN_LEFT;
            cell1.BorderWidth = 1;
            cell1.BorderWidthLeft = -8;
            cell1.FixedHeight = 24f;
            cell1.Colspan = 1;
            Tbl9.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase(dts.Rows[0]["ValidityDate"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
            Tbl9.AddCell(cell1);

            doc.Add(Tbl9);

            #endregion


            //#region space
            //iTextSharp.text.Table Tbl11 = new iTextSharp.text.Table(1);
            //Tbl11.Width = 100;
            //Tbl11.Alignment = Element.ALIGN_LEFT;
            //Tbl11.DefaultCell.Border = 0;
            //Tbl11.DefaultCellBorder = Rectangle.NO_BORDER;
            //Tbl11.Border = Rectangle.NO_BORDER;


            //cell = new Cell(new Phrase("" + " \n ", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            //cell.BorderWidth = 0;
            //cell.Colspan = 1;
            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //Tbl11.AddCell(cell);

            //doc.Add(Tbl11);

            //#endregion

            #region Signature and img

            iTextSharp.text.Table Tblsign = new iTextSharp.text.Table(1);
            Tblsign.Width = 100;
            Tblsign.Alignment = Element.ALIGN_RIGHT;
            Tblsign.DefaultCell.Border = 0;
            Tblsign.DefaultCellBorder = Rectangle.NO_BORDER;
            Tblsign.Border = Rectangle.NO_BORDER;

            cell = new Cell(new Phrase("Your Sincerely", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            Tblsign.AddCell(cell);

            cell = new Cell(new Phrase(dtx.Rows[0]["AgencyName"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            Tblsign.AddCell(cell);


            cell = new Cell(new Phrase("AUTHORISED SIGNATORY", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            Tblsign.AddCell(cell);
            doc.Add(Tblsign);



            iTextSharp.text.Table Tblfoot = new iTextSharp.text.Table(1);
            Tblfoot.Width = 100;
            Tblfoot.DefaultCell.Border = 0;
            Tblfoot.Alignment = Element.ALIGN_CENTER;
            Tblfoot.DefaultCellBorder = Rectangle.NO_BORDER;
            Tblfoot.Border = Rectangle.NO_BORDER;



            cell = new Cell(new Phrase("Note : This is computer generated document Hence No signature required", new Font(Font.HELVETICA, 8, Font.BOLD | Font.UNDERLINE, Color.BLACK)));
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
            //Response.AddHeader("content-disposition", "attachment;filename=EmptyReturnPDF.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();

        }


        public DataTable GetDOPrint(string id, string DoID)
        {
            string _Query = "select * from V_NVOImpDOPrintValueNew where Id =" + id + " AND DOID =" + DoID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetDoContDetails(string id, string DoID)
        {
            string _Query = " select distinct NVO_Containers.ID,(NVO_Containers.CntrNo+'/'+Size+'/'+CommodityType)as CntrDetails , " +
                        " convert(varchar, NVO_BOLDOCntrdtls.ValidityDate, 105) as ValidityDate from NVO_BOLDOCntrdtls INNER JOIN NVO_BOLCntrDetails ON NVO_BOLCntrDetails.BLID = NVO_BOLDOCntrdtls.BLID AND NVO_BOLDOCntrdtls.CntrID = NVO_BOLCntrDetails.CntrID " +
                       " inner join NVO_Booking On NVO_Booking.ID = NVO_BOLDOCntrdtls.BkgId inner join  NVO_Containers On NVO_Containers.ID = NVO_BOLDOCntrdtls.CntrID  where NVO_BOLDOCntrdtls.BLID=" + id + " AND DoID =" + DoID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetEmptyReturnPrint(string id, string bkgid, string DoId, string AgencyID)
        {
            string _Query = "select * from V_NVOImpDOPrintValueNew where Id =" + id + " and BkgID=" + bkgid + " and DOID=" + DoId;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetAgencyAddress(string AgencyID)
        {
            string _Query = " select * from NVO_AgencyMaster where ID =" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }



    }
}