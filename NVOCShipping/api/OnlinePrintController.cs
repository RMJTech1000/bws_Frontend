
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Mvc;
using DataTier;
using iTextSharp.text;
using QRCoder;
using DataManager;
using System.Net.Mail;
using System.Text;
using System.Data;
using System.IO;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using System.Web;


namespace NVOCShipping.api
{

    public class OnlinePrintController : ApiController
    {

        DocumentManager Manag = new DocumentManager();


        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getCroPFD")]
        public HttpResponseMessage getCroPFD([FromBody] MyInvoice Data)
        {

            MemoryStream memoryStream = new MemoryStream();
            DataTable dtc = GetAgencyDetails(Data.AgentId);
            DataTable dt = GetCROPDFValus(Data.ID.ToString());
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

                var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                img.Alignment = Element.ALIGN_LEFT;
                img.ScaleAbsolute(45f, 45f);
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
                cell1 = new PdfPCell(new Phrase("Release Order Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                cell1.BorderWidth = 0;
                cell1.Colspan = 1;
                Tbl3.AddCell(cell1);
                doc.Add(Tbl3);

                iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
                Tbl5.Width = 100;
                Tbl5.Alignment = Element.ALIGN_LEFT;
                Tbl5.Cellpadding = 1;
                Tbl5.BorderWidth = 0;

                //Caption
                cell = new Cell(new Phrase("Vessel & Voyage ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["VesVoy"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("ETA / ETD", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + dt.Rows[0]["ETADate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("Cut – Off Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["CUTDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("Line Code", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + dt.Rows[0]["Linecode"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                cell.Colspan = 3;
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


                DataTable dtCroDtls = GetCRODetailsPDFValues(Data.ID.ToString());
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


                cell = new Cell(new Phrase(" * Ensure the empty container received from our yard is in clean and sound condition. Costs for any subsequent rejection will be to your account.", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                Tbl7.AddCell(cell);

                cell = new Cell(new Phrase(" * Any loss or damage to the container while in custody of shipper, transporter, forwarder shal be fully identified for repair / replacement / reimbursement as notified by owner / hirer. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                Tbl7.AddCell(cell);

                cell = new Cell(new Phrase(" *  loading list needs to be sent to the shipping line by 48 hour prior to vessel cut off. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                Tbl7.AddCell(cell);

                cell = new Cell(new Phrase(" * For Non-availability of Containers kindly contact our Operations Incharge - Mr.Nivrutti - (M) +91-90223 45131 will be on your account.", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                Tbl7.AddCell(cell);

                cell = new Cell(new Phrase(" *FORM 13 to be collected & Shipping Bill to be handed over to our surveyor as mentioned in above. ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                cell.BorderWidth = 0;
                Tbl7.AddCell(cell);

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
                byte[] byteInfo = memoryStream.ToArray();
                memoryStream.Write(byteInfo, 0, byteInfo.Length);
                memoryStream.Position = 0;

            }


            // Create an HTTP response message
            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };

            // Set the content headers
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "yourfile.pdf"
            };

            return response;
        }



        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getBookingPFD")]
        public HttpResponseMessage getBookingPFD([FromBody] MyInvoice Data)
        {

            MemoryStream memoryStream = new MemoryStream();
            DataTable dtc = GetAgencyDetails(Data.AgentId);
            DataTable dtv = GetBkgPDFValus(Data.BkgID.ToString());
            if (dtv.Rows.Count > 0)
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

                DataTable dta = GetCompanyDetails();
               
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
                if (Data.AgentId == "3")
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
                if (Data.AgentId == "3")
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

                if (Data.AgentId != "3")
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

                if (Data.GeoLocID == "1" || Data.GeoLocID == "2" || Data.GeoLocID == "3")
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

                #endregion
                pdfWriter.CloseStream = false;
                doc.Close();
                byte[] byteInfo = memoryStream.ToArray();
                memoryStream.Write(byteInfo, 0, byteInfo.Length);
                memoryStream.Position = 0;

            }


            // Create an HTTP response message
            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };

            // Set the content headers
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "Booking.pdf"
            };

            return response;
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


        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getInvoivcepdf")]
        public HttpResponseMessage getInvoivcepdf([FromBody] MyInvoice Data)
        {
            MemoryStream memoryStream = new MemoryStream();


            MergeEx pdfmp = new MergeEx();
            DataTable dtv = GetInvPDFValus(Data.ID.ToString());
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 870);
                doc = new Document(rec);
                Paragraph para = new Paragraph();

                string _FileName = Data.ID.ToString() + 1;
                PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 830, 0);
                Data.AgentId = dtv.Rows[0]["AgentID"].ToString();
                DataTable dtc = GetAgencyDetails(Data.AgentId);

                DataTable _dtc = GetCompanyDetails();
                if (dtc.Rows.Count > 0)
                {
                    if (Data.AgentId == "13")
                    {

                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/neridaaddress1.png"));
                        png2.SetAbsolutePosition(400, 810);
                        png2.ScalePercent(37f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "79")
                    {
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/neridaaddress1.png"));
                        png2.SetAbsolutePosition(400, 810);
                        png2.ScalePercent(37f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "14")
                    {
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddressmundra.png"));
                        png2.SetAbsolutePosition(400, 810);
                        png2.ScalePercent(37f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "15")
                    {
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddressdelhi.png"));
                        png2.SetAbsolutePosition(400, 810);
                        png2.ScalePercent(37f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "16")
                    {
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/neridaaddress1.png"));
                        png2.SetAbsolutePosition(400, 810);
                        png2.ScalePercent(37f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "2")
                    {
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/ocnlogo.png"));
                        png2.SetAbsolutePosition(335, 810);
                        png2.ScalePercent(17f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "1")
                    {
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddresschennai.jpeg"));
                        png2.SetAbsolutePosition(400, 810);
                        png2.ScalePercent(37f);
                        doc.Add(png2);
                    }
                    if (Data.AgentId == "100" || Data.AgentId == "101")
                    {
                        //if (_dtc.Rows[0]["CompanyID"].ToString() == "1" || _dtc.Rows[0]["CompanyID"].ToString() == "2")
                        //{
                        iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddressjeblali.png"));
                        png2.SetAbsolutePosition(350, 805);
                        png2.ScalePercent(17f);
                        doc.Add(png2);
                        //}
                    }
                }


                cb.BeginText();
                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 15);
                cb.SetColorFill(Color.BLACK);
                //center
                //cb.MoveTo(345, 835);
                //cb.LineTo(345, 600);
                if (dtv.Rows[0]["InvTypes"].ToString() == "1")
                {

                    if (dtv.Rows[0]["FinalInvoice"].ToString() != "")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "TAX INVOICE", 200, 830, 0);
                    }
                    else
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "PROFORMA INVOICE", 200, 830, 0);
                    }
                }
                else
                {
                    if (dtv.Rows[0]["FinalInvoice"].ToString() != "")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "CREDIT NOTE", 200, 830, 0);
                    }
                    else
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "PROFORMA CREDIT NOTE", 200, 830, 0);
                    }
                }

                if (dtv.Rows[0]["FinalInvoice"].ToString() != "" && dtv.Rows[0]["SingnedQRCode"].ToString() != "")
                {
                    QRCodeGenerator _qrCode = new QRCodeGenerator();
                    QRCodeData _qrCodeData = _qrCode.CreateQrCode(dtv.Rows[0]["SingnedQRCode"].ToString(), QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(_qrCodeData);
                    System.Drawing.Bitmap qrCodeImage = qrCode.GetGraphic(20);


                    BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader, 14);
                    iTextSharp.text.Image png11 = iTextSharp.text.Image.GetInstance(BitmapToBytesCode(qrCodeImage));
                    png11.SetAbsolutePosition(30, 806);
                    png11.ScalePercent(2f);
                    doc.Add(png11);
                }

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
                //cb.MoveTo(10, 115);
                //cb.LineTo(660, 115);
                //left//
                cb.MoveTo(10, 805);
                cb.LineTo(10, 35);
                //right//      
                cb.MoveTo(660, 805);
                cb.LineTo(660, 35);

                //center
                //cb.MoveTo(330, 805);
                //cb.LineTo(330, 680);

                //cb.MoveTo(695, 935);
                //cb.LineTo(695, 842);
                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);



                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL TO", 15, 760, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 760, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "State Name", 15, 700, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 700, 0);


                if (Data.AgentId == "100" || Data.AgentId == "101")
                {

                    cb.SetFontAndSize(bfheader2, 8);
                    cb.SetColorFill(Color.BLACK);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TRN#:", 15, 688, 0);

                }
                else
                {

                    cb.SetFontAndSize(bfheader2, 8);
                    cb.SetColorFill(Color.BLACK);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GSTIN#:", 15, 688, 0);

                }
                //cb.SetFontAndSize(bfheader2, 8);
                //cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GSTIN#:", 15, 688, 0);
                //cb.EndText();
                //cb.BeginText();
                //cb.SetFontAndSize(bfheader2, 8);
                //cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 688, 0);





                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice No", 15, 795, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Invoice Date", 200, 795, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Due Date", 345, 795, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL Number", 500, 795, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel/Voyage", 15, 670, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 170, 670, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 250, 670, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 345, 670, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POO", 470, 670, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POF", 580, 670, 0);




                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage ETA", 15, 641, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage ETD", 170, 641, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Liner BL", 345, 641, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Volume", 470, 641, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gr.Wt (KGS/MT)", 560, 641, 0);




                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IRN :", 15, 615, 0);



                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Details :", 15, 602, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Line :", 450, 602, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Services", 15, 545, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SAC", 170, 545, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RatePer(Unit)", 220, 545, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Quantity", 290, 545, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Ex.Rate", 350, 545, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Taxable", 390, 545, 0);



                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Amount In " + dtv.Rows[0]["Currency"].ToString(), 500, 555, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 445, 535, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SGST", 480, 535, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 515, 535, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CGST", 550, 535, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "%", 595, 535, 0);

                cb.SetFontAndSize(bfheader2, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IGST", 630, 535, 0);

                //cb.BeginText();
                //cb.SetFontAndSize(bfheader2, 8);
                //cb.SetColorFill(Color.BLACK);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORDS", 15, 300, 0);
                //cb.EndText();






                DataTable dtvs = GetInvPDFDtlValues(dtv.Rows[0]["BLID"].ToString(), Data.ID.ToString());

                BaseFont bfheader31 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader31, 9);
                cb.SetColorFill(Color.BLACK);

                if (dtv.Rows[0]["TaxExemption"].ToString() == "YES")
                    if (Data.AgentId == "2" || Data.AgentId == "101" || Data.AgentId == "100")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 500, 760, 0);

                    }
                    else
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SEZ TYPE:", 500, 760, 0);
                    }


                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 9);
                cb.SetColorFill(Color.BLACK);

                if (dtv.Rows[0]["TaxExemption"].ToString() == "YES")
                    if (Data.AgentId == "2" || Data.AgentId == "101" || Data.AgentId == "100")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 550, 760, 0);
                    }
                    else
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SEZ PARTY", 550, 760, 0);
                    }

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["PartyName"].ToString(), 15, 745, 0);
                cb.SetFontAndSize(bfheader3, 7);
                int ColumnRows = 735;
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



                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Address"].ToString(), 15, 736, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 745, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 736, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["StateCode"].ToString().ToUpper(), 100, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["GSTIN"].ToString().ToUpper(), 100, 688, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 700, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 410, 688, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["FinalInvoice"].ToString().ToUpper(), 15, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDate"].ToString(), 200, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvDueDate"].ToString(), 345, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["BookingNo"].ToString().ToUpper(), 500, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["VesVoy"].ToString().ToUpper(), 15, 656, 0);
                var POLSplit = dtvs.Rows[0]["POL"].ToString().ToUpper().Split('-');
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, POLSplit[0].ToString(), 170, 656, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 345, 656, 0);

                var PODSplit = dtvs.Rows[0]["POD"].ToString().ToUpper().Split('-');
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, PODSplit[0].ToString(), 345, 656, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["POD"].ToString().ToUpper(), 345, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["POO"].ToString().ToUpper(), 450, 656, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["FPOD"].ToString().ToUpper(), 580, 656, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["ETA"].ToString(), 15, 631, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["ETD"].ToString(), 170, 631, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["IRNNo"].ToString(), 45, 615, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["CntrCount"].ToString(), 470, 631, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["GrsWt"].ToString(), 570, 631, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtvs.Rows[0]["CompanyName"].ToString(), 450, 585, 0);
                int ColumnRow = 15;
                int RowInx = 585;
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
                for (int i = 0; i < dtInvDtls.Rows.Count; i++)
                {
                    var splitDesc = SplitByLenght(dtInvDtls.Rows[i]["NarrationDescription"].ToString(), 33);
                    for (int k = 0; k < splitDesc.Length; k++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitDesc[k].ToString(), 15, RowGrd, 0);
                        if (splitDesc.Length >= 2)
                        {
                            if (k == 0)
                            {
                                RowGrd -= 10;
                            }
                        }
                    }
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["SACCode"].ToString(), 170, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["RatePerUnit"].ToString(), 257, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtInvDtls.Rows[i]["Currency"].ToString(), 262, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["Qty"].ToString() + " x " + dtInvDtls.Rows[i]["Size"].ToString(), 315, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["ROE"].ToString(), 375, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["LocalAmount"].ToString(), 430, RowGrd, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["SGSTPec"].ToString(), 458, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["SGSTAmt"].ToString(), 505, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["CGSTPec"].ToString(), 530, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["CGSTAmt"].ToString(), 585, RowGrd, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["IGSTPec"].ToString(), 610, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, dtInvDtls.Rows[i]["IGSTAmt"].ToString(), 655, RowGrd, 0);
                    SGSTAmt += decimal.Parse(dtInvDtls.Rows[i]["SGSTAmt"].ToString());
                    CGSTAmt += decimal.Parse(dtInvDtls.Rows[i]["CGSTAmt"].ToString());
                    IGSTAmt += decimal.Parse(dtInvDtls.Rows[i]["IGSTAmt"].ToString());

                    RowGrd -= 12;
                    if (i == 19)
                    {
                        break;
                    }
                }

                if (dtInvDtls.Rows.Count <= 20)
                {

                    int RowGrdv = 190;
                    DataTable _dtn = GetNotesClauses();
                    for (int d = 0; d < _dtn.Rows.Count; d++)
                    {
                        string[] ArrayNotes = Regex.Split(_dtn.Rows[d]["Notes"].ToString().Trim(), char.ConvertFromUtf32(13));
                        string[] Aaddsplitv;
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, (d + 1).ToString(), 20, RowGrdv, 0);
                        for (int x = 0; x < ArrayNotes.Length; x++)
                        {
                            Aaddsplitv = ArrayNotes[x].Split('\n');
                            for (int k = 0; k < Aaddsplitv.Length; k++)
                            {
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplitv[k].ToString(), 30, RowGrdv, 0);
                                RowGrdv -= 9;

                            }
                        }
                    }

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 130, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["CreatedBy"].ToString(), 15, 40, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Authorised Signatory", 530, 40, 0);

                    if (Data.AgentId == "14")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "# 04, 1st Floor, Kesar Arcade, Plot No.51, Sector 8, NR.Chamber Of Commerce, Gandhidham(Kutch), 370201 Gujarat, India", 175, 20, 0);
                    }
                    else if (Data.AgentId == "2")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 175, 20, 0);
                    }
                    else if (Data.AgentId == "1")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEW NO.171/OLD NO.83,2ND FLOOR,LINGHI CHETTY STREET,CHENNAI-600001,TAMIL NADU", 170, 20, 0);

                    }
                    else if (Data.AgentId == "100" || Data.AgentId == "101")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Al Abbas Building No 2, First Floor, Room No 114. P.O.Box No. 128661, Bur Dubai, Dubai.", 175, 20, 0);
                    }
                    else
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "4th Floor, Office No. 416-417, Avior Corporate Park Co-op Society Ltd,Nirmal Galaxy, LBS Road, Mulund West, Mumbai, Maharashtra, 400080.", 175, 20, 0);
                    }



                    BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader5, 8);
                    cb.SetColorFill(Color.BLACK);


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 270, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 350, 175, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 550, 170, 0);
                    BaseFont bfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader6, 8);
                    cb.SetColorFill(Color.BLACK);
                    decimal TotalAmount = 0;

                    TotalAmount = decimal.Parse(dtv.Rows[0]["InvTotal"].ToString());
                    decimal TotalW = Convert.ToDecimal(TotalAmount.ToString("#,#0.00"));
                    string Rupes = NumberConverWords.changeCurrencyToWords(TotalW.ToString());
                    string Ruppev = "";
                    if (dtvs.Rows[0]["CurrencyID"].ToString() == "146")
                    {
                        Ruppev = Rupes.Replace("Rupees", "DOLLAR");
                    }
                    if (dtvs.Rows[0]["CurrencyID"].ToString() == "118")
                    {
                        Ruppev = Rupes.Replace("Rupees", "RUBLE");
                    }
                    if (dtvs.Rows[0]["CurrencyID"].ToString() == "2")
                    {
                        Ruppev = Rupes.Replace("Rupees", "DIRHAM");
                    }
                    if (dtvs.Rows[0]["CurrencyID"].ToString() == "1")
                    {
                        Ruppev = Rupes.Replace("Rupees", "Rupees");
                    }




                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Ruppev.ToUpper(), 15, 280, 0);

                    BaseFont bfheader8 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader8, 8);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["InvAmount"].ToString(), 390, 280, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, SGSTAmt.ToString(), 470, 295, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, CGSTAmt.ToString(), 550, 295, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, IGSTAmt.ToString(), 623, 295, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TOTAL AMOUNT   " + dtv.Rows[0]["InvTotal"].ToString(), 470, 278, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Remarks :", 15, 260, 0);

                    cb.SetFontAndSize(bfheader2, 8);
                    cb.SetColorFill(Color.BLACK);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AMOUNT IN WORDS", 15, 300, 0);




                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Remarks"].ToString(), 15, 250, 0);
                    string[] splitremarks = Regex.Split(dtv.Rows[0]["Remarks"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(165));

                    ColumnRows = 249;
                    RowsColumn = 0;
                    string[] remarksplit;

                    for (int x = 0; x < splitremarks.Length; x++)
                    {
                        remarksplit = splitremarks[x].Split('\n');

                        for (int k = 0; k < remarksplit.Length; k++)
                        {

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, remarksplit[k].ToString(), 15, ColumnRows, 0);
                            ColumnRows -= 9;
                            RowsColumn++;
                        }
                    }


                    if (dtv.Rows[0]["CreditNoteType"].ToString() == "1")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Original Dr.Note No:   " + dtv.Rows[0]["DebitInvoice"].ToString(), 125, 250, 0);
                    }


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "E & O E", 15, 200, 0);

                    DataTable _dtb = GetInvPDFBankValues(dtv.Rows[0]["IntBank"].ToString());
                    if (_dtb.Rows.Count > 0)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name & Details", 15, 135, 0);

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Benificiary", 15, 120, 0);
                        if (Data.AgentId == "2")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OCEANUS CONTAINER LINES PTE LTD", 100, 120, 0);
                        }
                        else if (Data.AgentId == "100" || Data.AgentId == "101")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NERIDA SHIPPING SERVICES LLC", 100, 120, 0);
                        }
                        else
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NERIDA SHIPPING PVT LTD", 100, 120, 0);
                        }

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bank Name", 15, 105, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["BankName"].ToString(), 100, 105, 0);
                        if (Data.AgentId == "2")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SWIFT CODE", 15, 85, 0);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["SwiftCode"].ToString(), 100, 90, 0);
                        }
                        else if (Data.AgentId == "100" || Data.AgentId == "101")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IBAN Number", 15, 90, 0);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["IFSCCode"].ToString(), 100, 90, 0);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SWIFT CODE", 15, 75, 0);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["SwiftCode"].ToString(), 100, 75, 0);
                        }
                        else
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "IFSC", 15, 85, 0);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["IFSCCode"].ToString(), 100, 85, 0);
                        }

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Payment Ref No", 350, 115, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["FinalInvoice"].ToString(), 550, 115, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Acc No", 350, 100, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["AccountNo"].ToString(), 550, 100, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Branch", 350, 85, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtb.Rows[0]["BranchName"].ToString(), 550, 85, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 145, 0);
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Prepared By", 15, 55, 0);
                        if (Data.AgentId == "2")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "OCEANUS CONTAINER LINES PTE LTD", 480, 55, 0);
                        }
                        else if (Data.AgentId == "100" || Data.AgentId == "101")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for NERIDA SHIPPING SERVICSE LLC", 480, 55, 0);
                        }
                        else
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "for NERIDA SHIPPING PVT LTD", 480, 55, 0);
                        }
                        if (Data.AgentId == "2")
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 20, 0);
                        }
                        else
                        {
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMMUNICATION OFFICE ADDRESS :", 15, 20, 0);
                        }

                    }


                    //horizontal line2 big
                    cb.MoveTo(10, 776);
                    cb.LineTo(660, 776);

                    //horizontal line3 big
                    cb.MoveTo(10, 680);
                    cb.LineTo(660, 680);

                    cb.MoveTo(10, 625);
                    cb.LineTo(660, 625);

                    cb.MoveTo(10, 610);
                    cb.LineTo(660, 610);

                    cb.MoveTo(10, 595);
                    cb.LineTo(660, 595);

                    cb.MoveTo(10, 565);
                    cb.LineTo(660, 565);

                    cb.MoveTo(10, 530);
                    cb.LineTo(660, 530);

                    //Center line6 Small
                    cb.MoveTo(160, 564);
                    cb.LineTo(160, 310);
                    ///MMM
                    cb.MoveTo(210, 564);
                    cb.LineTo(210, 310);

                    cb.MoveTo(280, 564);
                    cb.LineTo(280, 310);

                    cb.MoveTo(340, 564);
                    cb.LineTo(340, 310);

                    cb.MoveTo(380, 564);
                    cb.LineTo(380, 270);

                    cb.MoveTo(440, 564);
                    cb.LineTo(440, 270);

                    cb.MoveTo(440, 550);
                    cb.LineTo(660, 550);

                    //center small
                    cb.MoveTo(460, 550);
                    cb.LineTo(460, 310);



                    cb.MoveTo(510, 550);
                    cb.LineTo(510, 290);
                    //Muthu



                    cb.MoveTo(535, 550);
                    cb.LineTo(535, 310);

                    cb.MoveTo(590, 550);
                    cb.LineTo(590, 290);

                    cb.MoveTo(615, 550);
                    cb.LineTo(615, 310);

                    cb.MoveTo(440, 290);
                    cb.LineTo(660, 290);

                    //horizontal line6 small
                    cb.MoveTo(10, 310);
                    cb.LineTo(660, 310);


                    cb.MoveTo(10, 270);
                    cb.LineTo(660, 270);

                    cb.MoveTo(10, 210);
                    cb.LineTo(660, 210);

                    cb.MoveTo(10, 150);
                    cb.LineTo(660, 150);


                    cb.MoveTo(10, 70);
                    cb.LineTo(660, 70);

                    cb.MoveTo(10, 35);
                    cb.LineTo(660, 35);
                }
                else
                {
                    cb.SetFontAndSize(bfheader2, 12);
                    cb.SetColorFill(Color.RED);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Continuity of charges & Total amount as Per Annexure of 2nd Page", 150, 20, 0);
                    cb.SetFontAndSize(bfheader2, 9);
                    cb.SetColorFill(Color.BLACK);

                    //horizontal line2 big
                    cb.MoveTo(10, 776);
                    cb.LineTo(660, 776);

                    //horizontal line3 big
                    cb.MoveTo(10, 680);
                    cb.LineTo(660, 680);

                    cb.MoveTo(10, 625);
                    cb.LineTo(660, 625);

                    cb.MoveTo(10, 610);
                    cb.LineTo(660, 610);

                    cb.MoveTo(10, 595);
                    cb.LineTo(660, 595);

                    cb.MoveTo(10, 565);
                    cb.LineTo(660, 565);

                    cb.MoveTo(10, 530);
                    cb.LineTo(660, 530);

                    //Center line6 Small
                    cb.MoveTo(160, 564);
                    cb.LineTo(160, 35);
                    ///MMM
                    ///
                    cb.MoveTo(280, 564);
                    cb.LineTo(280, 35);

                    cb.MoveTo(340, 564);
                    cb.LineTo(340, 35);

                    cb.MoveTo(380, 564);
                    cb.LineTo(380, 35);

                    cb.MoveTo(440, 564);
                    cb.LineTo(440, 35);

                    cb.MoveTo(440, 550);
                    cb.LineTo(660, 550);

                    //old
                    //cb.MoveTo(440, 550);
                    //cb.LineTo(660, 35);


                    //center small
                    cb.MoveTo(460, 550);
                    cb.LineTo(460, 35);

                    cb.MoveTo(510, 550);
                    cb.LineTo(510, 35);

                    cb.MoveTo(535, 550);
                    cb.LineTo(535, 35);

                    cb.MoveTo(590, 550);
                    cb.LineTo(590, 35);

                    cb.MoveTo(615, 550);
                    cb.LineTo(615, 35);

                    cb.MoveTo(10, 35);
                    cb.LineTo(660, 35);


                    ///

                    ////horizontal line6 small
                    //cb.MoveTo(10, 310);
                    //cb.LineTo(660, 310);


                    //cb.MoveTo(10, 270);
                    //cb.LineTo(660, 270);

                    //cb.MoveTo(10, 210);
                    //cb.LineTo(660, 210);

                    //cb.MoveTo(10, 150);
                    //cb.LineTo(660, 150);


                    //cb.MoveTo(10, 70);
                    //cb.LineTo(660, 70);

                    //cb.MoveTo(10, 35);
                    //cb.LineTo(660, 35);

                }


                cb.EndText();
                cb.Stroke();
                doc.NewPage();

                // Step 6: Add content to the second page
                doc.Add(new Paragraph("This is the second page"));
                writer.CloseStream = false;
                doc.Close();
                byte[] byteInfo = memoryStream.ToArray();
                memoryStream.Write(byteInfo, 0, byteInfo.Length);
                memoryStream.Position = 0;

            }



            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "Invoice.pdf"
            };

            return response;

        }



        public DataTable NewGetDisplayEstimateExisting(string ID, string CntrID)
        {
            string _Query = " select convert(varchar,LLimitDt, 103) as FDatev,convert(varchar,ULimitDt, 103) as TDatev,LLimit,ULimit,Rate,CurrID,Days,Amount,Total,NVO_ImpBLDDGCharges.ExRate,TotalAmount" +
                            " from NVO_ImpBLDDGChargedtls " +
                            " inner join NVO_ImpBLDDGCharges on NVO_ImpBLDDGCharges.ID=NVO_ImpBLDDGChargedtls.DDGID" +
                            " where DDGID = " + ID + " and NVO_ImpBLDDGChargedtls.CntrID=" + CntrID;
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

        public DataTable GetInvDetentionValues(string InvID)
        {
            string _Query = "select distinct (select top(1) isnull(DDGID,0) from NVO_BLCharges where Id = BLInvID) as DDGID from NVO_InvoiceCusBillingdtls where InvCusBillingID = " + InvID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetInvPDFBankValues(string InvID)
        {
            string _Query = "select * from NVO_FinBankMaster where Id = " + InvID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetNotesClauses()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=268";
            return Manag.GetViewData(_Query, "");
        }
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
        public DataTable GetInvPDFValus(string InvID)
        {
            string _Query = "select Id, InvoiceNo, convert(varchar, invDate, 103) as InvDate,AgentID,convert(varchar, InvDueDate, 103) as InvDueDate,PartyName,InvTotal,PartyTypes,PartyID, BLID,BranchID,TaxNo,Address,StateCode,InvTax,InvAmount,intBank,CurrencyID,(select top(1) GSTIN from NVO_CustomerMaster Inner Join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.ID where NVO_CusBranchLocation.CID = PartyID) as GSTIN," +
              " (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_InvoiceCusBilling.AgentID ) as Agency,(select  top(1) IRN from NVO_EInvoiceGeneration where InvID=NVO_InvoiceCusBilling.Id) as IRNNo,(select  top(1) SingnedQRCode from NVO_EInvoiceGeneration where InvID=NVO_InvoiceCusBilling.Id) as SingnedQRCode,intBank,(select top(1) CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.Id = NVO_InvoiceCusBilling.CurrencyID) as Currency,CurrencyID, " +
              " ( select top(1) UserName from NVO_UserDetails where ID = NVO_InvoiceCusBilling.UserID ) as CreatedBy,FinalInvoice,IsFinal,InvTypes,(select top(1) FinalInvoice from NVO_InvoiceCusBilling Inv where Inv.Id = NVO_InvoiceCusBilling.DrID) as DebitInvoice,Remarks,TaxExemption,isnull(CreditNoteType,0) CreditNoteType from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID=" + InvID;
            return Manag.GetViewData(_Query, "");
        }
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

        public DataTable GetInvPDFCntrValues(string BLID)
        {
            string _Query = " select NVO_Containers.CntrNo + ' / ' + size as CntrNos from NVO_BOLCntrDetails inner join NVO_Containers on NVO_Containers.ID = NVO_BOLCntrDetails.CntrID " +
                            " where NVO_BOLCntrDetails.BLID=" + BLID;
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

            string _Query = " Select distinct B.ID,NVO_BOL.BLNumber as BookingNo,CurrencyID, " +
                            "  case when NVO_InvoiceCusBilling.BLTypes = 2 then (select top(1)(select top(1)VesVoy from NVO_View_VoyageDetails  where NVO_View_VoyageDetails.ID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails where  NVO_BOLImpVoyageDetails.BLID=NVO_BOL.ID)  else VesVoy end as VesVoy, (select top 1 PortName from NVO_PortMaster where ID = NVO_BOL.POLID) as POL, " +
                            " (select top 1 PortName from NVO_PortMaster where ID = NVO_BOL.PODID) as POD,POO,  (select top 1 CityName from NVO_CityMaster where ID = NVO_BOL.FPODID) as FPOD, " +
                            " case when NVO_InvoiceCusBilling.BLTypes = 2 then (select top(1) convert(varchar, ETA, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BKgId=B.ID) else (select  top(1) Convert(varchar, ETA, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) end as ETA,  " +
                            " case when NVO_InvoiceCusBilling.BLTypes = 2 then(select top(1) convert(varchar, ETD, 103) from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BKgId = B.ID) else (select   top(1) Convert(varchar, ETD, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) end as ETD, " +
                            " '20' + ' x ' + convert(varchar, CTQ20) + ',' + '40' + ' x ' + convert(varchar, CTQ40) as CntrCount, " +
                            " (select top(1) GrsWt from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) as GrsWt1, " +
                            " (select top(1) GRWT from NVO_BLRelease where BLID = NVO_BOL.ID ) as GrsWt," +
                            " (select top(1) CntrNo from NVO_BOLCntrDetails where BLID = NVO_BOL.ID ) +'/' + (select top(1) size from NVO_BOLCntrDetails " +
                            " where BLID = NVO_BOL.ID ) +'/' + ServiceType as CntrSizeService,(SELECT top 1 CompanyName from NVO_NewCompnayDetails) AS CompanyName " +
                            " from NVO_Booking B INNER JOIN NVO_BOL ON NVO_BOL.BkgID = B.ID " +
                            " inner join NVO_InvoiceCusBilling on NVO_InvoiceCusBilling.BLID = NVO_BOL.ID " +
                            " and NVO_InvoiceCusBilling.BkgId = B.ID  where NVO_BOL.ID=" + BLID + " and NVO_InvoiceCusBilling.Id = " + InvID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetInvCusBillingdtls(string InvID)
        {
            string _Query = " Select  Id,InvCusBillingID,NarrationDescription,RatePerUnit, (select top(1) Size from NVO_tblCntrTypes where NVO_tblCntrTypes.ID = NVO_InvoiceCusBillingdtls.UnitID) as size, Qty,ROE,LocalAmount,(select top(1) SACCODE from NVO_ChargeTB where ID = NarrationID) as SACCode,(select top(1) CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.Id = CurrencyID) as Currency, " +
                            " isnull((select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID= NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID= 1 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as SGSTAmt, " +
                            " isnull((select Top(1) Tax_PCT from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 1 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as SGSTPec, " +
                            " isnull((select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 2 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as CGSTAmt, " +
                            " isnull((select Top(1) Tax_PCT from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 2 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as CGSTPec, " +
                            " isnull((select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 3 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as IGSTAmt, " +
                            " isnull((select Top(1) Tax_PCT from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 3 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID),0) as IGSTPec " +
                            " from NVO_InvoiceCusBillingdtls where InvCusBillingID =" + InvID;
            return Manag.GetViewData(_Query, "");
        }

        private static Byte[] BitmapToBytesCode(System.Drawing.Bitmap image)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return stream.ToArray();
            }
        }



        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getReceiptspdf")]
        public HttpResponseMessage getReceiptspdf([FromBody] MyInvoice Data)
        {
            MemoryStream memoryStream = new MemoryStream();
            MergeEx pdfmp = new MergeEx();
            DataTable dtv = GetInvPDFValus(Data.ID.ToString());
            if (dtv.Rows.Count > 0)
            {
                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();

                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
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

                DataTable dtc = GetAgencyDetails(Data.AgentId);
                if (dtc.Rows.Count > 0)
                {
                    if (Data.AgentId == "13")
                    {
                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/neridaaddress1.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);


                    }
                    if (Data.AgentId == "14")
                    {
                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddressmundra.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);

                    }
                    if (Data.AgentId == "15")
                    {
                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddressdelhi.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);

                    }
                    if (Data.AgentId == "16")
                    {

                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/neridaaddress1.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);
                    }

                    if (Data.AgentId == "1")
                    {

                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddresschennai.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);
                    }

                    if (Data.AgentId == "100" || Data.AgentId == "101")
                    {

                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/invaddressjeblali.png"));
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tbllogo.AddCell(cell);
                    }
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

                DataTable _dtv = GetReceiptDtls(Data.ID.ToString());
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

                    cell = new Cell(new Phrase(_dtv.Rows[0]["Remarks"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
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
                    cell.Colspan = 2;
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

                    cell = new Cell(new Phrase(_dtv.Rows[0]["Amount"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
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


                    DataTable _dtn = GetReceiptInvoiceDtls(Data.ID.ToString());

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

                        // cell = new Cell(new Phrase(DueAmt.ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        // cell.BorderWidth = 0.5f;
                        //// cell.Colspan = 1;
                        // cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        // TblInvoiceDtls.AddCell(cell);


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


                    // Step 6: Add content to the second page
                    doc.Add(new Paragraph("This is the second page"));
                    pdfWriter.CloseStream = false;
                    doc.Close();
                    byte[] byteInfo = memoryStream.ToArray();
                    memoryStream.Write(byteInfo, 0, byteInfo.Length);
                    memoryStream.Position = 0;

                }





            }

            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "Receipts.pdf"
            };
            return response;

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

        public DataTable GetReceiptInvoiceDtls(string idv)
        {
            string _Query = "select NVO_ReceiptBL.ReceiptID, (select Top(1) FinalInvoice from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId) as InvoiceNo, " +
                            " (select Top(1) Convert(varchar, InvDate, 103) from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId) as InvoiceDate, " +
                            " (select Top(1) InvTotal from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId and NVO_InvoiceCusBilling.InvTypes = 1) - isnull((select sum(InvTotal)  from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.InvTypes = 2 and NVO_InvoiceCusBilling.DRID = NVO_ReceiptBL.InvCusBillingId ),0) as InvoiceAmt, " +
                            " Amount as ReceivedAmt, " +
                            " (select Top(1)(select CurrencyCode from NVO_CurrencyMaster where NVO_CurrencyMaster.ID = NVO_InvoiceCusBilling.CurrencyID) " +
                            " from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId)as Currency, " +
                            " (select Top(1) TDSAmt from NVO_InvoiceCusBilling where NVO_InvoiceCusBilling.ID = NVO_ReceiptBL.InvCusBillingId) as TDSAmt,  isnull((select TOP 1 (GLCode + '-' + GLDesc) from NVO_GLMaster WHERE ID = TDS),'') as TDStype ,* " +
                            "  from NVO_ReceiptBL where ReceiptID = " + idv;
            return Manag.GetViewData(_Query, "");


        }

        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getBLPrintpdf")]
        public HttpResponseMessage getBLPrintpdf([FromBody] MyInvoice Data)
        {

            MemoryStream memoryStream = new MemoryStream();
            MergeEx pdfmp = new MergeEx();
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();

            DataTable _dt = GetBkgCustomer(Data.ID.ToString());
            if (_dt.Rows.Count > 0)
            {

                PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                PdfContentByte cb = writer.DirectContent;
                //cb.SetColorStroke(new Color(0, 0, 208));
                //cb.MoveTo(280, 860);
                //cb.LineTo(650, 860);
                int _Xp = 10, _Yp = 785, YDiff = 10;

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);
                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/pdfhead.png"));

                png1.SetAbsolutePosition(15, 837);
                png1.ScalePercent(17f);
                doc.Add(png1);

                //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/header.JPG"));
                //png2.SetAbsolutePosition(150, 842);
                //png2.ScalePercent(70f);
                //doc.Add(png2);

                BaseFont Crossbfheader = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetColorFill(Color.LIGHT_GRAY);
                cb.SetFontAndSize(Crossbfheader, 90);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   D R A F T   ", 100, 200, 45);



                //BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //cb.SetFontAndSize(bfheader21, 23);
                //cb.SetColorFill(new Color(0, 0, 128));
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Oceanus Container Lines Pvt.Ltd", 280, 870, 0);


                //BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //cb.SetFontAndSize(bfheader22, 8);
                //cb.SetColorFill(new Color(0, 0, 128));
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 280, 850, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 280, 835, 0);
                //cb.SetColorStroke(new Color(0, 0, 128));

                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(new Color(255, 200, 200));
                //center

                cb.BeginText();



                ////Top//
                cb.MoveTo(15, 825);
                cb.LineTo(650, 825);

                //left//
                cb.MoveTo(325, 825);
                cb.LineTo(325, 510);
                //left off
                cb.MoveTo(475, 825);
                cb.LineTo(475, 782);

                ////Top One off//
                cb.MoveTo(325, 782);
                cb.LineTo(650, 782);


                ////Top One//
                cb.MoveTo(15, 745);
                cb.LineTo(650, 745);

                ////Top Tow//
                cb.MoveTo(15, 665);
                cb.LineTo(650, 665);


                ////Top Three//
                cb.MoveTo(15, 580);
                cb.LineTo(650, 580);


                ////Top Four//
                cb.MoveTo(15, 545);
                cb.LineTo(650, 545);

                //left off
                cb.MoveTo(170, 580);
                cb.LineTo(170, 510);

                //left off Two
                cb.MoveTo(490, 580);
                cb.LineTo(490, 510);

                ////Top five//
                cb.MoveTo(15, 510);
                cb.LineTo(650, 510);

                ////Top five//
                cb.MoveTo(15, 485);
                cb.LineTo(650, 485);


                ////Top six//
                cb.MoveTo(15, 225);
                cb.LineTo(650, 225);

                ////Top Seven//
                cb.MoveTo(15, 30);
                cb.LineTo(650, 30);


                //left off Marks
                cb.MoveTo(190, 510);
                cb.LineTo(190, 260);


                cb.MoveTo(15, 260);
                cb.LineTo(260, 260);


                //left off Pakage
                cb.MoveTo(260, 510);
                cb.LineTo(260, 225);

                //left off Description
                cb.MoveTo(490, 510);
                cb.LineTo(490, 225);

                ////left off Description
                //cb.MoveTo(490, 485);
                //cb.LineTo(490, 225);


                //left off Mesu
                cb.MoveTo(560, 510);
                cb.LineTo(560, 225);



                cb.MoveTo(350, 225);
                cb.LineTo(350, 30);



                cb.SetFontAndSize(bfheader2, 11);
                cb.SetColorFill(Color.BLACK);

                //cb.EndText();
                //cb.BeginText();

                BaseFont bfheader23 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader23, 8);
                cb.SetColorFill(new Color(0, 0, 128));
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper", 15, 815, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Consignee (if 'To Order So Indicate')", 15, 733, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify Party (No claim shall attach for failure to notify)", 15, 650, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Delivery Agent", 345, 733, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify Party(2)", 345, 650, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Booking No.", 345, 810, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill of Lading No", 535, 810, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper's Ref:", 345, 760, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Receipt", 25, 570, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port of Loading", 180, 570, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Delivery", 340, 570, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Paid at", 500, 570, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel & Voyage No.", 25, 535, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port of Discharge", 180, 535, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Final Destination", 340, 535, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No.of Original Bill of Lading", 500, 535, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks & Numbers", 25, 495, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No of Pkgs. or", 195, 500, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipping Units", 195, 490, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description of Goods & Pkgs", 300, 495, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Cargo Weight", 500, 495, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Measurement", 570, 495, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " SHIPPERS STOW, COUNT, LOAD &SEALED", 270, 470, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " Gross Weight", 497, 460, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " Net Weight", 497, 410, 0);

                BaseFont bfheader24 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader24, 8);
                cb.SetColorFill(new Color(0, 0, 128));

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString().Trim().ToUpper(), 345, 795, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString().Trim().ToUpper(), 535, 795, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DRAFT", 420, 760, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POO"].ToString().Trim().ToUpper(), 25, 560, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POL"].ToString().Trim().ToUpper(), 180, 560, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper(), 340, 560, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FreightPaidAt"].ToString().Trim().ToUpper(), 500, 560, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["VesVoy"].ToString().Trim().ToUpper(), 25, 525, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POD"].ToString().Trim().ToUpper(), 180, 525, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper(), 340, 525, 0);






                int ColumnRows = 800; int RowsColumn = 0;
                RowsColumn = 0;
                string[] ArrayAddress = Regex.Split(_dt.Rows[0]["Shipper"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["ShipperAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
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

                ColumnRows = 720;
                RowsColumn = 0;
                string[] ArrayAddress1 = Regex.Split(_dt.Rows[0]["Consignee"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["Consignee"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                string[] Aaddsplit1;

                for (int x = 0; x < ArrayAddress1.Length; x++)
                {
                    Aaddsplit1 = ArrayAddress1[x].Split('\n');

                    for (int k = 0; k < Aaddsplit1.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit1[k].ToString(), 15, ColumnRows, 0);
                        ColumnRows -= 9;
                        RowsColumn++;
                    }
                }

                ColumnRows = 640;
                RowsColumn = 0;
                string[] ArrayAddress2 = Regex.Split(_dt.Rows[0]["Notify1"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["Notify1Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                string[] Aaddsplit2;

                for (int x = 0; x < ArrayAddress2.Length; x++)
                {
                    Aaddsplit2 = ArrayAddress2[x].Split('\n');

                    for (int k = 0; k < Aaddsplit2.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit2[k].ToString(), 15, ColumnRows, 0);
                        ColumnRows -= 9;
                        RowsColumn++;
                    }
                }


                ColumnRows = 720;
                RowsColumn = 0;

                string[] ArrayAddress3 = Regex.Split(_dt.Rows[0]["Agent"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["AgentAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                string[] Aaddsplit3;

                for (int x = 0; x < ArrayAddress3.Length; x++)
                {
                    Aaddsplit3 = ArrayAddress3[x].Split('\n');

                    for (int k = 0; k < Aaddsplit3.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit3[k].ToString(), 345, ColumnRows, 0);
                        ColumnRows -= 9;
                        RowsColumn++;
                    }
                }


                ColumnRows = 640;
                RowsColumn = 0;

                string[] ArrayAddress4 = Regex.Split(_dt.Rows[0]["Notify2"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["Notify2Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                string[] Aaddsplit4;

                for (int x = 0; x < ArrayAddress4.Length; x++)
                {
                    Aaddsplit4 = ArrayAddress4[x].Split('\n');

                    for (int k = 0; k < Aaddsplit4.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit4[k].ToString(), 345, ColumnRows, 0);
                        ColumnRows -= 9;
                        RowsColumn++;
                    }
                }



                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["Packages"].ToString().Trim().ToUpper(), 195, 450, 0);
                var Cargosplit = SplitByLenght(_dt.Rows[0]["CargoPakage"].ToString(), 9);
                int CRRow = 430;
                for (int z = 0; z < Cargosplit.Length; z++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Cargosplit[z].ToUpper(), 195, CRRow, 0);
                    CRRow -= 15;
                }
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["CargoPakage"].ToString().Trim().ToUpper(), 195, 430, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["GRWT"].ToString().Trim().ToUpper() + " KGS", 497, 440, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["NTWT"].ToString().Trim().ToUpper() + " KGS", 497, 380, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["CBM"].ToString().Trim().ToUpper() + " M3", 570, 450, 0);

                var MarkNo = _dt.Rows[0]["Marks"].ToString().Split('\n');
                var Description = _dt.Rows[0]["Description"].ToString().Split('\n');

                int RowMx = 470;

                for (int x = 0; x < MarkNo.Length; x++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, MarkNo[x].ToString(), 25, RowMx, 0);
                    RowMx -= 9;

                }

                int RowDec = 460;

                for (int x = 0; x < Description.Length; x++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Description[x].ToString(), 270, RowDec, 0);
                    RowDec -= 9;

                }
                BaseFont bfheader25 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader25, 8);
                cb.SetColorFill(new Color(0, 0, 128));
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Excess Value Declaration Refer to Clause 6 (3) (B) + (C)", 15, 245, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "on reverse side", 15, 235, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREIGHT: PREPAID", 270, 265, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["intFreedays"].ToString() + '-' + _dt.Rows[0]["ddlFreeday"].ToString(), 270, 250, 0);
                if (Description.Length > 16)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Continuity as Per Annexure Attached", 270, 235, 0);
                }

                BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader5, 6);
                cb.SetColorFill(new Color(0, 0, 128));
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "The term carriage by sea by defenition being the transport of goods, merchandise or their", 15, 210, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "packing inclusive of containers and/or goods of any type between one port and another port,", 15, 201, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "the carrier is not and shall not be responsible for:", 15, 192, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "a)     Any damage occasioned to the goods arising out of or in relation to the loading", 15, 172, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       and unloading of containers and/or goods on or off the vessel; and/or", 15, 163, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "b)     Any damage to containers and/or goods before the loading and after the", 15, 154, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       unloading of the said containers and/or goods from the vessel.", 15, 145, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "c)     Any damage caused to containers and/or goods of board the vessel by the other", 15, 136, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       container in the course of loading or unloading of those other containers and/or", 15, 127, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       goods on board the vessel by stevedores. And/or", 15, 118, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "d)     Any damage caused to containers and/or goods prior to the loading and", 15, 109, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       subsequent to the unloading of other containers and/or goods arising out of the", 15, 100, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       vessel’s ancillary equipment (or any part thereof) coming into contact with the", 15, 91, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       said Containers and/or goods lying on the quayside should the said containers", 15, 82, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       and/or goods to be stacked one on top of the other or improperly arranged on the", 15, 73, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       quayside.", 15, 64, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "e)     Any mis-information on the import General Manifest and re-export of import", 15, 55, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       containers and/or goods and where appropriate, the merchant shall furnish", 15, 46, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       guarantees to the Carrier’s agent if there is any breach.", 15, 37, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Received by the carrier the Goods as specified above in apparent good order and conditions", 355, 210, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "unless otherwise stated, to be trnspoted to such place agreed, authoried or permitted", 355, 201, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "herein and subject to all the terms and conditions appearing on the front and reverse of this", 355, 192, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill of Lading to which the Merchant agrees by accepting this Bill of Lading, any local", 355, 183, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "privilages and customs notwithstanding. The particulars given above are as stated by the", 355, 174, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "shipper and the weight, measure, quantity, condition, contents and value of the Goods are", 355, 165, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "unknown to the carrier. One of the original Bills of Lading shall be presented to the carrier", 355, 156, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "or his agent at destination before the", 355, 147, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipped on Board Date", 355, 120, 0);


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BlDatev"].ToString(), 475, 120, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place and Date of issue", 355, 100, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Signed on behalf of the Carrier - Oceanus Container Lines Pte. Ltd.   :", 355, 80, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "By", 355, 60, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 355, 40, 0);
                cb.MoveTo(450, 110);
                cb.LineTo(650, 110);

                cb.MoveTo(450, 90);
                cb.LineTo(650, 90);


                cb.EndText();
                cb.Stroke();


                writer.CloseStream = false;
                doc.Close();
                byte[] byteInfo = memoryStream.ToArray();
                memoryStream.Write(byteInfo, 0, byteInfo.Length);
                memoryStream.Position = 0;
            }

            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "BLPrint.pdf"
            };
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception("Failed to get a valid response.");
            }
            return response;
        }

        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getBLPrintpdfLive")]
        public HttpResponseMessage getBLPrintpdfLive([FromBody] myBLPrint Data)
        {

            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();

            DataTable _dt = GetBkgCustomer(Data.id.ToString());
            string pdfpath = HttpContext.Current.Server.MapPath("/pdfpath/");
            MergeEx pdfmp = new MergeEx();
            pdfmp.SourceFolder = pdfpath;

            Random random = new Random();
            int randomNumber = random.Next();

            pdfmp.DestinationFile = pdfpath + "Multiple-" + "BLNumber"+ randomNumber + "-" + Data.id + "BL.pdf";
            string FileHidpath = pdfpath + "Multiple-" + "BLNumber" + randomNumber +  "-" + Data.id + "BL.pdf";

            string _FileName = "BLNUmberrr" + Data.id + 1;
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));
            //PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();

            PdfContentByte cb = writer.DirectContent;
            //cb.SetColorStroke(new Color(0, 0, 208));
            //cb.MoveTo(280, 860);
            //cb.LineTo(650, 860);
            int _Xp = 10, _Yp = 785, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/assets/img/pdfhead.png"));
            png1.SetAbsolutePosition(15, 837);
            png1.ScalePercent(55f);
            doc.Add(png1);

            //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/oclheader.jpg"));
            //png2.SetAbsolutePosition(320, 835);
            //png2.ScalePercent(52f);
            //doc.Add(png2);

            BaseFont Crossbfheader = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(Color.LIGHT_GRAY);
            cb.SetFontAndSize(Crossbfheader, 70);
            cb.BeginText();

            if (Data.printvalue == "1")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   D R A F T   ", 200, 200, 45);
            }

            if (Data.printvalue == "2")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   FIRST ORIGINAL   ", 100, 200, 45);
            }
            if (Data.printvalue == "3")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SECOND ORIGINAL   ", 100, 200, 45);
            }
            if (Data.printvalue == "4")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   THIRD ORIGINAL   ", 100, 200, 45);
            }

            if (Data.printvalue == "6")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   EXPRESS RELEASE   ", 100, 200, 45);
            }


            BaseFont Crossbfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(Color.LIGHT_GRAY);
            cb.SetFontAndSize(Crossbfheader6, 40);

            if (Data.printvalue == "5")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SEAWAY BL - NON NEGOTIABLE  ", 100, 200, 45);
            }
            if (Data.printvalue == "11")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   BACK PAGE   ", 100, 200, 45);
            }
            if (Data.printvalue == "7")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SURRENDER BL  ", 200, 350, 45);
            }


            if (Data.printvalue == "5")
            {
                //cb..(PdfContentByte.ALIGN_LEFT, "   SEAWAY BL", 100, 200, 45);
            }

            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 100, 200, 45);
            }
            cb.Stroke();
            cb.EndText();

            cb.BeginText();
            BaseFont Crossbfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(Color.LIGHT_GRAY);
            cb.SetFontAndSize(Crossbfheader1, 40);

            if (Data.printvalue == "8")
            {
                if (Data.LocID == "25")
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RECEIVED FOR SHIPMENT", 100, 200, 45);
                    iTextSharp.text.Image png15 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/assets/img/Sign.png"));
                    png15.SetAbsolutePosition(400, 40);
                    png15.ScalePercent(20f);
                    doc.Add(png15);


                }
                else
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RECEIVED FOR SHIPMENT", 100, 200, 45);
                }

            }





            cb.EndText();



            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader2, 9);
            // cb.SetColorFill(new Color(255, 200, 200));
            //center

            cb.BeginText();



            ////Top//
            cb.MoveTo(15, 825);
            cb.LineTo(650, 825);

            //left//
            cb.MoveTo(325, 825);
            cb.LineTo(325, 510);
            //left off
            cb.MoveTo(475, 825);
            cb.LineTo(475, 782);

            ////Top One off//
            cb.MoveTo(325, 782);
            cb.LineTo(650, 782);


            ////Top One//
            cb.MoveTo(15, 745);
            cb.LineTo(650, 745);

            ////Top Tow//
            cb.MoveTo(15, 665);
            cb.LineTo(650, 665);


            ////Top Three//
            cb.MoveTo(15, 580);
            cb.LineTo(650, 580);


            ////Top Four//
            cb.MoveTo(15, 545);
            cb.LineTo(650, 545);

            //left off
            cb.MoveTo(170, 580);
            cb.LineTo(170, 510);

            //left off Two
            cb.MoveTo(490, 580);
            cb.LineTo(490, 510);

            ////Top five//
            cb.MoveTo(15, 510);
            cb.LineTo(650, 510);

            ////Top five//
            cb.MoveTo(15, 485);
            cb.LineTo(650, 485);


            ////Top six//
            cb.MoveTo(15, 225);
            cb.LineTo(650, 225);

            ////Top Seven//
            cb.MoveTo(15, 30);
            cb.LineTo(650, 30);


            //left off Marks
            cb.MoveTo(190, 510);
            cb.LineTo(190, 260);


            cb.MoveTo(15, 260);
            cb.LineTo(260, 260);


            //left off Pakage
            cb.MoveTo(260, 510);
            cb.LineTo(260, 225);

            //left off Description
            cb.MoveTo(490, 510);
            cb.LineTo(490, 225);

            ////left off Description
            //cb.MoveTo(490, 485);
            //cb.LineTo(490, 225);


            //left off Mesu
            cb.MoveTo(560, 510);
            cb.LineTo(560, 225);



            cb.MoveTo(350, 225);
            cb.LineTo(350, 30);



            cb.SetFontAndSize(bfheader2, 11);
            cb.SetColorFill(Color.BLACK);

            //cb.EndText();
            //cb.BeginText();

            BaseFont bfheader23 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader23, 8);
            cb.SetColorFill(new Color(0, 0, 0));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper", 15, 815, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Consignee (if 'To Order So Indicate')", 15, 733, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify Party (No claim shall attach for failure to notify)", 15, 650, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Delivery Agent", 345, 733, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify Party(2)", 345, 650, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Booking No.", 345, 810, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill of Lading No", 535, 810, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper's Ref:", 345, 760, 0);

            if (Data.printvalue == "12")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " NON NEGOTIABLE   ", 420, 760, 0);
            }


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Receipt", 25, 570, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port of Loading", 180, 570, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Delivery", 340, 570, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Paid at", 500, 570, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Vessel & Voyage No.", 25, 535, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port of Discharge", 180, 535, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Final Destination", 340, 535, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No.of Original Bill of Lading", 500, 535, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks & Numbers", 25, 495, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No of Pkgs. or", 195, 500, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipping Units", 195, 490, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description of Goods & Pkgs", 300, 495, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Cargo Weight", 500, 495, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Measurement", 570, 495, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " SHIPPERS STOW, COUNT, LOAD & SEALED", 270, 470, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " Gross Weight", 497, 460, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, " Net Weight", 497, 410, 0);

            BaseFont bfheader24 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader24, 8);
            cb.SetColorFill(new Color(0, 0, 0));

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString().Trim().ToUpper(), 345, 795, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString().Trim().ToUpper(), 535, 795, 0);
            string Valuesv = "";
            if (Data.printvalue == "1")
                Valuesv = "DRAFT";
            if (Data.printvalue == "2")
                Valuesv = "FIRST ORIGINAL";
            if (Data.printvalue == "3")
                Valuesv = "SECOND ORIGINAL";
            if (Data.printvalue == "4")
                Valuesv = "THIRD ORIGINAL";
            if (Data.printvalue == "5")
                Valuesv = "SEAWAY BL-NON NEGOTIABLE";
            if (Data.printvalue == "6")
                Valuesv = "EXPRESS RELEASE";
            if (Data.printvalue == "7")
                Valuesv = "SURRENDER BL";
            if (Data.printvalue == "8")
                Valuesv = "RFS FIRST ORIGINAL";
            if (Data.printvalue == "9")
                Valuesv = "RFS 2ND ORIGINAL";
            if (Data.printvalue == "10")
                Valuesv = "RFS 3RD ORIGINAL";
            if (Data.printvalue == "11")
                Valuesv = "BACK PAGE";
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Valuesv, 420, 760, 0);

            if (Data.printvalue == "2" || Data.printvalue == "3" || Data.printvalue == "4" || Data.printvalue == "5" || Data.printvalue == "7"
               || Data.printvalue == "8" || Data.printvalue == "9" || Data.printvalue == "10")
            {
                var AutoGen = Manag.GetBLPrint_Number("2025");
                string _Query = "select 'SEQ-' + RIGHT('0' + RTRIM(year(getdate())), 2) + RIGHT('0' + RTRIM(MONTH(getdate())), 2) + right('000' + convert(varchar(10), " + AutoGen + "), 4) as PrintSeq";
                DataTable _dtvx = Manag.GetViewData(_Query, "");
                if (_dtvx.Rows.Count > 0)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtvx.Rows[0]["PrintSeq"].ToString(), 550, 760, 0);

                }
            }
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Valuesv, 420, 760, 0);
            var POOIDv = _dt.Rows[0]["POO"].ToString().Trim().ToUpper().Split(',');
            if (_dt.Rows[0]["POO"].ToString().Length > 25)
            {
                int xRow = 560;
                for (int i = 0; i < POOIDv.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, POOIDv[i], 25, xRow, 0);
                    xRow -= 11;
                }
            }
            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POO"].ToString().Trim().ToUpper(), 25, 560, 0);
            }

            var POLIDv = _dt.Rows[0]["POL"].ToString().Trim().ToUpper().Split(',');
            if (_dt.Rows[0]["POO"].ToString().Length > 25)
            {
                int xRow = 560;
                for (int i = 0; i < POLIDv.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, POLIDv[i], 180, xRow, 0);
                    xRow -= 11;
                }
            }
            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POL"].ToString().Trim().ToUpper(), 180, 560, 0);
            }

            var FPODv = _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper().Split(',');
            if (_dt.Rows[0]["FPOD"].ToString().Length > 25)
            {
                int xRow = 560;
                for (int i = 0; i < FPODv.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, FPODv[i], 340, xRow, 0);
                    xRow -= 11;
                }
            }
            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper(), 340, 560, 0);
            }




            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["IssuedAt"].ToString().Trim().ToUpper(), 500, 560, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FreightPaidAt"].ToString().Trim().ToUpper(), 500, 560, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["VesVoy"].ToString().Trim().ToUpper(), 25, 525, 0);


            var PODv = _dt.Rows[0]["POD"].ToString().Trim().ToUpper().Split(',');
            if (_dt.Rows[0]["POD"].ToString().Length > 25)
            {
                int xRow = 525;
                for (int i = 0; i < PODv.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, PODv[i], 180, xRow, 0);
                    xRow -= 11;
                }
            }
            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POD"].ToString().Trim().ToUpper(), 180, 525, 0);
            }


            var FPODvv = _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper().Split(',');
            if (_dt.Rows[0]["FPOD"].ToString().Length > 25)
            {
                int xRow = 525;
                for (int i = 0; i < FPODvv.Length; i++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, FPODvv[i], 340, xRow, 0);
                    xRow -= 11;
                }
            }
            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper(), 340, 525, 0);
            }





            if (Data.printvalue == "5")
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "0", 520, 525, 0);
            else
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["NoofOriginal"].ToString().Trim().ToUpper(), 520, 525, 0);



            int ColumnRows = 800; int RowsColumn = 0;
            RowsColumn = 0;
            string[] ArrayAddress = Regex.Split(_dt.Rows[0]["Shipper"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["ShipperAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
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

            ColumnRows = 720;
            RowsColumn = 0;
            string[] ArrayAddress1 = Regex.Split(_dt.Rows[0]["Consignee"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["ConsigneeAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
            string[] Aaddsplit1;

            for (int x = 0; x < ArrayAddress1.Length; x++)
            {
                Aaddsplit1 = ArrayAddress1[x].Split('\n');

                for (int k = 0; k < Aaddsplit1.Length; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit1[k].ToString(), 15, ColumnRows, 0);
                    ColumnRows -= 9;
                    RowsColumn++;
                }
            }

            ColumnRows = 640;
            RowsColumn = 0;
            string[] ArrayAddress2 = Regex.Split(_dt.Rows[0]["Notify1"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["Notify1Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
            string[] Aaddsplit2;

            for (int x = 0; x < ArrayAddress2.Length; x++)
            {
                Aaddsplit2 = ArrayAddress2[x].Split('\n');

                for (int k = 0; k < Aaddsplit2.Length; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit2[k].ToString(), 15, ColumnRows, 0);
                    ColumnRows -= 9;
                    RowsColumn++;
                }
            }


            ColumnRows = 720;
            RowsColumn = 0;

            string[] ArrayAddress3 = Regex.Split(_dt.Rows[0]["Agent"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["AgentAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
            string[] Aaddsplit3;

            for (int x = 0; x < ArrayAddress3.Length; x++)
            {
                Aaddsplit3 = ArrayAddress3[x].Split('\n');

                for (int k = 0; k < Aaddsplit3.Length; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit3[k].ToString(), 345, ColumnRows, 0);
                    ColumnRows -= 9;
                    RowsColumn++;
                }
            }


            ColumnRows = 640;
            RowsColumn = 0;

            string[] ArrayAddress4 = Regex.Split(_dt.Rows[0]["Notify2"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["Notify2Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
            string[] Aaddsplit4;

            for (int x = 0; x < ArrayAddress4.Length; x++)
            {
                Aaddsplit4 = ArrayAddress4[x].Split('\n');

                for (int k = 0; k < Aaddsplit4.Length; k++)
                {

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit4[k].ToString(), 345, ColumnRows, 0);
                    ColumnRows -= 9;
                    RowsColumn++;
                }
            }



            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["Packages"].ToString().Trim().ToUpper(), 195, 450, 0);
            var Cargosplitcar = _dt.Rows[0]["CargoPakage"].ToString().Split(' ');
            int CRRow = 430;
            for (int k = 0; k < Cargosplitcar.Length; k++)
            {
                var Cargosplit = SplitByLenght(Cargosplitcar[k], 9);

                for (int z = 0; z < Cargosplit.Length; z++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Cargosplit[z].ToUpper(), 195, CRRow, 0);
                    CRRow -= 15;
                }
            }
            // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["CargoPakage"].ToString().Trim().ToUpper(), 195, 430, 0);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["GRWT"].ToString().Trim().ToUpper() + " " + _dt.Rows[0]["GrsWtType"].ToString().Trim().ToUpper(), 493, 440, 0);
            if (_dt.Rows[0]["NTWT"].ToString().Trim().ToUpper() != "0.000")
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["NTWT"].ToString().Trim().ToUpper() + " " + _dt.Rows[0]["NtWtType"].ToString().Trim().ToUpper(), 493, 380, 0);

            if (_dt.Rows[0]["CBM"].ToString().Trim().ToUpper() != "0.0000")
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["CBM"].ToString().Trim().ToUpper() + " M3", 570, 450, 0);

            string[] arrayMarks = new string[] { };
            string[] arrayDescription = new string[] { };
            string[] arrayCntrNo = new string[] { };


            List<string> ArrayMarksV = new List<string>();
            arrayMarks = _dt.Rows[0]["Marks"].ToString().Split('\n');
            int intMarkCount = arrayMarks.Length + 7;
            arrayDescription = _dt.Rows[0]["Description"].ToString().Split('\n');
            int intDescCount = arrayDescription.Length;





            int RowMx = 470;

            int TotalLine = 0;
            int ColumnCountMrks = 8;
            ColumnCountMrks = arrayMarks.Length;
            TotalLine = 8;
            for (int LineX = 0; LineX < TotalLine; LineX++)
            {
                if (arrayMarks.Length >= LineX + 1)

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayMarks[LineX].ToUpper(), 25, RowMx, 0);
                RowMx -= 10;
            }

            RowMx -= 20;

            TotalLine = 4;
            DataTable _dtCntr = GetContainerDetails(Data.id.ToString());
            var _dtCntrValues = _dtCntr.Rows[0]["CntrDtls"].ToString().Split('\n');
            int TotalColumnCntr = (_dtCntrValues.Length < TotalLine) ? _dtCntrValues.Length : TotalLine;
            if (_dtCntr.Rows.Count > 0)
            {
                TotalLine = TotalColumnCntr;
                for (int LineX = 0; LineX < TotalLine; LineX++)
                {
                    var arrayCntrNov = SplitByLenght(_dtCntrValues[LineX].ToString(), 30);
                    for (int d = 0; d < arrayCntrNov.Length; d++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayCntrNov[d].ToUpper(), 25, RowMx, 0);
                        RowMx -= 10;
                    }
                }
            }




            int RowDec = 460;

            TotalLine = 15;

            for (int LineX = 0; LineX < TotalLine; LineX++)
            {
                if (arrayDescription.Length >= LineX + 1)

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayDescription[LineX], 270, RowDec, 0);
                RowDec -= 10;
            }

            BaseFont bfheader25 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader25, 8);
            cb.SetColorFill(new Color(0, 0, 0));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Excess Value Declaration Refer to Clause 6 (3) (B) + (C)", 15, 245, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "on reverse side", 15, 235, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREIGHT: " + _dt.Rows[0]["FreightPayment"].ToString(), 270, 265, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["intFreedays"].ToString() + '-' + _dt.Rows[0]["ddlFreeday"].ToString(), 270, 250, 0);
            if (_dt.Rows[0]["ddlFreeday"].ToString() != "")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["intFreedays"].ToString() + '-' + _dt.Rows[0]["ddlFreeday"].ToString(), 270, 250, 0);
            }
            else
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 270, 250, 0);
            }

            int[] Anxarr = { arrayDescription.Length, intMarkCount };
            int Anxmax = Anxarr[0];
            for (int i = 1; i < Anxarr.Length; i++)
            {
                if (Anxarr[i] > Anxmax)
                    Anxmax = Anxarr[i];
            }


            if (Anxmax > 16)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Continuity as Per Annexure Attached", 270, 235, 0);
            }


            BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader5, 6);
            cb.SetColorFill(new Color(0, 0, 0));
            //cb.SetColorFill(new Color(0, 0, 128));
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "The term carriage by sea by defenition being the transport of goods, merchandise or their", 15, 210, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "packing inclusive of containers and/or goods of any type between one port and another port,", 15, 201, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "the carrier is not and shall not be responsible for:", 15, 192, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "a)     Any damage occasioned to the goods arising out of or in relation to the loading", 15, 172, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       and unloading of containers and/or goods on or off the vessel; and/or", 15, 163, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "b)     Any damage to containers and/or goods before the loading and after the", 15, 154, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       unloading of the said containers and/or goods from the vessel.", 15, 145, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "c)     Any damage caused to containers and/or goods of board the vessel by the other", 15, 136, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       container in the course of loading or unloading of those other containers and/or", 15, 127, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       goods on board the vessel by stevedores. And/or", 15, 118, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "d)     Any damage caused to containers and/or goods prior to the loading and", 15, 109, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       subsequent to the unloading of other containers and/or goods arising out of the", 15, 100, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       vessel’s ancillary equipment (or any part thereof) coming into contact with the", 15, 91, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       said Containers and/or goods lying on the quayside should the said containers", 15, 82, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       and/or goods to be stacked one on top of the other or improperly arranged on the", 15, 73, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       quayside.", 15, 64, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "e)     Any mis-information on the import General Manifest and re-export of import", 15, 55, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       containers and/or goods and where appropriate, the merchant shall furnish", 15, 46, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "       guarantees to the Carrier’s agent if there is any breach.", 15, 37, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Received by the carrier the Goods as specified above in apparent good order and conditions", 355, 210, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "unless otherwise stated, to be transported to such as place agreed, authoried or permitted", 355, 201, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "herein and subject to all the terms and conditions appearing on the front and reverse of this", 355, 192, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill of Lading to which the Merchant agrees by accepting this Bill of Lading, any local", 355, 183, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "privilages and customs notwithstanding. The particulars given above are as stated by the", 355, 174, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "shipper and the weight, measure, quantity, condition, contents and value of the Goods are", 355, 165, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "unknown to the carrier.", 355, 156, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "In witness whereof one original Bill of lading has been signed if not otherwise stated", 355, 147, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "above, the same being accomplished the other(s), if any, to be void.  One original Bill of", 355, 138, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Lading must be presented to the carrier in exchange for the Goods or delivery order.", 355, 130, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipped on Board Date", 355, 115, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place and Date of issue", 355, 95, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Signed on behalf of the Carrier - Blue Wave Shipping & Logistic Pte Ltd.   :", 355, 80, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "By", 355, 60, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "As Agent", 355, 40, 0);
            cb.MoveTo(450, 110);
            cb.LineTo(650, 110);

            cb.MoveTo(450, 90);
            cb.LineTo(650, 90);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["SOBDatev"].ToString(), 460, 120, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["IssuedAt"].ToString() + "   &  " + _dt.Rows[0]["BlDatev"].ToString(), 460, 100, 0);



            cb.EndText();
            cb.Stroke();
            doc.Close();
            pdfmp.AddFile(_FileName + ".pdf");

            //Largest  array number
            int[] arr = { arrayDescription.Length, intMarkCount, _dtCntrValues.Length + 10 };
            int max = arr[0];
            for (int i = 1; i < arr.Length; i++)
            {
                if (arr[i] > max)
                    max = arr[i];
            }

            // int TotalColumn = (_dt.Rows.Count < arrayDescription.Length) ? intMarkCount : intMarkCount;
            int TotalColumn = max;
            int WriteLine = TotalColumn - TotalLine;
            int AttachedsheetNo = int.Parse(Math.Ceiling((WriteLine / 40.00)).ToString());
            int Cot = 0;
            //int LineCount = 15 + Cot;
            int LineCount = Cot;
            int SheetNo = 1;
            string Filesv = "Attach" + Data.id;
            string _AttFileName = Filesv;
            int LIndex = 15;
            int LMarkindex = 8;
            int LCntrindex = 4;

            for (int k = 0; k < AttachedsheetNo; k++)
            {

                Document Attdocument = new Document(rec);
                PdfWriter Attwriter = PdfWriter.GetInstance(Attdocument, new FileStream(pdfpath + (_AttFileName + SheetNo) + ".pdf", FileMode.Create));
                Attdocument.Open();
                PdfContentByte Attcb = Attwriter.DirectContent;


                BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);
                iTextSharp.text.Image png11 = iTextSharp.text.Image.GetInstance( HttpContext.Current.Server.MapPath("~/assets/img/pdfhead.png"));
                png11.SetAbsolutePosition(15, 837);
                png11.ScalePercent(60f);
                Attdocument.Add(png11);

                //iTextSharp.text.Image png21 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/oclheader.jpg"));
                //png21.SetAbsolutePosition(320, 835);
                //png21.ScalePercent(52f);
                //Attdocument.Add(png21);

                BaseFont bfheader211 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                Attcb.SetFontAndSize(bfheader211, 23);
                Attcb.SetColorFill(new Color(0, 0, 0));
                //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Oceanus Container Lines Pvt.Ltd", 280, 870, 0);


                BaseFont bfheader222 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                Attcb.SetFontAndSize(bfheader222, 8);
                Attcb.SetColorFill(new Color(0, 0, 0));
                //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 280, 850, 0);
                //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 280, 835, 0);
                Attcb.SetColorStroke(new Color(0, 0, 0));


                #region Border
                Attcb.MoveTo(15, 825);
                Attcb.LineTo(650, 825);
                Attcb.MoveTo(15, 805);
                Attcb.LineTo(650, 805);

                Attcb.Stroke();
                #endregion

                Attcb.BeginText();

                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks and Nos.", 55, 815, 0);
                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Containers", 270, 815, 0);
                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description of Goods", 450, 815, 0);
                int yy = _Yp - YDiff * 6;
                int DeffLine = 0;
                int deffMark = 0;
                int LineX = 0;
                int LineMark = 0;
                int deffcntr = 0;
                int LineCntr = 0;

                DeffLine = 0;

                if (LineX <= 70)
                {

                    for (int Lines = LCntrindex; Lines < _dtCntrValues.Length; Lines++)
                    {
                        //var arrayCntrNov = _dtCntr.Rows[Lines]["CntrDtls"].ToString().Split('\n');
                        if (_dtCntrValues.Length >= LCntrindex + 1)
                        {
                            var arrayCntrNov = SplitByLenght(_dtCntrValues[Lines].ToString(), 30);
                            for (int d = 0; d < arrayCntrNov.Length; d++)
                            {
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayCntrNov[d].ToUpper(), 230, 780 - deffcntr, 0);
                                deffcntr += 10;
                                LineCntr++;
                            }
                            Cot++;
                            LCntrindex++;
                            if (LineCntr == 70)
                            {
                                deffcntr += LineCntr - 13;
                                break;
                            }
                        }
                    }
                    for (int Lines = LMarkindex; Lines < arrayMarks.Length; Lines++)
                    {
                        if (arrayMarks.Length >= LMarkindex + 1)
                        {
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayMarks[Lines], 20, 780 - deffMark, 0);
                            deffMark += 10;
                        }

                        Cot++;
                        LineMark++;
                        LMarkindex++;
                        if (LineMark == 70)
                        {
                            //LMarkindex += LineMark;
                            break;
                        }
                    }



                    for (int Lines = LIndex; Lines < arrayDescription.Length; Lines++)
                    {
                        if (arrayDescription.Length >= LIndex + 1)
                        {
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayDescription[Lines], 400, 780 - DeffLine, 0);
                            DeffLine += 10;
                        }
                        Cot++;
                        LineX++;
                        LIndex++;
                        if (LineX == 70)
                        {
                            //LIndex += LineX;
                            break;
                        }
                    }

                    int DeferentLine = (DeffLine < deffMark) ? deffcntr : deffcntr;
                    DataTable _dtns = GetNotes(Data.id.ToString());
                    if (_dtns.Rows.Count > 0)
                    {
                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL CLAUSES", 30, 700 - DeferentLine, 0);
                        DeferentLine += 20;
                        var Notes = _dtns.Rows[0]["Notes"].ToString().Split('\n'); ;
                        for (int t = 0; t < Notes.Length; t++)
                        {
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Notes[t].ToString(), 30, 700 - DeferentLine, 0);
                            DeferentLine += 10;
                        }
                    }
                }

                DeffLine += 10;
                Attcb.Stroke();
                int RowColm = 770 - DeffLine;

                LineCount += Cot;


                Attcb.EndText();
                Attdocument.Close();
                pdfmp.AddFile(_AttFileName + SheetNo + ".pdf");
                SheetNo++;
            }

            pdfmp.Execute();
            string str = FileHidpath;
            byte[] fileBytes = System.IO.File.ReadAllBytes(FileHidpath);
            string mime = MimeMapping.GetMimeMapping(FileHidpath);
            string fileName = Path.GetFileName(FileHidpath);
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new ByteArrayContent(fileBytes);
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(mime);
            response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
            {
                FileName = fileName
            };

            return response;
          

        }

        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlinepdf/getBLPrintpdfBackPageLive")]
        public HttpResponseMessage getBLPrintpdfBackPageLive([FromBody] myBLPrint Data)
        {
            MemoryStream memoryStream = new MemoryStream();
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            int _Xp = 10, _Yp = 785, YDiff = 10;
            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 300, 0);
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BLbackpage.png"));
            png1.SetAbsolutePosition(15, 1);
            png1.ScalePercent(38f);
            doc.Add(png1);
            cb.Stroke();
            writer.CloseStream = false;
            doc.Close();
            byte[] byteInfo = memoryStream.ToArray();
            memoryStream.Write(byteInfo, 0, byteInfo.Length);
            memoryStream.Position = 0;

            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "BakPagePrint.pdf"
            };
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception("Failed to get a valid response.");
            }
            return response;
        }


        public DataTable GetBkgCustomer(string BLID)
        {
            string _Query = " select convert(varchar,BlDate, 103) as BlDatev,convert(varchar,SOBDate, 103) as SOBDatev,(SELECT top(1) NoofOriginal FROM NVO_BOL  where Id = NVO_BLRelease.BLID) as NoofOriginal,(select top(1) Notes from NVO_BLNotesClauses where DocID= 264 and NID=NVO_BLRelease.FreeDays) as ddlFreeday," +
                " (select top (1) case when GrsWtType=1 then 'KGS' else case when GrsWtType=2 then 'MTS' end end from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID= NVO_BLRelease.BLID)  as GrsWtType, " +
                " (select top(1) case when NtWtType = 1 then 'KGS' else case when NtWtType = 2 then 'MTS' end end  from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID = NVO_BLRelease.BLID) as NtWtType, " +
                " (select top(1) (select top(1) PkgDescription from NVO_CargoPkgMaster where NVO_CargoPkgMaster.Id = PakgType) from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID= NVO_BLRelease.BLID) as CargoPakage,(select top(1) PODID from NVO_BOL where ID =NVO_BLRelease.BLID) as PODID,(select top(1) (SELECT top(1) EQTypeID FROM NVO_tblCntrTypes WHERE NVO_tblCntrTypes.Size= NVO_BOLCntrDetails.Size) from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID= NVO_BLRelease.BLID) as EqTypeId, * from NVO_BLRelease  where BLID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetContainerDetails(string BLID)
        {
            string _Query = "select CntrDetails as CntrDtls from NVO_BLRelease where BLID=" + BLID;
            //string _Query = " Select(select top(1) CntrNo from NVO_Containers where Id = NVO_BOLCntrDetails.CntrID) + '/ ' + size + '/ ' + SealNo + '/ \n/' + convert(varchar, convert(decimal(8, 3), GrsWt)) + ' - ' + (case when GrsWtType = 1 then 'KGS' else 'MT' end) + '/' + convert(varchar, convert(decimal(8, 3), NtWt))  + '- ' + (case when NtWtType = 1 then 'KGS' else 'MT' end) + '/' + convert(varchar, convert(decimal(8, 3), CBM)) + 'CBM' as CntrDtls from NVO_BOLCntrDetails where BLID = " + BLID;

            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetNotes(string BLID)
        {
            string _Query = "select Notes from NVO_BLNotesClauses inner join NVO_BOL on NVO_BOL.PODID=NVO_BLNotesClauses.PortID where Id=" + BLID;

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
                           " isnull((select top 1 VM.VesselID from NVO_VesselMaster VM inner  join NVO_Voyage on NVO_Voyage.VesselID = VM.ID   where NVO_Voyage.ID = NVO_Booking.VesVoyID ),'') AS VesselIDValue, " +
                           "  (select top(1) DepAddress from NVO_DepotMaster  where NVO_DepotMaster.Id = NVO_Booking.PickUpDepotID)  as DepAddress " +
                            " from NVO_Booking " +
                            " where NVO_Booking.ID=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetBkgCntrValus(string BkgId)
        {
            string _Query = " select BKgID,NVO_tblCntrTypes.Size,Qty,GeneralName as Commodity from NVO_BookingCntrTypes inner join NVO_tblCntrTypes on NVO_tblCntrTypes.ID = NVO_BookingCntrTypes.CntrTypes inner join NVO_GeneralMaster on NVO_GeneralMaster.ID = NVO_BookingCntrTypes.CommodityType where NVO_BookingCntrTypes.BkgID=" + BkgId;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetNotesClausesBooking()
        {
            string _Query = "select * from NVO_BLNotesClauses where DocID=266";
            return Manag.GetViewData(_Query, "");
        }

        private readonly string uploadFolder = HttpContext.Current.Server.MapPath("~/BLFileAttached");

     
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlineupload/BLFileUpload")]
        public IHttpActionResult BLuploadfile()
        {
            // Ensure there is a file in the request
            if (HttpContext.Current.Request.Files.Count == 0)
            {
                return BadRequest("No file uploaded.");
            }

            // Get the uploaded file from the request
            var file = HttpContext.Current.Request.Files[0];
            if (file == null || file.ContentLength == 0)
            {
                return BadRequest("File is empty.");
            }

            // Validate file type (optional)
            string[] allowedExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".pdf" };
            string fileExtension = Path.GetExtension(file.FileName).ToLower();

            if (Array.IndexOf(allowedExtensions, fileExtension) < 0)
            {
                return BadRequest("Invalid file type.");
            }

            // Generate a unique file name
            string fileName = $"{Path.GetFileNameWithoutExtension(file.FileName)}_{Guid.NewGuid()}{fileExtension}";

            // Define the full path to save the file
            string filePath = Path.Combine(uploadFolder, fileName);

            // Save the file to the server
            try
            {
                file.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }

            // Return the file path (relative or absolute URL)
            return Ok(new { filePath = $"/UploadedFiles/{fileName}" });
        }
    }
}

    

