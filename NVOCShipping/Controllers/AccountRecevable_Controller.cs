using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using DataManager;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Security.Cryptography;
namespace NVOCShipping.Controllers
{
    public class AccountRecevable_Controller : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: AccountRecevable_
        public ActionResult Index()
        {
            return View();
        }

        public void AccountRecevable_SummaryExcelReport(string PartyID)
        {

            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("AccountReceivableSummaryReport");

            int RowIndex = 2;
            int MaxCol = 6;
            ws.Cells[RowIndex, 1].Value = "ACCOUNT RECEIVABLE DETAILS";
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
            ws.Cells["A7"].Value = "S. No.";
            ws.Cells["B7"].Value = "Invoice Date";
            ws.Cells["C7"].Value = "Invoice No";
            ws.Cells["D7"].Value = "Invoice Amount";
            ws.Cells["E7"].Value = "Received Amount";
            ws.Cells["F7"].Value = "Balance Amount";
            ExcelRange r;
            r = ws.Cells["A7:f7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
            int RowLine = 8;
            DataTable dtx = AccRecInvDtls(PartyID);
            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                ws.Cells["A" + RowLine].Value = RowLine - 7;
                ws.Cells["B" + RowLine].Value = dtx.Rows[i]["InvDate"].ToString();
                ws.Cells["C" + RowLine].Value = dtx.Rows[i]["InvoiceNo"].ToString();
                ws.Cells["D" + RowLine].Value = dtx.Rows[i]["InvTotal"];
                ws.Cells["E" + RowLine].Value = dtx.Rows[i]["TotalReceived"];
                ws.Cells["F" + RowLine].Value = dtx.Rows[i]["OutStanding"];
                RowLine++;
            }

            ws.Cells["A7:f" + RowLine].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:f" + RowLine].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:F" + RowLine].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:F" + RowLine].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells.AutoFitColumns();


            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=AccountReceivableSummary.xlsx");
            Response.End();

        }


        public void AccountRecevable_SummaryPDFReport(string PartyID)
        {



            Document doc = new Document(PageSize.A4, 25, 25, 25, 15);
            string fileAttachedEmail = Session.SessionID.ToString() + 2 + ".Pdf";
            //PdfWriter pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/FileHRAppointment\\" + fileAttachedEmail), FileMode.Create));
            PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();


            PdfContentByte cb = pdfWriter.DirectContent;
            //cb.BeginText();
            //BaseFont Crossbfheader = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            //cb.SetColorFill(iTextSharp.text.Color.LIGHT_GRAY);
            //cb.SetFontAndSize(Crossbfheader, 15);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Salary", 50, 610, 90);
            //cb.EndText();


            PdfPTable tbllogo = new PdfPTable(7);
            tbllogo.WidthPercentage = 100;
            tbllogo.HorizontalAlignment = Element.ALIGN_LEFT;
            PdfPCell cell = new PdfPCell();
            doc.Add(tbllogo);


            //var imag = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
            //imag.SetAbsolutePosition(25, 750);
            //imag.ScalePercent(20f);
            //doc.Add(imag);

            cell = new PdfPCell(new Phrase(".", FontFactory.GetFont("'Oswald', sans-serif", 10, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.WHITE)));
            cell.BorderWidth = 0;
            cell.Colspan = 1;
            tbllogo.AddCell(cell);


            PdfPTable tbl3 = new PdfPTable(12);
            tbl3.WidthPercentage = 100;
            tbl3.HorizontalAlignment = Element.ALIGN_LEFT;

            cell = new PdfPCell(new Phrase("ACCOUNT RECEIVABLE SUMMARY", FontFactory.GetFont("'Oswald', sans-serif", 12, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 12;
            cell.BorderWidth = 0;
            cell.BorderWidthLeft = 1;
            cell.BorderWidthRight = 1;
            cell.BorderWidthTop = 1;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase(".", FontFactory.GetFont("'Oswald', sans-serif", 12, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.WHITE)));
            cell.Colspan = 12;
            cell.BorderWidth = 0;
            cell.BorderWidthLeft = 1;
            cell.BorderWidthRight = 1;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase("SNo", FontFactory.GetFont("'Oswald', sans-serif", 9, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 2;
            cell.BorderWidth = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase("Invoice Date", FontFactory.GetFont("'Oswald', sans-serif", 9, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 2;
            cell.BorderWidth = 1;
           
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase("Invoice No", FontFactory.GetFont("'Oswald', sans-serif", 9, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 2;
            cell.BorderWidth = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase("Invoice Amount", FontFactory.GetFont("'Oswald', sans-serif", 9, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 2;
            cell.BorderWidth = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase("Received Amount", FontFactory.GetFont("'Oswald', sans-serif", 9, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 2;
            cell.BorderWidth = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbl3.AddCell(cell);

            cell = new PdfPCell(new Phrase("Balance Amount", FontFactory.GetFont("'Oswald', sans-serif", 9, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLACK)));
            cell.Colspan = 2;
            cell.BorderWidth = 1;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            tbl3.AddCell(cell);
            int RowIndex = 1;
            DataTable dtx = AccRecInvDtls(PartyID);
            for (int i = 0; i < dtx.Rows.Count; i++)
            {

                cell = new PdfPCell(new Phrase("" + RowIndex, FontFactory.GetFont("'Oswald', sans-serif", 7, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE)));
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                cell.BackgroundColor = new iTextSharp.text.Color(220, 220, 220);
                tbl3.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtx.Rows[i]["InvDate"].ToString(), FontFactory.GetFont("'Oswald', sans-serif", 7, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE)));
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                cell.BackgroundColor = new iTextSharp.text.Color(220, 220, 220);
                tbl3.AddCell(cell);

                cell = new PdfPCell(new Phrase(dtx.Rows[i]["InvoiceNo"].ToString(), FontFactory.GetFont("'Oswald', sans-serif", 7, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE)));
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                cell.BackgroundColor = new iTextSharp.text.Color(220, 220, 220);
                tbl3.AddCell(cell);

                cell = new PdfPCell(new Phrase(decimal.Parse(dtx.Rows[i]["InvTotal"].ToString()).ToString("#,#0.00"), FontFactory.GetFont("'Oswald', sans-serif", 7, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE)));
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                cell.BackgroundColor = new iTextSharp.text.Color(220, 220, 220);
                tbl3.AddCell(cell);

                cell = new PdfPCell(new Phrase(decimal.Parse(dtx.Rows[i]["TotalReceived"].ToString()).ToString("#,#0.00"), FontFactory.GetFont("'Oswald', sans-serif", 7, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE)));
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                cell.BackgroundColor = new iTextSharp.text.Color(220, 220, 220);
                tbl3.AddCell(cell);

                cell = new PdfPCell(new Phrase(decimal.Parse(dtx.Rows[i]["OutStanding"].ToString()).ToString("#,#0.00"), FontFactory.GetFont("'Oswald', sans-serif", 7, iTextSharp.text.Font.BOLDITALIC, iTextSharp.text.Color.BLUE)));
                cell.Colspan = 2;
                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                cell.BackgroundColor = new iTextSharp.text.Color(220, 220, 220);
                tbl3.AddCell(cell);
                RowIndex++;
            }

            doc.Add(tbl3);

            pdfWriter.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(doc);
            Response.End();


        }

        public DataTable AccRecInvDtls(string PartyID)
        {
            string _Query = " select PartyID,InvId,InvoiceNo,CustomerName,convert(varchar, Invdate, 100) as InvDate,InvTotal,TotalReceived,OutStanding from NVO_V_Account_Recevable where PartyID=" + PartyID + " order by InvDate desc";
            return Manag.GetViewData(_Query, "");
        }
    }
}