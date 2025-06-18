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

    public class BLPrintController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: BLPrint
        public ActionResult Index()
        {
            return View();
        }

        //        public FileResult BLGLOBALNETWORKPDF(string id, string printvalue, string LocID)
        //        {
        //            Document doc = new Document();
        //            Rectangle rec = new Rectangle(670, 900);
        //            doc = new Document(rec);
        //            Paragraph para = new Paragraph();

        //            DataTable _dt = GetBkgCustomer(id);
        //            string pdfpath = Server.MapPath("./pdfpath/");
        //            MergeEx pdfmp = new MergeEx();
        //            pdfmp.SourceFolder = pdfpath;

        //            pdfmp.DestinationFile = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";
        //            string FileHidpath = "./pdfpath/Multiple-" + Session.SessionID.ToString() + "BL.pdf";

        //            string _FileName = Session.SessionID.ToString() + id + 1;
        //            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));
        //            //PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
        //            doc.Open();

        //            PdfContentByte cb = writer.DirectContent;

        //            int _Xp = 10, _Yp = 785, YDiff = 10;

        //            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/pdflogos/globalnetwork.png"));
        //            png1.SetAbsolutePosition(15, 800);
        //            png1.ScalePercent(30f);
        //            doc.Add(png1);

        //            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader, 10);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 380, 820, 0);

        //            cb.BeginText();

        //            cb.MoveTo(15, 840);
        //            cb.LineTo(650, 840);

        //            cb.MoveTo(15, 790);
        //            cb.LineTo(650, 790);

        //            cb.MoveTo(330, 840);
        //            cb.LineTo(330, 470);

        //            cb.MoveTo(330, 740);
        //            cb.LineTo(650, 740);

        //            cb.MoveTo(485, 840);
        //            cb.LineTo(485, 740);

        //            cb.MoveTo(15, 710);
        //            cb.LineTo(650, 710);

        //            cb.MoveTo(15, 630);
        //            cb.LineTo(650, 630);

        //            cb.MoveTo(330, 600);
        //            cb.LineTo(650, 600);

        //            cb.MoveTo(15, 550);
        //            cb.LineTo(650, 550);

        //            cb.MoveTo(15, 510);
        //            cb.LineTo(330, 510);

        //            cb.MoveTo(15, 470);
        //            cb.LineTo(650, 470);


        //            cb.MoveTo(15, 230);
        //            cb.LineTo(650, 230);

        //            cb.MoveTo(15, 190);
        //            cb.LineTo(650, 190);

        //            cb.MoveTo(160, 230);
        //            cb.LineTo(160, 70);

        //            cb.MoveTo(235, 230);
        //            cb.LineTo(235, 190);

        //            cb.MoveTo(300, 230);
        //            cb.LineTo(300, 10);

        //            cb.MoveTo(400, 230);
        //            cb.LineTo(400, 190);

        //            cb.MoveTo(500, 230);
        //            cb.LineTo(500, 190);

        //            cb.MoveTo(160, 160);
        //            cb.LineTo(300, 160);

        //            cb.MoveTo(15, 110);
        //            cb.LineTo(300, 110);

        //            cb.MoveTo(15, 70);
        //            cb.LineTo(300, 70);

        //            //cb.MoveTo(350, 40);
        //            //cb.LineTo(450, 40);

        //            BaseFont Crossbfheader = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetColorFill(Color.LIGHT_GRAY);
        //            cb.SetFontAndSize(Crossbfheader, 70);
        //            if (printvalue == "1")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   D R A F T   ", 200, 200, 45);
        //            }

        //            if (printvalue == "2")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   FIRST ORIGINAL   ", 100, 200, 45);
        //            }
        //            if (printvalue == "3")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SECOND ORIGINAL   ", 100, 200, 45);
        //            }
        //            if (printvalue == "4")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   THIRD ORIGINAL   ", 100, 200, 45);
        //            }

        //            if (printvalue == "6")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   EXPRESS RELEASE   ", 100, 200, 45);
        //            }

        //            BaseFont bfheader13 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader13, 12);
        //            cb.SetColorFill(Color.BLACK);
        //            string Valuesv = "";
        //            if (printvalue == "1")
        //                Valuesv = "DRAFT";
        //            if (printvalue == "2")
        //                Valuesv = "FIRST ORIGINAL";
        //            if (printvalue == "3")
        //                Valuesv = "SECOND ORIGINAL";
        //            if (printvalue == "4")
        //                Valuesv = "THIRD ORIGINAL";
        //            if (printvalue == "5")
        //                Valuesv = "SEAWAY BL-NON NEGOTIABLE";
        //            if (printvalue == "6")
        //                Valuesv = "EXPRESS RELEASE";
        //            if (printvalue == "7")
        //                Valuesv = "SURRENDER BL";
        //            if (printvalue == "8")
        //                Valuesv = "RFS 1ST ORIGINAL";
        //            if (printvalue == "9")
        //                Valuesv = "RFS 2ND ORIGINAL";
        //            if (printvalue == "10")
        //                Valuesv = "RFS 3RD ORIGINAL";
        //            if (printvalue == "11")
        //                Valuesv = "BACK PAGE";
        //            if (printvalue == "12")
        //                Valuesv = "NON NEGOTIABLE";
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Valuesv, 475, 725, 0);

        //            BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader1, 8);
        //            cb.SetColorFill(Color.BLACK);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper", 15, 775, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Consignee (negotiable only if consigned 'to order','to order of' a named Person or", 15, 695, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "'to order of bearer')", 15, 686, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify Party (see clause 22)", 15, 615, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL No.", 495, 825, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Booking No.", 335, 775, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Export References", 335, 725, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Destination Agent", 335, 695, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Onward inland routing (Not Part of carriage as certified in", 335, 620, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "clause 1. For account and risk of merchant)", 335, 612, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of receipt. Applicable only when document used as", 335, 590, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Multimodal Transport B/L.(See clause 1)", 335, 582, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Delivery Applicable only when document used as", 335, 540, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Multimodal Transport B/L. (See clause 1)", 335, 530, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "*Vessel (see clause 1+19) ", 15, 535, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Voyage No", 155, 535, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port Of Loading", 15, 495, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Port Of Discharge", 155, 495, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "kind of Packages; Description of goods; Marks and Numbers; container No./Seal No", 15, 440, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gross Weight", 515, 425, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Net Weight", 515, 350, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Measurement", 595, 425, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Total Package", 15, 425, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Above particulars are declared by shipper; but without responsibility of or representation by carrier (see clause 4)", 15, 240, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight & Charge", 15, 215, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Rate", 195, 215, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Unit", 240, 215, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Currency", 305, 215, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Prepaid", 405, 215, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Collect", 505, 215, 0);


        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "On Board Date :", 15, 25, 0);




        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR OCEAN ", 335, 825, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TRANSPORT OR MULTIMODAL", 335, 816, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TRANSPORT", 335, 809, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PARTICULARS FURNISHED BY SHIPPER - CARRIER NOT RESPONSIBLE", 170, 455, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Declared Value (see clause 7.3)", 165, 175, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Place of Issue of BL", 165, 145, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Date of Issue", 165, 95, 0);
        //            BaseFont bfheader12 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader12, 8);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString().Trim().ToUpper(), 495, 810, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString().Trim().ToUpper(), 335, 760, 0);


        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["IssuedAt"].ToString().Trim().ToUpper(), 165, 130, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BlDatev"].ToString().Trim().ToUpper(), 165, 80, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["SOBDateV"].ToString().Trim().ToUpper(), 100, 25, 0);



        //            int ColumnRows = 765; int RowsColumn = 0;
        //            RowsColumn = 0;
        //            string[] ArrayAddress = Regex.Split(_dt.Rows[0]["Shipper"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["ShipperAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
        //            string[] Aaddsplit;

        //            for (int x = 0; x < ArrayAddress.Length; x++)
        //            {
        //                Aaddsplit = ArrayAddress[x].Split('\n');

        //                for (int k = 0; k < Aaddsplit.Length; k++)
        //                {

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 15, ColumnRows, 0);
        //                    ColumnRows -= 9;
        //                    RowsColumn++;
        //                }
        //            }

        //            ColumnRows = 670;
        //            RowsColumn = 0;
        //            string[] ArrayAddress1 = Regex.Split(_dt.Rows[0]["Consignee"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["ConsigneeAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
        //            string[] Aaddsplit1;

        //            for (int x = 0; x < ArrayAddress1.Length; x++)
        //            {
        //                Aaddsplit1 = ArrayAddress1[x].Split('\n');

        //                for (int k = 0; k < Aaddsplit1.Length; k++)
        //                {

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit1[k].ToString(), 15, ColumnRows, 0);
        //                    ColumnRows -= 9;
        //                    RowsColumn++;
        //                }
        //            }

        //            ColumnRows = 600;
        //            RowsColumn = 0;
        //            string[] ArrayAddress2 = Regex.Split(_dt.Rows[0]["Notify1"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["Notify1Address"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
        //            string[] Aaddsplit2;

        //            for (int x = 0; x < ArrayAddress2.Length; x++)
        //            {
        //                Aaddsplit2 = ArrayAddress2[x].Split('\n');

        //                for (int k = 0; k < Aaddsplit2.Length; k++)
        //                {

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit2[k].ToString(), 15, ColumnRows, 0);
        //                    ColumnRows -= 9;
        //                    RowsColumn++;
        //                }
        //            }


        //            ColumnRows = 685;
        //            RowsColumn = 0;

        //            string[] ArrayAddress3 = Regex.Split(_dt.Rows[0]["Agent"].ToString().Trim().ToUpper() + "\r" + _dt.Rows[0]["AgentAddress"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
        //            string[] Aaddsplit3;

        //            for (int x = 0; x < ArrayAddress3.Length; x++)
        //            {
        //                Aaddsplit3 = ArrayAddress3[x].Split('\n');

        //                for (int k = 0; k < Aaddsplit3.Length; k++)
        //                {

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit3[k].ToString(), 335, ColumnRows, 0);
        //                    ColumnRows -= 9;
        //                    RowsColumn++;
        //                }
        //            }
        //            var VessVoy = _dt.Rows[0]["VesVoy"].ToString().Trim().Split('-');
        //            if (VessVoy[0].Length == 1)
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, VessVoy[0].ToString() + "-" + VessVoy[1].ToString(), 15, 520, 0);
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, VessVoy[2].ToString(), 155, 520, 0);
        //            }
        //            else
        //            {

        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, VessVoy[0].ToString(), 15, 520, 0);
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, VessVoy[1].ToString(), 155, 520, 0);
        //            }
        //            var POLIDv = _dt.Rows[0]["POL"].ToString().Trim().ToUpper().Split(',');
        //            if (_dt.Rows[0]["POL"].ToString().Length > 25)
        //            {
        //                int xRow = 483;
        //                for (int i = 0; i < POLIDv.Length; i++)
        //                {
        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, POLIDv[i], 15, xRow, 0);
        //                    xRow -= 11;
        //                }
        //            }
        //            else
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POL"].ToString().Trim().ToUpper(), 15, 483, 0);
        //            }

        //            //var FPODv = _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper().Split(',');
        //            //if (_dt.Rows[0]["FPOD"].ToString().Length > 25)
        //            //{
        //            //    int xRow = 482;
        //            //    for (int i = 0; i < FPODv.Length; i++)
        //            //    {
        //            //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, FPODv[i], 335, xRow, 0);
        //            //        xRow -= 11;
        //            //    }
        //            //}
        //            //else
        //            //{
        //            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper(), 335, 482, 0);
        //            //}
        //            //var POOIDv = _dt.Rows[0]["POO"].ToString().Trim().ToUpper().Split(',');
        //            //if (_dt.Rows[0]["POO"].ToString().Length > 25)
        //            //{
        //            //    int xRow = 565;
        //            //    for (int i = 0; i < POOIDv.Length; i++)
        //            //    {
        //            //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, POOIDv[i], 335, xRow, 0);
        //            //        xRow -= 11;
        //            //    }
        //            //}
        //            //else
        //            //{
        //            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POO"].ToString().Trim().ToUpper(), 335, 565, 0);
        //            //}

        //            //var PODv = _dt.Rows[0]["POD"].ToString().Trim().ToUpper().Split(',');
        //            //if (_dt.Rows[0]["POD"].ToString().Length > 25)
        //            //{
        //            //    int xRow = 480;
        //            //    for (int i = 0; i < PODv.Length; i++)
        //            //    {
        //            //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, PODv[i], 155, xRow, 0);
        //            //        xRow -= 11;
        //            //    }
        //            //}
        //            //else
        //            //{
        //            //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POD"].ToString().Trim().ToUpper(), 155, 480, 0);
        //            //}

        //            // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POL"].ToString().Trim().ToUpper(), 15, 480, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FPOD"].ToString().Trim().ToUpper(), 335, 482, 0);

        //             cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POO"].ToString().Trim().ToUpper(), 335, 565, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["POD"].ToString().Trim().ToUpper(), 155, 480, 0);


        //            string[] arrayMarks = new string[] { };
        //            string[] arrayDescription = new string[] { };
        //            string[] arrayCntrNo = new string[] { };


        //            List<string> ArrayMarksV = new List<string>();
        //            arrayMarks = _dt.Rows[0]["Marks"].ToString().Split('\n');
        //            int intMarkCount = arrayMarks.Length + 7;
        //            arrayDescription = _dt.Rows[0]["Description"].ToString().Split('\n');

        //            if (_dt.Rows[0]["ddlFreeday"].ToString() != "")
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["intFreedays"].ToString() + '-' + _dt.Rows[0]["ddlFreeday"].ToString(), 275, 260, 0);
        //            }
        //            else
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 275, 260, 0);
        //            }

        //            int RowMx = 410;

        //            int TotalLine = 0;
        //            int ColumnCountMrks = 8;


        //            ColumnCountMrks = arrayMarks.Length;
        //            TotalLine = 8;
        //            for (int LineX = 0; LineX < TotalLine; LineX++)
        //            {
        //                if (arrayMarks.Length >= LineX + 1)

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayMarks[LineX].ToUpper(), 15, RowMx, 0);
        //                RowMx -= 10;
        //            }

        //            RowMx -= 5;
        //            int Rowcntr = 410;
        //            TotalLine = 4;
        //            DataTable _dtCntr = GetContainerDetails(id);
        //            var _dtCntrValues = _dtCntr.Rows[0]["CntrDtls"].ToString().Split('\n');
        //            int TotalColumnCntr = (_dtCntrValues.Length < TotalLine) ? _dtCntrValues.Length : TotalLine;
        //            if (_dtCntr.Rows.Count > 0)
        //            {
        //                TotalLine = TotalColumnCntr;
        //                for (int LineX = 0; LineX < TotalLine; LineX++)
        //                {
        //                    var arrayCntrNov = SplitByLenght(_dtCntrValues[LineX].ToString(), 30);
        //                    for (int d = 0; d < arrayCntrNov.Length; d++)
        //                    {
        //                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayCntrNov[d].ToUpper(), 15, RowMx, 0);
        //                        RowMx -= 10;
        //                    }
        //                }
        //            }





        //            int RowDec = 410;

        //            TotalLine = 15;

        //            for (int LineX = 0; LineX < TotalLine; LineX++)
        //            {
        //                if (arrayDescription.Length >= LineX + 1)

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayDescription[LineX], 275, RowDec, 0);
        //                RowDec -= 10;
        //            }

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["Packages"].ToString().Trim().ToUpper(), 215, 415, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["CargoPakage"].ToString().Trim().ToUpper(), 215, 400, 0);

        //            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gross Weight", 495, 440, 0);
        //            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Measurement", 595, 440, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["GRWT"].ToString().Trim().ToUpper() + " " + _dt.Rows[0]["GrsWtType"].ToString().Trim().ToUpper(), 515, 400, 0);
        //            if (_dt.Rows[0]["NTWT"].ToString().Trim().ToUpper() != "0.000")
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["NTWT"].ToString().Trim().ToUpper() + " " + _dt.Rows[0]["NtWtType"].ToString().Trim().ToUpper(), 515, 320, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["CBM"].ToString().Trim().ToUpper(), 595, 400, 0);
        //            int ColumnRowsA = 180;
        //            int RowsColumnA = 0;

        //            BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader5, 7);
        //            cb.SetColorFill(Color.BLACK);
        //            string text = @"SHIPPED, as far as ascertained by reasonable means of checking in apparent good order and condition 
        //unless otherwise stated herein the total numnber of quantity of containers or other packages indicated in the 
        //box opposite entitled *carrier's receipt* for carriage from the port of loading(or the place of receipt,if 
        //mentioned above) to the port of discharge(or the palce of delivery,if mentioned above),such carriage being 
        //always subject to the terms, rights, defences, provisions,conditions,exceptions,limitation and liberties 
        //hereof. where the bill of lading is negotiable,the carrier will be entitled to give delivery of the goods against 
        //what he resonably belives to be genuine original bill of lading. Delivery as aforesaid is authorised and shall 
        //constitute due delivery hereunder and the merchant shall have no claim for loss or non delivery .In accepting 
        //the bill of lading, any local customs or previleges to the contrary notwithstanding, the merchant agrees to be 
        //bound by all terms and conditions stated herein whether written ,printed,stamped or incorporated on the face 
        //or the reverse side hereof as fully as if they were all signed by the Merchant.
        //In WITNESS WHEREOF then number of Bill of Lading stated on this side have been signed and wherever 
        //one original Bill of Lading has been surrendered any other shall be void";
        //            string[] Arrayterms = Regex.Split(text, char.ConvertFromUtf32(13));
        //            string[] Addsplit1;

        //            for (int x = 0; x < Arrayterms.Length; x++)
        //            {
        //                Addsplit1 = Arrayterms[x].Split('\n');

        //                for (int k = 0; k < Addsplit1.Length; k++)
        //                {

        //                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Addsplit1[k].ToString(), 305, ColumnRowsA, 0);
        //                    ColumnRowsA -= 4;
        //                    RowsColumnA++;
        //                }
        //            }


        //            BaseFont bfheader8 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //            cb.SetFontAndSize(bfheader8, 8);
        //            cb.SetColorFill(Color.BLACK);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Carrier's Receipt(see clauses 1 and 14)", 15, 175, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Total Number of Container or Packages", 15, 166, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "received by Carrier", 15, 157, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Number & Sequence of Original B(s)/L :", 15, 95, 0);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight collection mode :", 15, 55, 0);

        //            //Muthu

        //            string NumberConvert = NumberConverWords.changeCurrencyToWords(_dt.Rows[0]["NoOfOriginal"].ToString());

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "("+ _dt.Rows[0]["NoOfOriginal"].ToString() + ")   "  + NumberConvert.Replace("Only",""), 15, 80, 0);


        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["FreightPayment"].ToString(), 140, 55, 0);

        //            ColumnRows = 140;
        //            RowsColumn = 0;

        //            DataTable dtC = GetCntrCount(id);
        //            if(dtC.Rows.Count >0)
        //            {
        //                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtC.Rows[0]["CntrCount"].ToString() + " container ", 15, ColumnRows, 0);
        //            }
        //            //
        //           // string[] TotalPkg = Regex.Split(_dtCntr.Rows[0]["CntrDtls"].ToString().Split('\n').Length + " container ", char.ConvertFromUtf32(10));
        //            //string[] AddSplitpkg;

        //            //for (int x = 0; x < TotalPkg.Length; x++)
        //            //{
        //            //    AddSplitpkg = TotalPkg[x].Split('\n');

        //            //    for (int k = 0; k < AddSplitpkg.Length; k++)
        //            //    {

        //            //        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, AddSplitpkg[k].ToString(), 15, ColumnRows, 0);
        //            //        ColumnRows -= 9;
        //            //        RowsColumn++;
        //            //    }
        //            //}

        //            cb.SetFontAndSize(bfheader8, 10);
        //            cb.SetColorFill(Color.BLACK);
        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["Notes"].ToString(), 360, 65, 0);

        //            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "___________________________________", 350, 15, 0);

        //            cb.EndText();
        //            cb.Stroke();
        //            doc.Close();
        //            pdfmp.AddFile(_FileName + ".pdf");

        //            //Largest  array number
        //            int[] arr = { arrayDescription.Length, intMarkCount, _dtCntrValues.Length + 10 };
        //            int max = arr[0];
        //            for (int i = 1; i < arr.Length; i++)
        //            {
        //                if (arr[i] > max)
        //                    max = arr[i];
        //            }

        //            // int TotalColumn = (_dt.Rows.Count < arrayDescription.Length) ? intMarkCount : intMarkCount;
        //            int TotalColumn = max;

        //            if(TotalColumn == TotalLine)
        //            {
        //                TotalColumn += 1;
        //            }
        //            int WriteLine = TotalColumn - TotalLine;
        //            int AttachedsheetNo = int.Parse(Math.Ceiling((WriteLine / 80.00)).ToString());
        //            int Cot = 0;
        //            //int LineCount = 15 + Cot;
        //            int LineCount = Cot;
        //            int SheetNo = 1;
        //            string Filesv = "Attach" + id;
        //            string _AttFileName = Filesv;
        //            int LIndex = 15;
        //            int LMarkindex = 8;
        //            int LCntrindex = 4;

        //            for (int k = 0; k < AttachedsheetNo; k++)
        //            {

        //                Document Attdocument = new Document(rec);
        //                PdfWriter Attwriter = PdfWriter.GetInstance(Attdocument, new FileStream(pdfpath + (_AttFileName + SheetNo) + ".pdf", FileMode.Create));
        //                Attdocument.Open();
        //                PdfContentByte Attcb = Attwriter.DirectContent;
        //                Attcb.SetColorStroke(Color.BLACK);


        //                #region Border
        //                Attcb.MoveTo(15, 825);
        //                Attcb.LineTo(650, 825);
        //                Attcb.MoveTo(15, 805);
        //                Attcb.LineTo(650, 805);

        //                Attcb.Stroke();
        //                #endregion

        //                Attcb.BeginText();

        //                BaseFont bfheader23 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //                Attcb.SetFontAndSize(bfheader23, 9);
        //                Attcb.SetColorFill(Color.BLACK);
        //                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Annexure Sheet", 280, 880, 0);

        //                BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //                Attcb.SetFontAndSize(bfheader21, 8);
        //                Attcb.SetColorFill(Color.BLACK);
        //                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks and Nos.", 55, 815, 0);
        //                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Containers", 270, 815, 0);
        //                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description of Goods", 450, 815, 0);
        //                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL No:", 15, 840, 0);
        //                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dt.Rows[0]["BLNo"].ToString(), 60, 840, 0);
        //                int yy = _Yp - YDiff * 6;
        //                int DeffLine = 0;
        //                int deffMark = 0;
        //                int LineX = 0;
        //                int LineMark = 0;
        //                int deffcntr = 0;
        //                int LineCntr = 0;
        //                DeffLine = 0;

        //                BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //                Attcb.SetFontAndSize(bfheader22, 8);
        //                Attcb.SetColorFill(Color.BLACK);
        //                if (LineX <= 70)
        //                {

        //                    for (int Lines = LCntrindex; Lines < _dtCntrValues.Length; Lines++)
        //                    {
        //                        //var arrayCntrNov = _dtCntr.Rows[Lines]["CntrDtls"].ToString().Split('\n');
        //                        if (_dtCntrValues.Length >= LCntrindex + 1)
        //                        {
        //                            var arrayCntrNov = SplitByLenght(_dtCntrValues[Lines].ToString(), 30);
        //                            if (arrayCntrNov.Length>0)
        //                            {
        //                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayCntrNov[0].ToUpper(), 230, 780 - deffcntr, 0);
        //                                deffcntr += 10;
        //                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayCntrNov[1].ToUpper(), 230, 780 - deffcntr, 0);
        //                                deffcntr += 10;
        //                                Cot++;
        //                                LineCntr++;
        //                            }
        //                            if (LineCntr == 70)
        //                            {
        //                                deffcntr += LineCntr - 13;
        //                                break;
        //                            }
        //                        }
        //                    }

        //                    for (int Lines = LMarkindex; Lines < arrayMarks.Length; Lines++)
        //                    {
        //                        if (arrayMarks.Length >= LMarkindex + 1)

        //                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayMarks[Lines], 20, 780 - deffMark, 0);

        //                        deffMark += 10;
        //                        Cot++;
        //                        LineMark++;
        //                        if (LineMark == 70)
        //                        {
        //                            LMarkindex += LineMark;
        //                            break;
        //                        }
        //                    }



        //                    for (int Lines = LIndex; Lines < arrayDescription.Length; Lines++)
        //                    {
        //                        if (arrayDescription.Length >= TotalLine + 1)

        //                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, arrayDescription[Lines], 400, 780 - DeffLine, 0);

        //                        DeffLine += 10;
        //                        Cot++;
        //                        LineX++;
        //                        if (LineX == 70)
        //                        {
        //                            LIndex += LineX;
        //                            //LIndex += LineX - 13;
        //                            break;
        //                        }
        //                    }

        //                    int DeferentLine = (DeffLine < deffMark) ? deffcntr : deffcntr;
        //                    DataTable _dtns = GetNotes(id);
        //                    if (_dtns.Rows.Count > 0)
        //                    {
        //                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL CLAUSES", 30, 700 - DeferentLine, 0);
        //                        DeferentLine += 20;
        //                        var Notes = _dtns.Rows[0]["Notes"].ToString().Split('\n'); ;
        //                        for (int t = 0; t < Notes.Length; t++)
        //                        {
        //                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Notes[t].ToString(), 30, 700 - DeferentLine, 0);
        //                            DeferentLine += 10;
        //                        }
        //                    }

        //                }


        //                DeffLine += 10;
        //                Attcb.Stroke();
        //                int RowColm = 770 - DeffLine;

        //                LineCount += Cot;

        //                Attcb.EndText();
        //                Attdocument.Close();
        //                pdfmp.AddFile(_AttFileName + SheetNo + ".pdf");
        //                SheetNo++;
        //            }

        //            pdfmp.Execute();
        //            string str = FileHidpath;
        //            string mime = MimeMapping.GetMimeMapping(FileHidpath);
        //            return File(FileHidpath, mime);
        //        }


        public FileResult BLBWSPrintPDF(string id, string printvalue, string LocID, string SessionFinYear)
        {
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();

            DataTable _dt = GetBkgCustomer(id);
            string pdfpath = Server.MapPath("./pdfpath/");
            MergeEx pdfmp = new MergeEx();
            pdfmp.SourceFolder = pdfpath;

            pdfmp.DestinationFile = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";
            string FileHidpath = "./pdfpath/Multiple-" + Session.SessionID.ToString() + "BL.pdf";

            string _FileName = Session.SessionID.ToString() + id + 1;
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


            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/pdfhead.png"));
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

            if (printvalue == "1")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   D R A F T   ", 200, 200, 45);
            }

            if (printvalue == "2")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   FIRST ORIGINAL   ", 100, 200, 45);
            }
            if (printvalue == "3")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SECOND ORIGINAL   ", 100, 200, 45);
            }
            if (printvalue == "4")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   THIRD ORIGINAL   ", 100, 200, 45);
            }

            if (printvalue == "6")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   EXPRESS RELEASE   ", 100, 200, 45);
            }


            BaseFont Crossbfheader6 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(Color.LIGHT_GRAY);
            cb.SetFontAndSize(Crossbfheader6, 40);

            if (printvalue == "5")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SEAWAY BL - NON NEGOTIABLE  ", 100, 200, 45);
            }
            if (printvalue == "11")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   BACK PAGE   ", 100, 200, 45);
            }
            if (printvalue == "7")
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SURRENDER BL  ", 200, 350, 45);
            }


            if (printvalue == "5")
            {
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "   SEAWAY BL", 100, 200, 45);
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

            if (printvalue == "8")
            {
                if (LocID == "25")
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "RECEIVED FOR SHIPMENT", 100, 200, 45);
                    iTextSharp.text.Image png15 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/Sign.png"));
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

            if (printvalue == "12")
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
            if (printvalue == "1")
                Valuesv = "DRAFT";
            if (printvalue == "2")
                Valuesv = "FIRST ORIGINAL";
            if (printvalue == "3")
                Valuesv = "SECOND ORIGINAL";
            if (printvalue == "4")
                Valuesv = "THIRD ORIGINAL";
            if (printvalue == "5")
                Valuesv = "SEAWAY BL-NON NEGOTIABLE";
            if (printvalue == "6")
                Valuesv = "EXPRESS RELEASE";
            if (printvalue == "7")
                Valuesv = "SURRENDER BL";
            if (printvalue == "8")
                Valuesv = "RFS FIRST ORIGINAL";
            if (printvalue == "9")
                Valuesv = "RFS 2ND ORIGINAL";
            if (printvalue == "10")
                Valuesv = "RFS 3RD ORIGINAL";
            if (printvalue == "11")
                Valuesv = "BACK PAGE";
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Valuesv, 420, 760, 0);

            if (printvalue == "2" || printvalue == "3" || printvalue == "4" || printvalue == "5" || printvalue == "7"
               || printvalue == "8" || printvalue == "9" || printvalue == "10")
            {
                var AutoGen = Manag.GetBLPrint_Number(SessionFinYear);
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





            if (printvalue == "5")
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
            DataTable _dtCntr = GetContainerDetails(id);
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
            string Filesv = "Attach" + id;
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
                iTextSharp.text.Image png11 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/pdfhead.png"));
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
                    DataTable _dtns = GetNotes(id);
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
            string mime = MimeMapping.GetMimeMapping(FileHidpath);
            return File(FileHidpath, mime);

        }

        public ActionResult GlobalBackPage(string printvalue, string layout)
        {
            GlobalBack(printvalue, layout);
            return View();
        }
        public void GlobalBack(string printvalue, string layout)
        {
            Document doc = new Document();
            Rectangle rec = new Rectangle(670, 900);
            doc = new Document(rec);
            Paragraph para = new Paragraph();


            PdfWriter writer = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();
            PdfContentByte cb = writer.DirectContent;
            int _Xp = 10, _Yp = 785, YDiff = 10;
            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 300, 0);
            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BLbackpage.png"));
            png1.SetAbsolutePosition(15, 1);
            png1.ScalePercent(38f);
            doc.Add(png1);
            cb.Stroke();
            writer.CloseStream = false;
            doc.Close();
            Response.Buffer = true;
            Response.ContentType = "application/pdf";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.End();

        }
        public DataTable GetCompanyDetails()
        {

            string _Query = "Select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBLLayoutValues()
        {
            string _Query = "select * from NVO_BLPrintLayout";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBkgCustomer(string BLID)
        {
            string _Query = " select convert(varchar,BlDate, 103) as BlDatev,(select ServiceType from NVO_Booking where NVO_Booking.ID = NVO_BLRelease.BkgId) as ServiceType,convert(varchar,SOBDate, 103) as SOBDatev,(SELECT top(1) NoofOriginal FROM NVO_BOL  where Id = NVO_BLRelease.BLID) as NoofOriginal,(select top(1) Notes from NVO_BLNotesClauses where DocID= 264 and NID=NVO_BLRelease.FreeDays) as ddlFreeday," +
                " (select top (1) case when GrsWtType=1 then 'KGS' else case when GrsWtType=2 then 'MTS' end end from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID= NVO_BLRelease.BLID)  as GrsWtType, " +
                " (select top(1) case when NtWtType = 1 then 'KGS' else case when NtWtType = 2 then 'MTS' end end  from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID = NVO_BLRelease.BLID) as NtWtType,(select top(1) vesselName from NVO_VesselMaster where NVO_VesselMaster.ID = NVO_Voyage.VesselID)as Vessel, " +
                " (select top(1) (select top(1) PkgDescription from NVO_CargoPkgMaster where NVO_CargoPkgMaster.Id = PakgType) from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID= NVO_BLRelease.BLID) as CargoPakage,(select top(1) PODID from NVO_BOL where ID =NVO_BLRelease.BLID) as PODID,(select top(1) (SELECT top(1) EQTypeID FROM NVO_tblCntrTypes WHERE NVO_tblCntrTypes.Size= NVO_BOLCntrDetails.Size) from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID= NVO_BLRelease.BLID) as EqTypeId, " +
                " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 288 and VoyageID=(select top(1)VesVoyID from NVO_Booking where ID = NVO_BLRelease.BkgId))as PreCarVes, (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 289 and VoyageID = (select top(1)VesVoyID from NVO_Booking where ID = NVO_BLRelease.BkgId))as PreCarVoy, " +
                " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID = 296 and VoyageID = (select top(1)VesVoyID from NVO_Booking where ID = NVO_BLRelease.BkgId))as PreCarVoyNepal,VesVoy,(select top(1) Notes from NVO_BLNotesClauses where NID= NVO_BLRelease.SignedAs) as Notes, " +
                " * from NVO_BLRelease  " +
                " inner join NVO_BOL on NVO_BOL.ID = NVO_BLRelease.BLID  " +
                " inner join NVO_Voyage on NVO_Voyage.ID = NVO_BOL.BLVesVoyID " +
                " where BLID = " + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCntrCount(string BLID)
        {
            string _Query = "select count(CntrID) as CntrCount from NVO_BOLCntrDetails where BLID= " + BLID;
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

       
    }
}