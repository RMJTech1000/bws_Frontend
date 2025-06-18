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
    public class ExportReportPDFController : Controller
    {
        ExportReportManager RegMng = new ExportReportManager();
        // GET: ExportPDF
        public ActionResult Index()
        {
            return View();
        }

        //public ActionResult ExportFreightManifestPDF(string BLID,string VesVoyID)
        //{
        //    BindFreightManifestPDF(BLID,VesVoyID);
        //    return View();
        //}
        public FileResult ExportFreightManifestPDF(string BLID, string VesVoyID, string AgencyID)
        {

            MergeEx pdfmp = new MergeEx();
            string pdfpath = Server.MapPath("./pdfpath/");
            pdfmp.SourceFolder = pdfpath;
            pdfmp.DestinationFile = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";
            string FileHidpath = "./pdfpath/Multiple-" + Session.SessionID.ToString() + "BL.pdf";
            DataTable _dtm = GetFreghtManifestMainDtls(VesVoyID, BLID);
            for (int x = 0; x < _dtm.Rows.Count; x++)
            {

                Document doc = new Document();
                Rectangle rec = new Rectangle(700, 950);
                doc = new Document(rec);
                doc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());

                string _FileName = Session.SessionID.ToString() + _dtm.Rows[x]["BLNo"].ToString() + 1;
                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));
                doc.Open();
                PdfContentByte cb = pdfWriter.DirectContent;

                //cb.SetColorStroke(Color.BLACK);
                int _Xp = 10, _Yp = 785, YDiff = 10;
                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);

                DataTable dtc = GetAgencyDetails(AgencyID);
                //if (dtc.Rows.Count > 0)
                //{
                //    if (dtc.Rows[0]["LogoPath"].ToString() != "")
                //    {
                //        if ( AgencyID == "69" || AgencyID == "84" || AgencyID == "8" || AgencyID == "11" || AgencyID == "19" || AgencyID == "20" || AgencyID == "56" || AgencyID == "10" || AgencyID == "62")
                //        {
                //            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                //            png1.SetAbsolutePosition(15, 515);
                //            png1.ScalePercent(60f);
                //            doc.Add(png1);
                //        }
                //        if (AgencyID == "3" || AgencyID == "9" || AgencyID == "19" || AgencyID == "20")
                //        {
                //            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                //            png1.SetAbsolutePosition(15, 518);
                //            png1.ScalePercent(38f);
                //            doc.Add(png1);
                //        }
                //        if ( AgencyID == "13" || AgencyID == "14" || AgencyID == "15" || AgencyID == "16")
                //        {
                //            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                //            png1.SetAbsolutePosition(15, 515);
                //            png1.ScalePercent(7f);
                //            doc.Add(png1);
                //        }
                //        if (AgencyID == "2")
                //        {
                //            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                //            png1.SetAbsolutePosition(15, 515);
                //            png1.ScalePercent(7f);
                //            doc.Add(png1);
                //        }
                //    }
                //    else
                //    {
                //        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/blanklogo.png"));
                //        png1.SetAbsolutePosition(15, 515);
                //        png1.ScalePercent(7f);
                //        doc.Add(png1);
                //    }

                //}

                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.png"));
                png1.SetAbsolutePosition(15, 515);
                png1.ScalePercent(7f);
                doc.Add(png1);
                cb.BeginText();
                BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader21, 16);
                cb.SetColorFill(new Color(0, 0, 128));

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 350, 570, 0);


                int RowIndexAdd = 550;
                BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader22, 9);
                cb.SetColorFill(new Color(0, 0, 128));
                var splitadd = dtc.Rows[0]["Address"].ToString().Split('\n');
                for (int k = 0; k < splitadd.Length; k++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitadd[k].ToString(), 350, RowIndexAdd, 0);
                    RowIndexAdd -= 10;
                }
                RowIndexAdd -= 15;
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 480, 540, 0);
                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 480, 530, 0);

                cb.EndText();
                //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/oclheader.jpg"));
                //png2.SetAbsolutePosition(320, 530);
                //png2.ScalePercent(52f);
                //doc.Add(png2);
                cb.SetColorStroke(new Color(133, 210, 238));


                BaseFont bfheader31 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader31, 10);
                cb.BeginText();
                //cb.MoveTo(480, 550);
                //cb.LineTo(700, 550);
                //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/companyborder.png"));
                //png2.SetAbsolutePosition(480, 550);
                //png2.ScalePercent(50f);
                //doc.Add(png2);
                cb.EndText();

                cb.SetColorStroke(new Color(0, 0, 128));
                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader2, 10);
                cb.SetColorFill(Color.BLACK);

                cb.BeginText();

                cb.MoveTo(0, 510);
                cb.LineTo(850, 510);

                cb.MoveTo(0, 480);
                cb.LineTo(850, 480);

                cb.MoveTo(10, 460);
                cb.LineTo(820, 460);

                cb.MoveTo(10, 400);
                cb.LineTo(820, 400);
                cb.MoveTo(10, 460);
                cb.LineTo(10, 400);
                cb.MoveTo(820, 460);
                cb.LineTo(820, 400);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREIGHT MANIFEST", 330, 490, 0);
                cb.EndText();

                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.BeginText();
                cb.MoveTo(10, 755);
                cb.LineTo(660, 755);
                cb.MoveTo(10, 690);
                cb.LineTo(660, 690);
                cb.MoveTo(10, 755);
                cb.LineTo(10, 690);
                cb.MoveTo(660, 755);
                cb.LineTo(660, 690);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL", 15, 450, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD", 15, 430, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POL", 15, 410, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOYAGE", 300, 450, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA", 300, 430, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 300, 410, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN NO", 550, 450, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOAD TERMINAL", 550, 430, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["VesVoy"].ToString(), 80, 450, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["ETD"].ToString(), 80, 430, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["POL"].ToString(), 80, 410, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["VoyageNo"].ToString(), 400, 450, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["ETA"].ToString(), 400, 430, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["POD"].ToString(), 400, 410, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 650, 450, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtm.Rows[x]["LoadTerminal"].ToString(), 650, 430, 0);


                cb.EndText();

                //TABLE

                int RowIndex = 380;
                DataTable _dtBL = GetFrieghtManifestDtls(_dtm.Rows[x]["BLID"].ToString());
                if (_dtBL.Rows.Count > 0)
                {
                    cb.BeginText();
                    BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader4, 8);
                    cb.SetColorFill(Color.BLACK);
                    cb.MoveTo(10, 378);
                    cb.LineTo(820, 378);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Bill Of Ladding No", 15, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper/Consignee/Notify", 120, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Marks & Numbers", 350, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "No Of Package", 460, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Goods", 580, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Gross.Weight", 730, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CBM", 790, RowIndex, 0);
                    RowIndex -= 15;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["BLNo"].ToString(), 15, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Shipper", 120, RowIndex, 0);
                    RowIndex -= 5;

                    var splitMarks = _dtBL.Rows[0]["MarkNo"].ToString().Split('\n');

                    int RowIndexv = 350;
                    for (int j = 0; j < splitMarks.Length; j++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitMarks[j].ToString(), 350, RowIndexv, 0);
                        RowIndexv -= 12;
                    }
                    decimal intGrwet = 0;
                    decimal intPakage = 0;
                    decimal intCBM = 0;
                    for (int h = 0; h < _dtBL.Rows.Count; h++)
                    {
                        intPakage += decimal.Parse(_dtBL.Rows[h]["NoOfPkg"].ToString());

                    }
                    intGrwet += decimal.Parse(_dtBL.Rows[0]["GrsWt"].ToString());
                    //intCBM += decimal.Parse(_dtBL.Rows[0]["CBM"].ToString());
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, intPakage.ToString() + " " + _dtBL.Rows[0]["CargoPakage"].ToString(), 470, RowIndex, 0);

                    RowIndexv = 350;


                    var splitDesc = _dtBL.Rows[0]["CagoDescription"].ToString().Split('\n');

                    int TotalLine = 22;
                    int TotalColumn = (TotalLine > splitDesc.Length) ? splitDesc.Length : TotalLine;



                    for (int j = 0; j < TotalColumn; j++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitDesc[j].ToString(), 570, RowIndexv, 0);
                        RowIndexv -= 12;
                    }

                    //for (int j = 0; j < splitDesc.Length; j++)
                    //{
                    //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitDesc[j].ToString(), 570, RowIndexv, 0);
                    //    RowIndexv -= 12;
                    //}

                    // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[i]["CagoDescription"].ToString(), 630, RowIndex, 0);


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, intGrwet.ToString(), 730, RowIndex, 0);
                    //RowIndex = 350;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["GrsWtType"].ToString(), 730, 350, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Net.Weight", 730, 330, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["NetWeight"].ToString(), 730, 310, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["NtWtType"].ToString(), 730, 300, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["CBM"].ToString(), 790, RowIndex, 0);


                    RowIndex -= 10;
                    var splitShipper = _dtBL.Rows[0]["Shipper"].ToString().Split('\n');


                    for (int j = 0; j < splitShipper.Length; j++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitShipper[j].ToString(), 120, RowIndex, 0);
                        RowIndex -= 12;
                    }

                    var split = _dtBL.Rows[0]["ShipperAddress"].ToString().Split('\n');
                    for (int k = 0; k < split.Length; k++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, split[k].ToString(), 120, RowIndex, 0);
                        RowIndex -= 12;
                    }
                    RowIndex -= 15;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Consignee", 120, RowIndex, 0);
                    RowIndex -= 15;
                    var splitConsignee = _dtBL.Rows[0]["Consignee"].ToString().Split('\n');

                    for (int j = 0; j < splitConsignee.Length; j++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitConsignee[j].ToString(), 120, RowIndex, 0);
                        RowIndex -= 12;
                    }


                    var splitConsigneeAddress = _dtBL.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                    for (int k = 0; k < splitConsigneeAddress.Length; k++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitConsigneeAddress[k].ToString(), 120, RowIndex, 0);
                        RowIndex -= 12;
                    }
                    RowIndex -= 15;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Notify", 120, RowIndex, 0);
                    RowIndex -= 15;
                    var splitNotify = _dtBL.Rows[0]["Notify1"].ToString().Split('\n');

                    for (int j = 0; j < splitNotify.Length; j++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitNotify[j].ToString(), 120, RowIndex, 0);
                        RowIndex -= 12;
                    }

                    var splitNotifyAddress = _dtBL.Rows[0]["Notify1Address"].ToString().Split('\n');
                    for (int k = 0; k < splitNotifyAddress.Length; k++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitNotifyAddress[k].ToString(), 120, RowIndex, 0);
                        RowIndex -= 12;
                    }
                    RowIndex -= 40;

                    if (_dtBL.Rows[0]["ddlFreeday"].ToString() != "")
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["IntFreeDays"].ToString() + '-' + _dtBL.Rows[0]["ddlFreeday"].ToString(), 560, RowIndex, 0);
                    }
                    else
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 560, RowIndex, 0);
                    }



                    RowIndex -= 20;
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Contuinuity as per annexure attached", 600, RowIndex, 0);

                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container Size/Type/Seal", 15, RowIndex, 0);
                    //RowIndex -= 12;
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["Seal"].ToString(), 15, RowIndex, 0);
                    //RowIndex -= 15;
                    // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Details", 510, RowIndex, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Freight Collection Mode :" + _dtBL.Rows[0]["FreightPayment"].ToString(), 15, RowIndex, 0);

                    RowIndex -= 15;



                    cb.EndText();
                    cb.Stroke();
                    doc.Close();
                    pdfmp.AddFile(_FileName + ".pdf");


                    DataTable _dtCnt = GetCntrdisplay(_dtm.Rows[x]["BLID"].ToString());
                    int CntrCount = int.Parse(Math.Ceiling((_dtCnt.Rows.Count / 30.00)).ToString());
                   // int CntrCount = _dtCnt.Rows.Count/32;

                    int Attachrowcount = 0;
                    int AttachFull = 0;
                    int AttachedsheetNo = CntrCount;
                    int Cot = 0;
                    int LineCount = 15 + Cot;
                    int SheetNo = 1;
                    string Filesv = "Attach" + 1;
                    string _AttFileName = Filesv + _dtm.Rows[x]["BLID"].ToString();
                    int LIndex = 15;


                    for (int k = 0; k < AttachedsheetNo; k++)
                    {

                        Document Attdocument = new Document(rec);
                        Attdocument.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
                        PdfWriter Attwriter = PdfWriter.GetInstance(Attdocument, new FileStream(pdfpath + (_AttFileName + SheetNo) + ".pdf", FileMode.Create));
                        Attdocument.Open();
                        PdfContentByte Attcb = Attwriter.DirectContent;
                        Attcb.SetColorStroke(Color.BLACK);

                        BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        Attcb.SetFontAndSize(bfheader, 14);

                        //if (dtc.Rows[0]["LogoPath"].ToString() != "")
                        //{
                        //    if (AgencyID == "22" || AgencyID == "60" || AgencyID == "61" || AgencyID == "69" || AgencyID == "84" || AgencyID == "8" || AgencyID == "11" || AgencyID == "56" || AgencyID == "10" || AgencyID == "62")
                        //    {
                        //        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        //        png1.SetAbsolutePosition(15, 515);
                        //        png1.ScalePercent(60f);
                        //        Attdocument.Add(png1);
                        //    }
                        //    if (AgencyID == "3" || AgencyID == "9" || AgencyID == "19" || AgencyID == "20")
                        //    {
                        //        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        //        png1.SetAbsolutePosition(15, 518);
                        //        png1.ScalePercent(38f);
                        //        Attdocument.Add(png1);
                        //    }
                        //    if (AgencyID == "1" || AgencyID == "13" || AgencyID == "14" || AgencyID == "15" || AgencyID == "16")
                        //    {
                        //        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        //        png1.SetAbsolutePosition(15, 515);
                        //        png1.ScalePercent(7f);
                        //        Attdocument.Add(png1);
                        //    }
                        //    if (AgencyID == "2")
                        //    {
                        //        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        //        png1.SetAbsolutePosition(15, 515);
                        //        png1.ScalePercent(7f);
                        //        Attdocument.Add(png1);
                        //    }
                        //}
                        //else
                        //{
                        //    iTextSharp.text.Image png11 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/blanklogo.png"));
                        //    png11.SetAbsolutePosition(15, 517);
                        //    png11.ScalePercent(7f);
                        //    Attdocument.Add(png11);
                        //}

                        iTextSharp.text.Image png11 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.png"));
                        png11.SetAbsolutePosition(15, 515);
                        png11.ScalePercent(60f);
                        Attdocument.Add(png1);

                        Attcb.BeginText();
                        BaseFont bfheader25 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        Attcb.SetFontAndSize(bfheader25, 16);
                        Attcb.SetColorFill(new Color(0, 0, 128));

                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 350, 570, 0);


                        int RowIndexAdd1 = 550;
                        BaseFont bfheader26 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        Attcb.SetFontAndSize(bfheader26, 9);
                        Attcb.SetColorFill(new Color(0, 0, 128));


                        var splitadd1 = dtc.Rows[0]["Address"].ToString().Split('\n');
                        for (int v = 0; v < splitadd1.Length; v++)
                        {
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitadd[v].ToString(), 350, RowIndexAdd1, 0);
                            RowIndexAdd1 -= 10;
                        }
                        RowIndexAdd1 -= 15;
                        //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 480, 540, 0);
                        //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 480, 530, 0);

                        Attcb.EndText();
                        Attcb.SetColorStroke(new Color(133, 210, 238));


                        BaseFont bfheader32 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        Attcb.SetFontAndSize(bfheader32, 10);
                        Attcb.BeginText();
                        //iTextSharp.text.Image png21 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/companyborder.png"));
                        //png21.SetAbsolutePosition(480, 550);
                        //png21.ScalePercent(50f);
                        //Attdocument.Add(png21);
                        Attcb.EndText();

                        Attcb.BeginText();

                        BaseFont bfheader41 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        Attcb.SetFontAndSize(bfheader41, 8);
                        Attcb.SetColorFill(Color.BLACK);
                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CHARGES BREAKUP", 15, 490, 0);
                        Attcb.MoveTo(15, 475);
                        Attcb.LineTo(555, 475);

                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Container No/Seal", 20, 465, 0);
                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Seal No", 160, 465, 0);
                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Type", 300, 465, 0);


                        int inxColumn = 355;
                        if (_dtBL.Rows[0]["RRID"].ToString() != "")
                        {
                            DataTable _dtRnew = GetCntrChargesBreakupNew(_dtBL.Rows[0]["RRID"].ToString(), _dtm.Rows[x]["BLID"].ToString());
                            for (int z = 0; z < _dtRnew.Rows.Count; z++)
                            {
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtRnew.Rows[z]["ChgCode"].ToString(), inxColumn, 465, 0);
                                inxColumn += 50;

                            }
                        }
                        else
                        {
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", inxColumn, 465, 0);
                        }

                        //DataTable _dtR = GetCntrChargesBreakup(_dtBL.Rows[0]["RRID"].ToString());
                        //for (int z = 0; z < _dtR.Rows.Count; z++)
                        //{
                        //    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtR.Rows[z]["ChgCode"].ToString(), inxColumn, 465, 0);
                        //    inxColumn += 50;

                        //}
                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Total", inxColumn, 465, 0);


                        Attcb.MoveTo(15, 457);
                        Attcb.LineTo(555, 457);

                        //var cntrsplit = SplitByLenght(_dtBL.Rows[0]["cntrdetails"].ToString(),50);

                        //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["Seal"].ToString(), 20, 447, 0);
                        //Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtBL.Rows[0]["Size"].ToString(), 300, 447, 0);
                        decimal Totalv1 = 0;
                        int ColumnCntr = 447;
                        //int ColumnCntrint = 447;
                        int ColumnRow = 355;
                        string BLIdvs = _dtm.Rows[x]["BLID"].ToString();
                        // DataTable _dtCnt = GetCntrdisplay(_dtm.Rows[x]["BLID"].ToString());
                        Attachrowcount = 0;
                        for (int j = AttachFull; j < _dtCnt.Rows.Count; j++)
                        {
                           
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtCnt.Rows[j]["CntrNo"].ToString(), 20, ColumnCntr, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtCnt.Rows[j]["SealNo"].ToString(), 160, ColumnCntr, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtCnt.Rows[j]["size"].ToString(), 300, ColumnCntr, 0);
                            ColumnRow = 355;
                            Totalv1 = 0;

                            if (_dtBL.Rows[0]["RRID"].ToString() != "")
                            {
                                DataTable _dtRm = GetCntrChargesBreakupNew(_dtBL.Rows[0]["RRID"].ToString(), _dtm.Rows[x]["BLID"].ToString());

                                for (int y = 0; y < _dtRm.Rows.Count; y++)
                                {
                                    DataTable dtmr = GetCntrChargesBreakup(_dtBL.Rows[0]["RRID"].ToString(), _dtRm.Rows[y]["ChargeCodeID"].ToString(), _dtCnt.Rows[j]["TypeID"].ToString(), _dtm.Rows[x]["BLID"].ToString());
                                    for (int r = 0; r < dtmr.Rows.Count; r++)
                                    {
                                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtmr.Rows[r]["ManifRate"].ToString(), ColumnRow, ColumnCntr, 0);
                                        Totalv1 += decimal.Parse(dtmr.Rows[r]["ManifRate"].ToString());

                                        ColumnRow += 50;
                                    }
                                }
                            }
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Totalv1.ToString(), ColumnRow, ColumnCntr, 0);
                            ColumnCntr -= 15;
                            Attachrowcount++;
                            AttachFull++;
                            if ((Attachrowcount + 1) ==30)
                            {
                                break;
                            }
                        }


                        Attcb.MoveTo(15, 475);
                        Attcb.LineTo(15, ColumnCntr);

                        Attcb.MoveTo(140, 475);
                        Attcb.LineTo(140, ColumnCntr);


                        Attcb.MoveTo(295, 475);
                        Attcb.LineTo(295, ColumnCntr);
                        Attcb.MoveTo(350, 475);
                        Attcb.LineTo(350, ColumnCntr);
                        Attcb.MoveTo(400, 475);
                        Attcb.LineTo(400, ColumnCntr);
                        Attcb.MoveTo(450, 475);
                        Attcb.LineTo(450, ColumnCntr);
                        Attcb.MoveTo(555, 475);
                        Attcb.LineTo(555, ColumnCntr);
                        Attcb.MoveTo(15, ColumnCntr);
                        Attcb.LineTo(555, ColumnCntr);
                        RowIndex -= 100;

                        ColumnCntr -= 50;
                        int TotalRemondingColumn = (splitDesc.Length - TotalLine);
                        if (TotalLine < splitDesc.Length)
                        {
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Description Of Goods", 20, ColumnCntr, 0);
                            ColumnCntr -= 15;
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "(continuity)", 20, ColumnCntr, 0);
                            ColumnCntr -= 30;

                            for (int f = 22; f < splitDesc.Length; f++)
                            {
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitDesc[f].ToString(), 20, ColumnCntr, 0);
                                ColumnCntr -= 15;
                            }
                        }



                        Attcb.Stroke();
                        Attcb.EndText();
                        Attdocument.Close();
                        pdfmp.AddFile(_AttFileName + SheetNo + ".pdf");
                        SheetNo++;
                    }
                }
            }

            pdfmp.Execute();
            string str = FileHidpath;
            string mime = MimeMapping.GetMimeMapping(FileHidpath);
            return File(FileHidpath, mime);

        }
        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return RegMng.GetViewData(_Query, "");
        }
        public DataTable GetFreghtManifestMainDtls(string VesVoyID, string BLID)
        {
            //string _Query = " select NVO_BLRelease.ID, NVO_BOL.ID AS BLID,BLNo,Agent, AgentAddress, NVO_Ratesheet.ID  as RRID,  (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID " +
            // " order by NVO_VoyageRoute.RID desc) as VoyageNo,(Select top 1 (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) from NVO_Voyage V where V.ID = NVO_BOL.BLVesVoyID) as VesVoy, (select top 1 Convert(varchar, ETA, 106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID)as ETA, " +
            // " (select top 1 Convert(varchar, ETD, 106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID)as ETD,  NVO_BLRelease.POD,NVO_BLRelease.POL,NVO_BOL.BLVesVoyID, " +
            // " (select Top(1) TerminalName from NVO_TerminalMaster inner join NVO_VoyageRoute On NVO_VoyageRoute.TerminalID = NVO_TerminalMaster.ID " +
            // " where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) as LoadTerminal  from NVO_BLRelease inner join NVO_BOL On NVO_BOL.ID = NVO_BLRelease.BLID  inner join NVO_Booking On NVO_Booking.ID = NVO_BOL.BkgID left outer join NVO_Ratesheet On NVO_Ratesheet.ID = NVO_Booking.RRID " +
            // " where NVO_BOL.BLVesVoyID = " + VesVoyID;
            //if (BLID != "")
            //    _Query += " and NVO_BOL.ID in (" + BLID + ")";


            string _Query = " select NVO_BLRelease.ID, NVO_BOL.ID AS BLID,BLNo,Agent, AgentAddress, NVO_Ratesheet.ID as RRID,  " +
                           " (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = Nvo_View_BLVoyageLegDetails.VesVoyID " +
                           " order by NVO_VoyageRoute.RID desc) as VoyageNo, " +
                           " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) " +
                           " from NVO_Voyage V where V.ID = Nvo_View_BLVoyageLegDetails.VesVoyID) as VesVoy,  " +
                           " (select top 1 Convert(varchar, ETA, 106) from NVO_VoyageRoute " +
                           " where NVO_VoyageRoute.VoyageID = Nvo_View_BLVoyageLegDetails.VesVoyID)as ETA,   " +
                           " (select top 1 Convert(varchar, ETD, 106) from NVO_VoyageRoute " +
                           " where NVO_VoyageRoute.VoyageID = Nvo_View_BLVoyageLegDetails.VesVoyID)as ETD,  " +
                           " NVO_BLRelease.POD,NVO_BLRelease.POL, " +
                           " Nvo_View_BLVoyageLegDetails.VesVoyID,  " +
                           " (select Top(1) TerminalName from NVO_TerminalMaster " +
                           " inner join NVO_VoyageRoute On NVO_VoyageRoute.TerminalID = NVO_TerminalMaster.ID " +
                           " where NVO_VoyageRoute.VoyageID = Nvo_View_BLVoyageLegDetails.VesVoyID) as LoadTerminal " +
                           " from NVO_BLRelease " +
                           " inner join NVO_BOL On NVO_BOL.ID = NVO_BLRelease.BLID " +
                           " inner join NVO_Booking On NVO_Booking.ID = NVO_BOL.BkgID " +
                           " inner join NVO_Ratesheet On NVO_Ratesheet.ID = NVO_Booking.RRID " +
                           " inner join Nvo_View_BLVoyageLegDetails  on Nvo_View_BLVoyageLegDetails.BLID = NVO_BOL.ID " +
                           " where Nvo_View_BLVoyageLegDetails.VesVoyID = " + VesVoyID;
            if (BLID != "")
                _Query += " and NVO_BOL.ID in (" + BLID + ")";
            return RegMng.GetViewData(_Query, "");



            return RegMng.GetViewData(_Query, "");
        }
        public DataTable GetFrieghtManifestDtls(string BLID)
        {

            string _Query = "select distinct NVO_Ratesheet.Id as RRID,NVO_BLRelease.ID,NVO_BLRelease.POD,NVO_BLRelease.POL,NVO_BLRelease.BLID,BLNo,BLVesVoyID, NVO_BLRelease.Shipper,ShipperAddress,Consignee,ConsigneeAddress,NVO_BLRelease.IntFreeDays, " +
                            " Notify1,Notify1Address,NVO_BLRelease.Marks as MarkNo,NoOfPkg,NVO_BLRelease.GRWT AS GrsWt,NVO_BLRelease.Description as CagoDescription,  NVO_BLRelease.CBM,NVO_Containers.CntrNo,cntrdetails,(select top(1) Notes from NVO_BLNotesClauses where DocID= 264 and NID=NVO_BLRelease.FreeDays) as ddlFreeday,NVO_BLRelease.NTWT As NetWeight, " +
                            " (NVO_Containers.CntrNo + '/' + Size + '/' + SealNo) as Seal,FreightPayment, " +
                            " (select top(1) case when GrsWtType = 1 then 'KGS' else case when GrsWtType = 2 then 'MTS' end end from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID = NVO_BLRelease.BLID)  as GrsWtType,  " +
                           " (select top(1) case when NtWtType = 1 then 'KGS' else case when NtWtType = 2 then 'MTS' end end  from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID = NVO_BLRelease.BLID) as NtWtType, " +
                           " (select top(1)(select top(1) PkgDescription from NVO_CargoPkgMaster where NVO_CargoPkgMaster.Id = PakgType) from NVO_BOLCntrDetails where NVO_BOLCntrDetails.BLID = NVO_BLRelease.BLID) as CargoPakage " +
                            " from NVO_BLRelease " +
                            " inner join NVO_BOL On NVO_BOL.ID = NVO_BLRelease.BLID " +
                            " inner join NVO_BOLCntrDetails On NVO_BOLCntrDetails.BLID = NVO_BLRelease.BLID " +
                            " inner join NVO_Containers On NVO_Containers.ID = NVO_BOLCntrDetails.CntrID " +
                            " inner join NVO_Booking On NVO_Booking.ID = NVO_BOL.BkgID " +
                            " left outer join NVO_Ratesheet On NVO_Ratesheet.ID = NVO_Booking.RRID " +
                            " left outer join NVO_RatesheetCntrTypes on NVO_RatesheetCntrTypes.RRID = NVO_Ratesheet.ID " +
                            " where NVO_BLRelease.BLID in (" + BLID + ")";
            return RegMng.GetViewData(_Query, "");
        }

        public DataTable GetCntrChargesBreakup(string RID, string ChgcodeID, string CntrType, string BLID)
        {

            string _Query = "select * from NVO_V_FreightManifest_New where RRID=" + RID + " AND ChargeCodeID=" + ChgcodeID + " and CntrType =" + CntrType + " and BLID=" + BLID;
            return RegMng.GetViewData(_Query, "");
        }
        public DataTable GetCntrChargesBreakupNew(string RID, string BLID)
        {

            string _Query = "select  DISTINCT ChgCode,ChargeCodeID from NVO_V_FreightManifest_New where RRID=" + RID + " and BLID=" + BLID;
            return RegMng.GetViewData(_Query, "");
        }

        public DataTable GetCntrdisplay(string BLID)
        {

            string _Query = "select (select top(1) CntrNo from NVO_Containers where ID=CntrID) as CntrNo," +
                 " (select top(1) TypeID from NVO_Containers where ID = CntrID) as TypeID," +
                " SealNo,BkgId,Size from NVO_BOLCntrDetails  where BLID= " + BLID;
            return RegMng.GetViewData(_Query, "");
        }
        public DataTable GetCompanyDetails()
        {

            string _Query = "Select * from NVO_NewCompnayDetails";
            return RegMng.GetViewData(_Query, "");
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