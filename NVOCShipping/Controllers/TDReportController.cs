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

    public class TDReportController : Controller
    {
        ExportReportManager RegMng = new ExportReportManager();
        // GET: TDReport
        public ActionResult Index()
        {
            return View();
        }

        //public ActionResult ExportTerminalPDF(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID, string AgencyID)
        //{
        //    BindTerminalPDF(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID, AgencyID);
        //    return View();
        //}

        public FileResult ExportTerminalPDFOLD(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID, string AgencyID)
        {
            string FileHidpath = "";
            string str = FileHidpath;
            string mime = "";
            MergeEx pdfmp = new MergeEx();
            Document doc = new Document();
            Rectangle rec = new Rectangle(800, 1150);
            doc = new Document(rec);
            doc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            Paragraph para = new Paragraph();


            string pdfpath = Server.MapPath("./pdfpath/");

            pdfmp.SourceFolder = pdfpath;

            pdfmp.DestinationFile = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";
            FileHidpath = pdfpath + "Multiple-" + Session.SessionID.ToString() + "BL.pdf";

            string _FileName = Session.SessionID.ToString() + VesVoyID + 1;
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(pdfpath + _FileName + ".pdf", FileMode.Create));
            doc.Open();

            PdfContentByte cb = writer.DirectContent;
            cb.SetColorStroke(Color.BLACK);
            int _Xp = 10, _Yp = 785, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);
            DataTable dtc = GetAgencyDetails(AgencyID);
            if (dtc.Rows.Count > 0)
            {
                if (dtc.Rows[0]["LogoPath"].ToString() != "")
                {

                    if (AgencyID == "22" || AgencyID == "60" || AgencyID == "61" || AgencyID == "69" || AgencyID == "84" || AgencyID == "8" || AgencyID == "11" || AgencyID == "19" || AgencyID == "20" || AgencyID == "56" || AgencyID == "10" || AgencyID == "62")
                    {
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        png1.SetAbsolutePosition(15, 515);
                        png1.ScalePercent(7f);
                        doc.Add(png1);
                    }
                    if (AgencyID == "3" || AgencyID == "9" || AgencyID == "19" || AgencyID == "20")
                    {
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        png1.SetAbsolutePosition(15, 515);
                        png1.ScalePercent(30f);
                        doc.Add(png1);
                    }
                    if (AgencyID == "1" || AgencyID == "13" || AgencyID == "14" || AgencyID == "15" || AgencyID == "16")
                    {
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        png1.SetAbsolutePosition(15, 515);
                        png1.ScalePercent(7f);
                        doc.Add(png1);
                    }
            
                    if (AgencyID == "2")
                    {
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                        png1.SetAbsolutePosition(15, 515);
                        png1.ScalePercent(30f);
                        doc.Add(png1);
                    }




                }
                else
                {
                    iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/blanklogo.png"));
                    png1.SetAbsolutePosition(15, 515);
                    png1.ScalePercent(7f);
                    doc.Add(png1);
                }

            }
            cb.BeginText();
            BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader21, 15);
            cb.SetColorFill(new Color(0, 0, 128));

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 480, 560, 0);




            int RowIndexAdd = 540;
            BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader22, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            var splitadd = dtc.Rows[0]["Address"].ToString().Split('\n');
            for (int k = 0; k < splitadd.Length; k++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitadd[k].ToString(), 480, RowIndexAdd, 0);
                RowIndexAdd -= 10;
            }
            RowIndexAdd -= 15;
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 480, 540, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 480, 530, 0);
            cb.SetColorStroke(new Color(0, 0, 128));
            cb.EndText();
            cb.BeginText();
            //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/companyborder.png"));
            //png2.SetAbsolutePosition(480, 550);
            //png2.ScalePercent(50f);
            //doc.Add(png2);
            cb.EndText();
            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            cb.SetFontAndSize(bfheader2, 12);
            cb.SetColorFill(new Color(0, 0, 128));

            cb.BeginText();
            ///----Header-----///
            ///
            ///-TOP LINE-///
            cb.MoveTo(10, 500);
            cb.LineTo(830, 500);

            ///-BOTTOM-///
            cb.MoveTo(10, 460);
            cb.LineTo(830, 460);

            ///-LEFT-///
            cb.MoveTo(10, 500);
            cb.LineTo(10, 460);

            ///-Right-///
            cb.MoveTo(830, 500);
            cb.LineTo(830, 460);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMINAL DEPARTURE REPORT", 330, 480, 0);
            cb.SetRGBColorFill(0x00, 0x00, 0xFF);
            cb.EndText();

            ////---------TDR DETAILS------///
            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 9);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.BeginText();

            ///-------Lines----/////
            cb.MoveTo(10, 460);
            cb.LineTo(830, 460);

            cb.MoveTo(10, 435);
            cb.LineTo(830, 435);

            cb.MoveTo(10, 410);
            cb.LineTo(830, 410);

            cb.MoveTo(10, 385);
            cb.LineTo(830, 385);

            cb.MoveTo(10, 360);
            cb.LineTo(830, 360);

            cb.MoveTo(10, 335);
            cb.LineTo(830, 335);



            ///-LEFT-///
            cb.MoveTo(10, 460);
            cb.LineTo(10, 335);

            ///-Right-///
            cb.MoveTo(830, 460);
            cb.LineTo(830, 335);


            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING", 15, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NAME OF VESSEL", 15, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA POL", 15, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEXT PORT", 15, 365, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER OPERATOR", 15, 340, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 365, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 340, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMINAL NAME", 450, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOYAGE NUMBER", 450, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD POL", 450, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA NEXT PORT", 450, 365, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 365, 0);

            DataTable _dtv = GetTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);
            if (_dtv.Rows.Count > 0)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["POL"].ToString(), 130, 440, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VesselName"].ToString(), 130, 415, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["ETA"].ToString(), 130, 390, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["NextPort"].ToString(), 130, 365, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VesselOperator"].ToString(), 130, 340, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["TerminalName"].ToString(), 570, 440, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VoyageNo"].ToString(), 570, 415, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["ETD"].ToString(), 570, 390, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["NextPortETA"].ToString(), 570, 365, 0);
            }

            cb.EndText();

            //---Table---//
            cb.BeginText();

            ///-TOP-///
            cb.MoveTo(10, 330);
            cb.LineTo(830, 330);

            cb.MoveTo(10, 310);
            cb.LineTo(830, 310);

            ///-LEFT-///
            cb.MoveTo(10, 330);
            cb.LineTo(10, 30);

            ///-Right-///
            cb.MoveTo(830, 330);
            cb.LineTo(830, 30);

            ///-BOTTOM-///
            cb.MoveTo(10, 30);
            cb.LineTo(830, 30);

            //horizontallines//--

            cb.MoveTo(55, 330);
            cb.LineTo(55, 30);

            cb.MoveTo(73, 330);
            cb.LineTo(73, 30);

            cb.MoveTo(112, 330);
            cb.LineTo(112, 30);



            //cb.MoveTo(77, 330);
            //cb.LineTo(77, 30);

            cb.MoveTo(144, 330);
            cb.LineTo(144, 30);

            cb.MoveTo(230, 330);
            cb.LineTo(230, 30);

            cb.MoveTo(320, 330);
            cb.LineTo(320, 30);

            cb.MoveTo(410, 330);
            cb.LineTo(410, 30);

            cb.MoveTo(450, 330);
            cb.LineTo(450, 30);

            cb.MoveTo(480, 330);
            cb.LineTo(480, 30);

            cb.MoveTo(550, 330);
            cb.LineTo(550, 30);

            cb.MoveTo(690, 330);
            cb.LineTo(690, 30);

            cb.MoveTo(830, 330);
            cb.LineTo(830, 30);

            BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader4, 6);
            cb.SetColorFill(new Color(0, 0, 128));
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "S.NO", 12, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER#", 12, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SIZE", 57, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMMODITY", 74, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SERVICE", 115, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TS PORT", 145, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 232, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FINAL DEST", 322, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PICKUP DT", 412, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GW.KGS", 452, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL NUMBER", 482, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VSL OPERATOR", 552, 320, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD AGENT", 692, 320, 0);

            int RowGrd = 300;
            DataTable _dtC = GetCntrDtlsTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);

            //for (int k = 0; k < _dtC.Rows.Count; k++)
            //{





            //    RowGrd -= 10;

            //}
            RowGrd = 300;
            int cntrrowount = 13;
            for (int k = 0; k < _dtC.Rows.Count; k++)
            {
                string[] ArrayOperator = Regex.Split(_dtC.Rows[k]["Operator"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                string[] AaddsplitOP;

                for (int x = 0; x < ArrayOperator.Length; x++)
                {
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 12, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["CntrNo"].ToString(), 12, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["Size"].ToString(), 57, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["commodity"].ToString(), 74, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["ServiceType"].ToString(), 115, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["TSPORT"].ToString(), 145, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["POD"].ToString(), 232, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["FPOD"].ToString(), 322, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["PickUpDate"].ToString(), 412, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["GrsWt"].ToString(), 452, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["BLNumber"].ToString(), 482, RowGrd, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[k]["DestinationAgent"].ToString(), 692, RowGrd, 0);
                    AaddsplitOP = SplitByLenght(ArrayOperator[x].ToString(), 35);

                    for (int j = 0; j < AaddsplitOP.Length; j++)
                    {
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, AaddsplitOP[j].ToString(), 552, RowGrd, 0);
                        RowGrd -= 10;

                    }
                    

                }
                if (cntrrowount == k)
                {
                    break;
                }
            }




            cb.EndText();
            cb.Stroke();
            doc.Close();
            pdfmp.AddFile(_FileName + ".pdf");

            int AttachedsheetNo = 1;
            if (cntrrowount >= 13)
            {


                for (int k = 0; k < AttachedsheetNo; k++)
                {
                    int LineCount = 1;
                    int SheetNo = 1;
                    string Filesv = "Attach" + VesVoyID;
                    string _AttFileName = Filesv;
                    int LIndex = 15;
                    int LMarkindex = 8;
                    int LCntrindex = 4;

                    Rectangle attrec = new Rectangle(800, 1150);


                    Document Attdocument = new Document(attrec);
                    Attdocument.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
                    PdfWriter Attwriter = PdfWriter.GetInstance(Attdocument, new FileStream(pdfpath + (_AttFileName + SheetNo) + ".pdf", FileMode.Create));
                    Attdocument.Open();
                    PdfContentByte Attcb = Attwriter.DirectContent;


                    BaseFont bfheader1 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    Attcb.SetFontAndSize(bfheader, 14);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 830, 0);




                    Attcb.BeginText();

                    BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    Attcb.SetFontAndSize(bfheader5, 14);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);

                    if (dtc.Rows.Count > 0)
                    {
                        if (dtc.Rows[0]["LogoPath"].ToString() != "")
                        {

                            if (AgencyID == "22" || AgencyID == "60" || AgencyID == "61" || AgencyID == "69" || AgencyID == "84" || AgencyID == "8" || AgencyID == "11" || AgencyID == "19" || AgencyID == "20" || AgencyID == "56" || AgencyID == "10" || AgencyID == "62")
                            {
                                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                                png1.SetAbsolutePosition(15, 515);
                                png1.ScalePercent(7f);
                                Attdocument.Add(png1);
                            }
                            if (AgencyID == "3" || AgencyID == "9" || AgencyID == "19" || AgencyID == "20")
                            {
                                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                                png1.SetAbsolutePosition(15, 515);
                                png1.ScalePercent(30f);
                                Attdocument.Add(png1);
                            }
                            if (AgencyID == "1" || AgencyID == "13" || AgencyID == "14" || AgencyID == "15" || AgencyID == "16")
                            {
                                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                                png1.SetAbsolutePosition(15, 515);
                                png1.ScalePercent(7f);
                                Attdocument.Add(png1);
                            }

                            if (AgencyID == "2")
                            {
                                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/" + dtc.Rows[0]["LogoPath"].ToString()));
                                png1.SetAbsolutePosition(15, 515);
                                png1.ScalePercent(30f);
                                Attdocument.Add(png1);
                            }




                        }
                        else
                        {
                            iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/blanklogo.png"));
                            png1.SetAbsolutePosition(15, 515);
                            png1.ScalePercent(7f);
                            Attdocument.Add(png1);
                        }

                    }


                    BaseFont bfheader24 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    Attcb.SetFontAndSize(bfheader24, 15);
                    Attcb.SetColorFill(new Color(0, 0, 128));

                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 480, 560, 0);




                    int RowIndexAdd1 = 540;
                    BaseFont bfheader25 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    Attcb.SetFontAndSize(bfheader25, 8);
                    Attcb.SetColorFill(new Color(0, 0, 128));
                    var splitadd1 = dtc.Rows[0]["Address"].ToString().Split('\n');
                    for (int i = 0; i < splitadd1.Length; i++)
                    {
                        Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitadd1[i].ToString(), 480, RowIndexAdd1, 0);
                        RowIndexAdd1 -= 10;
                    }
                    RowIndexAdd1 -= 15;



                    ///-TOP-///
                    Attcb.MoveTo(10, 460);
                    Attcb.LineTo(830, 460);

                    Attcb.MoveTo(10, 440);
                    Attcb.LineTo(830, 440);

                    ///-LEFT-///
                    Attcb.MoveTo(10, 460);
                    Attcb.LineTo(10, 30);

                    ///-Right-///
                    Attcb.MoveTo(830, 460);
                    Attcb.LineTo(830, 30);

                    ///-BOTTOM-///
                    Attcb.MoveTo(10, 30);
                    Attcb.LineTo(830, 30);

                    //horizontallines//--

                    Attcb.MoveTo(55, 460);
                    Attcb.LineTo(55, 30);

                    Attcb.MoveTo(73, 460);
                    Attcb.LineTo(73, 30);

                    Attcb.MoveTo(112, 460);
                    Attcb.LineTo(112, 30);



                    //cb.MoveTo(77, 330);
                    //cb.LineTo(77, 30);

                    Attcb.MoveTo(144, 460);
                    Attcb.LineTo(144, 30);

                    Attcb.MoveTo(230, 460);
                    Attcb.LineTo(230, 30);

                    Attcb.MoveTo(320, 460);
                    Attcb.LineTo(320, 30);

                    Attcb.MoveTo(410, 460);
                    Attcb.LineTo(410, 30);

                    Attcb.MoveTo(450, 460);
                    Attcb.LineTo(450, 30);

                    Attcb.MoveTo(480, 460);
                    Attcb.LineTo(480, 30);

                    Attcb.MoveTo(550, 460);
                    Attcb.LineTo(550, 30);

                    Attcb.MoveTo(690, 460);
                    Attcb.LineTo(690, 30);

                    Attcb.MoveTo(830, 460);
                    Attcb.LineTo(830, 30);

                    BaseFont bfheader26 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    Attcb.SetFontAndSize(bfheader26, 6);
                    Attcb.SetColorFill(new Color(0, 0, 128));
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "S.NO", 12, 320, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER#", 12, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SIZE", 57, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "COMMODITY", 74, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SERVICE", 115, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TS PORT", 145, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD", 232, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FINAL DEST", 322, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PICKUP DT", 412, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "GW.KGS", 452, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BL NUMBER", 482, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VSL OPERATOR", 552, 450, 0);
                    Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "POD AGENT", 692, 450, 0);

                    int RowGrd1 = 300;
                    // DataTable _dtC = GetCntrDtlsTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);

                    //for (int k = 0; k < _dtC.Rows.Count; k++)
                    //{





                    //    RowGrd -= 10;

                    //}
                    RowGrd1 = 430;
                    for (int r = cntrrowount+1; r < _dtC.Rows.Count; r++)
                    {
                        string[] ArrayOperator1 = Regex.Split(_dtC.Rows[r]["Operator"].ToString().ToUpper().Trim(), char.ConvertFromUtf32(13));
                        string[] AaddsplitOP1;

                        for (int x = 0; x < ArrayOperator1.Length; x++)
                        {
                            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 12, RowGrd, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["CntrNo"].ToString(), 12, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["Size"].ToString(), 57, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["commodity"].ToString(), 74, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["ServiceType"].ToString(), 115, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["TSPORT"].ToString(), 145, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["POD"].ToString(), 232, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["FPOD"].ToString(), 322, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["PickUpDate"].ToString(), 412, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["GrsWt"].ToString(), 452, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["BLNumber"].ToString(), 482, RowGrd1, 0);
                            Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtC.Rows[r]["DestinationAgent"].ToString(), 692, RowGrd1, 0);
                            AaddsplitOP1 = SplitByLenght(ArrayOperator1[x].ToString(), 35);

                            for (int j = 0; j < AaddsplitOP1.Length; j++)
                            {
                                Attcb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, AaddsplitOP1[j].ToString(), 552, RowGrd1, 0);
                                RowGrd1 -= 10;

                            }

                        }
                    }




                    Attcb.Stroke();
                    Attcb.EndText();
                    Attdocument.Close();
                    pdfmp.AddFile(_AttFileName + SheetNo + ".pdf");
                    SheetNo++;


                }
            }
            pdfmp.Execute();
            mime = MimeMapping.GetMimeMapping(FileHidpath);
            return File(FileHidpath, mime);
            //writer.CloseStream = false;
            //doc.Close();


            //Response.Buffer = true;
            //Response.ContentType = "application/pdf";
            ////Response.AddHeader("content-disposition", "attachment;filename=TerminalDepartureReport.pdf");
            //Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //// Response.Write(doc);
            //Response.End();
        }


        public void ExportTerminalPDF(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID, string AgencyID)
        {
            

            Document doc = new Document();
            MergeEx pdfmp = new MergeEx();
            Rectangle rec = new Rectangle(800, 1150);
            doc = new Document(rec);
            doc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            Paragraph para = new Paragraph();


            //PdfWriter pdfWriter1 = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/BKGPDF\\" + dtv.Rows[0]["BookingNo"].ToString() + ".pdf"), FileMode.Create));
            PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
            doc.Open();

            PdfContentByte cb = pdfWriter.DirectContent;
            cb.SetColorStroke(Color.BLACK);
            int _Xp = 10, _Yp = 785, YDiff = 10;

            BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader, 14);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 15, 830, 0);
            DataTable dtc = GetAgencyDetails(AgencyID);
            if (dtc.Rows.Count > 0)
            {
                if (dtc.Rows[0]["LogoPath"].ToString() != "")
                {
                    if (dtc.Rows[0]["ID"].ToString() == "1" || dtc.Rows[0]["ID"].ToString() == "13" || dtc.Rows[0]["ID"].ToString() == "14" || dtc.Rows[0]["ID"].ToString() == "15" || dtc.Rows[0]["ID"].ToString() == "16")
                    {
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.png"));
                        png1.SetAbsolutePosition(15, 515);
                        png1.ScalePercent(10f);
                        doc.Add(png1);

                    }

                    if (dtc.Rows[0]["ID"].ToString() == "3")
                    {
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.png"));
                        png1.SetAbsolutePosition(15, 515);
                        png1.ScalePercent(10f);
                        doc.Add(png1);
                    }

                }
                else
                {
                    iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/agentlogo/BWSLOGO.png"));
                    png1.SetAbsolutePosition(15, 515);
                    png1.ScalePercent(10f);
                    doc.Add(png1);
                }

            }
            cb.BeginText();
            BaseFont bfheader21 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader21, 15);
            cb.SetColorFill(new Color(0, 0, 128));

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 480, 560, 0);




            int RowIndexAdd = 540;
            BaseFont bfheader22 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader22, 8);
            cb.SetColorFill(new Color(0, 0, 128));
            var splitadd = dtc.Rows[0]["Address"].ToString().Split('\n');
            for (int k = 0; k < splitadd.Length; k++)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitadd[k].ToString(), 480, RowIndexAdd, 0);
                RowIndexAdd -= 10;
            }

            RowIndexAdd -= 15;
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING FOR COMBINED TRANSPORT SHIPMENT OR", 480, 540, 0);
            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT TO PORT SHIPMENT NOT NEGOTIABLE UNLESS CONSIGNED 'TO ORDER''", 480, 530, 0);
            cb.SetColorStroke(new Color(0, 0, 128));
            cb.EndText();
            cb.BeginText();
            //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/companyborder.png"));
            //png2.SetAbsolutePosition(480, 550);
            //png2.ScalePercent(50f);
            //doc.Add(png2);
            cb.EndText();
            BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            cb.SetFontAndSize(bfheader2, 12);
            cb.SetColorFill(new Color(0, 0, 128));

            cb.BeginText();
            ///----Header-----///
            ///
            ///-TOP LINE-///
            cb.MoveTo(10, 507);
            cb.LineTo(830, 507);

            ///-BOTTOM-///
            cb.MoveTo(10, 480);
            cb.LineTo(830, 480);

            ///-LEFT-///
            cb.MoveTo(10, 507);
            cb.LineTo(10, 480);

            ///-Right-///
            cb.MoveTo(830, 507);
            cb.LineTo(830, 480);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMINAL DEPARTURE REPORT", 330, 490, 0);
            cb.SetRGBColorFill(0x00, 0x00, 0xFF);
            cb.EndText();

            ////---------TDR DETAILS------///
            BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetFontAndSize(bfheader3, 9);
            cb.SetColorFill(new Color(0, 0, 128));
            cb.BeginText();



            ///-------Lines----/////
            cb.MoveTo(10, 460);
            cb.LineTo(830, 460);

            cb.MoveTo(10, 435);
            cb.LineTo(830, 435);

            cb.MoveTo(10, 410);
            cb.LineTo(830, 410);

            cb.MoveTo(10, 385);
            cb.LineTo(830, 385);



            cb.MoveTo(10, 335);
            cb.LineTo(830, 335);



            ///-LEFT-///
            cb.MoveTo(10, 460);
            cb.LineTo(10, 335);

            ///-Right-///
            cb.MoveTo(830, 460);
            cb.LineTo(830, 335);



            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING", 15, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NAME OF VESSEL", 15, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA POL", 15, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEXT PORT", 15, 365, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONTAINER OPERATOR", 15, 340, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 365, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 340, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TERMINAL NAME", 450, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VOYAGE NUMBER", 450, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETD POL", 450, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA NEXT PORT", 450, 365, 0);

            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 440, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 415, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 390, 0);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 560, 365, 0);

            DataTable _dtv = GetTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);
            if (_dtv.Rows.Count > 0)
            {
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["POL"].ToString(), 130, 440, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VesselName"].ToString(), 130, 415, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["ETA"].ToString(), 130, 390, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["NextPort"].ToString(), 130, 365, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VesselOperator"].ToString(), 130, 340, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["TerminalName"].ToString(), 570, 440, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["VoyageNo"].ToString(), 570, 415, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["ETD"].ToString(), 570, 390, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, _dtv.Rows[0]["NextPortETA"].ToString(), 570, 365, 0);
            }
            cb.EndText();
            cb.Stroke();

            #region Container Details
            //----------------- Booking Details--------------//

            iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
            // Tbl3.Spacing = 300;
            Tbl3.Width = 100;
            Tbl3.Alignment = Element.ALIGN_LEFT;
            Tbl3.Cellpadding = 80;

            Tbl3.BorderWidth = 0;
            //Tbl3.Spacing =300f;
            Cell cell = new Cell();
            cell.Width = 10;

            //Sub Heading
            // cell = new Cell(new Phrase("Container Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

            cell.BorderWidth = 0;
            cell.Colspan = 1;
            Tbl3.AddCell(cell);
            doc.Add(Tbl3);


            iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
            Tbl5.Width = 106;
            // Tbl5.Alignment = Element.ALIGN_LEFT;
            Color navyBlue = new Color(0, 0, 128);
            Tbl5.Cellpadding = 1;
            Tbl5.BorderWidthTop = 1;
            Tbl5.BorderColor = navyBlue;

            //Tbl5.BorderWidthBottom = 1;

            DataTable _dtC = GetCntrDtlsTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);

            for (int i = 0; i < _dtC.Rows.Count; i++)
            {
                //Caption
                cell = new Cell(new Phrase("  " + "CONTAINER NO ", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["CntrNo"].ToString() + " / " + _dtC.Rows[i]["Size"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("COMMODITY", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Value
                cell = new Cell(new Phrase(" :  " + _dtC.Rows[i]["commodity"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("SERVICE / SLOT TERM", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value

                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["ServiceType"].ToString() + " / " + _dtC.Rows[i]["SlotTerm"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                ////Caption
                //cell = new Cell(new Phrase("SLOT TERM", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 2;
                //Tbl5.AddCell(cell);
                ////Value

                //cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["SlotTerm"].ToString(), new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                //cell.BorderWidth = 0;
                //cell.Colspan = 2;
                //Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("  " + "POD", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["POD"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("FINAL DEST", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value

                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["FPOD"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("TSPORT", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["TSPORT"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("  " + "PICKUP DT", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["PickUpDate"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                //Caption
                cell = new Cell(new Phrase("GW.KGS", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["GrsWt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);

                //Caption
                cell = new Cell(new Phrase("BL NUMBER", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["BLNumber"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                Tbl5.AddCell(cell);


                cell = new Cell(new Phrase("  " + "OPERATOR", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["Operator"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase(" " + "POD AGENT", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 1;
                Tbl5.AddCell(cell);
                //Value
                cell = new Cell(new Phrase(" : " + _dtC.Rows[i]["DestinationAgent"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                Tbl5.AddCell(cell);

                cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 8, Font.BOLD, navyBlue)));
                cell.BorderWidth = 0;
                cell.Colspan = 12;
                cell.BorderColor = navyBlue;
                cell.BorderWidthBottom = 1;
                Tbl5.AddCell(cell);


            }


            doc.Add(Tbl5);


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

        public DataTable GetTerminalDepPdfValues(string VesVoy, string POD, string TsPort, string Agent, string VesselOperator)
        {
            string strWhere = "";


            string _Query = "Select Top(1)  NVO_Booking.ID,NVO_BOL.ID as BLID, NVO_Booking.VesVoyID, " +
                            " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) from NVO_Voyage V where V.ID = NVO_Booking.VesVoyID)as VesselName, " +
                            " (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as VoyageNo, " +
                            " NVO_Booking.POL AS BkgPOL,(select  top(1) Convert(varchar,ETA,106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID)as ETA,(select   top(1) Convert(varchar,ETD,106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID)as ETD, " +
                            " (select top(1) PortName from NVO_VoyageRoute " +
                            " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as NextPort, " +
                             " (select top(1) PortName from NVO_VoyageRoute " +
                            " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID asc)as POL, " +
                              " (select top(1) TerminalName from NVO_VoyageRoute inner join NVO_TerminalMaster On NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID asc) as TerminalName, " +
                            " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc),106) as NextPortETA, " +
                            " ( select top(1) CompanyName from NVO_NewCompnayDetails) as VesselOperator, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster " +
                            " where NVO_AgencyMaster.ID = NVO_Booking.DestinationAgentID) as PODAgent from NVO_Booking " +
                            " inner join NVO_Voyage On NVO_Voyage.ID = NVO_Booking.VesVoyID" +
                            " inner join NVO_BOL On NVO_BOL.BkgID = NVO_Booking.ID ";

            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.VesVoyID=" + VesVoy;
                else

                    strWhere += " AND NVO_Booking.VesVoyID=" + VesVoy;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.PODID=" + POD;
                else
                    strWhere += " AND NVO_Booking.PODID=" + POD;

            if (TsPort != "" && TsPort != "0" && TsPort != "null" && TsPort != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.TSPORTID=" + TsPort;
                else
                    strWhere += " AND NVO_Booking.TSPORTID=" + TsPort;

            if (Agent != "" && Agent != "0" && Agent != "null" && Agent != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.DestinationAgentID=" + Agent;
                else
                    strWhere += " AND NVO_Booking.DestinationAgentID=" + Agent;

            if (VesselOperator != "" && VesselOperator != "0" && VesselOperator != "null" && VesselOperator != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.SlotOperatorID=" + VesselOperator;
                else
                    strWhere += " AND NVO_Booking.SlotOperatorID=" + VesselOperator;

            if (strWhere == "")
                strWhere = _Query;

            strWhere += " UNION  " +
                             " Select Top(1)  NVO_Booking.ID,NVO_BOL.ID as BLID,NVO_Booking.VesVoyID, " +
                            " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) from NVO_Voyage V where V.ID = NVO_Booking.VesVoyID)as VesselName, " +
                            " (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as VoyageNo, " +
                            " NVO_Booking.POL AS BkgPOL,(select  top(1) Convert(varchar,ETA,106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID)as ETA,(select   top(1) Convert(varchar,ETD,106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID)as ETD, " +
                            " (select top(1) PortName from NVO_VoyageRoute " +
                            " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc)as NextPort, " +
                             " (select top(1) PortName from NVO_VoyageRoute " +
                            " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID asc)as POL, " +
                              " (select top(1) TerminalName from NVO_VoyageRoute inner join NVO_TerminalMaster On NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID asc) as TerminalName, " +
                            " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID desc),106) as NextPortETA, " +
                            " ( select top(1) CompanyName from NVO_NewCompnayDetails) as VesselOperator, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster " +
                            " where NVO_AgencyMaster.ID = NVO_Booking.DestinationAgentID) as PODAgent from NVO_Booking " +
                            " inner join NVO_Voyage On NVO_Voyage.ID = NVO_Booking.VesVoyID" +
                            " inner join NVO_BOL On NVO_BOL.BkgID = NVO_Booking.ID inner join NVO_Gateway_VoyageBL GT ON GT.BkgID = NVO_Booking.ID ";

            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.VesVoyID=" + VesVoy;
                else

                    strWhere += " AND NVO_Booking.VesVoyID=" + VesVoy;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.PODID=" + POD;
                else
                    strWhere += " AND NVO_Booking.PODID=" + POD;

            if (TsPort != "" && TsPort != "0" && TsPort != "null" && TsPort != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.TSPORTID=" + TsPort;
                else
                    strWhere += " AND NVO_Booking.TSPORTID=" + TsPort;

            if (Agent != "" && Agent != "0" && Agent != "null" && Agent != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.DestinationAgentID=" + Agent;
                else
                    strWhere += " AND NVO_Booking.DestinationAgentID=" + Agent;

            if (VesselOperator != "" && VesselOperator != "0" && VesselOperator != "null" && VesselOperator != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.SlotOperatorID=" + VesselOperator;
                else
                    strWhere += " AND NVO_Booking.SlotOperatorID=" + VesselOperator;

            if (strWhere == "")
                strWhere = _Query;



            strWhere += " UNION  " +
                           " Select Top(1)  NVO_Booking.ID,NVO_BOL.ID as BLID,VL.VesVoyID, " +
                          " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) from NVO_Voyage V where V.ID = VL.VesVoyID)as VesselName, " +
                          " (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VL.VesVoyID order by NVO_VoyageRoute.RID desc)as VoyageNo, " +
                          " NVO_Booking.POL AS BkgPOL,(select  top(1) Convert(varchar,ETA,106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VL.VesVoyID)as ETA,(select   top(1) Convert(varchar,ETD,106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VL.VesVoyID)as ETD, " +
                          " (select top(1) PortName from NVO_VoyageRoute " +
                          " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                          " where NVO_VoyageRoute.VoyageID = VL.VesVoyID order by NVO_VoyageRoute.RID desc)as NextPort, " +
                           " (select top(1) PortName from NVO_VoyageRoute " +
                          " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                          " where NVO_VoyageRoute.VoyageID = VL.VesVoyID order by NVO_VoyageRoute.RID asc)as POL, " +
                            " (select top(1) TerminalName from NVO_VoyageRoute inner join NVO_TerminalMaster On NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID where NVO_VoyageRoute.VoyageID = VL.VesVoyID order by NVO_VoyageRoute.RID asc) as TerminalName, " +
                          " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VL.VesVoyID order by NVO_VoyageRoute.RID desc),106) as NextPortETA, " +
                          " ( select top(1) CompanyName from NVO_NewCompnayDetails) as VesselOperator, " +
                          " (select top(1) AgencyName from NVO_AgencyMaster " +
                          " where NVO_AgencyMaster.ID = NVO_Booking.DestinationAgentID) as PODAgent from NVO_Booking " +
                          " inner join NVO_BOL On NVO_BOL.BkgID = NVO_Booking.ID inner join NVO_VoyageAllocationDtls VD ON VD.BLID = NVO_BoL.ID " +
                           " inner join NVO_VoyageAllocation VL on VL.ID = VD.VoyAllocID  inner join NVO_Voyage On NVO_Voyage.ID = VL.VesVoyID ";

            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  VL.VesVoyID=" + VesVoy;
                else

                    strWhere += " AND VL.VesVoyID=" + VesVoy;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.PODID=" + POD;
                else
                    strWhere += " AND NVO_Booking.PODID=" + POD;

            if (TsPort != "" && TsPort != "0" && TsPort != "null" && TsPort != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.TSPORTID=" + TsPort;
                else
                    strWhere += " AND NVO_Booking.TSPORTID=" + TsPort;

            if (Agent != "" && Agent != "0" && Agent != "null" && Agent != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.DestinationAgentID=" + Agent;
                else
                    strWhere += " AND NVO_Booking.DestinationAgentID=" + Agent;

            if (VesselOperator != "" && VesselOperator != "0" && VesselOperator != "null" && VesselOperator != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Booking.SlotOperatorID=" + VesselOperator;
                else
                    strWhere += " AND NVO_Booking.SlotOperatorID=" + VesselOperator;

            if (strWhere == "")
                strWhere = _Query;

            return RegMng.GetViewData(strWhere, "");
        }
        public DataTable GetCntrDtlsTerminalDepPdfValues(string VesVoy, string POD, string TsPort, string Agent, string VesselOperator)
        {
            string strWhere = "";
            string _Query = "Select * from NVO_ViewTerminalDepReport WHERE  VesVoyID=" + VesVoy + "";

            if (POD != "" && POD != "0" && POD != null && POD != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) PODID   from NVO_Booking where ID = BkgId) =" + POD;
                else
                    strWhere += " and (select top(1) PODID   from NVO_Booking where ID = BkgId) =" + POD;


            if (TsPort != "" && TsPort != "0" && TsPort != null && TsPort != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) TSPORTID   from NVO_Booking where ID = BkgId)=" + TsPort;
                else
                    strWhere += " and (select top(1) TSPORTID   from NVO_Booking where ID = BkgId)=" + TsPort;

            if (Agent != "" && Agent != "0" && Agent != null && Agent != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) DestinationAgentID   from NVO_Booking where ID = BkgId)=" + Agent;
                else
                    strWhere += " and (select top(1) DestinationAgentID   from NVO_Booking where ID = BkgId)=" + Agent;


            if (VesselOperator != "" && VesselOperator != "0" && VesselOperator != null && VesselOperator != "?")

                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) SlotOperatorID   from NVO_Booking where ID = BkgId)=" + VesselOperator;
                else
                    strWhere += " and (select top(1) SlotOperatorID   from NVO_Booking where ID = BkgId)=" + VesselOperator;


            if (strWhere == "")
                strWhere = _Query;

            return RegMng.GetViewData(strWhere, "");
        }

        public DataTable GetCompanyDetails()
        {

            string _Query = "Select * from NVO_NewCompnayDetails";
            return RegMng.GetViewData(_Query, "");
        }

        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
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