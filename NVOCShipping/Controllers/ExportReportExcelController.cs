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
using System.Text;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace NVOCShipping.Controllers
{
    public class ExportReportExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: ExportReportExcel
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult FreightManifestReport(string BLID, string VesVoyID)
        {
            BindFreightManifestReport(BLID, VesVoyID);
            return View();
        }
        public void BindFreightManifestReport(string BLID, string VesVoyID)
        {


            StringBuilder SB = new StringBuilder();
            string AllHeader = "<td align='left' style='color:#000; font-family:Arial; font-weight: bold; font-size:17px;'>";
            string AllValues = "<td align='left' style='color:#000; font-family:Arial; font-weight: normal; font-size:17px;'>";
            SB.Append("<table align=\"center\" style='margin-top:20px;'>");
            SB.Append("<tr>");
            SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
            SB.Append("</tr>");
            SB.Append("</table>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("<table align=\"center\" style='margin-top:20px;'>");
            SB.Append("<tr>");
            SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'>BLUE WAVE SHIPPING (M) SDN. BHD.</td>");
            SB.Append("</tr>");
            SB.Append("</table>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("<table align=\"center\" style='margin-top:20px;'>");
            SB.Append("<tr>");
            SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
            SB.Append("</tr>");
            SB.Append("</table>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("<table align=\"center\" style='margin-top:20px;'>");
            SB.Append("<tr>");
            SB.Append("<td align=\"center\" colspan=\"9\" style='font-size:20px; font-family:Arial;'><B>FREIGHT MANIFEST REPORT</B></td>");
            SB.Append("</tr>");
            SB.Append("</table>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("<table align=\"center\" style='margin-top:20px;'>");
            SB.Append("<tr>");
            SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
            SB.Append("</tr>");
            SB.Append("</table>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            SB.Append("</br>");
            DataTable _dtm = GetFreghtManifestMainDtls(VesVoyID, BLID);
            for (int x = 0; x < _dtm.Rows.Count; x++)
            {
                if (_dtm.Rows.Count > 0)
                {
                    SB.Append("<table style='margin-left:20px;'>");
                    SB.Append("<tr>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "VESSEL</td>");
                    SB.Append(AllValues + _dtm.Rows[x]["VesVoy"].ToString() + "</td>");
                    SB.Append(AllHeader + " VOYAGE </td>");
                    SB.Append(AllValues + _dtm.Rows[x]["VoyageNo"].ToString() + "</td>");
                    SB.Append(AllHeader + "LOAD TERMINAL</td>");
                    SB.Append(AllValues + _dtm.Rows[x]["LoadTerminal"].ToString() + "</td>");
                    SB.Append("</tr>");
                    SB.Append("<tr></tr>");
                    SB.Append("<tr>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "ETD</td>");
                    SB.Append(AllValues + _dtm.Rows[x]["ETD"].ToString() + "</td>");
                    SB.Append(AllHeader + "ETA</td>");
                    SB.Append(AllValues + _dtm.Rows[x]["ETA"].ToString() + "</td>");
                    SB.Append(AllHeader + "POD</td>");
                    SB.Append(AllValues + _dtm.Rows[x]["POD"].ToString() + "</td>");
                    SB.Append(AllHeader + "POL</td>");
                    SB.Append(AllValues + _dtm.Rows[x]["POL"].ToString() + "</td>");
                    SB.Append("</tr>");
                    SB.Append("</table>");
                }
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("<table align=\"center\" style='margin-top:20px;'>");
                SB.Append("<tr>");
                SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
                SB.Append("</tr>");
                SB.Append("</table>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                SB.Append("</br>");
                int RowIndex = 11;
                DataTable _dtBL = GetFrieghtManifestDtls(_dtm.Rows[x]["BLID"].ToString());

                if (_dtBL.Rows.Count > 0)
                {
                    //for (int i = 0; i < _dtBL.Rows.Count; i++)
                    //{
                    SB.Append("<table style='margin-left:20px;'>");
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Bill Of Lading</td>");
                    SB.Append(AllHeader + "Shipper /Consignee/Notify</td>");
                    SB.Append(AllHeader + "Marks & Numbers</td>");
                    SB.Append(AllHeader + "No Of Package</td>");
                    SB.Append(AllHeader + "Description of Goods</td>");
                    SB.Append(AllHeader + "Gross Weight</td>");
                    SB.Append(AllHeader + "CBM</td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["BLNo"].ToString() + "</td>");
                    SB.Append(AllHeader + "Shipper</td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["MarkNo"].ToString() + "</td>");
                    // SB.Append(AllValues + _dtBL.Rows[0]["NoOfPkg"].ToString() + "</td>");
                    //SB.Append(AllHeader + "SHIPPERS STOW,COUNT,LOAD&SEALED</td>");
                    //SB.Append(AllValues + _dtBL.Rows[0]["GrsWt"].ToString() + "</td>");
                    //SB.Append(AllValues + _dtBL.Rows[0]["CBM"].ToString() + "</td>");


                    decimal intGrwet = 0;
                    decimal intPakage = 0;
                    decimal intCBM = 0;
                    for (int h = 0; h < _dtBL.Rows.Count; h++)
                    {
                        intPakage += decimal.Parse(_dtBL.Rows[h]["NoOfPkg"].ToString());

                    }
                    intGrwet += decimal.Parse(_dtBL.Rows[0]["GrsWt"].ToString());
                    //intCBM += decimal.Parse(_dtBL.Rows[0]["CBM"].ToString());


                    SB.Append(AllValues + intPakage + " " + _dtBL.Rows[0]["CargoPakage"].ToString() + "</td>");
                    SB.Append(AllValues + "" + "</td>");
                    SB.Append(AllValues + intGrwet + " " + _dtBL.Rows[0]["GrsWtType"].ToString() + "</td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["CBM"].ToString() + "</td>");

                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Net.Weight" + "</td>");
                    SB.Append("</tr>");
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["NetWeight"].ToString() + " " + _dtBL.Rows[0]["NtWtType"].ToString() + "</td>");
                    SB.Append("</tr>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    var splitv = _dtBL.Rows[0]["Shipper"].ToString().Split('\n');
                    var tdsh = "";
                    for (int m = 0; m < splitv.Length; m++)
                    {
                        tdsh += splitv[m].ToString() + "<br/>";

                    }
                    SB.Append("<td style='vertical-align:top;'>");
                    SB.Append("<table>");
                    SB.Append("<tr>");

                    SB.Append(AllValues + tdsh + "\n\n" + "</td>");
                    SB.Append("</tr>");
                    SB.Append("<tr>");
                    var split = _dtBL.Rows[0]["ShipperAddress"].ToString().Split('\n');
                    var tdv = "";
                    for (int k = 0; k < split.Length; k++)
                    {
                        tdv += split[k].ToString() + "<br/>";

                    }
                    SB.Append("<td>" + tdv + "</td>\n");
                    SB.Append("</tr>");
                    SB.Append("</table>");
                    SB.Append("</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["CagoDescription"].ToString() + "</td>");
                    SB.Append("</tr>");
                    RowIndex++;

                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");

                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Consignee</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    var splitc = _dtBL.Rows[0]["Consignee"].ToString().Split('\n');
                    var tdc = "";
                    for (int m = 0; m < splitc.Length; m++)
                    {
                        tdc += splitc[m].ToString() + "<br/>";

                    }
                    SB.Append(AllValues + tdc + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    var splitcn = _dtBL.Rows[0]["ConsigneeAddress"].ToString().Split('\n');
                    var tdcn = "";
                    for (int k = 0; k < splitcn.Length; k++)
                    {
                        tdcn += splitcn[k].ToString() + "<br/>";

                    }
                    SB.Append(AllValues + tdcn + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Notify</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    var splitn = _dtBL.Rows[0]["Notify1"].ToString().Split('\n');
                    var tdn = "";
                    for (int m = 0; m < splitn.Length; m++)
                    {
                        tdn += splitn[m].ToString() + "<br/>";

                    }
                    SB.Append(AllValues + tdn + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    var splitnv = _dtBL.Rows[0]["Notify1Address"].ToString().Split('\n');
                    var tdnv = "";
                    for (int k = 0; k < splitnv.Length; k++)
                    {
                        tdnv += splitnv[k].ToString() + "<br/>";

                    }
                    SB.Append(AllValues + tdnv + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["IntFreeDays"].ToString() + '-' + _dtBL.Rows[0]["ddlFreeday"].ToString() + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Continuity as per annexure attached</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;
                    //SB.Append("<tr " + RowIndex + ">");
                    //SB.Append("<td></td>");
                    //SB.Append("<td></td>");
                    //SB.Append(AllHeader + "Container Size/Type/Seal</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Freight Collection Mode:</td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    RowIndex++;

                    //SB.Append("<tr " + RowIndex + ">");
                    //SB.Append("<td></td>");
                    //SB.Append("<td></td>");
                    //SB.Append(AllValues + _dtBL.Rows[0]["Seal"].ToString() + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append("<td></td>");
                    SB.Append(AllValues + _dtBL.Rows[0]["FreightPayment"].ToString() + "</td>");
                    SB.Append("<td></td>");
                    SB.Append("</tr>");
                    SB.Append("</table>");

                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("<table align=\"center\" style='margin-top:20px;'>");
                    SB.Append("<tr>");
                    SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
                    SB.Append("</tr>");
                    SB.Append("</table>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("<table style='margin-top:20px;'>");
                    SB.Append("<tr>");
                    SB.Append("<td></td>");
                    SB.Append(AllHeader + "Container No</td>");
                    SB.Append(AllHeader + "Seal No</td>");
                    SB.Append(AllHeader + "Type</td>");
                    DataTable dtx = GetCntrChargesBreakupNew(_dtBL.Rows[0]["RRID"].ToString(), _dtm.Rows[x]["BLID"].ToString());
                    // DataTable dtx = GetCntrChargesBreakup(_dtBL.Rows[0]["RRID"].ToString());
                    for (int y = 0; y < dtx.Rows.Count; y++)
                    {
                        SB.Append(AllHeader + dtx.Rows[y]["ChgCode"].ToString() + "</td>");
                    }

                    SB.Append(AllHeader + "TOTAL</td>");
                    SB.Append("</tr>");
                    SB.Append("<tr " + RowIndex + ">");
                    SB.Append("<td></td>");

                    decimal Totalv1 = 0;
                    DataTable _dtCnt = GetCntrdisplay(_dtm.Rows[x]["BLID"].ToString());
                    for (int j = 0; j < _dtCnt.Rows.Count; j++)
                    {
                        SB.Append(AllValues + _dtCnt.Rows[j]["CntrNo"].ToString() + "</td>");
                        SB.Append(AllValues + _dtCnt.Rows[j]["SealNo"].ToString() + "</td>");
                        SB.Append(AllValues + _dtCnt.Rows[j]["size"].ToString() + "</td>");
                        Totalv1 = 0;
                        DataTable _dtRm = GetCntrChargesBreakupNew(_dtBL.Rows[0]["RRID"].ToString(), _dtm.Rows[x]["BLID"].ToString());

                        for (int y = 0; y < _dtRm.Rows.Count; y++)
                        {
                            DataTable dtmr = GetCntrChargesBreakup(_dtBL.Rows[0]["RRID"].ToString(), _dtRm.Rows[y]["ChargeCodeID"].ToString(), _dtCnt.Rows[j]["TypeID"].ToString(), _dtm.Rows[x]["BLID"].ToString());
                            for (int r = 0; r < dtmr.Rows.Count; r++)
                            {
                                SB.Append(AllValues + dtmr.Rows[r]["ManifRate"].ToString() + "</td>");

                                Totalv1 += decimal.Parse(dtmr.Rows[r]["ManifRate"].ToString());

                            }

                        }

                        SB.Append(AllValues + Totalv1.ToString() + "</td>");
                        SB.Append("</tr>");
                        SB.Append("<td></td>");
                    }





                    //SB.Append(AllValues + _dtBL.Rows[0]["cntrdetails"].ToString() + "</td>");
                    ////  SB.Append(AllValues + _dtBL.Rows[0]["size"].ToString() + "</td>");

                    //for (int y = 0; y < dtx.Rows.Count; y++)
                    //{
                    //    SB.Append(AllHeader + dtx.Rows[y]["ManifRate"].ToString() + "</td>");
                    //}


                    //SB.Append(AllValues + _dtBL.Rows[0]["Total"].ToString() + "</td>");
                    SB.Append("</tr>");
                    SB.Append("</table>");

                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("<table align=\"center\" style='margin-top:20px;border-bottom:2px dashed #ccc;'>");
                    SB.Append("<tr>");
                    SB.Append("<td></td>");
                    SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
                    SB.Append("</tr>");
                    SB.Append("</table>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("</br>");
                    SB.Append("<table align=\"center\"");
                    SB.Append("<tr>");
                    SB.Append("<td></td>");
                    SB.Append("<td align=\"center\" colspan=\"9\" style='font-family:Arial;font-size:25px; color:#21277a;font-weight:bold;'></td>");
                    SB.Append("</tr>");
                    SB.Append("</table>");
                }



            }
            Response.Write(SB.ToString());
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=FreightManifestReport.xls");
            Response.Charset = "";
            Response.ContentType = "application/vnd.xls";
            System.IO.StringWriter stringWrite = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
            htmlWrite.Write(SB.ToString());
            Response.Write(stringWrite.ToString());
            Response.End();
            //using (ExcelPackage _xlPV = new ExcelPackage())
            //{
            //    _xlPV.Workbook.Properties.Author = "FREIGHT MANIFEST REPORT";
            //    _xlPV.Workbook.Properties.Title = "FREIGHT MANIFEST REPORT";
            //    _xlPV.Workbook.Worksheets.Add("FREIGHT MANIFEST REPORT");
            //    ExcelWorksheet ws = _xlPV.Workbook.Worksheets[1];
            //    double maximumSizeJob = 17;
            //    double maximumSizeCustomer = 26;
            //    double maximumSizeContNo = 14;
            //    double maximumSizeContType = 12;

            //    ws.Name = "FREIGHT MANIFEST REPORT"; //Setting Sheet's name
            //    ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
            //    int MaxCol = 5;
            //    int RowIndex = 4;
            //    ws.Cells["A2"].Value = "OCEANUS CONTAINER LINES";
            //    //ws.Cells[RowIndex, 1].Value = "STATEMENT OF ACCOUNT RECEIVABLE";
            //    ws.Cells["A2"].Style.Font.Bold = true;
            //    ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    ExcelRange r1 = ws.Cells["A1:L1"];
            //    r1.Merge = true;
            //    ExcelRange r2 = ws.Cells["A2:L2"];
            //    r2.Merge = true;
            //    r2.Style.Font.Size = 14;
            //    r2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    r2.Style.Fill.BackgroundColor.SetColor(Color.White);
            //    r2.Style.Font.Color.SetColor(Color.Black);
            //    ws.Cells["A2:L2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //    ws.Cells["A2:L2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //    ws.Cells["A2:L2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //    ws.Cells["A2:L2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //    ExcelRange r3 = ws.Cells["A3:L3"];
            //    r3.Merge = true;
            //    ExcelRange r4 = ws.Cells["A4:L4"];
            //    r4.Merge = true;
            //    ws.Cells["A4"].Value = "FREIGHT MANIFEST REPORT";
            //    ws.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    r4.Style.Font.Size = 12;
            //    r4.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    r4.Style.Fill.BackgroundColor.SetColor(Color.White);
            //    r4.Style.Font.Color.SetColor(Color.Navy);
            //    r4.Style.Font.Bold = true;

            //    ws.Cells["A4:L4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //    ws.Cells["A4:L4"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //    ws.Cells["A4:L4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //    ws.Cells["A4:L4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;



            //    DataTable _dtm = GetFreghtManifestMainDtls(VesVoyID);
            //    if (_dtm.Rows.Count > 0)
            //    {

            //        ws.Cells["A5"].Value = "VESSEL";
            //        ws.Cells["A5"].AutoFitColumns();
            //        ws.Cells["A5"].Style.Font.Bold = true;
            //        ws.Cells["B5:C5"].Value = _dtm.Rows[0]["VesVoy"].ToString();
            //        ws.Cells["B5:C5"].Merge = true;
            //        ws.Cells["B5:C5"].Style.Font.Bold = false;
            //        ws.Cells["D5"].Value = "VOYAGE";
            //        ws.Cells["D5"].AutoFitColumns();
            //        ws.Cells["D5"].Style.Font.Bold = true;
            //        ws.Cells["E5"].Value = _dtm.Rows[0]["VoyageNo"].ToString();
            //        ws.Cells["E5"].AutoFitColumns();
            //        ws.Cells["E5"].Style.Font.Bold = false;
            //        ws.Cells["G5"].Value = "LOAD TERMINAL";
            //        ws.Cells["G5"].AutoFitColumns();
            //        ws.Cells["G5"].Style.Font.Bold = true;
            //        ws.Cells["H5:I5"].Value = _dtm.Rows[0]["LoadTerminal"].ToString();
            //        ws.Cells["H5:I5"].AutoFitColumns();
            //        ws.Cells["H5:I5"].Merge = true;
            //        ws.Cells["H5:I5"].Style.Font.Bold = false;

            //        ws.Cells["A7"].Value = "ETD";
            //        ws.Cells["A7"].AutoFitColumns();
            //        ws.Cells["A7"].Style.Font.Bold = true;
            //        ws.Cells["B7:C7"].Value = _dtm.Rows[0]["ETD"].ToString();
            //        ws.Cells["B7:C7"].AutoFitColumns();
            //        ws.Cells["B7:C7"].Merge = true;
            //        ws.Cells["B7:C7"].Style.Font.Bold = false;
            //        ws.Cells["D7"].Value = "ETA";
            //        ws.Cells["D7"].AutoFitColumns();
            //        ws.Cells["D7"].Style.Font.Bold = true;
            //        ws.Cells["E7:F7"].Value = _dtm.Rows[0]["ETA"].ToString();
            //        ws.Cells["E7:F7"].AutoFitColumns();
            //        ws.Cells["E7:F7"].Merge = true;
            //        ws.Cells["E7:F7"].Style.Font.Bold = false;


            //        ws.Cells["G7"].Value = "POD";
            //        ws.Cells["G7"].AutoFitColumns();
            //        ws.Cells["G7"].Style.Font.Bold = true;
            //        ws.Cells["H7:I7"].Value = _dtm.Rows[0]["POD"].ToString();
            //        ws.Cells["H7:I7"].AutoFitColumns();
            //        ws.Cells["H7:I7"].Merge = true;
            //        ws.Cells["H7:I7"].Style.Font.Bold = false;
            //        ws.Cells["J7"].Value = "POL";
            //        ws.Cells["J7"].AutoFitColumns();
            //        ws.Cells["J7"].Style.Font.Bold = true;
            //        ws.Cells["K7:L7"].Value = _dtm.Rows[0]["POL"].ToString();
            //        ws.Cells["K7:L7"].AutoFitColumns();
            //        ws.Cells["K7:L7"].Merge = true;
            //        ws.Cells["K7:L7"].Style.Font.Bold = false;
            //    }
            //    RowIndex = 9;

            //    RowIndex++;
            //    DataTable _dtBL = GetFrieghtManifestDtls(BLID);
            //    if(_dtBL.Rows.Count>0)
            //    {
            //        for(int i=0; i<_dtBL.Rows.Count; i++)
            //        {
            //            ws.Cells["A" + RowIndex].Value = "BILL OF LADING NO";
            //            ws.Cells["B" + RowIndex].Value = "SHIPPER/CONSIGNEE/NOTIFY";
            //            //ws.Cells["C" + RowIndex].Value = "MARKS & NUMBERS";
            //            //ws.Cells["D" + RowIndex].Value = "NO OF PACKAGE";
            //            //ws.Cells["E" + RowIndex].Value = "DESCRIPTION OF GOODS";
            //            //ws.Cells["F" + RowIndex].Value = "GROSSS.WEIGHT";
            //            //ws.Cells["G" + RowIndex].Value = "CBM";

            //            RowIndex += 1;
            //            ws.Cells["A" + RowIndex].Value = _dtBL.Rows[i]["BLNo"].ToString();                      
            //            ws.Cells["B:F" + RowIndex].Value = "SHIPPER";
            //            RowIndex += 1;
            //            ws.Cells["B:F" + RowIndex].Value = _dtBL.Rows[i]["Shipper"].ToString();
            //            RowIndex += 1;
            //            ws.Cells["B:F" + RowIndex].Value = _dtBL.Rows[i]["ShipperAddress"].ToString();
            //            ws.Cells["B:F" + RowIndex].AutoFitColumns();

            //            RowIndex += 3;

            //            ws.Cells["B:F" + RowIndex].Value = "CONSIGNEE";
            //            RowIndex += 1;
            //            ws.Cells["B:F" + RowIndex].Value = _dtBL.Rows[i]["Consignee"].ToString();
            //            RowIndex += 1;
            //            ws.Cells["B:F" + RowIndex].Value = _dtBL.Rows[i]["ConsigneeAddress"].ToString();

            //            RowIndex += 3;

            //            ws.Cells["B:F" + RowIndex].Value = "NOTIFY";
            //            RowIndex += 1;
            //            ws.Cells["B:F" + RowIndex].Value = _dtBL.Rows[i]["Notify1"].ToString();
            //            RowIndex += 1;
            //            ws.Cells["B:F" + RowIndex].Value = _dtBL.Rows[i]["Notify1Address"].ToString();



            //        }
            //        RowIndex++;
            //    }

            //    string _filePath = Server.MapPath("~/PreXL/");
            //    string _fileName = "FreightManifestReport " + Session.SessionID + ".xlsx";
            //    Byte[] bin = _xlPV.GetAsByteArray();
            //    System.IO.File.WriteAllBytes(_filePath + "\\" + _fileName, bin);
            //    Response.Clear();
            //    Response.AppendHeader("content-disposition", "attachment; filename=" + _fileName);
            //    Response.ContentType = "application/octet-stream";
            //    Response.WriteFile(_filePath + "\\" + _fileName);
            //    Response.Flush();
            //    Response.End();
            //}

        }
        public DataTable GetFreghtManifestMainDtls(string VesVoyID, string BLID)
        {
            //string _Query = " select NVO_BLRelease.ID, NVO_BOL.ID AS BLID,BLNo,Agent, AgentAddress, NVO_Ratesheet.ID  as RRID,  (select top(1)  ExportVoyageCd from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID " +
            // " order by NVO_VoyageRoute.RID desc) as VoyageNo, (Select top 1 (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) from NVO_Voyage V where V.ID = NVO_BOL.BLVesVoyID) as VesVoy , (select top 1 Convert(varchar, ETA, 106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID)as ETA, " +
            // " (select top 1 Convert(varchar, ETD, 106) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID)as ETD,  NVO_BLRelease.POD,NVO_BLRelease.POL,NVO_BOL.BLVesVoyID, " +
            // " (select Top(1) TerminalName from NVO_TerminalMaster inner join NVO_VoyageRoute On NVO_VoyageRoute.TerminalID = NVO_TerminalMaster.ID " +
            // " where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID) as LoadTerminal  from NVO_BLRelease inner join NVO_BOL On NVO_BOL.ID = NVO_BLRelease.BLID  inner join NVO_Booking On NVO_Booking.ID = NVO_BOL.BkgID inner join NVO_Ratesheet On NVO_Ratesheet.ID = NVO_Booking.RRID " +
            // " where NVO_BOL.BLVesVoyID = " + VesVoyID;
            //if (BLID != "")
            //    _Query += " and NVO_BOL.ID in (" + BLID + ")";



            //return Manag.GetViewData(_Query, "");


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
            return Manag.GetViewData(_Query, "");

        }
        /////     public DataTable GetFreghtManifestMainDtls(string VesVoyID)
        //     {
        //         string _Query = "select NVO_BLRelease.ID, BLNo,Agent, AgentAddress, VesVoy,NVO_Ratesheet.ID as RRID, " +
        //                         " (select top(1)  ExportVoyageCd from NVO_VoyageRoute Inner join NVO_BOL ON NVO_BOL.ID = NVO_VoyageRoute.RID where NVO_VoyageRoute.VoyageID = NVO_BOL.BLVesVoyID order by NVO_VoyageRoute.RID desc) as VoyageNo, " +
        //                         " (select Convert(varchar, ETA, 106) from NVO_VoyageRoute where NVO_VoyageRoute.RID = NVO_BLRelease.ID) as ETA, " +
        //                         " (select Convert(varchar, ETD, 106) from NVO_VoyageRoute where NVO_VoyageRoute.RID = NVO_BLRelease.ID)as ETD,POD,POL,NVO_BOL.BLVesVoyID, " +
        //                         " (select Top(1) TerminalName from NVO_TerminalMaster inner join NVO_VoyageRoute On NVO_VoyageRoute.TerminalID = NVO_TerminalMaster.ID where NVO_VoyageRoute.RID = NVO_BLRelease.ID)as LoadTerminal " +
        //                         " from NVO_BLRelease inner join NVO_Ratesheet On NVO_Ratesheet.ID = NVO_BLRelease.BLID inner join NVO_BOL On NVO_BOL.ID = NVO_BLRelease.BLID where NVO_BOL.BLVesVoyID = " + VesVoyID;
        //         return Manag.GetViewData(_Query, "");
        //     }
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
                            " inner join NVO_Ratesheet On NVO_Ratesheet.ID = NVO_Booking.RRID " +
                            " left outer join NVO_RatesheetCntrTypes on NVO_RatesheetCntrTypes.RRID = NVO_Ratesheet.ID " +
                            " where NVO_BLRelease.BLID in (" + BLID + ")";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCntrChargesBreakup(string RID)
        {
            string _Query = "select * from NVO_V_FreightManifest where RRID=" + RID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCntrChargesBreakup(string RID, string ChgcodeID, string CntrType, string BLID)
        {

            string _Query = "select * from NVO_V_FreightManifest where RRID=" + RID + " AND ChargeCodeID=" + ChgcodeID + " and CntrType =" + CntrType + " and BLID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCntrChargesBreakupNew(string RID, string BLID)
        {

            string _Query = "select  DISTINCT ChgCode,ChargeCodeID from NVO_V_FreightManifest where RRID=" + RID + " and BLID=" + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCntrdisplay(string BLID)
        {

            string _Query = "select (select top(1) CntrNo from NVO_Containers where ID=CntrID) as CntrNo," +
                 " (select top(1) TypeID from NVO_Containers where ID = CntrID) as TypeID," +
                " SealNo,BkgId,Size from NVO_BOLCntrDetails  where BLID= " + BLID;
            return Manag.GetViewData(_Query, "");
        }
    }
}