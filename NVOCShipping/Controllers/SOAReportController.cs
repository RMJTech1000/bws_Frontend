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

namespace NVOCShipping.Controllers
{
    public class SOAReportController : Controller
    {
        ExportReportManager RegManage = new ExportReportManager();
        public ActionResult Index()
        {
            return View();
        }

        public void SOAReportExport(string DtFrom, string DtTo, string AgencyID)
        {
            DataTable dtx = GetSOAExport(DtFrom, DtTo, AgencyID);
            if (dtx.Rows.Count > 0)
            {
                DataTable _dtAgnt = GetAgenctDetails(AgencyID);


                ExcelPackage pck = new ExcelPackage();

                #region  Export
                var ws = pck.Workbook.Worksheets.Add("Export");

                ExcelRange r;



                ws.Cells["A1"].Value = "Statement of Billing  :";
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["A2"].Value = "For the month of      :";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A3"].Value = "Agent                 :";
                ws.Cells["A3"].Style.Font.Bold = true;
                ws.Cells["A4"].Value = "Location              :";
                ws.Cells["A4"].Style.Font.Bold = true;

                ws.Cells["B1"].Value = "Outward / Exports";
                ws.Cells["B1"].Style.Font.Bold = true;
                ws.Cells["B2"].Value = DtFrom + " to " + DtTo;
                ws.Cells["B2"].Style.Font.Bold = true;
                ws.Cells["B3"].Value = _dtAgnt.Rows[0]["AgencyName"].ToString();
                ws.Cells["B3"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = _dtAgnt.Rows[0]["GeoLocation"].ToString();
                ws.Cells["B4"].Style.Font.Bold = true;

                //Record Headers
                int rw;



                int sl3 = 1;

                rw = 11;
                ws.Cells["N7:W7"].Value = "REVENUE IN USD ";
                ws.Cells["N7:W7"].Merge = true;
                ws.Cells["N7:W7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["X7:AC7"].Value = "COST IN USD ";
                ws.Cells["X7:AC7"].Merge = true;
                ws.Cells["X7:AC7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                ws.Cells["A8"].Value = "S. No.";
                ws.Cells["B8"].Value = "Vessel / Voy";
                ws.Cells["C8"].Value = "Sailing";
                ws.Cells["D8"].Value = "Port Of ";
                ws.Cells["E8"].Value = "Port Of ";
                ws.Cells["F8"].Value = "Feeder ";
                ws.Cells["G8"].Value = "BL Number";
                ws.Cells["H8"].Value = "Container No";
                ws.Cells["I8:K8"].Value = "Quantity";
                ws.Cells["I8:K8"].Merge = true;
                ws.Cells["L8"].Value = "Type";
                ws.Cells["M8"].Value = "Terms";
                ws.Cells["N8:O8"].Value = "Ocean Freight";
                ws.Cells["N8:O8"].Merge = true;
                ws.Cells["P8:Q8"].Value = "BAF/CAF/FAF";
                ws.Cells["P8:Q8"].Merge = true;
                ws.Cells["R8"].Value = "DGS";
                ws.Cells["S8"].Value = "OTHER COLLECTION";
                ws.Cells["T8"].Value = "MUC";
                ws.Cells["U8"].Value = "TOLL";
                ws.Cells["V8"].Value = "THC COLLECTION";
                ws.Cells["V8"].Merge = true;
                ws.Cells["W8"].Value = "Total";
                ws.Cells["X8"].Value = "Commission";
                ws.Cells["Y8"].Value = "Logistics Fees";

                ws.Cells["Z8"].Value = "COMM ON GST";

                ws.Cells["AA8"].Value = " Thc / Ihc";
                ws.Cells["AB8"].Value = "Feeder ";
                ws.Cells["AC8"].Value = " Other";
                ws.Cells["AD8"].Value = " Total ";
                ws.Cells["AE8"].Value = " Nett Due ";


                ws.Cells["C9"].Value = " Date";
                ws.Cells["D9"].Value = "Loading";
                ws.Cells["E9"].Value = "Discharge";
                ws.Cells["F9"].Value = "Operator";


                ws.Cells["I9"].Value = " 20'";
                ws.Cells["J9"].Value = " 40'";
                ws.Cells["K9"].Value = "40' RF";
                ws.Cells["L9"].Value = "DG/GEN/MTY";
                ws.Cells["M9"].Value = "CFS/CY";
                ws.Cells["N9"].Value = "Prepaid";
                ws.Cells["O9"].Value = "Collect";
                ws.Cells["P9"].Value = "Prepaid";
                ws.Cells["Q9"].Value = "Collect";
                ws.Cells["R9"].Value = "PrePaid";
                ws.Cells["S9"].Value = "PCS / ISPS";
                ws.Cells["T9"].Value = "Prepaid";
                ws.Cells["U9"].Value = "Prepaid";
                ws.Cells["V9"].Value = "Prepaid";
                ws.Cells["W9"].Value = "Revenue In";
                ws.Cells["X9"].Value = "5%";
                ws.Cells["Y9"].Value = "@ USD 5.00";
                ws.Cells["Z9"].Value = "Cost vs THC";
                ws.Cells["AA9"].Value = "Cost";
                ws.Cells["AB9"].Value = " Cost if any";
                ws.Cells["AC9"].Value = "In";
                ws.Cells["AD9"].Value = "Cost";
                ws.Cells["AE10"].Value = "US $";
                ws.Cells["N10"].Value = "US $";
                ws.Cells["O10"].Value = "US $";
                ws.Cells["P10"].Value = "US $";
                ws.Cells["Q10"].Value = "US $";
                ws.Cells["R10"].Value = "US $";
                ws.Cells["T10"].Value = "US $";
                ws.Cells["U10"].Value = "US $";
                ws.Cells["V10"].Value = "US $";
                ws.Cells["W10"].Value = "US $";
                ws.Cells["X9"].Value = "PER UNIT";
                ws.Cells["AA10"].Value = "US $";
                ws.Cells["AB10"].Value = "USD";

                r = ws.Cells["A7:AD10"];
                r.Style.Font.Bold = true;
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                r = ws.Cells["N7:W10"];
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                r = ws.Cells["X7:AD10"];
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                r = ws.Cells["AE7:AE10"];
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.BlanchedAlmond);
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                for (int i = 0; i < dtx.Rows.Count; i++)
                {

                    ws.Cells["A" + rw].Value = sl3;
                    ws.Cells["B" + rw].Value = dtx.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtx.Rows[i]["SailingDate"].ToString();
                    ws.Cells["D" + rw].Value = dtx.Rows[i]["POL"].ToString();
                    ws.Cells["E" + rw].Value = dtx.Rows[i]["POD"].ToString();
                    ws.Cells["F" + rw].Value = dtx.Rows[i]["SlotOperator"].ToString();
                    ws.Cells["G" + rw].Value = dtx.Rows[i]["BookingNo"].ToString();
                    ws.Cells["H" + rw].Value = dtx.Rows[i]["CntrNo"].ToString();
                    ws.Cells["I" + rw].Value = dtx.Rows[i]["Size20"].ToString();
                    ws.Cells["J" + rw].Value = dtx.Rows[i]["Size40"].ToString();
                    ws.Cells["K" + rw].Value = dtx.Rows[i]["SizeRF40"].ToString();
                    ws.Cells["L" + rw].Value = dtx.Rows[i]["Commodity"].ToString();
                    ws.Cells["M" + rw].Value = dtx.Rows[i]["slotTerm"].ToString();

                    ws.Cells["N" + rw].Value = dtx.Rows[i]["OFTPreAmt"];
                    ws.Cells["O" + rw].Value = dtx.Rows[i]["OFTCollAmt"];
                    ws.Cells["P" + rw].Value = dtx.Rows[i]["BAFPreAmt"];
                    ws.Cells["Q" + rw].Value = dtx.Rows[i]["BAFCollAmt"];
                    ws.Cells["R" + rw].Value = dtx.Rows[i]["DGSAmt"];
                    ws.Cells["S" + rw].Value = dtx.Rows[i]["PCSISPSAmt"];
                    ws.Cells["T" + rw].Value = dtx.Rows[i]["MUCAmt"];
                    ws.Cells["U" + rw].Value = dtx.Rows[i]["TOLLAmt"];
                    ws.Cells["V" + rw].Value = dtx.Rows[i]["THCPreAmt"];
                    ws.Cells["W" + rw].Formula = "=SUM(N" + rw + " : V" + rw + ")";

                    decimal ExCommt = decimal.Parse(dtx.Rows[i]["ExCommAmt"].ToString());
                    if (ExCommt <= 10)
                        ws.Cells["X" + rw].Value = 10;
                    else
                        ws.Cells["X" + rw].Value = dtx.Rows[i]["ExCommAmt"];

                    // ws.Cells["W" + rw].Value = dtx.Rows[i]["ExCommAmt"];
                    ws.Cells["Y" + rw].Value = "5";
                    ws.Cells["Z" + rw].Formula = "=SUM(X" + rw + " +Y" + rw + ")*18%";
                    ws.Cells["AA" + rw].Value = dtx.Rows[i]["THCCostAmt"];
                    ws.Cells["AB" + rw].Value = dtx.Rows[i]["SlotAmt"];
                    ws.Cells["AC" + rw].Value = "";
                    ws.Cells["AD" + rw].Formula = "=SUM(X" + rw + " :AC" + rw + ")";
                    ws.Cells["AE" + rw].Formula = "=SUM(N" + rw + " : V" + rw + ") - SUM(X" + rw + " : AC" + rw + ")";
                    sl3++;
                    rw += 1;
                }


                rw -= 1;

                ws.Cells["A7:AE" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AE" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AE" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AE" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A11:AE" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[1, 1, rw, 40].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=SOAExportReport.xlsx");
                Response.End();

            }

        }

        public void SOAReportImport(string DtFrom, string DtTo, string AgencyID)
        {
            DataTable dtx = GetSOAImport(DtFrom, DtTo, AgencyID);
            if (dtx.Rows.Count > 0)
            {
                DataTable _dtAgnt = GetAgenctDetails(AgencyID);
                ExcelPackage pck = new ExcelPackage();

                #region  Import
                var ws = pck.Workbook.Worksheets.Add("Import");

                ExcelRange r;


                ws.Cells["A1"].Value = "Statement of Billing  :";
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["A2"].Value = "For the month of      :";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A3"].Value = "Agent                 :";
                ws.Cells["A3"].Style.Font.Bold = true;
                ws.Cells["A4"].Value = "Location              :";
                ws.Cells["A4"].Style.Font.Bold = true;

                ws.Cells["B1"].Value = "Inward / Imports 	";
                ws.Cells["B1"].Style.Font.Bold = true;
                ws.Cells["B2"].Value = DtFrom + " - " + DtTo;
                ws.Cells["B2"].Style.Font.Bold = true;
                if (_dtAgnt.Rows.Count > 0)
                {
                    ws.Cells["B3"].Value = _dtAgnt.Rows[0]["AgencyName"].ToString();
                    ws.Cells["B4"].Value = _dtAgnt.Rows[0]["GeoLocation"].ToString();
                }
                else
                {
                    ws.Cells["B3"].Value = "";
                    ws.Cells["B4"].Value = "";
                }
                ws.Cells["B3"].Style.Font.Bold = true;
                ws.Cells["B4"].Style.Font.Bold = true;


                //Record Headers


                int sl2 = 1;

                int rw = 11;

                ws.Cells["O7:AB7"].Value = "REVENUE IN USD ";
                ws.Cells["O7:AB7"].Merge = true;
                ws.Cells["O7:AB7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["AC7:AG7"].Value = "COST IN USD ";
                ws.Cells["AC7:AG7"].Merge = true;
                ws.Cells["AC7:AG7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["A8"].Value = "S. No.";
                ws.Cells["B8"].Value = "POL Agecny Name.";
                ws.Cells["C8"].Value = "Vessel / Voy";
                ws.Cells["D8"].Value = "Arrival";
                ws.Cells["E8"].Value = "Port Of";
                ws.Cells["F8"].Value = "Port Of ";
                ws.Cells["G8"].Value = "Destination Agent";
                ws.Cells["H8"].Value = " Final Destination";
               
                ws.Cells["I8"].Value = "Feeder";
                ws.Cells["J8"].Value = "BL Number";
                ws.Cells["K8"].Value = "Container No";
                ws.Cells["L8:N8"].Value = "Quantity";
                ws.Cells["L8:N8"].Merge = true;
                ws.Cells["O8"].Value = "Type";
                ws.Cells["P8"].Value = "Terms";
                ws.Cells["Q8:R8"].Value = "Ocean Freight";
                ws.Cells["Q8:R8"].Merge = true;
                ws.Cells["S8:T8"].Value = "BAF/CAF/FAF";
                ws.Cells["S8:T8"].Merge = true;
                ws.Cells["U8"].Value = "PCS ";
                ws.Cells["V8"].Value = "ISPS ";

                ws.Cells["W8"].Value = "Do Fees";
                ws.Cells["X8"].Value = "MUC";
                ws.Cells["Y8"].Value = "Toll";

                ws.Cells["Z8:AA8"].Value = "THC";
                ws.Cells["Z8:AA8"].Merge = true;
                ws.Cells["AB8"].Value = "Total";
                ws.Cells["AC8"].Value = "Commission";
                ws.Cells["AD8"].Value = " Thc / Ihc";
                ws.Cells["AE8"].Value = "Feeder";
                ws.Cells["AF8"].Value = " Other";
                ws.Cells["AG8"].Value = " TotalCost ";
                ws.Cells["AH8"].Value = " Nett Due ";

                ws.Cells["D9"].Value = " Date";
                ws.Cells["E9"].Value = "Loading";
                ws.Cells["F9"].Value = "Discharge";
                ws.Cells["I9"].Value = "Operator";


                ws.Cells["L9"].Value = " 20'";
                ws.Cells["M9"].Value = " 40'";
                ws.Cells["N9"].Value = "40' RF";
                ws.Cells["O9"].Value = "DG/GEN/MTY";
                ws.Cells["P9"].Value = "CFS/CY";
                ws.Cells["Q9"].Value = "Prepaid";
                ws.Cells["R9"].Value = "Collect";
                ws.Cells["S9"].Value = "Prepaid";
                ws.Cells["T9"].Value = "Collect";
                ws.Cells["U9"].Value = "Collect";
                ws.Cells["V9"].Value = "Collect";
                ws.Cells["W9"].Value = "Prepaid";
                ws.Cells["X9"].Value = "Collect";
                ws.Cells["Z9"].Value = "Collect";
                ws.Cells["Z9"].Value = "";
                ws.Cells["AA9"].Value = "Collect";
                ws.Cells["AB9"].Value = "Revenue In";
                ws.Cells["AC9"].Value = "2.50%";
                ws.Cells["AD9"].Value = "Cost vs THC";
                ws.Cells["AE9"].Value = "Cost";
                ws.Cells["AF9"].Value = " Cost if any";
                ws.Cells["AG9"].Value = "In";
                ws.Cells["AH9"].Value = "To Blue Wave Shipping";

                ws.Cells["Q10"].Value = "US $";
                ws.Cells["R10"].Value = "US $";
                ws.Cells["S10"].Value = "US $";
                ws.Cells["T10"].Value = "US $";
                ws.Cells["U10"].Value = "US $";
                ws.Cells["W10"].Value = "US $";
                ws.Cells["X10"].Value = "US $";
                ws.Cells["Z10"].Value = "US $";
                ws.Cells["AD10"].Value = "US $";
                ws.Cells["AE10"].Value = "USD";

                r = ws.Cells["A7:AE10"];
                r.Style.Font.Bold = true;

                r = ws.Cells["Q7:AB10"];
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                r = ws.Cells["AC7:AG10"];
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);

                r = ws.Cells["AH7:AH10"];
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightPink);



                for (int i = 0; i < dtx.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl2;
                    ws.Cells["B" + rw].Value = dtx.Rows[i]["POLAgency"].ToString();
                    ws.Cells["C" + rw].Value = dtx.Rows[i]["VesVoy"].ToString();
                    ws.Cells["D" + rw].Value = dtx.Rows[i]["ArrivalDate"].ToString();
                    ws.Cells["E" + rw].Value = dtx.Rows[i]["POL"].ToString();
                    ws.Cells["F" + rw].Value = dtx.Rows[i]["POD"].ToString();
                    ws.Cells["G" + rw].Value = dtx.Rows[i]["DestinationAgent"].ToString();
                    ws.Cells["H" + rw].Value = dtx.Rows[i]["FinalDestination"].ToString();
                    ws.Cells["I" + rw].Value = dtx.Rows[i]["SlotOperator"].ToString();
                    ws.Cells["J" + rw].Value = dtx.Rows[i]["BookingNo"].ToString();
                    ws.Cells["K" + rw].Value = dtx.Rows[i]["CntrNo"].ToString();
                    ws.Cells["L" + rw].Value = dtx.Rows[i]["Size20"].ToString();
                    ws.Cells["M" + rw].Value = dtx.Rows[i]["Size40"].ToString();
                    ws.Cells["N" + rw].Value = dtx.Rows[i]["SizeRF40"].ToString();
                    ws.Cells["O" + rw].Value = dtx.Rows[i]["Commodity"].ToString();
                    ws.Cells["P" + rw].Value = dtx.Rows[i]["slotTerm"].ToString();

                    ws.Cells["Q" + rw].Value = dtx.Rows[i]["OFTPreAmt"];
                    ws.Cells["R" + rw].Value = dtx.Rows[i]["OFTCollAmt"];
                    ws.Cells["S" + rw].Value = dtx.Rows[i]["BAFPreAmt"];
                    ws.Cells["T" + rw].Value = dtx.Rows[i]["BAFCollAmt"];
                    ws.Cells["U" + rw].Value = dtx.Rows[i]["PCSISPSAmt"];

                    ws.Cells["V" + rw].Value = dtx.Rows[i]["ISPSCollAmt"];

                    ws.Cells["W" + rw].Value = dtx.Rows[i]["DOFAmt"];
                    ws.Cells["X" + rw].Value = dtx.Rows[i]["MUCAmt"];
                    ws.Cells["Y" + rw].Value = dtx.Rows[i]["TOLLAmt"];

                    ws.Cells["Z" + rw].Value = dtx.Rows[i]["THCPreAmt"];
                    ws.Cells["AA" + rw].Value = dtx.Rows[i]["THCCollAmt"];
                    ws.Cells["AB" + rw].Formula = "=SUM(Q" + rw + " : AA" + rw + ")";

                    //  ws.Cells["Z" + rw].Value = dtx.Rows[i]["ICommAmt"];


                    decimal ExCommt = decimal.Parse(dtx.Rows[i]["ICommAmt"].ToString());
                    if (ExCommt <= 10)
                        ws.Cells["AC" + rw].Value = 10;
                    else
                        ws.Cells["AC" + rw].Value = dtx.Rows[i]["ICommAmt"];



                    ws.Cells["AD" + rw].Value = dtx.Rows[i]["THCCostAmt"];
                    ws.Cells["AE" + rw].Value = dtx.Rows[i]["FeederCostAmt"];
                    ws.Cells["AF" + rw].Value = "";
                    ws.Cells["AG" + rw].Formula = "=SUM(AC" + rw + " :aF" + rw + ")";
                    ws.Cells["AH" + rw].Formula = "=SUM(Q" + rw + " : AA" + rw + ") - SUM(AC" + rw + " : AF" + rw + ")";


                    sl2++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:AH" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AH" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AH" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AH" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells[1, 1, rw, 43].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=SOAImportReport.xlsx");
                Response.End();

            }

        }
        public void SOAReportValues(string DtFrom, string DtTo, string AgencyID)
        {
            //DataTable dtv = GetSOAImport(DtFrom, DtTo, AgencyID);
            //if (dtv.Rows.Count > 0)
            //{

            ExcelPackage pck = new ExcelPackage();

            #region 1ST SHEET DETENTION

            var ws = pck.Workbook.Worksheets.Add("Det-Det");

            ws.Cells["A1"].Value = "DEMURRAGE / DETENTION COLLECTIONS ";
            ExcelRange r = ws.Cells["A1:R1"];
            r.Merge = true;
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A2"].Value = "EX-RATE";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["B2"].Value = "1.00";
            ws.Cells["B2"].Style.Font.Bold = true;
            int rw = 4;

            ws.Cells["A4:G4"].Value = "";
            ws.Cells["A4:G4"].Merge = true;
            ws.Cells["N4:P4"].Value = "DETENTION";
            ws.Cells["N4:P4"].Merge = true;
            ws.Cells["A5:A7"].Value = "S.NO";
            ws.Cells["A5:A7"].Merge = true;
            ws.Cells["A5:A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["B5:B7"].Value = "VESSEL/VOYAGE";
            ws.Cells["B5:B7"].Merge = true;
            ws.Cells["B5:B7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["C5:C7"].Value = "POL";
            ws.Cells["C5:C7"].Merge = true;
            ws.Cells["C5:C7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["D5:D7"].Value = "POD";
            ws.Cells["D5:D7"].Merge = true;
            ws.Cells["D5:D7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["E5:E7"].Value = "B/L NUMBER";
            ws.Cells["E5:E7"].Merge = true;
            ws.Cells["E5:E7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["F5:F7"].Value = "CONTAINER NO";
            ws.Cells["F5:F7"].Merge = true;
            ws.Cells["F5:F7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["G5:G7"].Value = "TYPE-SIZE";
            ws.Cells["G5:G7"].Merge = true;
            ws.Cells["G5:G7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["H4"].Value = "Landing Date";
            ws.Cells["H5"].Value = "Arrival";
            ws.Cells["H6"].Value = "Date";
            ws.Cells["H7"].Value = "A(FV)";
            ws.Cells["I5"].Value = "Free";
            ws.Cells["I6"].Value = "Time";
            ws.Cells["I7"].Value = "B";
            ws.Cells["J4"].Value = "Start Date";
            ws.Cells["J5"].Value = "Free";
            ws.Cells["J6"].Value = "Time";
            ws.Cells["J7"].Value = "B";
            ws.Cells["K4"].Value = "End Date";
            ws.Cells["K5"].Value = "DTN From";
            ws.Cells["K6"].Value = "(FV + F/Time)";
            ws.Cells["K7"].Value = "A+B = C";
            ws.Cells["L4"].Value = "Total Days";
            ws.Cells["L5"].Value = "DTN Upto";
            ws.Cells["L6"].Value = "Mty";
            ws.Cells["L7"].Value = "D (MA)";
            ws.Cells["M5"].Value = "DTN";
            ws.Cells["M6"].Value = "Days";
            ws.Cells["M7"].Value = "D-C";
            ws.Cells["N5:N6"].Value = "SLAB1";
            ws.Cells["N5:N6"].Merge = true;
            ws.Cells["N7"].Value = "00- 00 DAYS";
            ws.Cells["O5:O6"].Value = "SLAB2";
            ws.Cells["O5:O6"].Merge = true;
            ws.Cells["O7"].Value = "00- 00 DAYS";
            ws.Cells["P5:P6"].Value = "SLAB2";
            ws.Cells["P5:P6"].Merge = true;
            ws.Cells["P7"].Value = "00- 00 DAYS";
            ws.Cells["Q5"].Value = "TOTAL";
            ws.Cells["Q6"].Value = "LCY";
            ws.Cells["Q7"].Value = "COLLECTION";
            ws.Cells["R5"].Value = "TOTAL";
            ws.Cells["R6"].Value = "USD";
            ws.Cells["R7"].Value = "COLLECTION";

            r = ws.Cells["A4:R7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            //for (int i = 0; i < dtv.Rows.Count; i++)
            //{
            ws.Cells["B" + rw].Value = "";
            ws.Cells["C" + rw].Value = "";
            ws.Cells["D" + rw].Value = "";
            ws.Cells["E" + rw].Value = "";
            int sl = 1;
            sl++;
            rw += 1;
            //  }

            rw -= 1;

            ws.Cells["A4:R4" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A4:R4" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A4:R4" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A4:R4" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[1, 1, rw, 18].AutoFitColumns();

            #endregion

            #region 2nd SHEET Import
            ws = pck.Workbook.Worksheets.Add("Import");

            ws.Cells["A1"].Value = "Statement of Billing  :";
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A2"].Value = "For the month of      :";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A3"].Value = "Agent                 :";
            ws.Cells["A3"].Style.Font.Bold = true;
            ws.Cells["A4"].Value = "Location              :";
            ws.Cells["A4"].Style.Font.Bold = true;

            ws.Cells["A5"].Value = "Rate Of Exchange      :";
            ws.Cells["A5"].Style.Font.Bold = true;


            ws.Cells["B1"].Value = "Inward / Imports 	";
            ws.Cells["B1"].Style.Font.Bold = true;
            ws.Cells["B2"].Value = "";
            ws.Cells["B2"].Style.Font.Bold = true;
            ws.Cells["B3"].Value = "";
            ws.Cells["B3"].Style.Font.Bold = true;
            ws.Cells["B4"].Value = "";
            ws.Cells["B4"].Style.Font.Bold = true;

            ws.Cells["B5"].Value = "";
            ws.Cells["B5"].Style.Font.Bold = true;
            //Record Headers




            int sl2 = 1;

            rw = 11;

            ws.Cells["N7:V7"].Value = "REVENUE IN USD ";
            ws.Cells["N7:V7"].Merge = true;
            ws.Cells["N7:V7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["W7:AA7"].Value = "COST IN USD ";
            ws.Cells["W7:AA7"].Merge = true;
            ws.Cells["W7:AA7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A8"].Value = "S. No.";
            ws.Cells["B8"].Value = "Vessel / Voy";
            ws.Cells["C8"].Value = "Arrival";
            ws.Cells["D8"].Value = "Port Of";
            ws.Cells["E8"].Value = "Port Of ";
            ws.Cells["F8"].Value = "Feeder";
            ws.Cells["G8"].Value = "BL Number";
            ws.Cells["H8"].Value = "Container No";
            ws.Cells["I8:K8"].Value = "Quantity";
            ws.Cells["I8:K8"].Merge = true;
            ws.Cells["L8"].Value = "Type";
            ws.Cells["M8"].Value = "Terms";
            ws.Cells["N8:O8"].Value = "Ocean Freight";
            ws.Cells["N8:O8"].Merge = true;
            ws.Cells["P8:Q8"].Value = "BAF/CAF/FAF";
            ws.Cells["P8:Q8"].Merge = true;
            ws.Cells["R8"].Value = "PCS / ISPS ";
            ws.Cells["S8"].Value = "Do Fees";
            ws.Cells["T8:U8"].Value = "Thc Collections";
            ws.Cells["T8:U8"].Merge = true;
            ws.Cells["V8"].Value = "Total";
            ws.Cells["W8"].Value = "Commission";
            ws.Cells["X8"].Value = " Thc / Ihc";
            ws.Cells["Y8"].Value = "Feeder";
            ws.Cells["Z8"].Value = " Other";
            ws.Cells["AA8"].Value = " TotalCost ";
            ws.Cells["AB8"].Value = " Nett Due ";

            ws.Cells["C9"].Value = " Date";
            ws.Cells["D9"].Value = "Loading";
            ws.Cells["E9"].Value = "Discharge";
            ws.Cells["F9"].Value = "Operator";


            ws.Cells["I9"].Value = " 20'";
            ws.Cells["J9"].Value = " 40'";
            ws.Cells["K9"].Value = "40' RF";
            ws.Cells["L9"].Value = "DG/GEN/MTY";
            ws.Cells["M9"].Value = "CFS/CY";
            ws.Cells["N9"].Value = "Prepaid";
            ws.Cells["O9"].Value = "Collected";
            ws.Cells["P9"].Value = "Prepaid";
            ws.Cells["Q9"].Value = "Collected";
            ws.Cells["R9"].Value = "Collected";
            ws.Cells["T9"].Value = "Prepaid";
            ws.Cells["U9"].Value = "Collected";
            ws.Cells["V9"].Value = "Revenue In";
            ws.Cells["W9"].Value = "2.50%";
            ws.Cells["X9"].Value = "Cost vs THC";
            ws.Cells["Y9"].Value = "Cost";
            ws.Cells["Z9"].Value = " Cost if any";
            ws.Cells["AA9"].Value = "In";
            ws.Cells["AB9"].Value = "To OCEANUS";
            ws.Cells["N10"].Value = "US $";
            ws.Cells["O10"].Value = "US $";
            ws.Cells["P10"].Value = "US $";
            ws.Cells["Q10"].Value = "US $";
            ws.Cells["R10"].Value = "US $";
            ws.Cells["T10"].Value = "US $";
            ws.Cells["U10"].Value = "US $";
            ws.Cells["V10"].Value = "US $";
            ws.Cells["AA10"].Value = "US $";
            ws.Cells["AB10"].Value = "USD";

            r = ws.Cells["A7:AB10"];
            r.Style.Font.Bold = true;

            r = ws.Cells["N7:V10"];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            r = ws.Cells["W7:AA10"];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);


            DataTable dtv = GetSOAImport(DtFrom, DtTo, AgencyID);

            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                ws.Cells["A" + rw].Value = sl2;
                ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["ArrivalDate"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["POL"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["POD"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["Size20"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["Size40"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["SizeRF40"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["Commodity"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["slotTerm"].ToString();

                ws.Cells["N" + rw].Value = dtv.Rows[i]["OFTPreAmt"];
                ws.Cells["O" + rw].Value = dtv.Rows[i]["OFTCollAmt"];
                ws.Cells["P" + rw].Value = dtv.Rows[i]["BAFPreAmt"];
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["BAFCollAmt"];
                ws.Cells["R" + rw].Value = dtv.Rows[i]["PCSAmt"];
                ws.Cells["S" + rw].Value = dtv.Rows[i]["DOFAmt"];
                ws.Cells["T" + rw].Value = dtv.Rows[i]["THCPreAmt"];
                ws.Cells["U" + rw].Value = dtv.Rows[i]["THCCollAmt"];
                ws.Cells["V" + rw].Formula = "=SUM(N" + rw + " : U" + rw + ")";
                ws.Cells["W" + rw].Value = dtv.Rows[i]["ICommAmt"];
                ws.Cells["X" + rw].Value = dtv.Rows[i]["THCCostAmt"];
                ws.Cells["Y" + rw].Value = dtv.Rows[i]["FeederCostAmt"];
                ws.Cells["Z" + rw].Value = "";
                ws.Cells["AA" + rw].Formula = "=SUM(W" + rw + " :Z" + rw + ")";
                ws.Cells["AB" + rw].Formula = "=SUM(N" + rw + " : U" + rw + ") - SUM(W" + rw + " : Z" + rw + ")";


                sl2++;
                rw += 1;
            }

            rw -= 1;

            ws.Cells["A7:AB" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AB" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AB" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AB" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;


            ws.Cells[1, 1, rw, 40].AutoFitColumns();
            #endregion

            #region 3rd SHEET Export
            ws = pck.Workbook.Worksheets.Add("Export");

            ws.Cells["A1"].Value = "Statement of Billing  :";
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A2"].Value = "For the month of      :";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A3"].Value = "Agent                 :";
            ws.Cells["A3"].Style.Font.Bold = true;
            ws.Cells["A4"].Value = "Location              :";
            ws.Cells["A4"].Style.Font.Bold = true;

            ws.Cells["A5"].Value = "Rate Of Exchange      :";
            ws.Cells["A5"].Style.Font.Bold = true;

            ws.Cells["B1"].Value = "Outward / Exports";
            ws.Cells["B1"].Style.Font.Bold = true;
            ws.Cells["B2"].Value = "";
            ws.Cells["B2"].Style.Font.Bold = true;
            ws.Cells["B3"].Value = "";
            ws.Cells["B3"].Style.Font.Bold = true;
            ws.Cells["B4"].Value = "";
            ws.Cells["B4"].Style.Font.Bold = true;

            ws.Cells["B5"].Value = "";
            ws.Cells["B5"].Style.Font.Bold = true;
            //Record Headers




            int sl3 = 1;

            rw = 11;
            ws.Cells["N7:V7"].Value = "REVENUE IN USD ";
            ws.Cells["N7:V7"].Merge = true;
            ws.Cells["N7:V7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["W7:AB7"].Value = "COST IN USD ";
            ws.Cells["W7:AB7"].Merge = true;
            ws.Cells["W7:AB7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


            ws.Cells["A8"].Value = "S. No.";
            ws.Cells["B8"].Value = "Vessel / Voy";
            ws.Cells["C8"].Value = "Sailing";
            ws.Cells["D8"].Value = "Port Of ";
            ws.Cells["E8"].Value = "Port Of ";
            ws.Cells["F8"].Value = "Feeder ";
            ws.Cells["G8"].Value = "BL Number";
            ws.Cells["H8"].Value = "Container No";
            ws.Cells["I8:K8"].Value = "Quantity";
            ws.Cells["I8:K8"].Merge = true;
            ws.Cells["L8"].Value = "Type";
            ws.Cells["M8"].Value = "Terms";
            ws.Cells["N8:O8"].Value = "Ocean Freight";
            ws.Cells["N8:O8"].Merge = true;
            ws.Cells["P8:Q8"].Value = "BAF/CAF/FAF";
            ws.Cells["P8:Q8"].Merge = true;
            ws.Cells["R8"].Value = "PCS / ISPS ";
            ws.Cells["S8"].Value = "Do Fees";
            ws.Cells["T8:U8"].Value = "Thc Collections";
            ws.Cells["T8:U8"].Merge = true;
            ws.Cells["V8"].Value = "Total";
            ws.Cells["W8"].Value = "Commission";
            ws.Cells["X8"].Value = "Logistics Fees";
            ws.Cells["Y8"].Value = " Thc / Ihc";
            ws.Cells["Z8"].Value = "Feeder ";
            ws.Cells["AA8"].Value = " Other";
            ws.Cells["AB8"].Value = " Total ";
            ws.Cells["AC8"].Value = " Nett Due ";


            ws.Cells["C9"].Value = " Date";
            ws.Cells["D9"].Value = "Loading";
            ws.Cells["E9"].Value = "Discharge";
            ws.Cells["F9"].Value = "Operator";


            ws.Cells["I9"].Value = " 20'";
            ws.Cells["J9"].Value = " 40'";
            ws.Cells["K9"].Value = "40' RF";
            ws.Cells["L9"].Value = "DG/GEN/MTY";
            ws.Cells["M9"].Value = "CFS/CY";
            ws.Cells["N9"].Value = "Prepaid";
            ws.Cells["O9"].Value = "Collected";
            ws.Cells["P9"].Value = "Prepaid";
            ws.Cells["Q9"].Value = "Collected";
            ws.Cells["R9"].Value = "Collected";
            ws.Cells["T9"].Value = "Prepaid";
            ws.Cells["U9"].Value = "Collected";
            ws.Cells["V9"].Value = "Revenue In";
            ws.Cells["W9"].Value = "2.50%";
            ws.Cells["X9"].Value = "@ USD 5.00";
            ws.Cells["Y9"].Value = "Cost vs THC";
            ws.Cells["Z9"].Value = "Cost";
            ws.Cells["AA9"].Value = " Cost if any";
            ws.Cells["AB9"].Value = "In";
            ws.Cells["AC9"].Value = "To OCEANUS";
            ws.Cells["N10"].Value = "US $";
            ws.Cells["O10"].Value = "US $";
            ws.Cells["P10"].Value = "US $";
            ws.Cells["Q10"].Value = "US $";
            ws.Cells["R10"].Value = "US $";
            ws.Cells["T10"].Value = "US $";
            ws.Cells["U10"].Value = "US $";
            ws.Cells["V10"].Value = "US $";
            ws.Cells["X9"].Value = "PER UNIT";
            ws.Cells["AA10"].Value = "US $";
            ws.Cells["AB10"].Value = "USD";

            r = ws.Cells["A7:AC10"];
            r.Style.Font.Bold = true;

            r = ws.Cells["N7:V10"];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            r = ws.Cells["W7:AB10"];
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);

            DataTable dtx = GetSOAExport(DtFrom, DtTo, AgencyID);
            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                ws.Cells["A" + rw].Value = sl3;
                ws.Cells["B" + rw].Value = dtx.Rows[i]["VesVoy"].ToString();
                ws.Cells["C" + rw].Value = dtx.Rows[i]["SailingDate"].ToString();
                ws.Cells["D" + rw].Value = dtx.Rows[i]["POL"].ToString();
                ws.Cells["E" + rw].Value = dtx.Rows[i]["POD"].ToString();
                ws.Cells["F" + rw].Value = dtx.Rows[i]["SlotOperator"].ToString();
                ws.Cells["G" + rw].Value = dtx.Rows[i]["BookingNo"].ToString();
                ws.Cells["H" + rw].Value = dtx.Rows[i]["CntrNo"].ToString();
                ws.Cells["I" + rw].Value = dtx.Rows[i]["Size20"].ToString();
                ws.Cells["J" + rw].Value = dtx.Rows[i]["Size40"].ToString();
                ws.Cells["K" + rw].Value = dtx.Rows[i]["SizeRF40"].ToString();
                ws.Cells["L" + rw].Value = dtx.Rows[i]["Commodity"].ToString();
                ws.Cells["M" + rw].Value = dtx.Rows[i]["slotTerm"].ToString();

                ws.Cells["N" + rw].Value = dtx.Rows[i]["OFTPreAmt"];
                ws.Cells["O" + rw].Value = dtx.Rows[i]["OFTCollAmt"];
                ws.Cells["P" + rw].Value = dtx.Rows[i]["BAFPreAmt"];
                ws.Cells["Q" + rw].Value = dtx.Rows[i]["BAFCollAmt"];
                ws.Cells["R" + rw].Value = dtx.Rows[i]["PCSISPSAmt"];
                ws.Cells["S" + rw].Value = dtx.Rows[i]["DOFAmt"];
                ws.Cells["T" + rw].Value = dtx.Rows[i]["THCPreAmt"];
                ws.Cells["U" + rw].Value = dtx.Rows[i]["THCCollAmt"];
                ws.Cells["V" + rw].Formula = "=SUM(N" + rw + " : U" + rw + ")";
                ws.Cells["W" + rw].Value = dtx.Rows[i]["ExCommAmt"];
                ws.Cells["X" + rw].Value = "5%";
                ws.Cells["Y" + rw].Value = dtx.Rows[i]["THCCostAmt"];
                ws.Cells["Z" + rw].Value = dtx.Rows[i]["FeederCostAmt"];
                ws.Cells["AA" + rw].Value = "";
                ws.Cells["AB" + rw].Formula = "=SUM(W" + rw + " :AA" + rw + ")";
                ws.Cells["AC" + rw].Formula = "=SUM(N" + rw + " : U" + rw + ") - SUM(W" + rw + " : AA" + rw + ")";
                sl3++;
                rw += 1;
            }


            rw -= 1;

            ws.Cells["A7:AC" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AC" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AC" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AC" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 40].AutoFitColumns();
            #endregion

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=SOAFullReport.xlsx");
            Response.End();

            // }

        }



        public DataTable GetSOAExport(string DtFrom, string DtTo, string AgencyID)
        {
            string strWhere = "";


            string _Query = " select distinct VesVoy,SailingDate,POL,POD,SlotOperator,BookingNo,CntrNo,Size20,Size40,SizeRF40,Commodity,slotTerm,cast(isnull((OFTPREPAID / (case when OFTPRECURR = 146 then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.OFTPRECURR)  end)),0) as decimal(10, 2)) as OFTPreAmt ,cast(isnull((OFTCOLLECT / (case when OFTCOLLCURR = 146  then 1 else (select top(1) ExRate from  " +
                            " NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.OFTCOLLCURR)  end)),0) as decimal(10, 2)) as OFTCollAmt, " +
                            " cast(isnull((BAFPrepaid / (case when BAFPreCurrID = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.BAFPreCurrID)  end)),0) as decimal(10, 2)) as BAFPreAmt , cast(isnull((BAFCollect / (case when BAFCollCurrID = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.BAFCollCurrID)  end)),0) as decimal(10, 2)) as BAFCollAmt, cast(isnull((SURChargesColl / (case when SURChargesCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.SURChargesCurr)  end)),0) as decimal(10, 2)) as PCSISPSAmt, cast(isnull((DOFCharge / (case when DOFCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.DOFCurr)  end)),0) as decimal(10, 2)) as DOFAmt, cast(isnull((THCPrepid / (case when THCPreCurr = 146  then 1 else (select top(1) ExRate from  " +
                            " NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.THCPreCurr)  end)),0) as decimal(10, 2)) as THCPreAmt1, " +
                            " cast(isnull(((select Revenu from View_LTHCAmountPort_Tariff  Tr where Tr.Id = NVO_V_SOAExportALLDataReport.Id) / (case when THCPreCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.THCPreCurr)  end)),0) as decimal(10,2)) as THCPreAmt," +
                            " cast(isnull((MUCPrepid / (case when MUCPreCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.MUCPreCurr)  end)),0) as decimal(10, 2)) as MUCAmt, " +
                            " cast(isnull((TOLLPrepid / (case when TOLLPreCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.TOLLPreCurr)  end)),0) as decimal(10, 2)) as TOLLAmt, " +
                            " cast(isnull((DGSPrepaid / (case when DGSPreCurrID = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.DGSPreCurrID)  end)),0) as decimal(10, 2)) as DGSAmt, " +
                            " cast(isnull((THCColl / (case when THCCollCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.THCCollCurr)  end)),0) as decimal(10, 2)) as THCCollAmt, " +
                            " cast(isnull((EXCOMM / (case when EXCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.EXCOMMCurrId)  end)),0) as decimal(10, 2)) as ExCommAmt, " +
                            " cast(isnull((THCCost / (case when THCCostCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.THCCostCurr)  end)),0) as decimal(10, 2)) as THCCostAmt1, " +
                            " cast(isnull(((select Cost from View_LTHCAmountPort_Tariff  Tr where Tr.Id = NVO_V_SOAExportALLDataReport.Id) / (case when THCPreCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAExportALLDataReport.THCPreCurr)  end)),0) as decimal(10,2)) as THCCostAmt," +
                            " cast(isnull((FeederCost / (case when FeederCostCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
                            " = NVO_V_SOAExportALLDataReport.FeederCostCurrId)  end)),0) as decimal(10, 2)) as FeederCostAmt,SlotAmt from  NVO_V_SOAExportALLDataReport  ";



            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,DtMovement,23) between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and convert(varchar,DtMovement,23) between '" + DtFrom + "' and '" + DtTo + "'";

            if (AgencyID != "" && AgencyID != "null" && AgencyID != "?" && AgencyID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where AgentID=" + AgencyID;
                else
                    strWhere += " and AgentID =" + AgencyID;
            if (strWhere == "")
                strWhere = _Query;
            return RegManage.GetViewData(strWhere, "");
        }

        public DataTable GetSOAImport(string DtFrom, string DtTo, string AgencyID)
        {
            string strWhere = "";


            string _Query = "select distinct VesVoy,ArrivalDate,POL,POD,DestinationAgent,POLAgency,SlotOperator,BookingNo,CntrNo,Size20,Size40,SizeRF40,Commodity,slotTerm,FinalDestination,cast(isnull((OFTPREPAID / (case when OFTPRECURR = 146 then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
             " = NVO_V_SOAImportALLDataReport.OFTPRECURR)  end)),0) as decimal(10, 2)) as OFTPreAmt ,cast(isnull((OFTCOLLECT / (case when OFTCOLLCURR = 146  then 1 else (select top(1) ExRate from  " +
             " NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.OFTCOLLCURR)  end)),0) as decimal(10, 2)) as OFTCollAmt, " +
           " cast(isnull((BAFPrepaid / (case when BAFPreCurrID = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
            " = NVO_V_SOAImportALLDataReport.BAFPreCurrID)  end)),0) as decimal(10, 2)) as BAFPreAmt , cast(isnull((BAFCollect / (case when BAFCollCurrID = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
           " = NVO_V_SOAImportALLDataReport.BAFCollCurrID)  end)),0) as decimal(10, 2)) as BAFCollAmt, cast(isnull((SURChargesColl / (case when SURChargesCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
          " = NVO_V_SOAImportALLDataReport.SURChargesCurr)  end)),0) as decimal(10, 2)) as PCSISPSAmt, cast(isnull((DOFCharge / (case when DOFCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
          " = NVO_V_SOAImportALLDataReport.DOFCurr)  end)),0) as decimal(10, 2)) as DOFAmt, cast(isnull((THCPrepid / (case when THCPreCurr = 146  then 1 else (select top(1) ExRate from  " +
          " NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.THCPreCurr)  end)),0) as decimal(10, 2)) as THCPreAmt, " +

       " cast(isnull((THCColl / (case when THCCollCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.THCCollCurr)  end)),0) as decimal(10, 2)) as THCCollAmt1, " +
       " cast(isnull(((select Revenu from View_DTHCImpAmountPort_Tariff  Tr where Tr.BLID = NVO_V_SOAImportALLDataReport.BLID) / (case when THCCollCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.THCCollCurr)  end)),0) as decimal(10,2)) as THCCollAmt, " +

        " cast(isnull((ISPSChargesColl / (case when ISPSChargesCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.ISPSChargesCurr)  end)),0) as decimal(10, 2)) as ISPSCollAmt, " +
        " cast(isnull((ISPSChargesColl / (case when ISPSChargesCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.ISPSChargesCurr)  end)),0) as decimal(10, 2)) as ISPSCollAmt, " +

          " cast(isnull((MUCColl / (case when MUCCollCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.MUCCollCurr)  end)),0) as decimal(10, 2)) as MUCAmt, " +
         " cast(isnull((TOLLColl / (case when TOLLCollCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.TOLLCollCurr)  end)),0) as decimal(10, 2)) as TOLLAmt, " +

        " cast(isnull((ICOMM / (case when ICOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.ICOMMCurrId)  end)),0) as decimal(10, 2)) as ICommAmt, " +
       " cast(isnull((THCCost / (case when THCCostCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
      " = NVO_V_SOAImportALLDataReport.THCCostCurr)  end)),0) as decimal(10, 2)) as THCCostAmt1, " +
      " cast(isnull(((select Cost from View_DTHCImpAmountPort_Tariff  Tr where Tr.BLID = NVO_V_SOAImportALLDataReport.BLID) / (case when THCCollCurr = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_SOAImportALLDataReport.THCCollCurr)  end)),0) as decimal(10,2)) as THCCostAmt, " +
      "  cast(isnull((FeederCost / (case when FeederCostCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id " +
      " = NVO_V_SOAImportALLDataReport.FeederCostCurrId)  end)),0) as decimal(10, 2)) as FeederCostAmt from NVO_V_SOAImportALLDataReport ";

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " where cast(DtMovement as Date) between '" + DtFrom + "' and '" + DtTo + "'";

            if (AgencyID != "" && AgencyID != "null" && AgencyID != "0" && AgencyID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where AgentID=" + AgencyID;
                else
                    strWhere += " and AgentID =" + AgencyID;
            if (strWhere == "")
                strWhere = _Query;

            return RegManage.GetViewData(strWhere, "");
        }


        public DataTable GetAgenctDetails(string AgencyID)
        {
            string strWhere = " select (select top(1) GeoLocation from NVO_GeoLocations where NVO_GeoLocations.Id = NVO_AgencyMaster.GeoLocationID)as GeoLocation,* from NVO_AgencyMaster where ID=" + AgencyID;
            return RegManage.GetViewData(strWhere, "");
        }


    }
}
