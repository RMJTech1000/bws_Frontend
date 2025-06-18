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

namespace NVOCShipping.Controllers
{
    public class EQCReportController : Controller
    {
        // GET: EQCReport
        MasterManager Manag = new MasterManager();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult LongStayReport()
        {
            return View();
        }
        public ActionResult DCMRLocationWise()
        {
            return View();
        }
        public ActionResult EQCLiftingReport()
        {
            return View();
        }

        public ActionResult DCMRGlobalReport()
        {
            return View();
        }
        public ActionResult EQCStatusWiseReport()
        {
            return View();
        }
        public ActionResult StockReport()
        {
            return View();
        }
        public ActionResult CntrTurnAroundReport()
        {
            return View();
        }
        public ActionResult EQCIdlingReport()
        {
            return View();
        }
        public ActionResult EQCAgewiseReport()
        {
            return View();
        }
        public ActionResult CntrHistory()
        {
            return View();
        }
        public ActionResult ImportDetentionReport()
        {
            return View();
        }

        public ActionResult ExportDetentionReport()
        {
            return View();
        }

        public ActionResult ExportDetentionSummaryReport()
        {
            return View();
        }

        public ActionResult ImportDetentionSummaryReport()
        {
            return View();
        }

        public ActionResult DetentionDemurageReport()
        {
            return View();
        }

        public ActionResult EQC_Detention_Dumarage_Report()
        {
            return View();
        }
        public ActionResult InventoryStatusReport()
        {
            return View();
        }

        public void DCMRGlobalExcelReport(string User)
        {
            DataTable dtv = GetDCMRGlobalExcelReport();
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                #region 1ST SHEET
                var ws = pck.Workbook.Worksheets.Add("DCMRGlobal");

                ws.Cells["A2"].Value = "DCMR Global Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:E2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A6:B6"].Value = "SUMMARY TYPE";
                ws.Cells["A6:B6"].Merge = true;
                ws.Cells["A6:B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["C6:H6"].Value = "IMPORT FULL";
                ws.Cells["C6:H6"].Merge = true;
                ws.Cells["C6:H6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["I6:N6"].Value = "WITH CNEE";
                ws.Cells["I6:N6"].Merge = true;
                ws.Cells["I6:N6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["O6:T6"].Value = "DAMAGE";
                ws.Cells["O6:T6"].Merge = true;
                ws.Cells["O6:T6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["U6:Z6"].Value = "AVAILABLE";
                ws.Cells["U6:Z6"].Merge = true;
                ws.Cells["U6:Z6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AA6:AF6"].Value = "WITH SHIPPER";
                ws.Cells["AA6:AF6"].Merge = true;
                ws.Cells["AA6:AF6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AG6:AL6"].Value = "EXP LDN";
                ws.Cells["AG6:AL6"].Merge = true;
                ws.Cells["AG6:AL6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AM6:AR6"].Value = "SAILED";
                ws.Cells["AM6:AR6"].Merge = true;
                ws.Cells["AM6:AR6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AS6:AX6"].Value = "TRANSHIPMENT";
                ws.Cells["AS6:AX6"].Merge = true;
                ws.Cells["AS6:AX6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AY6:BD6"].Value = "MT REPO";
                ws.Cells["AY6:BD6"].Merge = true;
                ws.Cells["AY6:BD6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                r = ws.Cells["A6:BF6"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);

                ws.Cells["A7"].Value = "AGENCY CODE";
                ws.Cells["B7"].Value = "GEOLOCATION";
                ws.Cells["C7"].Value = "LOCATIONS(Port)";

                ws.Cells["D7"].Value = "20'GP";
                ws.Cells["E7"].Value = "40'HC";
                ws.Cells["F7"].Value = "20'OT";
                ws.Cells["G7"].Value = "40'OT";
                ws.Cells["H7"].Value = "20'FR";
                ws.Cells["I7"].Value = "40'FR";

                ws.Cells["J7"].Value = "20'GP";
                ws.Cells["K7"].Value = "40'HC";
                ws.Cells["L7"].Value = "20'OT";
                ws.Cells["M7"].Value = "40'OT";
                ws.Cells["N7"].Value = "20'FR";
                ws.Cells["O7"].Value = "40'FR";


                ws.Cells["P7"].Value = "20'GP";
                ws.Cells["Q7"].Value = "40'HC";
                ws.Cells["R7"].Value = "20'OT";
                ws.Cells["S7"].Value = "40'OT";
                ws.Cells["T7"].Value = "20'FR";
                ws.Cells["U7"].Value = "40'FR";

                ws.Cells["V7"].Value = "20'GP";
                ws.Cells["W7"].Value = "40'HC";
                ws.Cells["X7"].Value = "20'OT";
                ws.Cells["Y7"].Value = "40'OT";
                ws.Cells["Z7"].Value = "20'FR";
                ws.Cells["AA7"].Value = "40'FR";

                ws.Cells["AB7"].Value = "20'GP";
                ws.Cells["AC7"].Value = "40'HC";
                ws.Cells["AD7"].Value = "20'OT";
                ws.Cells["AE7"].Value = "40'OT";
                ws.Cells["AF7"].Value = "20'FR";
                ws.Cells["AG7"].Value = "40'FR";

                ws.Cells["AH7"].Value = "20'GP";
                ws.Cells["AI7"].Value = "40'HC";
                ws.Cells["AJ7"].Value = "20'OT";
                ws.Cells["AK7"].Value = "40'OT";
                ws.Cells["AL7"].Value = "20'FR";
                ws.Cells["AM7"].Value = "40'FR";

                ws.Cells["AN7"].Value = "20'GP";
                ws.Cells["AO7"].Value = "40'HC";
                ws.Cells["AP7"].Value = "20'OT";
                ws.Cells["AQ7"].Value = "40'OT";
                ws.Cells["AR7"].Value = "20'FR";
                ws.Cells["AS7"].Value = "40'FR";

                ws.Cells["AT7"].Value = "20'GP";
                ws.Cells["AU7"].Value = "40'HC";
                ws.Cells["AV7"].Value = "20'OT";
                ws.Cells["AW7"].Value = "40'OT";
                ws.Cells["AX7"].Value = "20'FR";
                ws.Cells["AY7"].Value = "40'FR";

                ws.Cells["AZ7"].Value = "20'GP";
                ws.Cells["BA7"].Value = "40'HC";
                ws.Cells["BB7"].Value = "20'OT";
                ws.Cells["BC7"].Value = "40'OT";
                ws.Cells["BD7"].Value = "20'FR";
                ws.Cells["BE7"].Value = "40'FR";

                ws.Cells["BF7"].Value = "Grand Total";
                r = ws.Cells["A7:BF7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;
                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    ws.Cells["A" + rw].Value = dtv.Rows[i]["AgencyName"];
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["GeoLocationV"];
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["GeoLocation"];
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ImpFull20GP"];
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ImpFull40HC"];
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["ImpFull20OT"];
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["ImpFull40OT"];
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ImpFull20FR"];
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["ImpFull40FR"];

                    ws.Cells["J" + rw].Value = dtv.Rows[i]["DeStuff20GP"];
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["DeStuff40HC"];
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["DeStuff20OT"];
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["DeStuff40OT"];
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["DeStuff20FR"];
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["DeStuff40FR"];


                    ws.Cells["P" + rw].Value = dtv.Rows[i]["DmgUr20GP"];
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["DmgUr40HC"];
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["DmgUr20OT"];
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["DmgUr40OT"];
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["DmgUr20FR"];
                    ws.Cells["U" + rw].Value = dtv.Rows[i]["DmgUr40FR"];


                    ws.Cells["V" + rw].Value = dtv.Rows[i]["Avail20GP"];
                    ws.Cells["W" + rw].Value = dtv.Rows[i]["Avail40HC"];
                    ws.Cells["X" + rw].Value = dtv.Rows[i]["Avail20OT"];
                    ws.Cells["Y" + rw].Value = dtv.Rows[i]["Avail40OT"];
                    ws.Cells["Z" + rw].Value = dtv.Rows[i]["Avail20FR"];
                    ws.Cells["AA" + rw].Value = dtv.Rows[i]["Avail40FR"];


                    ws.Cells["AB" + rw].Value = dtv.Rows[i]["OutTo20GP"];
                    ws.Cells["AC" + rw].Value = dtv.Rows[i]["OutTo40HC"];
                    ws.Cells["AD" + rw].Value = dtv.Rows[i]["OutTo20OT"];
                    ws.Cells["AE" + rw].Value = dtv.Rows[i]["OutTo40OT"];
                    ws.Cells["AF" + rw].Value = dtv.Rows[i]["OutTo20FR"];
                    ws.Cells["AG" + rw].Value = dtv.Rows[i]["OutTo40FR"];

                    ws.Cells["AH" + rw].Value = dtv.Rows[i]["ExpFull20GP"];
                    ws.Cells["AI" + rw].Value = dtv.Rows[i]["ExpFull40HC"];
                    ws.Cells["AJ" + rw].Value = dtv.Rows[i]["ExpFull20OT"];
                    ws.Cells["AK" + rw].Value = dtv.Rows[i]["ExpFull40OT"];
                    ws.Cells["AL" + rw].Value = dtv.Rows[i]["ExpFull20FR"];
                    ws.Cells["AM" + rw].Value = dtv.Rows[i]["ExpFull40FR"];

      

                    ws.Cells["AN" + rw].Value = dtv.Rows[i]["MtyTransit20GP"];
                    ws.Cells["AO" + rw].Value = dtv.Rows[i]["MtyTransit40HC"];
                    ws.Cells["AP" + rw].Value = dtv.Rows[i]["MtyTransit20OT"];
                    ws.Cells["AQ" + rw].Value = dtv.Rows[i]["MtyTransit40OT"];
                    ws.Cells["AR" + rw].Value = dtv.Rows[i]["MtyTransit20FR"];
                    ws.Cells["AS" + rw].Value = dtv.Rows[i]["MtyTransit40FR"];


                    ws.Cells["AT" + rw].Value = dtv.Rows[i]["Trans20GP"];
                    ws.Cells["AU" + rw].Value = dtv.Rows[i]["Trans40HC"];
                    ws.Cells["AV" + rw].Value = dtv.Rows[i]["Trans20OT"];
                    ws.Cells["AW" + rw].Value = dtv.Rows[i]["Trans40OT"];
                    ws.Cells["AX" + rw].Value = dtv.Rows[i]["Trans20FR"];
                    ws.Cells["AY" + rw].Value = dtv.Rows[i]["Trans40FR"];

                    ws.Cells["AZ" + rw].Value = dtv.Rows[i]["MtyRepo20GP"];
                    ws.Cells["BA" + rw].Value = dtv.Rows[i]["MtyRepo40HC"];
                    ws.Cells["BB" + rw].Value = dtv.Rows[i]["MtyRepo20OT"];
                    ws.Cells["BC" + rw].Value = dtv.Rows[i]["MtyRepo40OT"];
                    ws.Cells["BD" + rw].Value = dtv.Rows[i]["MtyRepo20FR"];
                    ws.Cells["BE" + rw].Value = dtv.Rows[i]["MtyRepo40FR"];

                    ws.Cells["BF" + rw].Formula = "=SUM(D" + rw + " :BE" + rw + ")";
                    sl++;
                    rw += 1;
                }
                ws.Cells["BE" + rw].Value = "Total";
                r = ws.Cells["BE" + rw];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["BF" + rw].Formula = "=SUM(BF8:BF" + (rw-1)+")";
                r = ws.Cells["BF" + rw];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["BE" + rw + ":BF" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["BE" + rw + ":BF" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["BE" + rw + ":BF" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["BE" + rw + ":BF" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw -= 1;

                ws.Cells["A6:BF" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A6:BF" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A6:BF" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A6:BF" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 60].AutoFitColumns();
                #endregion


                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=DCMRGlobalReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetDCMRGlobalExcelReport()
        {

            string _Query = "select * from NVO_V_DCMRGlobalViewNew ";
            return Manag.GetViewData(_Query, "");
        }

        public void DCMRLocWiseExcelReport(string GeoLoc, string User)
        {
            DataTable dtv = GetDCMRLocWiseExcelReport(GeoLoc);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                #region 1ST SHEET
                var ws = pck.Workbook.Worksheets.Add("DCMRSummary");

                ws.Cells["A2"].Value = "DCMR Summary Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:F2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7:A8"].Value = "CURRENT STATUS";
                ws.Cells["A7:A8"].Merge = true;
                ws.Cells["A7:A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["B7:E7"].Value = "EQUIPMENT";
                ws.Cells["B7:E7"].Merge = true;
                ws.Cells["B7:E7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["B8"].Value = "20' GP";
                ws.Cells["C8"].Value = "20' HC";
                ws.Cells["D8"].Value = "40' GP";
                ws.Cells["E8"].Value = "40' HC";
                ws.Cells["F8"].Value = "TOTAL";
                r = ws.Cells["A7:F7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                r = ws.Cells["A8:F8"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 9;


                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["Description"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["DCMRLoc20GP"];
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["DCMRLoc20HC"];
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["DCMRLoc40GP"];
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DCMRLoc40HC"];
                    ws.Cells["F" + rw].Formula = "=SUM(" + ws.Cells["B" + rw] + ":" + ws.Cells["E" + rw] + ")";
                    sl++;
                    rw += 1;
                }

                rw -= 1;
                //ws.Cells["A17"].Value = "SUB TOTAL";
                //ws.Cells["A17"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //ws.Cells["A17"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                //ws.Cells["A18"].Value = "BOOKING (BKK)";
                ws.Cells["A19"].Value = "TOTAL";
                ws.Cells["A19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A19"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                r = ws.Cells["A9:A19"];
                r.Style.Font.Bold = true;

                ws.Cells["B19"].Formula = "=SUM(B9:B17)";
                ws.Cells["B19"].Style.Font.Bold = true;

                ws.Cells["C19"].Formula = "=SUM(C9:C17)";
                ws.Cells["C19"].Style.Font.Bold = true;

                ws.Cells["D19"].Formula = "=SUM(D9:D17)";
                ws.Cells["D19"].Style.Font.Bold = true;

                ws.Cells["E19"].Formula = "=SUM(E9:E17)";
                ws.Cells["E19"].Style.Font.Bold = true;

                ws.Cells["F19"].Formula = "=SUM(F9:F17)";
                ws.Cells["F19"].Style.Font.Bold = true;

                ws.Cells["A7:F" + 20].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + 20].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + 20].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + 20].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();
                #endregion

                #region 2ND SHEET
                ws = pck.Workbook.Worksheets.Add("DCMR");
                ws.Cells["A2"].Value = "DCMR Detail Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r = ws.Cells["A2:T2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers
                //Record Headers

                ws.Cells["A6:M6"].Value = "IMPORT";
                ws.Cells["A6:M6"].Merge = true;
                ws.Cells["A6:M6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["N6:T6"].Value = "EXPORT";
                ws.Cells["N6:T6"].Merge = true;
                ws.Cells["N6:T6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                r = ws.Cells["A6:M6"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                r = ws.Cells["N6:T6"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Container Type";
                ws.Cells["D7"].Value = "POL";
                ws.Cells["E7"].Value = "Status(Full/Mty)";
                ws.Cells["F7"].Value = "BLNO";
                ws.Cells["G7"].Value = "Vessel / Voyage";
                ws.Cells["H7"].Value = "Arrival Date(ETA)";
                ws.Cells["I7"].Value = "Landed Date(FV)";
                ws.Cells["J7"].Value = "Moved for Devanning (FU)";
                ws.Cells["K7"].Value = "Empty In (MA)";
                ws.Cells["L7"].Value = "Status";
                ws.Cells["M7"].Value = "Depot Name";
                ws.Cells["N7"].Value = "Empty Out to Shipper (MS)";
                ws.Cells["O7"].Value = "Laden In @ Terminal (FL)";
                ws.Cells["P7"].Value = "Vessel / Voyage";
                ws.Cells["Q7"].Value = "POL";
                ws.Cells["R7"].Value = "POD";
                ws.Cells["S7"].Value = "BLNO";
                ws.Cells["T7"].Value = "REMARKS";
                r = ws.Cells["A7:T7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int slno = 1;
                int row = 8;
                DataTable dtL = GetDCMRLocWiseDetailReport(GeoLoc);
                for (int i = 0; i < dtL.Rows.Count; i++)
                {
                    ws.Cells["A" + row].Value = slno;
                    ws.Cells["B" + row].Value = dtL.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + row].Value = dtL.Rows[i]["Size"].ToString();
                    ws.Cells["D" + row].Value = dtL.Rows[i]["POL"].ToString();
                    ws.Cells["E" + row].Value = "";
                    ws.Cells["F" + row].Value = dtL.Rows[i]["BookingNo"].ToString();
                    ws.Cells["G" + row].Value = dtL.Rows[i]["VesVoy"].ToString();
                    ws.Cells["H" + row].Value = "";
                    ws.Cells["I" + row].Value = "";
                    ws.Cells["J" + row].Value = "";
                    ws.Cells["K" + row].Value = "";
                    ws.Cells["L" + row].Value = "";
                    ws.Cells["M" + row].Value = "";
                    ws.Cells["N" + row].Value = "";
                    ws.Cells["O" + row].Value = "";
                    ws.Cells["P" + row].Value = "";
                    ws.Cells["Q" + row].Value = "";
                    ws.Cells["R" + row].Value = "";
                    ws.Cells["S" + row].Value = "";
                    ws.Cells["T" + row].Value = "";

                    slno++;
                    row += 1;
                }

                row -= 1;

                ws.Cells["A6:T" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A6:T" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A6:T" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A6:T" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, row, 24].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=DCMRLocationWiseReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetDCMRLocWiseExcelReport(string GeoLoc)
        {


            if (GeoLoc == "?" || GeoLoc == "null")
            {
                string _Query = " select Description,(select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 1 ) as DCMRLoc20GP, " +
                " (select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 10  ) as DCMRLoc20HC, " +
               " (select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 2 ) as DCMRLoc40GP, " +
             " (select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 3  )  as DCMRLoc40HC from NVO_DCMRLocwiseColumn DC ";
                return Manag.GetViewData(_Query, "");
            }
            else
            {
                string _Query = "select Description,(select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 1 AND(select top 1 ID from NVO_GeoLocations where Id = A.GeolocationID) = " + GeoLoc + ") as DCMRLoc20GP, " +
               " (select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 10 AND(select top 1 ID from NVO_GeoLocations where Id = A.GeolocationID) = " + GeoLoc + " )    as DCMRLoc20HC, " +
               " (select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 2 AND(select top 1 ID from NVO_GeoLocations where Id = A.GeolocationID) = " + GeoLoc + ") as DCMRLoc40GP, " +
              " (select COUNT(distinct C.ID) from NVO_Containers C INNER Join NVO_AgencyMaster  A on  A.ID = C.AgencyID WHERE StatusCode IN(DC.Statuscode,DC.Statuscode2,DC.Statuscode3) and TypeID = 3 AND(select top 1 ID from NVO_GeoLocations where Id = A.GeolocationID) = " + GeoLoc + ") as DCMRLoc40HC from NVO_DCMRLocwiseColumn DC";

                return Manag.GetViewData(_Query, "");

            }



        }

        public DataTable GetDCMRLocWiseDetailReport(string GeoLoc)
        {


            if (GeoLoc == "?" || GeoLoc == "null")
            {
                string _Query = "select C.ID,C.CntrNo,(Select top 1(Type + '-' + Size) from NVO_tblCntrTypes where ID = C.TypeID) As Size, c.StatusCode, isnull(B.BookingNo, '') as BookingNo, " +
                                " isnull((Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = B.VesVoyID)  ,'') As VesVoy, ISNULL(B.POL, '') POL " +
                                " from NVO_Containers C INNER Join NVO_AgencyMaster A on A.ID = C.AgencyID left outer join NVO_Booking B on B.ID = isnull((select top(1) BLNumber from NVO_ContainerTxns INNER JOIN NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber where " +
                                " ContainerID = C.ID order by DtMovement desc),0) WHERE StatusCode IN('FV', 'FVICD', 'FVCFS', 'DV', 'FU', 'MA', 'DL', 'UR', 'AV', 'MS', 'FC', 'FL', 'PL', 'MB', 'MV', 'TZ', 'TZFB','FB') and C.TypeID IN(1,2,3,10) ";

                return Manag.GetViewData(_Query, "");
            }
            else
            {
                string _Query = "select C.ID,C.CntrNo,(Select top 1(Type + '-' + Size) from NVO_tblCntrTypes where ID = C.TypeID) As Size, c.StatusCode, isnull(B.BookingNo, '') as BookingNo, " +
                  " isnull((Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = B.VesVoyID)  ,'')As VesVoy, ISNULL(B.POL, '') POL " +
                  " from NVO_Containers C INNER Join NVO_AgencyMaster A on A.ID = C.AgencyID left outer join NVO_Booking B on B.ID = isnull((select top(1) BLNumber from NVO_ContainerTxns INNER JOIN NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber where " +
                  " ContainerID = C.ID order by DtMovement desc),0) WHERE StatusCode IN('FV', 'FVICD', 'FVCFS', 'DV', 'FU', 'MA', 'DL', 'UR', 'AV', 'MS', 'FC', 'FL', 'PL', 'MB', 'MV', 'TZ', 'TZFB','FB')  and A.GeoLocationID = " + GeoLoc + " and C.TypeID IN(1,2,3,10) ";

                return Manag.GetViewData(_Query, "");

            }



        }

        public void LongStayExcelReport(string PortID, string StatusCode, string Type, string DaysV, string Days, string User)
        {
            DataTable dtv = GetLongStayExcelReport(PortID, StatusCode, Type, DaysV, Days);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("LongStayReport");

                ws.Cells["A2"].Value = "LongStay Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:M2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Type-Size";
                ws.Cells["D7"].Value = "GeoLocation";
                ws.Cells["E7"].Value = "Days";
                ws.Cells["F7"].Value = "DtMovement";
                ws.Cells["G7"].Value = "StatusCode";
                //ws.Cells["H7"].Value = "DtArrival";
                ws.Cells["H7"].Value = "BL Number";
                ws.Cells["I7"].Value = "BL Date";
                ws.Cells["J7"].Value = "Cargo";
                ws.Cells["K7"].Value = "POL";
                ws.Cells["L7"].Value = "POD";
                ws.Cells["M7"].Value = "LINE NO/ITEM No";
                ws.Cells["N7"].Value = "IGM NO";
                ws.Cells["O7"].Value = "IGM Date";
                ws.Cells["P7"].Value = "Shipper/ Address";
                ws.Cells["Q7"].Value = "Consignee/ Address";
                ws.Cells["R7"].Value = "Forwarder/ Address";
                ws.Cells["S7"].Value = "Nominated";
                ws.Cells["T7"].Value = "CurrentPort";
                r = ws.Cells["A7:T7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CntrType"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["GeoLoc"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Days"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["DtMovement"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["StatusCode"].ToString();
                    // ws.Cells["H" + rw].Value = "";
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["BkgDate"].ToString();
                    ws.Cells["J" + rw].Value = "";
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["M" + rw].Value = "";
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["IGMNo"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["IGMDate"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["Consignee"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["Notify"].ToString();
                    ws.Cells["S" + rw].Value = "";
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["CurrentPort"].ToString();

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:T" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:T" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:T" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:T" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=LongStayReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetLongStayExcelReport(string PortID, string StatusCode, string Type, string DaysV, string Days)
        {
            var symbol = "";

            if (DaysV == "1")
                symbol = "<";
            if (DaysV == "2")
                symbol = ">";
            if (DaysV == "3")
                symbol = "=";

            string _Query = " Select DISTINCT C.ID,C.CntrNo,CT.Type +'-'+ CT.Size as CntrType, Datediff(d, (select top(1) DtMovement   from NVO_ContainerTxns where ContainerID = C.ID order by NVO_ContainerTxns.ID  desc),GETDATE()) Days, isnull((select top 1 GeoLocation from NVO_GeoLocations Where id = NVO_AgencyMaster.GeoLocationID) ,'')AS GeoLoc, " +
           " isnull((select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by NVO_ContainerTxns.ID desc),'') StatusCode,isnull((select top(1) DtMovement from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc),'') DtMovement,BookingNo,BkgDate,POD,POL,isnull((select top 1 IGMNo from NVO_VoyageManifestDtls WHERE VoyageID = NVO_Booking.VesVoyID),'') As IGMNo, " +
          "convert(varchar, (select top 1 IGMDate from NVO_VoyageManifestDtls WHERE VoyageID = NVO_Booking.VesVoyID),103)  As IGMDate, ISNULL((select top 1 upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id INNER JOIN NVO_BOLCustomerDetails ON NVO_BOLCustomerDetails.BkgID " +
           " = NVO_Booking.ID  WHERE  NVO_CusBranchLocation.CID = NVO_BOLCustomerDetails.PartID  AND PartyTypeID = 1   ),'') As Shipper," +
          "ISNULL((select top 1 upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id INNER JOIN NVO_BOLCustomerDetails ON NVO_BOLCustomerDetails.BkgID  = NVO_Booking.ID  WHERE  NVO_CusBranchLocation.CID = NVO_BOLCustomerDetails.PartID  AND PartyTypeID = 2   ),'') As Consignee, " +
        "ISNULL((select top 1 upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster inner join NVO_CusBranchLocation  on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id INNER JOIN NVO_BOLCustomerDetails ON NVO_BOLCustomerDetails.BkgID   = NVO_Booking.ID  WHERE  NVO_CusBranchLocation.CID = NVO_BOLCustomerDetails.PartID  AND PartyTypeID = 3   ),'') As Notify, " +
       " isnull((select top 1 PARTYADDRESS   from NVO_BOLCustomerDetails where BkgId = NVO_Booking.ID AND PartyTypeID = 2),'') AS ConsigneeAddress, isnull((select top 1 PARTYADDRESS   from NVO_BOLCustomerDetails where BkgId = NVO_Booking.ID AND PartyTypeID = 1),'') AS ShipperAddress, " +
        "isnull((select top 1 PARTYADDRESS   from NVO_BOLCustomerDetails where BkgId = NVO_Booking.ID AND PartyTypeID = 3),'') AS NotifyAddress, isnull((select top 1 PortName   from NVO_PortMaster where ID = C.CurrentPortID ),'') AS CurrentPort " +
        " FROM NVO_Containers C INNER JOIN NVO_tblCntrTypes CT ON CT.ID = C.TypeID  left outer join NVO_AgencyMaster on NVO_AgencyMaster.ID = C.AgencyID left outer join NVO_Booking on NVO_Booking.ID = isnull((select top(1) BLNumber from NVO_ContainerTxns INNER JOIN NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber where  ContainerID = C.ID order by DtMovement desc),0) WHERE (select top(1) StatusCode  from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc)  NOT IN('PENDING') ";

            string strWhere = "";

            if (DaysV != "" && DaysV != "0" && DaysV != "null" && DaysV != "?" || Days != "" && Days != "undefined")

                if (strWhere == "")
                    strWhere += _Query + " and  Datediff(d, (select top(1) DtMovement from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc), GETDATE()) " + symbol + "" + Days;
                else
                    strWhere += " and  Datediff(d, (select top(1) DtMovement from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc), GETDATE())  " + symbol + "" + Days;


            if (StatusCode != "" && StatusCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) ='" + StatusCode + "'";
                else
                    strWhere += " and (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) ='" + StatusCode + "'";



            if (PortID != "" && PortID != "0" && PortID != "null" && PortID != "?")
                if (strWhere == "")
                    strWhere += _Query + " and CurrentPortID=" + PortID;
                else
                    strWhere += " and CurrentPortID =" + PortID;

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
        public void EQCStatusWiseExcelReport(string LocID, string StatusCode, string FromDt, string ToDt, string User)
        {
            DataTable dtv = GetEQCStatusWiseReport(LocID, StatusCode, FromDt, ToDt);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("EQCStatusWiseReport");

                ws.Cells["A2"].Value = "EQCStatusWise Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:K2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Downloaded Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;

                ws.Cells["A5"].Value = "From Date :";
                ws.Cells["A5"].Style.Font.Bold = true;
                ws.Cells["B5"].Value = FromDt;
                ws.Cells["B5"].Style.Font.Bold = true;
                ws.Cells["C5"].Value = "To Date :";
                ws.Cells["C5"].Style.Font.Bold = true;
                ws.Cells["D5"].Value = ToDt;
                ws.Cells["D5"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Type-Size";
                ws.Cells["D7"].Value = "StatusCode";
                ws.Cells["E7"].Value = "DtMovement";
                ws.Cells["F7"].Value = "Agency Name";
                ws.Cells["G7"].Value = "Location";
                ws.Cells["H7"].Value = "Vessel / Voyage";
                ws.Cells["I7"].Value = "Transit Mode";
                ws.Cells["J7"].Value = "Depot";
                ws.Cells["K7"].Value = "BLNumber";
                //ws.Cells["L7"].Value = "ORGIN";
                //ws.Cells["M7"].Value = "POL";
                //ws.Cells["N7"].Value = "POD";
                //ws.Cells["O7"].Value = "Final Destination";
                //ws.Cells["P7"].Value = "CRO NUMBER";
                //ws.Cells["Q7"].Value = "Reference Number";
                //ws.Cells["R7"].Value = "Customer";
                //ws.Cells["S7"].Value = "Vendor";
                //ws.Cells["T7"].Value = "CreatedBy";
                //ws.Cells["U7"].Value = "Created On";
                r = ws.Cells["A7:K7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["typesize"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusCode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DtMovement"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["location"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["VESVOY"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Transitmode"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Depot"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    //ws.Cells["L" + rw].Value = "";
                    //ws.Cells["M" + rw].Value = "";
                    //ws.Cells["N" + rw].Value = "";
                    //ws.Cells["O" + rw].Value = "";
                    //ws.Cells["P" + rw].Value = "";
                    //ws.Cells["Q" + rw].Value = "";
                    //ws.Cells["R" + rw].Value = "";
                    //ws.Cells["S" + rw].Value = "";
                    //ws.Cells["T" + rw].Value = "";
                    //ws.Cells["U" + rw].Value = "";

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:K" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=EQCStatusWiseReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetEQCStatusWiseReport(string LocID, string StatusCode, string FromDt, string ToDt)
        {


            string _Query = "select cmd1.ID, ContainerID, (select c.cntrno from NVO_Containers c where c.ID = cmd1.ContainerID) as cntrno, " +
        " (select c.TypeID from NVO_Containers c where c.ID = cmd1.ContainerID) as typeid,cmd1.DtMovement,cmd1.StatusCode, " +
     " NVO_PortMaster.PortName as location,(select typ.Type + '-' + typ.Size from NVO_Containers c inner join NVO_tblcntrtypes typ on typ.id = c.typeid where c.ID = cmd1.ContainerID) as typesize," +
       " isnull((Select top 1 AgencyName from NVO_AgencyMaster where ID = cmd1.agencyID ),'') As Agency," +
        "isnull((Select top 1 GeneralName from NVO_GeneralMaster where ID = cmd1.ModeOfTransportID),'' ) As Transitmode, " +
       " isnull((Select top 1 BookingNo from NVO_Booking where ID = cmd1.BLNumber),'' ) As BLNumber," +
       " isnull((Select top 1 DepName from NVO_DepotMaster where ID = cmd1.DepotID),'' ) As Depot , " +
       " isnull(( Select top 1 (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where cmd1.VesVoyID = V.ID Order by DtMovement desc)  ,'')As VESVOY " +
       " from NVO_ContainerTxns cmd1 inner join NVO_PortMaster on NVO_PortMaster.ID = cmd1.locationid ";

            string strWhere = "";

            if (LocID != "undefined" && LocID != "" && LocID != "0" && LocID != "null" && LocID != "?")

                if (strWhere == "")
                    strWhere += _Query + " WHERE cmd1.locationid=" + LocID;
                else
                    strWhere += " and cmd1.locationid = " + LocID;


            if (StatusCode != "" && StatusCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE cmd1.StatusCode ='" + StatusCode + "'";
                else
                    strWhere += " and cmd1.StatusCode ='" + StatusCode + "'";



            if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE  convert(varchar, cmd1.DtMovement , 23) between '" + FromDt + "' and '" + ToDt + "'";
                else
                    strWhere += "  and convert(varchar, cmd1.DtMovement , 23)  between '" + FromDt + "' and '" + ToDt + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        //public void EQCLiftingReportValues(string Type, string FromDt, string ToDt, string AgencyID, string User)
        //{


        //    ExcelPackage pck = new ExcelPackage();

        //    #region 1ST SHEET
        //    var ws = pck.Workbook.Worksheets.Add("SUMMARY");

        //    ws.Cells["A2"].Value = "EQCLifting Report List";
        //    ws.Cells["A2"].Style.Font.Bold = true;
        //    ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //    ExcelRange r = ws.Cells["A2:E2"];
        //    r.Merge = true;
        //    r.Style.Font.Size = 12;
        //    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

        //    ws.Cells["A4"].Value = "User :";
        //    ws.Cells["A4"].Style.Font.Bold = true;
        //    ws.Cells["B4"].Value = User;
        //    ws.Cells["B4"].Style.Font.Bold = true;
        //    ws.Cells["C4"].Value = "Downloaded Date :";
        //    ws.Cells["C4"].Style.Font.Bold = true;
        //    ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
        //    ws.Cells["D4"].Style.Font.Bold = true;

        //    //ws.Cells["A5"].Value = "From Date :";
        //    //ws.Cells["A5"].Style.Font.Bold = true;
        //    //ws.Cells["B5"].Value = FromDt;
        //    //ws.Cells["B5"].Style.Font.Bold = true;
        //    //ws.Cells["C5"].Value = "To Date :";
        //    //ws.Cells["C5"].Style.Font.Bold = true;
        //    //ws.Cells["D5"].Value = ToDt;
        //    //ws.Cells["D5"].Style.Font.Bold = true;


        //    ws.Cells["A6"].Value = "Shipment Type :";
        //    ws.Cells["A6"].Style.Font.Bold = true;
        //    var ShipmentType = "";

        //    if (Type == "1")
        //    {
        //        ShipmentType = "EXPORT";
        //    }
        //    if (Type == "2")
        //    {
        //        ShipmentType = "IMPORT";
        //    }
        //    ws.Cells["B6"].Value = ShipmentType;
        //    ws.Cells["B6"].Style.Font.Bold = true;
        //    //Record Headers

        //    ws.Cells["A8:A9"].Value = "AGENCY";
        //    ws.Cells["A8:A9"].Merge = true;
        //    ws.Cells["A8:A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //    ws.Cells["A8:A9"].AutoFitColumns();
        //    ws.Cells["B8:B9"].Value = "GEO LOCATION";
        //    ws.Cells["B8:B9"].Merge = true;
        //   // ws.Cells["B8:B9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //    ws.Cells["B8:B9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //    ws.Cells["B8:B9"].AutoFitColumns();
        //    ws.Cells["C8:E9"].Value = "CONTAINER COUNT";
        //    ws.Cells["C8:E9"].Merge = true;
        //    ws.Cells["C8:E9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        //    r = ws.Cells["A8:E8"];
        //    r.Style.Font.Bold = true;
        //    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

        //    ws.Cells["C10"].Value = "20";
        //    ws.Cells["D10"].Value = "40";
        //    ws.Cells["E10"].Value = "TOTAL";
        //    ws.Cells["C10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //    ws.Cells["D10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //    ws.Cells["E10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //    r = ws.Cells["A10:E10"];
        //    r.Style.Font.Bold = true;
        //    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //    r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


        //    int rw = 11;
        //    int Startrw = 11;
        //    int endrow = 0;
        //    DataTable dtvs = GetEQCLiftingReport(Type, FromDt, ToDt, AgencyID);
        //    if (dtvs.Rows.Count > 0)
        //    {
        //        for (int i = 0; i < dtvs.Rows.Count; i++)
        //        {
        //            ws.Cells["A" + rw].Value = dtvs.Rows[i]["AgencyName"].ToString();
        //            ws.Cells["B" + rw].Value = dtvs.Rows[i]["GeoLocation"].ToString();
        //            ws.Cells["C" + rw].Value = dtvs.Rows[i]["Lifting20GP"];
        //            ws.Cells["D" + rw].Value = dtvs.Rows[i]["Lifting40GP"];
        //            ws.Cells["E" + rw].Formula = "=SUM(" + ws.Cells["C" + rw] + ":" + ws.Cells["D" + rw] + ")";
        //            ws.Cells["A" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            ws.Cells["B" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            ws.Cells["C" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            ws.Cells["E" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            rw += 1;
        //        }
        //        endrow = rw;

        //        ws.Cells["B" + rw].Value = "TOTAL";
        //        ws.Cells["B" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        ws.Cells["B" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

        //        ws.Cells["C" + rw].Formula = "=sum(" + ws.Cells["C" + Startrw] + ":" + ws.Cells["C" + (rw-1)] + ")";
        //        ws.Cells["C" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        ws.Cells["C" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


        //        ws.Cells["D" + rw].Formula = "=sum(" + ws.Cells["D" + Startrw] + ":" + ws.Cells["D" + (rw - 1)] + ")";
        //        ws.Cells["D" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        ws.Cells["D" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

        //        ws.Cells["E" + rw].Formula = "=sum(" + ws.Cells["E" + Startrw] + ":" + ws.Cells["E" + (rw - 1)] + ")";
        //        ws.Cells["E" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        ws.Cells["E" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        //        ws.Cells["C" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //        ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //        ws.Cells["E" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //        ws.Cells["A8:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        //        ws.Cells["A8:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //        ws.Cells["A8:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        //        ws.Cells["A8:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

        //        ws.Cells[1, 1, rw, 24].AutoFitColumns();
        //        #endregion

        //        #region 2ND SHEET
        //        ws = pck.Workbook.Worksheets.Add("EQCLiftingDetail");
        //        ws.Cells["A2"].Value = "EQCLifting Detail Report List";
        //        ws.Cells["A2"].Style.Font.Bold = true;
        //        ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //        r = ws.Cells["A2:X2"];
        //        r.Merge = true;
        //        r.Style.Font.Size = 12;
        //        r.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

        //        ws.Cells["A4"].Value = "User :";
        //        ws.Cells["A4"].Style.Font.Bold = true;
        //        ws.Cells["B4"].Value = User;
        //        ws.Cells["B4"].Style.Font.Bold = true;
        //        ws.Cells["C4"].Value = "Date :";
        //        ws.Cells["C4"].Style.Font.Bold = true;
        //        ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
        //        ws.Cells["D4"].Style.Font.Bold = true;

        //        //Record Headers


        //        ws.Cells["A7"].Value = "S. No.";
        //        ws.Cells["B7"].Value = "Container No";
        //        ws.Cells["C7"].Value = "Container Type";
        //        ws.Cells["D7"].Value = "StatusCode";
        //        ws.Cells["E7"].Value = "Date Movement";
        //        ws.Cells["F7"].Value = "AgencyName";
        //        ws.Cells["G7"].Value = "Location";
        //        ws.Cells["H7"].Value = "Vessel / Voyage";
        //        ws.Cells["I7"].Value = "Transit";
        //        ws.Cells["J7"].Value = "Depot";
        //        ws.Cells["K7"].Value = "BLNumber";
        //        ws.Cells["L7"].Value = "Orgin";
        //        ws.Cells["M7"].Value = "POL";
        //        ws.Cells["N7"].Value = "POD";
        //        ws.Cells["O7"].Value = "Final Destination";
        //        ws.Cells["P7"].Value = "Container Owner";
        //        ws.Cells["Q7"].Value = "Leasing Partner";
        //        ws.Cells["R7"].Value = "Leasing Term";

        //        r = ws.Cells["A7:R7"];
        //        r.Style.Font.Bold = true;
        //        r.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

        //        int slno = 1;
        //        int row = 8;
        //        DataTable dtv = GetEQCLiftingDetailReport(Type, FromDt, ToDt, AgencyID);
        //        for (int i = 0; i < dtv.Rows.Count; i++)
        //        {
        //            ws.Cells["A" + row].Value = slno;
        //            ws.Cells["B" + row].Value = dtv.Rows[i]["CntrNo"].ToString();
        //            ws.Cells["C" + row].Value = dtv.Rows[i]["Size"].ToString();
        //            ws.Cells["D" + row].Value = dtv.Rows[i]["StatusCode"].ToString();
        //            ws.Cells["E" + row].Value = dtv.Rows[i]["LastDtMovement"].ToString();
        //            ws.Cells["F" + row].Value = dtv.Rows[i]["Agency"].ToString();
        //            ws.Cells["G" + row].Value = dtv.Rows[i]["CurrentPort"].ToString();
        //            ws.Cells["H" + row].Value = dtv.Rows[i]["VesVoy"].ToString();
        //            ws.Cells["I" + row].Value = dtv.Rows[i]["Transit"].ToString();
        //            ws.Cells["J" + row].Value = dtv.Rows[i]["CurrentDepot"].ToString();
        //            ws.Cells["K" + row].Value = dtv.Rows[i]["BookingNo"].ToString();
        //            ws.Cells["L" + row].Value = dtv.Rows[i]["POO"].ToString();
        //            ws.Cells["M" + row].Value = dtv.Rows[i]["POL"].ToString();
        //            ws.Cells["N" + row].Value = dtv.Rows[i]["POD"].ToString();
        //            ws.Cells["O" + row].Value = dtv.Rows[i]["FPOD"].ToString();
        //            ws.Cells["P" + row].Value = dtv.Rows[i]["CntrOwner"].ToString();
        //            ws.Cells["Q" + row].Value = dtv.Rows[i]["LeasePartner"].ToString();
        //            ws.Cells["R" + row].Value = dtv.Rows[i]["LeaseTerm"].ToString();

        //            slno++;
        //            row += 1;
        //        }

        //        row -= 1;


        //        ws.Cells["A7:R" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        //        ws.Cells["A7:R" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //        ws.Cells["A7:R" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        //        ws.Cells["A7:R" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        //        ws.Cells[1, 1, row, 24].AutoFitColumns();
        //        #endregion

        //        pck.SaveAs(Response.OutputStream);
        //        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        Response.AddHeader("content-disposition", "attachment;  filename=EQCLifting.xlsx");
        //        Response.End();

        //    }

        //}



        public void EQCLiftingReportValues(string Type, string FromDt, string ToDt, string AgencyID, string User)
        {

            var sw = 1;
            ExcelPackage pck = new ExcelPackage();

            #region 1ST SHEET
            var ws = pck.Workbook.Worksheets.Add("SUMMARY");

            ws.Cells["A2"].Value = "EQCLifting Report List";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:E2"];
            r.Merge = true;
            r.Style.Font.Size = 12;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            ws.Cells["A4"].Value = "User :";
            ws.Cells["A4"].Style.Font.Bold = true;
            ws.Cells["B4"].Value = User;
            ws.Cells["B4"].Style.Font.Bold = true;
            ws.Cells["C4"].Value = "Downloaded Date :";
            ws.Cells["C4"].Style.Font.Bold = true;
            ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
            ws.Cells["D4"].Style.Font.Bold = true;

            //ws.Cells["A5"].Value = "From Date :";
            //ws.Cells["A5"].Style.Font.Bold = true;
            //ws.Cells["B5"].Value = FromDt;
            //ws.Cells["B5"].Style.Font.Bold = true;
            //ws.Cells["C5"].Value = "To Date :";
            //ws.Cells["C5"].Style.Font.Bold = true;
            //ws.Cells["D5"].Value = ToDt;
            //ws.Cells["D5"].Style.Font.Bold = true;


            ws.Cells["A6"].Value = "Shipment Type :";
            ws.Cells["A6"].Style.Font.Bold = true;
            var ShipmentType = "";
            var PortDesc = "";

            if (Type == "1")
            {
                ShipmentType = "EXPORT";
                PortDesc = "POD";
            }
            if (Type == "2")
            {
                ShipmentType = "IMPORT";
                PortDesc = "POL";
            }
            ws.Cells["B6"].Value = ShipmentType;
            ws.Cells["B6"].Style.Font.Bold = true;
            //Record Headers

            ws.Cells["A8:A9"].Value = "AGENCY";
            ws.Cells["A8:A9"].Merge = true;
            ws.Cells["A8:A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A8:A9"].AutoFitColumns();

            ws.Cells["B8:B9"].Value = "GEO LOCATION";
            ws.Cells["B8:B9"].Merge = true;
            ws.Cells["B8:B9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["B8:B9"].AutoFitColumns();

            ws.Cells["C8:C9"].Value = PortDesc;
            ws.Cells["C8:C9"].Merge = true;
            ws.Cells["C8:C9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["C8:C9"].AutoFitColumns();

            ws.Cells["D8:F9"].Value = "CONTAINER COUNT";
            ws.Cells["D8:F9"].Merge = true;
            ws.Cells["D8:F9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            r = ws.Cells["A8:F8"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            ws.Cells["D10"].Value = "20";
            ws.Cells["E10"].Value = "40";
            ws.Cells["F10"].Value = "TOTAL";
            ws.Cells["D10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["E10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["F10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            r = ws.Cells["A10:F10"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


            int rw = 11;
            int Startrw = 11;
            int endrow = 0;
            DataTable dtvs = GetEQCLiftingReport(Type, FromDt, ToDt, AgencyID);
            if (dtvs.Rows.Count > 0)
            {
                for (int i = 0; i < dtvs.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = dtvs.Rows[i]["AgencyName"].ToString();
                    ws.Cells["B" + rw].Value = dtvs.Rows[i]["GeoLocation"].ToString();
                    if (Type == "1")
                        ws.Cells["C" + rw].Value = dtvs.Rows[i]["POD"].ToString();
                    else
                        ws.Cells["C" + rw].Value = dtvs.Rows[i]["POL"].ToString();

                    ws.Cells["D" + rw].Value = dtvs.Rows[i]["GP20"];
                    ws.Cells["E" + rw].Value = dtvs.Rows[i]["GP40"];
                    ws.Cells["F" + rw].Formula = "=SUM(" + ws.Cells["D" + rw] + ":" + ws.Cells["E" + rw] + ")";
                    ws.Cells["A" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["B" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["C" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["E" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["F" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    rw += 1;
                }
                endrow = rw;

                ws.Cells["C" + rw].Value = "TOTAL";
                ws.Cells["C" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["C" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["D" + rw].Formula = "=sum(" + ws.Cells["D" + Startrw] + ":" + ws.Cells["D" + (rw - 1)] + ")";
                ws.Cells["D" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["D" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


                ws.Cells["E" + rw].Formula = "=sum(" + ws.Cells["E" + Startrw] + ":" + ws.Cells["E" + (rw - 1)] + ")";
                ws.Cells["E" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["E" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["F" + rw].Formula = "=sum(" + ws.Cells["F" + Startrw] + ":" + ws.Cells["F" + (rw - 1)] + ")";
                ws.Cells["F" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["F" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["C" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["E" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A8:F" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:F" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:F" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:F" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();
                #endregion

                #region 2ND SHEET
                ws = pck.Workbook.Worksheets.Add("EQCLiftingDetail");
                ws.Cells["A2"].Value = "EQCLifting Detail Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r = ws.Cells["A2:X2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;

                //Record Headers


                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Container Type";
                ws.Cells["D7"].Value = "StatusCode";
                ws.Cells["E7"].Value = "Date Movement";
                ws.Cells["F7"].Value = "AgencyName";
                ws.Cells["G7"].Value = "Location";
                ws.Cells["H7"].Value = "Vessel / Voyage";
                ws.Cells["I7"].Value = "Transit";
                ws.Cells["J7"].Value = "Depot";
                ws.Cells["K7"].Value = "BLNumber";
                ws.Cells["L7"].Value = "Orgin";
                ws.Cells["M7"].Value = "POL";
                ws.Cells["N7"].Value = "POD";
                ws.Cells["O7"].Value = "Final Destination";
                ws.Cells["P7"].Value = "Container Owner";
                ws.Cells["Q7"].Value = "Leasing Partner";
                ws.Cells["R7"].Value = "Leasing Term";

                r = ws.Cells["A7:R7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int slno = 1;
                int row = 8;
                DataTable dtv = GetEQCLiftingDetailReport(Type, FromDt, ToDt, AgencyID);
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + row].Value = slno;
                    ws.Cells["B" + row].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + row].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["D" + row].Value = dtv.Rows[i]["StatusCode"].ToString();
                    ws.Cells["E" + row].Value = dtv.Rows[i]["DtMovement"].ToString();
                    ws.Cells["F" + row].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["G" + row].Value = dtv.Rows[i]["GeoLocation"].ToString();
                    ws.Cells["H" + row].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["I" + row].Value = dtv.Rows[i]["ModeOfTransport"].ToString();
                    ws.Cells["J" + row].Value = dtv.Rows[i]["Depot"].ToString();
                    ws.Cells["K" + row].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["L" + row].Value = dtv.Rows[i]["POO"].ToString();
                    ws.Cells["M" + row].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["N" + row].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["O" + row].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["P" + row].Value = dtv.Rows[i]["CntrOwner"].ToString();
                    ws.Cells["Q" + row].Value = dtv.Rows[i]["LeasePartner"].ToString();
                    ws.Cells["R" + row].Value = dtv.Rows[i]["LeaseTerm"].ToString();

                    slno++;
                    row += 1;
                }

                row -= 1;


                ws.Cells["A7:R" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:R" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:R" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:R" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, row, 24].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=EQCLifting_" + DateTime.Now.Date + "_.xlsx");
                Response.End();



            }

        }

        public DataTable GetEQCLiftingReport(string Type, string FromDt, string ToDt, string AgencyID)
        {
            string fromdate = ""; string ToDate = "";
            if (FromDt != "")
                fromdate = (DateTime.Parse(FromDt)).ToString("MM/dd/yyyy");
            if (ToDt != "")
                ToDate = (DateTime.Parse(ToDt)).ToString("MM/dd/yyyy");
            string strWhere = ""; string _Query = "";


            if (Type == "1")
            {

                _Query = " select distinct POD,AgencyName,GeoLocation,PODID,GeoLocationID, " +
                            " (select count(CntrNo)  from v_CntrStatusCount CV " +
                            " where CV.AgencyID = v_CntrStatusCount.AgencyID and CV.StatusCode = v_CntrStatusCount.StatusCode and CV.PODID = v_CntrStatusCount.PODID " +
                            " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "' " +
                            " and CV.GeoLocationID = v_CntrStatusCount.GeoLocationID and CV.TypeId in((select ID from NVO_tblCntrTypes where Teus = 1))) GP20, " +
                            " (select count(CntrNo)  from v_CntrStatusCount CV " +
                            " where CV.AgencyID = v_CntrStatusCount.AgencyID and CV.StatusCode = v_CntrStatusCount.StatusCode and CV.PODID = v_CntrStatusCount.PODID " +
                            " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "' " +
                            " and CV.GeoLocationID = v_CntrStatusCount.GeoLocationID and CV.TypeId in((select ID from NVO_tblCntrTypes where Teus = 2))) GP40 " +
                            " from v_CntrStatusCount " +
                            " where StatusCode = 'FB'";


                if (AgencyID != "" && AgencyID != "0" && AgencyID != "null" && AgencyID != "?")

                    if (strWhere == "")
                        strWhere += _Query + " and AgencyID=" + AgencyID;
                    else
                        strWhere += " and AgencyID = " + AgencyID;

                if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                    if (strWhere == "")
                        strWhere += _Query + " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";
                    else
                        strWhere += "  and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";


                if (strWhere == "")
                    strWhere = _Query;
                return Manag.GetViewData(strWhere, "");
            }
            else
            {
                _Query = " select distinct POL,AgencyName,GeoLocation,POLID,GeoLocationID, " +
                           " (select count(CntrNo)  from v_CntrStatusCount CV " +
                           " where CV.AgencyID = v_CntrStatusCount.AgencyID and CV.StatusCode = v_CntrStatusCount.StatusCode and CV.POLID = v_CntrStatusCount.POLID " +
                           " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "' " +
                           " and CV.GeoLocationID = v_CntrStatusCount.GeoLocationID and CV.TypeId in((select ID from NVO_tblCntrTypes where Teus = 1))) GP20, " +
                           " (select count(CntrNo)  from v_CntrStatusCount CV " +
                           " where CV.AgencyID = v_CntrStatusCount.AgencyID and CV.StatusCode = v_CntrStatusCount.StatusCode and CV.POLID = v_CntrStatusCount.POLID " +
                           " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "' " +
                           " and CV.GeoLocationID = v_CntrStatusCount.GeoLocationID and CV.TypeId in((select ID from NVO_tblCntrTypes where Teus = 2))) GP40 " +
                           " from v_CntrStatusCount " +
                           " where StatusCode = 'FV'";


                if (AgencyID != "" && AgencyID != "0" && AgencyID != "null" && AgencyID != "?")

                    if (strWhere == "")
                        strWhere += _Query + " and AgencyID=" + AgencyID;
                    else
                        strWhere += " and AgencyID = " + AgencyID;

                if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                    if (strWhere == "")
                        strWhere += _Query + " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";
                    else
                        strWhere += "  and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";


                if (strWhere == "")
                    strWhere = _Query;
                return Manag.GetViewData(strWhere, "");

            }



        }
        public DataTable GetEQCLiftingDetailReport(string Type, string FromDt, string ToDt, string AgencyID)
        {
            string fromdate = ""; string ToDate = "";
            if (FromDt != "")
                fromdate = (DateTime.Parse(FromDt)).ToString("MM/dd/yyyy");
            if (ToDt != "")
                ToDate = (DateTime.Parse(ToDt)).ToString("MM/dd/yyyy");
            string _Query = "";
            string strWhere = "";

            if (Type == "1")
            {
                _Query = " select CntrNo,(select top(1) Size from NVO_tblCntrTypes where Id=v_CntrStatusCount.TypeId) as size,StatusCode, " +
                               " convert(varchar, DtMovement, 103) as DtMovement,AgencyName,GeoLocation,VesVoy, ModeOfTransport,Depot,BookingNo,POO, " +
                               " POL,POD,FPOD,CntrOwner,LeasePartner,LeaseTerm from " +
                               " v_CntrStatusCount " +
                               " where StatusCode = 'FB'";
                if (strWhere == "")
                    strWhere = _Query;

                if (AgencyID != "" && AgencyID != "0" && AgencyID != "null" && AgencyID != "?")

                    if (strWhere == "")
                        strWhere += _Query + " and AgencyID=" + AgencyID;
                    else
                        strWhere += " and AgencyID = " + AgencyID;

                if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                    if (strWhere == "")
                        strWhere += _Query + " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";
                    else
                        strWhere += "  and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";

                return Manag.GetViewData(strWhere, "");
            }
            else
            {
                _Query = " select CntrNo,(select top(1) Size from NVO_tblCntrTypes where Id=v_CntrStatusCount.TypeId) as size,StatusCode, " +
                            " convert(varchar, DtMovement, 103) as DtMovement,AgencyName,GeoLocation,VesVoy, ModeOfTransport,Depot,BookingNo,POO, " +
                            " POL,POD,FPOD,CntrOwner,LeasePartner,LeaseTerm from " +
                            " v_CntrStatusCount " +
                            " where StatusCode = 'FV'";
                if (strWhere == "")
                    strWhere = _Query;

                if (AgencyID != "" && AgencyID != "0" && AgencyID != "null" && AgencyID != "?")

                    if (strWhere == "")
                        strWhere += _Query + " and AgencyID=" + AgencyID;
                    else
                        strWhere += " and AgencyID = " + AgencyID;

                if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                    if (strWhere == "")
                        strWhere += _Query + " and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";
                    else
                        strWhere += "  and CONVERT(DATE, DtMovement) between '" + fromdate + "' and '" + ToDate + "'";


                return Manag.GetViewData(strWhere, "");
            }

        }

        public void StockCategoryReport(string DateV, string User)
        {
            DataTable dtv = GetStockCategorySummaryReport(DateV);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                #region 1ST SHEET
                var ws = pck.Workbook.Worksheets.Add("StockCategory");

                ws.Cells["A2"].Value = "StockCategory Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:E2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;

                //ws.Cells["C4"].Value = "Downloaded Date :";
                //ws.Cells["C4"].Style.Font.Bold = true;
                //ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                //ws.Cells["D4"].Style.Font.Bold = true;
                if (DateV == "undefined")
                {
                    ws.Cells["A5"].Value = "Stock Category Date :";
                    ws.Cells["A5"].Style.Font.Bold = true;
                    ws.Cells["B5"].Value = System.DateTime.Today.Date.ToShortDateString();
                }
                else
                {
                    ws.Cells["A5"].Value = "Stock Category Date :";
                    ws.Cells["A5"].Style.Font.Bold = true;
                    ws.Cells["B5"].Value = DateV;
                    ws.Cells["B5"].Style.Font.Bold = true;
                }


                //Record Headers

                ws.Cells["A8:A10"].Value = "COUNTRY";
                ws.Cells["A8:A10"].Merge = true;
                ws.Cells["A8:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["A8:A10"].AutoFitColumns();
                ws.Cells["B8:B10"].Value = "PORTS";
                ws.Cells["B8:B10"].Merge = true;
                ws.Cells["B8:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["B8:B10"].AutoFitColumns();


                ws.Cells["C8:J8"].Value = "OFF-HIRE UNITS SUMMARY";
                ws.Cells["C8:J8"].Merge = true;
                ws.Cells["C8:J8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["K8:R8"].Value = "ACTIVE UNITS SUMMARY";
                ws.Cells["K8:R8"].Merge = true;
                ws.Cells["K8:R8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["C9:D9"].Value = "DRY";
                ws.Cells["C9:D9"].Merge = true;
                ws.Cells["C9:D9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["E9:F9"].Value = "Reefers";
                ws.Cells["E9:F9"].Merge = true;
                ws.Cells["E9:F9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["G9:H9"].Value = "Flat Rack";
                ws.Cells["G9:H9"].Merge = true;
                ws.Cells["G9:H9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["I9:J9"].Value = "Open Top";
                ws.Cells["I9:J9"].Merge = true;
                ws.Cells["I9:J9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["K9:L9"].Value = "DRY";
                ws.Cells["K9:L9"].Merge = true;
                ws.Cells["K9:L9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["M9:N9"].Value = "Reefers";
                ws.Cells["M9:N9"].Merge = true;
                ws.Cells["M9:N9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["O9:P9"].Value = "Flat Rack";
                ws.Cells["O9:P9"].Merge = true;
                ws.Cells["O9:P9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["Q9:R9"].Value = "Open Top";
                ws.Cells["Q9:R9"].Merge = true;
                ws.Cells["Q9:R9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["C10"].Value = "20'";
                ws.Cells["D10"].Value = "40'";
                ws.Cells["E10"].Value = "20";
                ws.Cells["F10"].Value = "40";
                ws.Cells["G10"].Value = "20";
                ws.Cells["H10"].Value = "40";
                ws.Cells["I10"].Value = "20";
                ws.Cells["J10"].Value = "40";
                ws.Cells["K10"].Value = "20";
                ws.Cells["L10"].Value = "40";
                ws.Cells["M10"].Value = "20";
                ws.Cells["N10"].Value = "40";
                ws.Cells["O10"].Value = "20";
                ws.Cells["P10"].Value = "40";
                ws.Cells["Q10"].Value = "20";
                ws.Cells["R10"].Value = "40";

                r = ws.Cells["A8:B10"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                r = ws.Cells["A8:R8"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                r = ws.Cells["C9:J10"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightPink);

                r = ws.Cells["K9:R10"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);



                int rw = 11;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["Country"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Ports"].ToString();

                    ws.Cells["C" + rw].Value = dtv.Rows[i]["OFFDRY20GP"];
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["OFFDRY40GP"];
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["OFFREF20GP"];
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["OFFREF40GP"];
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["OFFFLAT20GP"];
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["OFFFLAT40GP"];
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["OFFOT20GP"];
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["OFFOT40GP"];

                    ws.Cells["K" + rw].Value = dtv.Rows[i]["ACTDRY20GP"];
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["ACTDRY40GP"];
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["ACTREF20GP"];
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["ACTREF40GP"];
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["ACTFLAT20GP"];
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["ACTFLAT40GP"];
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["ACTOT20GP"];
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["ACTOT40GP"];

                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A8:R" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:R" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:R" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:R" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();
                #endregion

                #region 2ND SHEET
                ws = pck.Workbook.Worksheets.Add("StockBreakUp");
                ws.Cells["A2"].Value = "Stock Category Detail Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r = ws.Cells["A2:X2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;

                //Record Headers


                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Container Type";
                ws.Cells["D7"].Value = "Last StatusCode";
                ws.Cells["E7"].Value = "Agency";
                ws.Cells["F7"].Value = "Date Movement";
                ws.Cells["G7"].Value = "Container Ownership";
                ws.Cells["H7"].Value = "Lease Term";
                ws.Cells["I7"].Value = "Leasing Partner";

                r = ws.Cells["A7:I7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int slno = 1;
                int row = 8;

                DataTable dtL = GetStockCategoryReport(DateV);
                for (int i = 0; i < dtL.Rows.Count; i++)
                {
                    ws.Cells["A" + row].Value = slno;
                    ws.Cells["B" + row].Value = dtL.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + row].Value = dtL.Rows[i]["CntrType"].ToString();
                    ws.Cells["D" + row].Value = dtL.Rows[i]["StatusCode"].ToString();
                    ws.Cells["E" + row].Value = dtL.Rows[i]["AgencyName"].ToString();
                    ws.Cells["F" + row].Value = dtL.Rows[i]["DtMovement"].ToString();
                    ws.Cells["G" + row].Value = dtL.Rows[i]["BoxOwner"].ToString();
                    ws.Cells["H" + row].Value = dtL.Rows[i]["LeaseTerm"].ToString();
                    ws.Cells["I" + row].Value = dtL.Rows[i]["LeasingPartner"].ToString();

                    slno++;
                    row += 1;
                }

                row -= 1;

                ws.Cells["A7:I" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, row, 24].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=StockCategoryReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetStockCategorySummaryReport(string DateV)
        {
            string _Query = "Select * from NVO_V_STOCKReport";
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetStockCategoryReport(string DateV)
        {


            string _Query = " Select DISTINCT C.ID,C.CntrNo,CT.Type +'-'+ CT.Size as CntrType, Datediff(d, (select top(1) DtMovement   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc),GETDATE()) Days,  " +
             " (select top(1) DtMovement   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) DtMovement, (select top(1) A.AgencyName from NVO_AgencyMaster A inner join NVO_ContainerTxns ct on ct.containerID = C.ID where A.ID = ct.AgencyID order by ct.DtMovement desc) AgencyName , " +
             " (select top(1) P.PortName from NVO_PortMaster P inner join NVO_ContainerTxns ct on ct.containerID = C.ID where P.ID = ct.NextPortID order by ct.DtMovement desc) LOCATION , " +
             " C.StatusCode,  (select top(1) GeneralName from NVO_GeneralMaster where ID = C.LeaseTermID ) LeaseTerm, " +
             " (select top 1 CustomerName  from NVO_view_CustomerDetails Where CID = c.BoxOwnerID) BoxOwner, (select top 1 upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster " +
            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where NVO_CusBranchLocation.CID = c.LeasingPartnerID) LeasingPartner  FROM NVO_Containers C  INNER Join NVO_PortMaster P on P.ID = C.CurrentPortID " +
            " INNER JOIN NVO_tblCntrTypes CT ON CT.ID = C.TypeID  WHERE C.StatusCode NOT IN('PENDING') and isnull(C.CurrentPortID,0 )!= 0 ";

            string strWhere = "";

            //if (PortID != "" && PortID != "0" && PortID != "null" && PortID != "?")

            //    if (strWhere == "")
            //        strWhere += _Query + " and (select top(1) NextPortID from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc)=" + PortID;
            //    else
            //        strWhere += " and (select top(1) NextPortID from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) = " + PortID;


            //if (StatusCode != "" && StatusCode != "undefined")
            //    if (strWhere == "")
            //        strWhere += _Query + " and (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) ='" + StatusCode + "'";
            //    else
            //        strWhere += " and (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) ='" + StatusCode + "'";



            //if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
            //    if (strWhere == "")
            //        strWhere += _Query + " and (select top(1) DtMovement from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) between '" + FromDt + "' and '" + ToDt + "'";
            //    else
            //        strWhere += "  and (select top(1) DtMovement from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) between '" + FromDt + "' and '" + ToDt + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public void EQCTurnAroundReport(string FromDt, string ToDt, string StatusCode, string User, string St, string GeoLoc)
        {
            DataTable Dt = GetEQCTurnAroundReport(FromDt, ToDt, StatusCode, GeoLoc);
            if (Dt.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TurnAroundReport");

                ws.Cells["A2"].Value = "TurnAround Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:O2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Downloaded Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;

                ws.Cells["F4"].Value = "StatusCode";
                ws.Cells["F4"].Style.Font.Bold = true;
                ws.Cells["G4"].Value = St;
                ws.Cells["G4"].Style.Font.Bold = true;

                ws.Cells["A5"].Value = "From Date :";
                ws.Cells["A5"].Style.Font.Bold = true;
                ws.Cells["B5"].Value = FromDt;
                ws.Cells["B5"].Style.Font.Bold = true;
                ws.Cells["C5"].Value = "To Date :";
                ws.Cells["C5"].Style.Font.Bold = true;
                ws.Cells["D5"].Value = ToDt;
                ws.Cells["D5"].Style.Font.Bold = true;
                //Record Headers

                //Record Headers
                ws.Cells["A7"].Value = "S.No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Size";
                ws.Cells["D7"].Value = "From Status";
                ws.Cells["E7"].Value = "From Location";
                ws.Cells["F7"].Value = "From Geo Location";
                ws.Cells["G7"].Value = "From Agency";
                ws.Cells["H7"].Value = "Date";
                ws.Cells["I7"].Value = "To Status";
                ws.Cells["J7"].Value = "To Location";
                ws.Cells["K7"].Value = "To Geo Location";
                ws.Cells["L7"].Value = "To Agency";
                ws.Cells["M7"].Value = "Date";
                ws.Cells["N7"].Value = "Days";
                ws.Cells["O7"].Value = "BLNUMBER";
                ws.Cells["A7:O7"].Style.Font.Bold = true;

                r = ws.Cells["A7:O7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;
                int rw_b = rw;
                int days;
                for (int i = 0; i < Dt.Rows.Count; i++)
                {
                    DateTime dtTrans = new DateTime();


                    ws.Cells["A" + rw].Value = (i + 1);
                    ws.Cells["B" + rw].Value = Dt.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = Dt.Rows[i]["Size"].ToString();
                    ws.Cells["D" + rw].Value = Dt.Rows[i]["FromStatus"].ToString();
                    ws.Cells["E" + rw].Value = Dt.Rows[i]["FromLocation"].ToString();
                    ws.Cells["F" + rw].Value = Dt.Rows[i]["FromGeoLocation"].ToString();
                    ws.Cells["G" + rw].Value = Dt.Rows[i]["FromAgency"].ToString();


                    if (Dt.Rows[i]["FromDate"].ToString() == "")
                    {
                        ws.Cells["H" + rw].Value = "";
                    }
                    else
                    {
                        DateTime.TryParse(Dt.Rows[i]["FromDate"].ToString(), out dtTrans);
                        ws.Cells["H" + rw].Value = dtTrans;
                        ws.Cells["H" + rw].Style.Numberformat.Format = "dd/MM/yyyy";
                    }

                    ws.Cells["H" + rw].Style.Numberformat.Format = "dd/MM/yyyy";
                    ws.Cells["I" + rw].Value = Dt.Rows[i]["ToStatus"].ToString();
                    ws.Cells["J" + rw].Value = Dt.Rows[i]["ToLocation"].ToString();
                    ws.Cells["K" + rw].Value = Dt.Rows[i]["ToGeoLocation"].ToString();
                    ws.Cells["L" + rw].Value = Dt.Rows[i]["ToAgency"].ToString();
                    if (Dt.Rows[i]["ToDate"].ToString() == "")
                    {
                        ws.Cells["M" + rw].Value = "";
                    }
                    else
                    {
                        DateTime.TryParse(Dt.Rows[i]["ToDate"].ToString(), out dtTrans);
                        ws.Cells["M" + rw].Value = dtTrans;
                        ws.Cells["M" + rw].Style.Numberformat.Format = "dd/MM/yyyy";
                    }

                    //ws.Cells["K" + rw].Value = DtTo;
                    //ws.Cells["K" + rw].Style.Numberformat.Format = "dd/MM/yyyy";


                    int.TryParse(Dt.Rows[i]["NDays"].ToString(), out days);
                    ws.Cells["N" + rw].Value = days;
                    ws.Cells["O" + rw].Value = Dt.Rows[i]["BookingNo"].ToString();
                    rw++;
                }

                rw -= 1;

                ws.Cells["A7:O" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TurnAroundReport.xlsx");
                Response.End();

            }


        }


        public DataTable GetEQCTurnAroundReport(string FromDt, string ToDt, string StatusCode, string GeoLoc)
        {
            string FromStatus = "", ToStatus = "", ToStatus2 = "", WC = "";
            //PortTypeID = " tbllocs.PortTypeID IN(1,5) AND ";

            string strWhere = "";

            switch (StatusCode)
            {
                //--CASE 1---
                case "238":
                    FromStatus = "'FV'";
                    ToStatus = "'FU'";
                    break;
                //--CASE 2---
                case "243":
                    FromStatus = "'FV'";
                    ToStatus = "'MA'";
                    break;
                //--CASE 3---
                case "239":
                    FromStatus = "'FU'";
                    ToStatus = "'MA'";
                    // PortTypeID = "";
                    break;
                //--CASE 4---
                case "244":
                    FromStatus = "'FV'";
                    ToStatus = "'DL'";
                    break;
                //--CASE 5---
                case "245":
                    FromStatus = "'MA'";
                    ToStatus = "'FB'";
                    //  PortTypeID = "";
                    WC += " AND CTTO.ID NOT IN (SELECT TOP 1 ID FROM NVO_Containertxns D WHERE D.ContainerID=CTTO.ContainerID AND D.Statuscode IN('ST','FB')  AND D.ID>(SELECT TOP 1 ID FROM NVO_Containertxns B WHERE B.Statuscode IN('TZ') AND B.ContainerID=CTTO.ContainerID) ORDER BY D.DtMovement) ";
                    break;
                //--CASE 6---
                case "246":
                    FromStatus = "'MA'";
                    ToStatus = "'MB'";
                    break;
                //--CASE 7---
                case "241":
                    FromStatus = "'MA'";
                    ToStatus = "'MS'";
                    // PortTypeID = "";
                    break;
                //--CASE 8---
                case "242":
                    FromStatus = "'MS'";
                    ToStatus = "'FL'";
                    //PortTypeID = "";//To
                    break;
                //--CASE 9---
                case "247":
                    FromStatus = "'FL'";
                    ToStatus = "'FB'";
                    // PortTypeID = "A.PortTypeID=5 AND ";
                    WC += " AND CTTO.ID NOT IN (SELECT TOP 1 ID FROM NVO_Containertxns D WHERE D.ContainerID=CTTO.ContainerID AND D.Statuscode IN ('ST','FB')  AND D.ID>(SELECT TOP 1 ID FROM NVO_Containertxns B WHERE B.Statuscode IN('TZ') AND B.ContainerID=CTTO.ContainerID) ORDER BY D.DtMovement) ";
                    break;
                //--CASE 10---
                case "10":
                    FromStatus = "'MB'";
                    ToStatus = "'MA'";
                    /// PortTypeID = "A.PortTypeID=5 AND ";
                    break;
                //--CASE 11---
                case "11":
                    FromStatus = "'FV'";
                    ToStatus = "'FI'";
                    break;
                //--CASE 12---
                case "12":
                    FromStatus = "'MA'";
                    ToStatus = "'MI'";
                    //   PortTypeID = "";
                    break;
                //--CASE 13---
                case "13":
                    FromStatus = "'MS'";
                    ToStatus = "'MA'";
                    // PortTypeID = "";
                    break;
                //--CASE 14---
                case "14":
                    FromStatus = "'FB'";
                    ToStatus = "'FV'";
                    //PortTypeID = "A.PortTypeID=5 AND ";
                    break;
                //--CASE 15---
                case "15":
                    FromStatus = "'FV'";
                    ToStatus = "'RC'";
                    //   PortTypeID = "A.PortTypeID=5 AND ";
                    // PortTypeID = "";

                    break;
                //--CASE 16---
                case "16":
                    FromStatus = "'RC'";
                    ToStatus = "'ST','FB'";
                    // PortTypeID = "";
                    ///  PortTypeID = "A.PortTypeID=5 AND ";
                    break;
                //--CASE 17---
                case "17":
                    FromStatus = "'MS'";
                    ToStatus = "'FB'";
                    ToStatus2 = "'MA'";
                    // PortTypeID = "";
                    WC += " AND CTTO.ID NOT IN (SELECT TOP 1 ID FROM NVO_Containertxns D WHERE D.ContainerID=CTTO.ContainerID AND D.Statuscode IN ('ST','FB')  AND D.ID>(SELECT TOP 1 ID FROM NVO_Containertxns B WHERE B.Statuscode IN('TZ') AND B.ContainerID=CTTO.ContainerID) ORDER BY D.DtMovement) ";
                    break;

                //--CASE 18---
                case "240":
                    FromStatus = "'DL'";
                    ToStatus = "'MA'";
                    //   PortTypeID = "A.PortTypeID=5 AND ";
                    // PortTypeID = "";
                    break;
            }

            string Qr = "  SELECT DISTINCT C.CntrNo,NVO_tblCntrTypes.Size, FmMove.StatusCode FromStatus, FmMove.DtMovement FromDate,BookingNo, " +
            " FLoc.PortName as FromLocation, FLoc.GeoLocation as FromGeoLocation, FLoc.AgencyName As FromAgency, CTTO.Statuscode ToStatus, AG.AgencyName As ToAgency,CTTO.Dtmovement ToDate, " +
            " DateDiff(dd, FmMove.DtMovement, CTTO.Dtmovement)+1 NDays, t1.PortName as ToLocation, togl.GeoLocation as ToGeoLocation  FROM NVO_Containers C " +
            " INNER JOIN NVO_tblCntrTypes ON NVO_tblCntrTypes.ID = C.TypeID INNER JOIN NVO_ContainerTxns CTTO ON CTTO.ContainerID = C.ID " +
            " INNER JOIN NVO_PortMaster t1 ON t1.ID = CTTO.LocationID INNER JOIN NVO_GeoLocations togl on togl.ID = t1.GeoLocID " +
            "   INNER JOIN NVO_AgencyMaster AG ON AG.ID = CTTO.AgencyID " +
            " OUTER APPLY(SELECT TOP 1 NVO_PortMaster.PortName, gl.GeoLocation,AG1.AgencyName FROM NVO_ContainerTxns A " +
            " INNER JOIN NVO_PortMaster ON NVO_PortMaster.ID = A.LocationID INNER JOIN NVO_AgencyMaster AG1 ON AG1.ID = A.AgencyID INNER JOIN NVO_GeoLocations gl on gl.ID = NVO_PortMaster.GeoLocID WHERE " +
            " A.StatusCode IN (" + FromStatus + ")  AND A.ContainerID = C.ID  AND A.DtMovement <= CTTO.DtMovement AND NVO_PortMaster.CountryID = t1.CountryID ORDER BY A.DtMovement Desc ) as FLoc " +
            " OUTER APPLY (SELECT TOP 1 DtMovement, StatusCode,NVO_Booking.BookingNo FROM NVO_ContainerTxns A left outer JOIN NVO_Booking ON NVO_Booking.ID = A.BLNumber  INNER JOIN NVO_PortMaster ON NVO_PortMaster.ID = A.LocationID WHERE A.StatusCode IN (" + FromStatus + ")  AND A.ContainerID = C.ID " +
            " AND A.DtMovement <= CTTO.DtMovement AND NVO_PortMaster.CountryID = t1.CountryID ORDER BY  A.DtMovement Desc ) as FmMove  WHERE CTTO.StatusCode IN (" + ToStatus + ")  ";

            if (GeoLoc != "" && GeoLoc != "0" && GeoLoc != "null" && GeoLoc != "?")
                if (strWhere == "")
                    strWhere += Qr + " and togl.ID=" + GeoLoc;
                else
                    strWhere += " and togl.ID =" + GeoLoc;

            if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                if (strWhere == "")
                    strWhere += Qr + " and convert(varchar, CTTO.DtMovement, 23)  between '" + FromDt + "' AND '" + ToDt + "' " + WC +
                        " ORDER BY 6,4 ";
                else
                    strWhere += " and convert(varchar, CTTO.DtMovement, 23) between '" + FromDt + "' AND '" + ToDt + "' " + WC +
                        " ORDER BY 6,4 ";

            string stat = StatusCode;

            //if (stat == "239" || stat == "242" || stat == "247" || stat == "10" || stat == "14" || stat == "16")
            //{
            //    Qr = " SELECT DISTINCT C.CntrNo,NVO_tblCntrTypes.Size,  CTTO.Statuscode FromStatus, " +
            //       " CTTO.Dtmovement FromDate, t1.PortName as FromLocation, togl.GeoLocation as FromGeoLocation,  " +
            //     " ToMove.StatusCode ToStatus, ToMove.DtMovement ToDate, ToMove.Name as ToLocation,ToMove.GeoLocation as ToGeoLocation,  DateDiff(dd, CTTO.Dtmovement, ToMove.DtMovement) + 1 NDays " +
            //    " FROM NVO_Containers C INNER JOIN NVO_tblCntrTypes ON NVO_tblCntrTypes.ID = C.TypeID  INNER JOIN NVO_ContainerTxns CTTO ON CTTO.ContainerID = C.ID " +
            //   " INNER JOIN NVO_PortMaster t1 ON t1.ID = CTTO.LocationID INNER JOIN NVO_GeoLocations togl on togl.ID = t1.GeoLocID " +
            //  " OUTER APPLY(SELECT TOP 1 DtMovement,StatusCode, PortName as Name, GeoLocation FROM NVO_ContainerTxns A  WHERE A.StatusCode IN " +
            //   " (" + ToStatus + ") AND A.ContainerID = C.ID  AND A.DtMovement >= CTTO.DtMovement  ORDER BY A.DtMovement)  as ToMove  WHERE CTTO.StatusCode IN(" + FromStatus + ") ";


            //    if (GeoLoc != "" && GeoLoc != "0" && GeoLoc != "null" && GeoLoc != "?")
            //        if (strWhere == "")
            //            strWhere += Qr + " and togl.ID=" + GeoLoc;
            //        else
            //            strWhere += " and togl.ID =" + GeoLoc;

            //    if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
            //        if (strWhere == "")
            //            strWhere += Qr + " and CTTO.DtMovement between '" + FromDt + "' AND '" + ToDt + "' " + WC +
            //                " ORDER BY 6,4 ";
            //        else
            //            strWhere += " and CTTO.DtMovement between '" + FromDt + "' AND '" + ToDt + "' " + WC +
            //                " ORDER BY 6,4 ";
            //}
            //else if (stat == "17")
            //{
            //    Qr = "SELECT DISTINCT C.CntrNo,tblCntrTypes.Size, " +
            //        " CTTO.Statuscode FromStatus, CTTO.Dtmovement FromDate, " +
            //        " t1.Name as FromLocation, togl.LocDesc as FromGeoLocation, " +
            //        " ToMove.StatusCode ToStatus, ToMove.DtMovement ToDate, " +
            //        " ToMove.Name as ToLocation, ToMove.LocDesc as ToGeoLocation, " +
            //        " DateDiff(dd, CTTO.Dtmovement, ToMove.DtMovement) + 1 NDays " +
            //        " FROM Containers C " +
            //        " INNER JOIN tblCntrTypes ON tblCntrTypes.ID = C.TypeID " +
            //        " INNER JOIN ContainerTxns CTTO ON CTTO.ContainerID = C.ID " +
            //        " INNER JOIN tblLocs t1 ON t1.ID = CTTO.LocationID " +
            //        " INNER JOIN tblGeoLocations togl on togl.ID = t1.GeoLocID " +
            //        " OUTER APPLY " +
            //        " (SELECT TOP 1 DtMovement, StatusCode, Location as Name, LocDesc FROM vContainerTxns A " +
            //        " WHERE " + PortTypeID + " A.StatusCode IN (" + ToStatus + ") AND A.ContainerID = C.ID  AND A.DtMovement >= CTTO.DtMovement " +
            //        //--AND tblLocs.CtryID = t1.CtryID " +
            //        " ORDER BY A.DtMovement) as ToMove " +
            //        " OUTER APPLY " +
            //        " (SELECT TOP 1 DtMovement, StatusCode, Location as Name, LocDesc FROM vContainerTxns B " +
            //        " WHERE " + PortTypeID + " B.StatusCode IN (" + ToStatus2 + ") AND B.ContainerID = C.ID  AND B.DtMovement >= CTTO.DtMovement " +
            //        " ORDER BY B.DtMovement) as ToMove2 " +
            //        " WHERE CTTO.StatusCode IN (" + FromStatus + ") AND ToMove.DtMovement between '" + DtFrom + "' AND '" + DtTo + "' " + WC +
            //        " ORDER BY 6,4";
            //}





            if (strWhere == "")
                strWhere = Qr;


            return Manag.GetViewData(strWhere, "");
        }

        public void InventoryStatusReportData(string User)
        {
            DataTable dtv = GetWeeklyEQCReport();
            if (dtv.Rows.Count > 0)
            {


                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("EQCGlobalReport");

                ws.Cells["A2"].Value = "EQC Global Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:J2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                //Record Headers

                ws.Cells["A7"].Value = "Statuscode";
                ws.Cells["B7"].Value = "Status";
                ws.Cells["C7"].Value = "Agency";
                ws.Cells["D7"].Value = "Dry-20GP";
                ws.Cells["E7"].Value = "Dry-20HC";
                ws.Cells["F7"].Value = "Dry-40GP";
                ws.Cells["G7"].Value = "Dry-40HC";
                ws.Cells["H7"].Value = "Reefer-20RF";
                ws.Cells["I7"].Value = "Reefer-40RF";
                ws.Cells["J7"].Value = "Grand Total";

                r = ws.Cells["A7:J7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;
                int frowid = 0;
                int rw = 8;
                int rw_end = 0;
                int rw_bgn_sub = rw;
                string StatusCodes = "";
                StringBuilder sbSub = new StringBuilder();
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    if (StatusCodes.Trim() != dtv.Rows[i]["StatusCode"].ToString().Trim())
                    {

                        if (StatusCodes != "" && i > 0)
                        {
                            if (dtv.Rows.Count > 0)
                            {
                                rw++;
                                rw_end = rw - 1;
                                sbSub.AppendLine(rw_end.ToString());

                                r = ws.Cells["A" + rw_end + ":C" + rw_end];
                                r.Value = "Sub Total - " + StatusCodes;
                                r.Merge = true;
                                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                r = ws.Cells["A" + rw_end + ":J" + rw_end];

                                ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Merge = true;
                                ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                                //End Region merge

                                //Subtotal Begin 13-sept-2019 rgs
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Font.Bold = true;
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Font.Size = 9;
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Font.Color.SetColor(Color.Black);
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                ws.Cells["D" + rw_end].Formula = string.Format("=SUM(D{0}:D{1})", rw_bgn_sub, (rw_end - 1));
                                ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub, (rw_end - 1));
                                ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub, (rw_end - 1));
                                ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub, (rw_end - 1));
                                ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub, (rw_end - 1));
                                ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub, (rw_end - 1));
                                ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub, (rw_end - 1));
                                //frowid = rw;

                                // r = ws.Cells["A" + rw_bgn_sub + ":J" + rw_bgn_sub];
                                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                r.Style.Fill.BackgroundColor.SetColor(Color.LightGray);


                            }

                        }
                        StatusCodes = dtv.Rows[i]["StatusCode"].ToString();
                        ws.Cells["A" + rw].Value = StatusCodes;
                        rw_bgn_sub = rw;
                    }


                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["StatusCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Dry20GP"];
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Dry20HC"];
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Dry40GP"];
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Dry40HC"];
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["Reefer20RF"];
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Reefer40RF"];
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["GrandTotal"];
                    rw++;

                    ws.Cells[1, 1, rw, 29].AutoFitColumns();
                    ws.Cells["A7:J" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A7:J" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A7:J" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A7:J" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                #region Final sub-total

                if (StatusCodes.Trim() != "")
                {
                    if (StatusCodes != "")
                    {
                        rw++;
                        rw_end = rw - 1;
                        sbSub.AppendLine(rw_end.ToString());
                        if (dtv.Rows.Count > 0)
                        {
                            r = ws.Cells["A" + rw_end + ":C" + rw_end];
                            r.Value = "Sub Total - " + StatusCodes;
                            r.Merge = true;
                            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            r = ws.Cells["A" + rw_end + ":J" + rw_end];

                            ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                            //End Region merge

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":J" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            ws.Cells["D" + rw_end].Formula = string.Format("=SUM(D{0}:D{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub, (rw_end - 1));
                            //frowid = rw;

                            // r = ws.Cells["A" + rw_bgn_sub + ":J" + rw_bgn_sub];
                            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            r.Style.Fill.BackgroundColor.SetColor(Color.LightGray);


                        }

                    }


                }


                #endregion



                rw_end = rw - 1;

                #region foooter
                r = ws.Cells["A" + rw + ":J" + rw];
                r.Style.Font.Bold = true;
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                string row_nos = "";
                string[] sub_rows = sbSub.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                foreach (var ln in sub_rows)
                {
                    if (ln != "")
                        row_nos += "+D" + ln;
                }

                //Total 13-sep-2019 rgs
                r = ws.Cells["A" + rw + ":C" + rw];
                r.Value = "GRAND TOTAL";
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r.Merge = true;

                ws.Cells["A" + rw + ":J" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":J" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":J" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":J" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":J" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":J" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // ws.Cells["C" + rw].Value = "TOTAL";

                ws.Cells["D" + rw].Formula = "=" + row_nos;
                ws.Cells["E" + rw].Formula = "=" + row_nos.Replace("D", "E");
                ws.Cells["F" + rw].Formula = "=" + row_nos.Replace("D", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos.Replace("D", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos.Replace("D", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos.Replace("D", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos.Replace("D", "J");
                r = ws.Cells["D" + rw + ":J" + rw];
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=EQCGlobalStatusReport.xlsx");
                Response.End();

            }

        }


     
        public DataTable GetWeeklyEQCReport()
        {
            string _Query = " Select Distinct CS.ID,SeqNoGlobal,CS.StatusCode,A.AgencyName, " +
                   " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID = 1 AND A.ID = AgencyID) as Dry20GP, " +
                   " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID = 10 AND A.ID = AgencyID )  as Dry20HC,  " +
                   " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID = 2  AND A.ID = AgencyID )  as Dry40GP,  " +
                   " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID = 3 AND A.ID = AgencyID )  as Dry40HC,  " +
                    " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID = 8 AND A.ID = AgencyID )  as Reefer20RF,  " +
                   " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID = 9 AND A.ID = AgencyID  )  as Reefer40RF, " +
                   " (select COUNT(distinct ID) from NVO_Containers WHERE StatusCode = CS.StatusCode and TypeID IN(1, 2, 3, 8, 9, 10) AND A.ID = AgencyID )  as GrandTotal " +
                   " from NVO_ContainerStatusCodes CS LEFT OUTER Join NVO_Containers C on C.StatusCode = CS.StatusCode " +
                   " INNER Join NVO_AgencyMaster A on A.ID = C.AgencyID WHERE ISNULL(SeqNoGlobal,0)!= 0 ORDER BY SeqNoGlobal,CS.StatusCode,CS.ID ";

            return Manag.GetViewData(_Query, "");
        }
    }
}