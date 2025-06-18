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
    public class ExportBookingReportController : Controller
    {
        // GET: ExportBookingReport
        ExportReportManager RegManage = new ExportReportManager();
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult ExportBookingSummaryReport(string DtFrom, string DtTo)
        {
            ExportBookingSummaryReportv(DtFrom, DtTo);
            return View();
        }

        public void ExportBookingSummaryReportv(string DtFrom, string ToDate)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("Booking Summary Report");

            ws.Cells["A2"].Value = "BOOKING SUMMARY REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:S2"];
            r.Merge = true;
            r.Style.Font.Size = 12;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            ws.Cells["A4"].Value = "User :";
            ws.Cells["A4"].Style.Font.Bold = true;
            ws.Cells["B4"].Value = "";
            ws.Cells["B4"].Style.Font.Bold = true;
            ws.Cells["C4"].Value = "Date :";
            ws.Cells["C4"].Style.Font.Bold = true;
            ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
            ws.Cells["D4"].Style.Font.Bold = true;
            //Record Headers

            int RowIndex = 6;
            Color colum1 = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
            ws.Cells[RowIndex, 1, RowIndex, 30].Value = "DESC";
            ws.Cells[RowIndex, 1, RowIndex, 30].Merge = true;
            ws.Cells[RowIndex, 1, RowIndex, 30].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 1, RowIndex, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 1, RowIndex, 30].Style.Fill.BackgroundColor.SetColor(colum1);


            Color Colum2 = System.Drawing.ColorTranslator.FromHtml("#AC9EF4");
            ws.Cells[RowIndex, 31, RowIndex, 35].Value = "Other Expenses";
            ws.Cells[RowIndex, 31, RowIndex, 35].Merge = true;
            ws.Cells[RowIndex, 31, RowIndex, 35].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 31, RowIndex, 35].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 31, RowIndex, 35].Style.Fill.BackgroundColor.SetColor(Colum2);

            Color Colum3 = System.Drawing.ColorTranslator.FromHtml("#F7E3E3");
            ws.Cells[RowIndex, 36, RowIndex, 38].Value = "";
            ws.Cells[RowIndex, 36, RowIndex, 38].Merge = true;
            ws.Cells[RowIndex, 36, RowIndex, 38].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 36, RowIndex, 38].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 36, RowIndex, 38].Style.Fill.BackgroundColor.SetColor(Colum3);


            ws.Cells["A7"].Value = "S. No.";
            ws.Cells["B7"].Value = "WEEK";
            ws.Cells["C7"].Value = "PRIN";
            ws.Cells["D7"].Value = "SOC/COC";
            ws.Cells["E7"].Value = "ST";
            ws.Cells["F7"].Value = "BL.";
            ws.Cells["G7"].Value = "CUSTOMER NAME";
            ws.Cells["H7"].Value = "TERMINAL";
            ws.Cells["I7"].Value = "POL";
            ws.Cells["J7"].Value = "POD";
            ws.Cells["K7"].Value = "FPOD";
            ws.Cells["L7"].Value = "VESSEL NAME";
            ws.Cells["M7"].Value = "SCN";
            ws.Cells["N7"].Value = "SLOT OPR";
            ws.Cells["O7"].Value = "ETA";
            ws.Cells["P7"].Value = "ETD";
            ws.Cells["Q7"].Value = "CARGO TYPE";
            ws.Cells["R7"].Value = "WGT (MT)";
            ws.Cells["S7"].Value = "CONT SIZE";
            ws.Cells["T7"].Value = "QTY OF CONT";

            ws.Cells["U7"].Value = "OFRT";

            ws.Cells["V7"].Value = "SLOT";

            ws.Cells["W7"].Value = "COLLECTION MODE";
            ws.Cells["X7"].Value = "BC RELEASED DATE";
            ws.Cells["Y7"].Value = "BL TYPE (TELEX/OBL)";
            ws.Cells["Z7"].Value = "LOCAL CHARGES";
            ws.Cells["AA7"].Value = "INVOICE NO";
            ws.Cells["AB7"].Value = "INVOICE DATE";
            ws.Cells["AC7"].Value = "JOB STATUS";

            ws.Cells["AD7"].Value = "SSR";
            ws.Cells["AE7"].Value = "ITT";
            ws.Cells["AF7"].Value = "STORAGE";
            ws.Cells["AG7"].Value = "REMOVAL";
            ws.Cells["AH7"].Value = "REEFER MON";
            ws.Cells["AI7"].Value = "VENDOR NAME";
            ws.Cells["AJ7"].Value = "OTHER COST";
            ws.Cells["AK7"].Value = "JOB HANDLED BY";




            r = ws.Cells["A7:AK7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            Color Colum4 = System.Drawing.ColorTranslator.FromHtml("#FBE4D5");
            r.Style.Fill.BackgroundColor.SetColor(Colum4);

            int sl = 1;

            DataTable dtv = GetExportBookingValues(DtFrom, ToDate);


            int rw = 8;
            int frowid = 0;

            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":S" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = dtv.Rows[i]["weeks"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["Prin"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["SOC"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["St"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["CustomerName"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["Terminal"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["POL"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["POD"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["SCN"].ToString();

                ws.Cells["N" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                ws.Cells["O" + rw].Value = dtv.Rows[i]["ETA"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["ETD"].ToString();
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["Commodity"].ToString();
                ws.Cells["R" + rw].Value = dtv.Rows[i]["WGT"].ToString();
                ws.Cells["S" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                ws.Cells["T" + rw].Value = dtv.Rows[i]["Qty"].ToString();
                ws.Cells["U" + rw].Value = dtv.Rows[i]["OFTR"].ToString();
                ws.Cells["V" + rw].Value = dtv.Rows[i]["SlotAmt"].ToString();

                ws.Cells["w" + rw].Value = dtv.Rows[i]["CollectionMode"].ToString();
                ws.Cells["x" + rw].Value = dtv.Rows[i]["BCRELEASEDDATE"].ToString();
                ws.Cells["Y" + rw].Value = dtv.Rows[i]["BLTyes"].ToString();
                ws.Cells["Z" + rw].Value = dtv.Rows[i]["LocalCharges"].ToString();
                ws.Cells["AA" + rw].Value = dtv.Rows[i]["INVOICENO"].ToString();
                ws.Cells["AB" + rw].Value = dtv.Rows[i]["INVOICEDATE"].ToString();
                ws.Cells["AC" + rw].Value = dtv.Rows[i]["jobstatus"].ToString();
                ws.Cells["AD" + rw].Value = dtv.Rows[i]["SSR"].ToString();
                ws.Cells["AE" + rw].Value = dtv.Rows[i]["ITT"].ToString();
                ws.Cells["AF" + rw].Value = dtv.Rows[i]["Storage"].ToString();
                ws.Cells["AG" + rw].Value = dtv.Rows[i]["Removal"].ToString();
                ws.Cells["AH" + rw].Value = dtv.Rows[i]["ReeferMON"].ToString();
                ws.Cells["AI" + rw].Value = dtv.Rows[i]["VENDORName"].ToString();
                ws.Cells["AJ" + rw].Value = dtv.Rows[i]["OtherCost"].ToString();
                ws.Cells["AK" + rw].Value = dtv.Rows[i]["JOBHandledBy"].ToString();
               


                ws.Cells["K" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["L" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["N" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["O" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:Ak" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Ak" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Ak" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Ak" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 40].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=BookingSummaryReport.xlsx");
            Response.End();

        }




        public DataTable GetExportBookingValues(string DtFrom, string ToDate)
        {
            string strWhere = "";
            string _Query = " select datepart(week, getdate()) as weeks,'BWS' as Prin,'SOC' as SOC, 'EXP' St,NVO_BOL.BLNumber, " +
                            " BkgParty as CustomerName,(select top(1)(select top(1) TerminalName from NVO_TerminalMaster where ID = NVO_VoyageRoute.TerminalID) " +
                            " from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID) as Terminal, " +
                            " POL,POD, FPOD,VesVoy,(select  top(1) Notes  from  NVO_VoyageNotesDtls WHERE NotesTypeID =282 and VoyageID = NVO_Booking.VesVoyID) as SCN,  " +
                            " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_Booking.SlotOperatorID) as SlotOperator, " +
                            " (select top(1) ETA from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID) as ETA, " +
                            " (select top(1) ETD from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID) as ETD, " +
                            " (select top(1) Commodity from NVO_V_BookingMultipleCntrTypes " +
                            " where NVO_V_BookingMultipleCntrTypes.BKgID = NVO_Booking.ID) as Commodity, " +
                            " (select top(1) CntrTypes from NVO_V_BookingMultipleCntrTypes " +
                            " where NVO_V_BookingMultipleCntrTypes.BKgID = NVO_Booking.ID) as CntrTypes, " +
                            " (select sum(GrsWt) from NVO_BOLCntrDetails where BLID = NVO_BOL.ID) as WGT, " +
                            " (select sum(Qty) from NVO_V_BookingMultipleCntrTypes " +
                            " where NVO_V_BookingMultipleCntrTypes.BKgID = NVO_Booking.ID) as Qty, " +
                           // " (select sum(CustomerRate) from NVO_RatesheetCharges  where NVO_RatesheetCharges.RRID = NVO_Booking.RRID AND  NVO_RatesheetCharges.ChargeCodeID IN (1) ) as OFTR, " +
                           " (select sum(ManifRate) from NVO_RatesheetCharges  where NVO_RatesheetCharges.RRID = NVO_Booking.RRID AND  NVO_RatesheetCharges.TariffTypeID IN (135) ) as OFTR, " +
                            " (SlotAmt20 + SlotAmt40) as SlotAmt, " +
                            "  case when (select top(1) rc.PaymentModeID from NVO_RatesheetCharges rc where RRID = NVO_Booking.RRID) =18 then 'Prepaid' " +
                            "  when(select top(1) rc.PaymentModeID from NVO_RatesheetCharges rc where RRID = NVO_Booking.RRID) = 19 then 'Collect' else '' end" +
                            " as CollectionMode,  case when (select count(lg.ID) from NVO_BLPrintLog lg WHERE TYPE in(3,12,9,2,8,11,7,10,4) AND lg.BkgID=NVO_Booking.ID)>0 THEN  " +
                           " (select TOP 1(lg.PrintedOn) from NVO_BLPrintLog lg WHERE TYPE in(3,12,9,2,8,11,7,10,4) AND lg.BkgID = NVO_Booking.ID) ELSE '' end as BCRELEASEDDATE, " +
                            " case when(select count(lg.ID) from NVO_BLPrintLog lg WHERE TYPE = 7 AND lg.BkgID = NVO_Booking.ID) > 0 Then 'Telex' else 'OBL' END BLTyes, '' LocalCharges, " +
                            " '' as INVOICENO, '' as INVOICEDATE,'' AS jobstatus, '' as SSR,'' as ITT,'' as Storage, '' as Removal, " +
                            " '' as ReeferMON,'' as VENDORName,'' as OtherCost, (select top(1) UserName from NVO_UserDetails where ID = NVO_Booking.PreparedBYID) as JOBHandledBy " +
                            " from NVO_ContainerTxns  " +
                            " inner join NVO_Booking on NVO_Booking.ID=NVO_ContainerTxns.BLNumber " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID  " +
                            "  where NVO_ContainerTxns.StatusCode in ('FB','MB')";

            if (DtFrom != "" && DtFrom != "0" && DtFrom != null && DtFrom != "?" && DtFrom != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " AND NVO_ContainerTxns.DtMovement between '" + DtFrom + "' and '" + ToDate + "'";
                else
                    strWhere += "  and NVO_ContainerTxns.DtMovement '" + DtFrom + "' and '" + ToDate + "'";
            if (strWhere == "")
                strWhere = _Query;
            return RegManage.GetViewData(strWhere, "");
        }



    }
}