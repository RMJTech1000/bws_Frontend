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
    public class LeaseBillingReportController : Controller
    {
        // GET: LeaseBillingReport

        MasterManager Manag = new MasterManager();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult LeaseBillingReport()
        {
            return View();
        }
        public ActionResult MonthlyLadenStorageReport()
        {
            return View();
        }
        public void MonthlyLadenStorageReportValues(string RptMth, string RptMthTxt, string RptYear, string RptYearTxt, string Ctry, string CtryID, string GeoLocID, string GeoLoc, string Port, string PortID)
        {
            DataTable dtReport = vMonthlyLadenStorage(RptMth, RptYearTxt, CtryID, GeoLocID, PortID);


            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("LeaseBill");

            ws.Cells["A2"].Value = "Monthly Laden Report";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:P2"];
            r.Merge = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


            ws.Cells["A3"].Value = "Country :";
            ws.Cells["B3"].Value = Ctry;
            ws.Cells["A4"].Value = "GeoLocation :";
            ws.Cells["B4"].Value = GeoLoc;
            ws.Cells["C4"].Value = "PORT :";
            ws.Cells["D4"].Value = Port;
            ws.Cells["A5"].Value = "FOR :";
            ws.Cells["B5"].Value = RptMthTxt + " - " + RptYearTxt;//3

            ws.Cells["A3:D5"].Style.Font.Bold = true;

            //add blank row
            //next row
            ws.Cells["A7"].Value = "S.No";
            ws.Cells["B7"].Value = "IN Date.";
            ws.Cells["C7"].Value = "Cntr#";
            ws.Cells["D7"].Value = "CntrType";
            ws.Cells["E7"].Value = "Geo Location";
            ws.Cells["F7"].Value = "Agency";
            ws.Cells["G7"].Value = "Depot";
            ws.Cells["H7"].Value = "Storage From";
            ws.Cells["I7"].Value = "Storage To";
            ws.Cells["J7"].Value = "Lift On";
            ws.Cells["K7"].Value = "Lift Of";
            ws.Cells["L7"].Value = "Storage Days";
            ws.Cells["M7"].Value = "Billable Days";
            //colspan=3
            ws.Cells["N7:P7"].Value = "SLAB";//F,G,H
            r = ws.Cells["N7:P7"];
            r.Merge = true;
            //pickup sub components
            ws.Cells["N8"].Value = "From Date";
            ws.Cells["O8"].Value = "To Date";
            ws.Cells["P8"].Value = "Days";


            ws.Cells["Q7"].Value = "Removal Charge (USD)";
            ws.Cells["R7"].Value = "Ex.Rate (MYR)";
            ws.Cells["S7"].Value = "Removal Charge (MYR)";
            ws.Cells["T7"].Value = "Storage Amount";
            ws.Cells["U7"].Value = "Total Amount";
            ws.Cells["V7"].Value = "Currency";
            ws.Cells["W7"].Value = "From Status";
            ws.Cells["X7"].Value = "To Status";


            ws.Cells["A7:X8"].Style.Font.Bold = true;
            ws.Cells["A7:X8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A7:X8"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            if (dtReport.Rows.Count > 0)
            {
                int sl = 1;
                //add row
                float totAmt = 0.00f;
                int rw_begin = 7, rw = 9, c = 1;

                for (int i = 0; i < dtReport.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl++;
                    ws.Cells["B" + rw].Value = dtReport.Rows[i]["FromDate"].ToString();
                    ws.Cells["C" + rw].Value = dtReport.Rows[i]["CntrNo"].ToString();
                    ws.Cells["D" + rw].Value = dtReport.Rows[i]["Size"].ToString();
                    ws.Cells["E" + rw].Value = dtReport.Rows[i]["ToGeoLocation"].ToString();
                    ws.Cells["F" + rw].Value = dtReport.Rows[i]["ToAgency"].ToString();
                    ws.Cells["G" + rw].Value = "";
                    ws.Cells["H" + rw].Value = dtReport.Rows[i]["FromDate"].ToString();
                    ws.Cells["I" + rw].Value = dtReport.Rows[i]["ToDate"].ToString();
                    ws.Cells["J" + rw].Value = "";
                    ws.Cells["K" + rw].Value = "";
                    ws.Cells["L" + rw].Value = dtReport.Rows[i]["NDays"];
                    ws.Cells["M" + rw].Value = dtReport.Rows[i]["NDays"];
                    ws.Cells["N" + rw].Value = dtReport.Rows[i]["FromDate"].ToString();
                    ws.Cells["O" + rw].Value = dtReport.Rows[i]["ToDate"].ToString();
                    ws.Cells["P" + rw].Value = dtReport.Rows[i]["NDays"];
                    ws.Cells["Q" + rw].Value = "";
                    ws.Cells["R" + rw].Value = "";
                    ws.Cells["S" + rw].Value = "";
                    ws.Cells["T" + rw].Value = "";
                    ws.Cells["U" + rw].Value = "";
                    ws.Cells["V" + rw].Value = "";
                    ws.Cells["W" + rw].Value = dtReport.Rows[i]["FromStatus"].ToString();
                    ws.Cells["X" + rw].Value = dtReport.Rows[i]["ToStatus"].ToString();
                    rw++;
                }

                //ws.Cells["I" + rw_begin + ":X" + rw].Style.Numberformat.Format = "#,##0.00";
                ws.Cells["I" + rw].Value = "Total";
                ws.Cells["J" + rw].Formula = "=SUM(J" + (rw_begin + 1) + ":J" + (rw - 1) + ")";
                ws.Cells["K" + rw].Formula = "=SUM(K" + (rw_begin + 1) + ":K" + (rw - 1) + ")";
                ws.Cells["L" + rw].Formula = "=SUM(L" + (rw_begin + 1) + ":L" + (rw - 1) + ")";
                ws.Cells["M" + rw].Formula = "=SUM(M" + (rw_begin + 1) + ":M" + (rw - 1) + ")";
                ws.Cells["P" + rw].Formula = "=SUM(P" + (rw_begin + 1) + ":P" + (rw - 1) + ")";
                ws.Cells["Q" + rw].Formula = "=SUM(Q" + (rw_begin + 1) + ":Q" + (rw - 1) + ")";
                ws.Cells["S" + rw].Formula = "=SUM(S" + (rw_begin + 1) + ":S" + (rw - 1) + ")";
                ws.Cells["T" + rw].Formula = "=SUM(T" + (rw_begin + 1) + ":T" + (rw - 1) + ")";
                ws.Cells["U" + rw].Formula = "=SUM(U" + (rw_begin + 1) + ":U" + (rw - 1) + ")";

                ws.Cells["I" + rw + ":U" + rw].Style.Font.Bold = true;

                ws.Cells["A" + rw_begin + ":X" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":X" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":X" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":X" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, rw, 24].AutoFitColumns();





            }

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=MonthlyLadenReport.xlsx");
            Response.End();
        }

        public DataTable vMonthlyLadenStorage(string RptMth, string RptYearTxt, string CtryID, string GeoLocID, string PortID)
        {
            int Mth = int.Parse(RptMth);
            int Yr = int.Parse(RptYearTxt);
            string RptFrom = RptMth + "/01/" + RptYearTxt;

            string RptTill = RptMth + "/" + DateTime.DaysInMonth(Yr, Mth).ToString("0#") + "/" + RptYearTxt + " 23:59";
            string strWhere = "";

            string Qr = "  SELECT DISTINCT C.CntrNo,NVO_tblCntrTypes.Size, FmMove.StatusCode FromStatus,  " +
                            " FmMove.DtMovement FromDate, BookingNo, FLoc.PortName as FromLocation, FLoc.GeoLocation " +
                            " as FromGeoLocation, FLoc.AgencyName As FromAgency, CTTO.Statuscode ToStatus, AG.AgencyName As ToAgency," +
                            " CTTO.Dtmovement ToDate, DateDiff(dd, FmMove.DtMovement,CTTO.Dtmovement)+1 NDays, t1.PortName as ToLocation, " +
                            " togl.GeoLocation as ToGeoLocation FROM NVO_Containers C INNER JOIN NVO_tblCntrTypes ON NVO_tblCntrTypes.ID = C.TypeID " +
                            " INNER JOIN NVO_ContainerTxns CTTO ON CTTO.ContainerID = C.ID  INNER JOIN NVO_PortMaster " +
                            "  t1 ON t1.ID = CTTO.LocationID INNER JOIN NVO_GeoLocations togl on togl.ID = t1.GeoLocID " +
                            " INNER JOIN NVO_AgencyMaster AG ON AG.ID = CTTO.AgencyID  OUTER APPLY(SELECT TOP 1 NVO_PortMaster.PortName, gl.GeoLocation,AG1.AgencyName " +
                            "  FROM NVO_ContainerTxns A  INNER JOIN NVO_PortMaster ON NVO_PortMaster.ID = A.LocationID INNER JOIN NVO_AgencyMaster AG1 ON AG1.ID = A.AgencyID " +
                             " INNER JOIN NVO_GeoLocations gl on gl.ID = NVO_PortMaster.GeoLocID WHERE A.StatusCode IN('FV')  AND A.ContainerID = C.ID  AND A.DtMovement " +
                             " <= CTTO.DtMovement AND NVO_PortMaster.CountryID = t1.CountryID ORDER BY A.DtMovement Desc ) as FLoc " +
                            "  OUTER APPLY(SELECT TOP 1 DtMovement, StatusCode,NVO_Booking.BookingNo FROM NVO_ContainerTxns A left outer JOIN NVO_Booking ON NVO_Booking.ID = " +
                           "  A.BLNumber  INNER JOIN NVO_PortMaster ON NVO_PortMaster.ID = A.LocationID WHERE A.StatusCode IN('FV')  AND A.ContainerID = C.ID  AND A.DtMovement " +
                          " <= CTTO.DtMovement AND NVO_PortMaster.CountryID = t1.CountryID ORDER BY A.DtMovement Desc ) as FmMove WHERE CTTO.StatusCode IN('FU') ";



            if (GeoLocID != "" && GeoLocID != "0" && GeoLocID != "null" && GeoLocID != "?")
                if (strWhere == "")
                    strWhere += Qr + " and togl.ID=" + GeoLocID;
                else
                    strWhere += " and togl.ID =" + GeoLocID;

            if (RptFrom != "" && RptFrom != "undefined" || RptTill != "" && RptTill != "undefined")
                if (strWhere == "")
                    strWhere += Qr + " and CTTO.DtMovement between '" + RptFrom + "' AND '" + RptTill + "' ";
                else
                    strWhere += " and CTTO.DtMovement between '" + RptFrom + "' AND '" + RptTill + "' ";

            if (strWhere == "")
                strWhere = Qr;


            strWhere += " union  SELECT DISTINCT C.CntrNo,NVO_tblCntrTypes.Size, FmMove.StatusCode FromStatus,  " +
                      " FmMove.DtMovement FromDate, BookingNo, FLoc.PortName as FromLocation, FLoc.GeoLocation " +
                      " as FromGeoLocation, FLoc.AgencyName As FromAgency, CTTO.Statuscode ToStatus, AG.AgencyName As ToAgency," +
                      " CTTO.Dtmovement ToDate, DateDiff(dd, FmMove.DtMovement,CTTO.Dtmovement)+1 NDays, t1.PortName as ToLocation, " +
                      " togl.GeoLocation as ToGeoLocation FROM NVO_Containers C INNER JOIN NVO_tblCntrTypes ON NVO_tblCntrTypes.ID = C.TypeID " +
                      " INNER JOIN NVO_ContainerTxns CTTO ON CTTO.ContainerID = C.ID  INNER JOIN NVO_PortMaster " +
                      "  t1 ON t1.ID = CTTO.LocationID INNER JOIN NVO_GeoLocations togl on togl.ID = t1.GeoLocID " +
                      " INNER JOIN NVO_AgencyMaster AG ON AG.ID = CTTO.AgencyID  OUTER APPLY(SELECT TOP 1 NVO_PortMaster.PortName, gl.GeoLocation,AG1.AgencyName " +
                      "  FROM NVO_ContainerTxns A  INNER JOIN NVO_PortMaster ON NVO_PortMaster.ID = A.LocationID INNER JOIN NVO_AgencyMaster AG1 ON AG1.ID = A.AgencyID " +
                       " INNER JOIN NVO_GeoLocations gl on gl.ID = NVO_PortMaster.GeoLocID WHERE A.StatusCode IN('TZ')  AND A.ContainerID = C.ID  AND A.DtMovement " +
                       " <= CTTO.DtMovement AND NVO_PortMaster.CountryID = t1.CountryID ORDER BY A.DtMovement Desc ) as FLoc " +
                      "  OUTER APPLY(SELECT TOP 1 DtMovement, StatusCode,NVO_Booking.BookingNo FROM NVO_ContainerTxns A left outer JOIN NVO_Booking ON NVO_Booking.ID = " +
                     "  A.BLNumber  INNER JOIN NVO_PortMaster ON NVO_PortMaster.ID = A.LocationID WHERE A.StatusCode IN('TZ')  AND A.ContainerID = C.ID  AND A.DtMovement " +
                    " <= CTTO.DtMovement AND NVO_PortMaster.CountryID = t1.CountryID ORDER BY A.DtMovement Desc ) as FmMove WHERE CTTO.StatusCode IN('TZFB') ";

            if (GeoLocID != "" && GeoLocID != "0" && GeoLocID != "null" && GeoLocID != "?")
                if (strWhere == "")
                    strWhere += Qr + " and togl.ID=" + GeoLocID;
                else
                    strWhere += " and togl.ID =" + GeoLocID;

            if (RptFrom != "" && RptFrom != "undefined" || RptTill != "" && RptTill != "undefined")
                if (strWhere == "")
                    strWhere += Qr + " and CTTO.DtMovement between '" + RptFrom + "' AND '" + RptTill + "' " +
                        " ORDER BY 6,4 ";
                else
                    strWhere += " and CTTO.DtMovement between '" + RptFrom + "' AND '" + RptTill + "' " +
                        " ORDER BY 6,4 ";

            //if (strWhere == "")
            //    strWhere = Qr;

            return Manag.GetViewData(strWhere, "");



        }




        public void LeaseBillingReportValues(string LPID, string LP, string CntrType, string CntrTypeID, string CtRefNo, string FromDate, string ToDate, string RptYearTxt, string LeaseType, string LeaseTypeID)
        {
            DataTable dtReport = vLeaseBilling(LPID, CntrTypeID, CtRefNo, FromDate, ToDate, LeaseTypeID);


            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("LeaseBill");

            ws.Cells["A2"].Value = "LEASE BILLING";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:P2"];
            r.Merge = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


            ws.Cells["A3"].Value = "LEASING COMPANY :";//2
            ws.Cells["B3"].Value = LP;
            ws.Cells["A4"].Value = "EQUIPMENT TYPE :";
            ws.Cells["B4"].Value = CntrType;//2
            ws.Cells["C4"].Value = "DATE :";//2
            ws.Cells["D4"].Value = FromDate + " To " + ToDate;//3
            ws.Cells["A5"].Value = "TERM :";//2
            ws.Cells["B5"].Value = LeaseType;//3
            ws.Cells["A3:D5"].Style.Font.Bold = true;

            //add blank row
            //next row
            ws.Cells["A7"].Value = "S.No";
            ws.Cells["B7"].Value = "Cntr No.";
            ws.Cells["C7"].Value = "Type-Size";
            ws.Cells["D7"].Value = "Lease Term";
            ws.Cells["E7"].Value = "Contract-Release Ref#";

            ws.Cells["F7"].Value = "Free Days";
            ws.Cells["G7"].Value = "PUC Type";
            ws.Cells["H7"].Value = "PUC Charges";
            //colspan=3
            ws.Cells["I7"].Value = "Pick Up";//F,G,H
            r = ws.Cells["I7:K7"];
            r.Merge = true;
            //pickup sub components
            ws.Cells["I8"].Value = "Date";
            ws.Cells["J8"].Value = "Geo Loc";
            ws.Cells["K8"].Value = "Location";

            //colspan=3
            ws.Cells["L7"].Value = "Drop Off";
            r = ws.Cells["L7:N7"];
            r.Merge = true;

            //dropoff sub components
            ws.Cells["L8"].Value = "Date";
            ws.Cells["M8"].Value = "Geo Loc";
            ws.Cells["N8"].Value = "Location";

            //colspan=2
            ws.Cells["O7"].Value = "Period From the Date of On-Hire to till Date";
            r = ws.Cells["O7:P7"];
            r.Merge = true;

            //period sub components
            ws.Cells["O8"].Value = "From";
            ws.Cells["P8"].Value = "To";

            ws.Cells["Q7"].Value = "Overall";
            ws.Cells["Q8"].Value = "Lease Rent";
            //colspan=2
            ws.Cells["R7"].Value = "Monthly Rental for Selected Month";
            r = ws.Cells["R7:S7"];
            r.Merge = true;

            //duration sub components
            ws.Cells["R8"].Value = "From";
            ws.Cells["S8"].Value = "To";

            ws.Cells["T7"].Value = "Cost Head";
            ws.Cells["U7"].Value = "Rate";
            ws.Cells["V7"].Value = "Amount";

            ws.Cells["W7"].Value = "Current Status";
            ws.Cells["X7"].Value = "Current Location";
            ws.Cells["Y7"].Value = "Current Iding Ageing";
            ws.Cells["Z7"].Value = "Leasing Partner";
            ws.Cells["A7:Z8"].Style.Font.Bold = true;
            ws.Cells["A7:Z8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A7:Z8"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            if (dtReport.Rows.Count > 0)
            {
                int sl = 1;
                //add row
                float totAmt = 0.00f;
                int rw_begin = 7, rw = 9, c = 1;

                for (int i = 0; i < dtReport.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl++;
                    ws.Cells["B" + rw].Value = dtReport.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtReport.Rows[i]["TypeSize"].ToString();
                    ws.Cells["D" + rw].Value = dtReport.Rows[i]["LeaseTerm"].ToString();
                    ws.Cells["E" + rw].Value = dtReport.Rows[i]["RefNo"].ToString();
                    ws.Cells["F" + rw].Value = dtReport.Rows[i]["FreeDays"].ToString();
                    ws.Cells["G" + rw].Value = dtReport.Rows[i]["PucType"].ToString();
                    ws.Cells["H" + rw].Value = dtReport.Rows[i]["PucAmount"].ToString();

                    ws.Cells["I" + rw].Value = dtReport.Rows[i]["DtPkUp"].ToString();
                    ws.Cells["J" + rw].Value = dtReport.Rows[i]["GeoLocPkUp"].ToString();
                    ws.Cells["K" + rw].Value = dtReport.Rows[i]["LocPkUpNew"].ToString();

                    if (dtReport.Rows[i]["CheckDropOff"].ToString() == "1")
                    {
                        ws.Cells["L" + rw].Value = dtReport.Rows[i]["DtDpOff"].ToString();
                    }
                    else
                    {
                        ws.Cells["L" + rw].Value = "";
                    }

                    ws.Cells["M" + rw].Value = dtReport.Rows[i]["GeoLocDpOff"].ToString();
                    ws.Cells["N" + rw].Value = dtReport.Rows[i]["LocDpOff"].ToString();
                    ws.Cells["O" + rw].Value = dtReport.Rows[i]["DtPkUp"].ToString();
                    ws.Cells["P" + rw].Value = dtReport.Rows[i]["Toend"].ToString();
                    ws.Cells["Q" + rw].Value = " Rental (  " + dtReport.Rows[i]["Slab"].ToString() + "  )";
                    ws.Cells["R" + rw].Value = FromDate;
                    ws.Cells["S" + rw].Value = ToDate;
                    ws.Cells["T" + rw].Value = " Rental (  " + dtReport.Rows[i]["SelectedSlab"].ToString() + "  )";
                    ws.Cells["U" + rw].Value = double.Parse("0" + dtReport.Rows[i]["Rate"].ToString());
                    ws.Cells["V" + rw].Value = double.Parse("0" + dtReport.Rows[i]["SelectedAmount"].ToString());


                    ws.Cells["W" + rw].Value = dtReport.Rows[i]["Statuscode"].ToString();
                    ws.Cells["X" + rw].Value = dtReport.Rows[i]["CurrentLocation"].ToString();
                    ws.Cells["Y" + rw].Value = dtReport.Rows[i]["TotalDays"].ToString();
                    ws.Cells["Z" + rw].Value = dtReport.Rows[i]["LeasingPartner"].ToString();
                    float fVar = 0;
                    float.TryParse(dtReport.Rows[i]["SelectedAmount"].ToString(), out fVar);
                    totAmt += fVar;
                    rw++;
                }

                ws.Cells["N" + rw_begin + ":V" + rw].Style.Numberformat.Format = "#,##0.00";
                ws.Cells["U" + rw].Value = "Total";
                ws.Cells["V" + rw].Formula = "=SUM(V" + (rw_begin + 1) + ":V" + (rw - 1) + ")";
                ws.Cells["U" + rw + ":V" + rw].Style.Font.Bold = true;

                ws.Cells["A" + rw_begin + ":Z" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":Z" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":Z" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":Z" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                rw++;
                ws.Cells[1, 1, rw, 30].AutoFitColumns();

                DataTable copyToDataTable = dtReport.Copy();
                copyToDataTable.DefaultView.Sort = "LeaseTerm, TypeID";

                DataView view = new DataView(dtReport);
                string[] strFlds = { "TypeID", "LeaseTerm" };
                DataTable distinctTable = view.ToTable(true, strFlds);

                rw++;
                rw_begin = rw;
                ws.Cells["A" + rw].Value = "Term";
                ws.Cells["B" + rw].Value = "Type-Size";
                ws.Cells["C" + rw].Value = "Count";
                ws.Cells["D" + rw].Value = "Amount";

                ws.Cells["A" + rw + ":D" + rw].Style.Font.Bold = true;

                rw++;

                DataTable tempDt = null;

                for (int i = 0; i < distinctTable.Rows.Count; i++)
                {
                    decimal Amt = 0.00M;
                    DataView view1 = new DataView(copyToDataTable);
                    string filter1 = "TypeID = '" + distinctTable.Rows[i]["TypeID"].ToString() + "' AND LeaseTerm = '" + distinctTable.Rows[i]["LeaseTerm"].ToString() + "'";
                    //String.Format("TypeID = {0}", distinctTable.Rows[i]["TypeID"].ToString())
                    view1.RowFilter = filter1;
                    tempDt = view1.ToTable();

                    if (tempDt.Rows.Count > 0)
                    {
                        ws.Cells["A" + rw].Value = tempDt.Rows[0]["LeaseTerm"].ToString();
                        ws.Cells["B" + rw].Value = tempDt.Rows[0]["TypeSize"].ToString();

                        //Amt = tempDt.AsEnumerable().Sum(x => Convert.ToDecimal(x["Amount"]));
                        var amt1 = tempDt.Compute("Sum(SelectedAmount)", "[SelectedAmount] IS NOT NULL");
                        decimal.TryParse(amt1.ToString(), out Amt);

                        DataView view2 = new DataView(tempDt);
                        string[] strFlds2 = { "CntrNo" };
                        DataTable dtCntrs = view2.ToTable(true, strFlds2);

                        ws.Cells["C" + rw].Value = dtCntrs.Compute("Count(CntrNo)", "CntrNo IS NOT NULL");
                        ws.Cells["D" + rw].Value = decimal.Parse(Amt.ToString());
                        rw++;
                    }

                    view1 = null;
                    tempDt = null;
                }

                ws.Cells["A" + rw_begin + ":D" + (rw - 1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":D" + (rw - 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":D" + (rw - 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw_begin + ":D" + (rw - 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();



            }

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=LeaseBilling.xlsx");
            Response.End();
        }

        public DataTable vLeaseBilling(string LPID, string CntrTypeID, string CtRefNo, string FromDate, string ToDate, string LeaseTypeID)
        {
            //int Mth = int.Parse(RptMth);
            //int Yr = int.Parse(RptYearTxt);
            //string RptFrom = RptMth + "/01/" + RptYearTxt;

            //string RptTill = RptMth + "/" + DateTime.DaysInMonth(Yr, Mth).ToString("0#") + "/" + RptYearTxt + " 23:59";
            string strWhere = "";

            string strSel = " select Case when DtDpOff ='1900-01-01 00:00:00' then DATEDIFF(day, DtPkUp, Toend)+1 else DATEDIFF(day, DtPkUp, DtDpOff) + 1 end As Slab, " +
                            " Case when DtDpOff = '1900-01-01 00:00:00' then DATEDIFF(day, '" + FromDate + "', '" + ToDate + "')+1  when Convert(date, DtDpOff) > '" + ToDate + "' " +
                            " THEN DATEDIFF(day, '" + FromDate + "', '" + ToDate + "')+1  else DATEDIFF(day, '" + FromDate + "', DtDpOff) + 1 end As SelectedSlab, " +
                  " (select top 1 FreeDays from NVO_LeaseDetails LD WHERE LD.LeaseContractID = PickUpRefID and LD.CntrTypeID=TypeID ) AS FreeDays," +
                  " (select top 1 PickUpDebit from NVO_LeaseDetails LD WHERE LD.LeaseContractID = PickUpRefID and LD.CntrTypeID = TypeID) AS PucType, " +
                  " isnull((select top 1 PickUpTariffAmt from NVO_LeaseDetails LD WHERE LD.LeaseContractID = PickUpRefID and LD.CntrTypeID = TypeID ),0.00) AS PucAmount," +
                  "  (select top 1 PerDiemAmt  from NVO_LeaseDetails WHERE LeaseContractID = PickUpRefID AND NVO_LeaseDetails.CntrTypeID =TypeID ) as Rate," +
                              " Case when DtDpOff ='1900-01-01 00:00:00' then DATEDIFF(day, DtPkUp, Toend)*(select top 1 PerDiemAmt from NVO_LeaseDetails WHERE LeaseContractID = PickUpRefID  AND NVO_LeaseDetails.CntrTypeID = TypeID) " +
                              " else DATEDIFF(day, DtPkUp, dtdpoff)*(select top 1 PerDiemAmt from NVO_LeaseDetails WHERE LeaseContractID = PickUpRefID AND NVO_LeaseDetails.CntrTypeID =TypeID ) end as FullAmount," +
                              "   Case when DtDpOff ='1900-01-01 00:00:00' then (DATEDIFF(day, '" + FromDate + "','" + ToDate + "')+1) * (select top 1 PerDiemAmt from NVO_LeaseDetails WHERE LeaseContractID = PickUpRefID  AND NVO_LeaseDetails.CntrTypeID = TypeID) " +
                              "   when Convert(date, DtDpOff) > '" + ToDate + "'  then (DATEDIFF(day, '" + FromDate + "', '" + ToDate + "') + 1) * (select top 1 PerDiemAmt from NVO_LeaseDetails " +
                              " WHERE LeaseContractID = PickUpRefID  AND NVO_LeaseDetails.CntrTypeID = TypeID)  else " +
                              " (DATEDIFF(day, '" + FromDate + "', DtDpOff) + 1) * (select top 1 PerDiemAmt from NVO_LeaseDetails   WHERE LeaseContractID = PickUpRefID  AND NVO_LeaseDetails.CntrTypeID = TypeID) end SelectedAmount," +
                              "  case when LocPkUpID in (160,162,161,163,164,724)   then   'NHAVA SHEVA' when LocPkUpID in (232, 237, 354)  then   'PORT KLANG' else " +
                             " (select top 1 PM.PortName from NVO_PortMaster PM where PM.ID = LocPkUpID) end As LocPkUpNew,Case when DtDpOff ='1900-01-01 00:00:00' THEN 0 ELSE 1 END CheckDropOff,* from NVO_View_LeaseBilling " +

                           " where  ((Convert(date, DtPkUp)  <='" + ToDate + "')  or (Convert(date, DtDpOff) between '" + FromDate + "' and '" + ToDate + "') " +
                           " or(Convert(date, DtPkUp) <= '" + ToDate + "' and Convert(date, DtDpOff) > '" + ToDate + "'))  and  (Case when DtDpOff = '1900-01-01 00:00:00'  "+
                          " then DATEDIFF(day, '" + FromDate + "', '" + ToDate + "' )+1 else DATEDIFF(day, '" + FromDate + "', DtDpOff) + 1 end) > 1";



            if (LPID != "" && LPID != "0" && LPID != "null" && LPID != "?")
                if (strWhere == "")
                    strWhere += " and LeasingPartnerID=" + LPID;

            if (CntrTypeID != "" && CntrTypeID != "0" && CntrTypeID != "null" && CntrTypeID != "?")
                if (strWhere == "")
                    strWhere += " and TypeID=" + CntrTypeID;

            if (CtRefNo != "" && CtRefNo != "undefined")
                if (strWhere == "")
                    strWhere += " and RefNo like '" + CtRefNo + "'";


            if (LeaseTypeID != "" && LeaseTypeID != "0" && LeaseTypeID != "null" && LeaseTypeID != "?")
                if (strWhere == "")
                    strWhere += " and LeaseTermID = " + LeaseTypeID;



            string _query = strSel + strWhere + "";



            return Manag.GetViewData(_query, "");
        }

    }
}