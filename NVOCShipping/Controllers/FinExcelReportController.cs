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
    public class FinExcelReportController : Controller
    {
        // GET: FinExcelReport
        DocumentManager Manag = new DocumentManager();
        public ActionResult Index()
        {
            return View();
        }
        public void FinanceOutStaingReport(string DtFrom, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("FinanceOutStandingReport");

            ws.Cells["A2"].Value = "CUSTOMER OUTSTANDING REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:S2"];
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
            ws.Cells["B7"].Value = "GEOLOCATION";
            ws.Cells["C7"].Value = "AGENCY NAME";
            ws.Cells["D7"].Value = "INVOICE PARTY";
            ws.Cells["E7"].Value = "SHIPPER NAME";
            ws.Cells["F7"].Value = "BL/BOOKING No.";
            ws.Cells["G7"].Value = "DOC NO#";
            ws.Cells["H7"].Value = "SHIPMENT TYPE";
            ws.Cells["I7"].Value = "DATE";
            ws.Cells["J7"].Value = "VSL-VOY";
            ws.Cells["K7"].Value = "DOC AMOUNT";
            ws.Cells["L7"].Value = "RECIPT AMOUNT";
            ws.Cells["M7"].Value = "TDS AMOUNT";
            ws.Cells["N7"].Value = "WRITE OFF AMOUNT";
            ws.Cells["O7"].Value = "BALANCE";
            ws.Cells["P7"].Value = "AGE";
            ws.Cells["Q7"].Value = "CREDIT DAYS";
            ws.Cells["R7"].Value = "CREDIT AMOUNT";
            ws.Cells["S7"].Value = "DO NUMBER";


            r = ws.Cells["A7:S7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetOutStandingCustomerReport(DtFrom, PartyID);


            int rw = 8;
            int frowid = 0;

            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":S" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = dtv.Rows[i]["GeoLoc"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["shipperName"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["FinalInvoice"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["InvDate"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["VslVoy"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["ReceiptAmount"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["TDSAmount"].ToString();
                ws.Cells["N" + rw].Value = 0.00;
                ws.Cells["O" + rw].Value = dtv.Rows[i]["OutStanding"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["Age"].ToString();
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["CreditDays"].ToString();
                ws.Cells["R" + rw].Value = dtv.Rows[i]["CreditLimit"].ToString();
                ws.Cells["S" + rw].Value = "";


                ws.Cells["K" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["L" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["N" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["O" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:S" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:S" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:S" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:S" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 24].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=FinanceOutStandingReport.xlsx");
            Response.End();

        }


        public void FinanceOutStaingReportCN(string DtFrom, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("FinanceOutStandingReport");

            ws.Cells["A2"].Value = "CUSTOMER OUTSTANDING CREDIT SUMMARY";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:Q2"];
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
            ws.Cells["B7"].Value = "INVOICE PARTY";
            ws.Cells["C7"].Value = "SHIPPER NAME";
            ws.Cells["D7"].Value = "BL/BOOKING No.";
            ws.Cells["E7"].Value = "DOC No#";
            ws.Cells["F7"].Value = "SHIPMENT TYPE";
            ws.Cells["G7"].Value = "DATE";
            ws.Cells["H7"].Value = "VSL-VOY";
            ws.Cells["I7"].Value = "DOC AMOUNT";
            ws.Cells["J7"].Value = "PAYMENT AMOUNT";
            ws.Cells["K7"].Value = "TDS AMOUNT";
            ws.Cells["L7"].Value = "WRITE OFF AMOUNT";
            ws.Cells["M7"].Value = "BALANCE";
            ws.Cells["N7"].Value = "AGE";
            ws.Cells["O7"].Value = "CREDIT DAYS";
            ws.Cells["P7"].Value = "CREDIT AMOUNT";
            ws.Cells["Q7"].Value = "DO NUMBER";


            r = ws.Cells["A7:Q7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetOutStandingCustomerReportCN(DtFrom, PartyID);


            int rw = 8;
            int frowid = 0;

            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":P" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["shipperName"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["FinalInvoice"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["InvDate"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["VslVoy"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["ReceiptAmount"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["TDSAmount"].ToString();
                ws.Cells["L" + rw].Value = 0.00;
                ws.Cells["M" + rw].Value = "-" + dtv.Rows[i]["OutStanding"].ToString();
                ws.Cells["N" + rw].Value = dtv.Rows[i]["Age"].ToString();
                ws.Cells["O" + rw].Value = dtv.Rows[i]["CreditDays"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["CreditLimit"].ToString();
                ws.Cells["Q" + rw].Value = "";


                ws.Cells["I" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["J" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["K" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["L" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:Q" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Q" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Q" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Q" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 17].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=FinanceOutStandingReportCN.xlsx");
            Response.End();

        }


        public DataTable GetOutStandingCustomerReport(string DtFrom, string PartyID)
        {
            //    string _Query = " select InvoiceNo, FinalInvoice,convert(varchar,InvDate, 103) as InvDate,PartyName,PartyID,  (select top(1) PartyName from  NVO_BOLCustomerDetails where PartyTypeID=1 and NVO_BOLCustomerDetails.BLID = CS.BLID) as shipperName, (select top(1) BLNumber from NVO_BOL where ID = cs.BLID) as BLNumber, " +
            //                               " (select top(1) SalesPerson from NVO_Booking where ID = BkgID) as SalesPerson,(select top(1) BLVesVoy from NVO_BOL where NVO_BOL.ID = CS.BLID) as VslVoy, " +
            //                               " (select top(1) CurrencyCode from NVO_CurrencyMaster where Id = CurrencyID) as Currency, DATEDIFF(Day, CS.InvDate, GETDATE()) AS Age, " +
            //                               " isnull(InvTotal, 0) -isnull((select top(1)InvTotal  from NVO_InvoiceCusBilling CR where CR.InvTypes = 2 and CR.DRID = CS.ID),0) as InvTotal, " +
            //                               " (isnull(InvTotal, 0) - isnull((select top(1)InvTotal  from NVO_InvoiceCusBilling CR where CR.InvTypes = 2 " +
            //                               " and CR.DRID = CS.ID), 0)) -isnull((SELECT SUM(ISNULL(Amount, 0)) FROM NVO_ReceiptBL WHERE NVO_ReceiptBL.InvCusBillingID = CS.ID " +
            //                               " AND ReceiptID  IN(SELECT R.ID FROM NVO_Receipts R WHERE  R.PartyID = CS.PartyID  and R.ReceiptStatus = 1)),0) as OutStanding, " +
            //                               " isnull((select top(1) Amount from NVO_ReceiptBL where NVO_ReceiptBL.InvCusBillingId = cs.ID),0.00) as ReceiptAmount, " +
            //                               " isnull((select top(1) TDSAmt from NVO_ReceiptBL where NVO_ReceiptBL.InvCusBillingId = cs.ID),0.00) as TDSAmount, " +
            //                               " (select top(1) CusCreditDays from NVO_FinCustomerCreditControl where NVO_FinCustomerCreditControl.PartyID = CS.PartyID)  as CreditDays, " +
            //                               " (select top(1) CusCreditLimit from NVO_FinCustomerCreditControl where NVO_FinCustomerCreditControl.PartyID = CS.PartyID)  as CreditLimit " +
            //                               " from NVO_InvoiceCusBilling CS where InvTypes = 1  and IsFinal = 1 and " +
            //                               " ABS((SELECT isnull(dnd.InvTotal, 0) - isnull((select top(1)InvTotal  from NVO_InvoiceCusBilling CR where CR.InvTypes = 2 " +
            //                               " and CR.DRID = CS.ID), 0) FROM NVO_InvoiceCusBilling dnd WHERE dnd.ID = CS.ID) -ISNULL((SELECT SUM(ISNULL(Amount, 0)) FROM v_ReceiptSetOffDR " +
            //                               " WHERE  v_ReceiptSetOffDR.InvID = CS.ID and ReceiptStatus = 1),0)) > 1";

            string strWhere = "";

            string _Query = "select *  from NVO_V_CustomerOutstandingReport ";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where PartyID= " + PartyID.ToString();
                else
                    strWhere += " and PartyID = " + PartyID.ToString();

            if (DtFrom != "" && DtFrom != "0" && DtFrom != null && DtFrom != "?" && DtFrom != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where InvDatev <='" + DtFrom + "'";
                else
                    strWhere += " and InvDatev <='" + DtFrom + "'";


            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }


        public DataTable GetOutStandingCustomerReportCN(string DtFrom, string PartyID)
        {
         
            string strWhere = "";

            string _Query = "select *  from NVO_V_CustomerOutstandingReportCN ";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where PartyID= " + PartyID.ToString();
                else
                    strWhere += " and PartyID = " + PartyID.ToString();

            if (DtFrom != "" && DtFrom != "0" && DtFrom != null && DtFrom != "?" && DtFrom != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where InvDatev <='" + DtFrom + "'";
                else
                    strWhere += " and InvDatev <='" + DtFrom + "'";


            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }



        public void FinanceInvoiceSummaryReport(string DtFrom, string DtTo, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("FinanceInvoiceSummaryReport");

            ws.Cells["A2"].Value = "INVOICE SUMMARY REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:AA2"];
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
            ws.Cells["B7"].Value = "LOCATION";
            ws.Cells["C7"].Value = "BRANCH GST";
            ws.Cells["D7"].Value = "DOC TYPE";
            ws.Cells["E7"].Value = "SHIPMENT TYPE";
            ws.Cells["F7"].Value = "DOC No#";
            ws.Cells["G7"].Value = "BL NUMBER";
            ws.Cells["H7"].Value = "VSL-VOY";
            ws.Cells["I7"].Value = "PRINCIPAL";
            ws.Cells["J7"].Value = "DOC DATE";
            ws.Cells["K7"].Value = "SUPPLY TO";
            ws.Cells["L7"].Value = "SUPPLY TO TAX NUMBER";
            ws.Cells["M7"].Value = "SUPPLY STATE";
            ws.Cells["N7"].Value = "SUPPLY COUNTRY";
            ws.Cells["O7"].Value = "CHARGE DESCRIPTION";
            ws.Cells["P7"].Value = "SAC CODE";
            ws.Cells["Q7"].Value = "CURRENCY";
            ws.Cells["R7"].Value = "DOC AMOUNT";
            ws.Cells["S7"].Value = "LOCAL AMOUNT";
            ws.Cells["T7"].Value = "GST%";
            ws.Cells["U7"].Value = "SGST";
            ws.Cells["V7"].Value = "CGST";
            ws.Cells["W7"].Value = "IGST";
            ws.Cells["X7"].Value = "TAX TOTAL";
            ws.Cells["Y7"].Value = "GRAND TOTAL";
            ws.Cells["Z7"].Value = "PAYMENT STATUS";
            ws.Cells["AA7"].Value = "RECEIPT NUMBER";
            ws.Cells["AB7"].Value = "UTR REFERENCE";
            ws.Cells["AC7"].Value = "CLEARENCE DATE";
            ws.Cells["AD7"].Value = "CONTAINER";


            r = ws.Cells["A7:AD7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetInvoiceSummaryReport(DtFrom, DtTo, PartyID);


            int rw = 8;
            int frowid = 0;
            string PartyInvoiceNo = "";
            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":AD" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                if (dtv.Rows[i]["FinalInvoice"].ToString() != PartyInvoiceNo)
                {

                    ws.Cells["A" + rw].Value = sl++;
                }
                ws.Cells["B" + rw].Value = dtv.Rows[i]["Location"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["BranchGSTNo"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["InvTypes"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["FinalInvoice"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["VslVoy"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["Princible"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["InvDate"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["TaxNo"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["StateCode"].ToString();
                ws.Cells["N" + rw].Value = dtv.Rows[i]["Country"].ToString();
               
                ws.Cells["Z" + rw].Value = dtv.Rows[i]["PatmntStatus"].ToString();
                ws.Cells["AA" + rw].Value = dtv.Rows[i]["Receipts"].ToString();
                ws.Cells["AB" + rw].Value = dtv.Rows[i]["UTRNumber"].ToString();
                ws.Cells["AC" + rw].Value = dtv.Rows[i]["DtReceipts"].ToString();
                ws.Cells["AD" + rw].Value = dtv.Rows[i]["CntrNos"].ToString();

                //}
                //else
                //{

                //    ws.Cells["A" + rw].Value = "";
                //    ws.Cells["B" + rw].Value = "";
                //    ws.Cells["C" + rw].Value = "";
                //    ws.Cells["D" + rw].Value = "";
                //    ws.Cells["E" + rw].Value = "";
                //    ws.Cells["F" + rw].Value = "";
                //    ws.Cells["G" + rw].Value = "";
                //    ws.Cells["H" + rw].Value = "";
                //    ws.Cells["I" + rw].Value = "";
                //    ws.Cells["J" + rw].Value = "";
                //    ws.Cells["K" + rw].Value = "";
                //    ws.Cells["L" + rw].Value = "";
                //    ws.Cells["M" + rw].Value = "";
                //    ws.Cells["X" + rw].Value = "";
                //    ws.Cells["Y" + rw].Value = "";
                //    ws.Cells["Z" + rw].Value = "";
                //    ws.Cells["AA" + rw].Value = "";
                //    ws.Cells["AB" + rw].Value = "";
                //    ws.Cells["AC" + rw].Value = "";
                //}

                ws.Cells["O" + rw].Value = dtv.Rows[i]["NarrationDescription"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["SACCode"].ToString();
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                if(dtv.Rows[i]["InvTypes"].ToString() =="DR")
                {
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["Amount"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["GSTper"].ToString();
                    ws.Cells["U" + rw].Value = dtv.Rows[i]["SGSTAmt"].ToString();
                    ws.Cells["V" + rw].Value = dtv.Rows[i]["CGSTAmt"].ToString();
                    ws.Cells["W" + rw].Value = dtv.Rows[i]["IGSTAmt"].ToString();
                    ws.Cells["X" + rw].Value = dtv.Rows[i]["TaxAmt"].ToString();
                    ws.Cells["Y" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                }
                else
                {
                    ws.Cells["R" + rw].Value = "-"+dtv.Rows[i]["Amount"].ToString();
                    ws.Cells["S" + rw].Value = "-" + dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["T" + rw].Value = "-" + dtv.Rows[i]["GSTper"].ToString();
                    ws.Cells["U" + rw].Value = "-" + dtv.Rows[i]["SGSTAmt"].ToString();
                    ws.Cells["V" + rw].Value = "-" + dtv.Rows[i]["CGSTAmt"].ToString();
                    ws.Cells["W" + rw].Value = "-" + dtv.Rows[i]["IGSTAmt"].ToString();
                    ws.Cells["X" + rw].Value = "-" + dtv.Rows[i]["TaxAmt"].ToString();
                    ws.Cells["Y" + rw].Value = "-" + dtv.Rows[i]["InvTotal"].ToString();
                }
              
                //ws.Cells["X" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                //ws.Cells["Y" + rw].Value = dtv.Rows[i]["PatmntStatus"].ToString();
                //ws.Cells["Z" + rw].Value = dtv.Rows[i]["Receipts"].ToString();
                //ws.Cells["AA" + rw].Value = dtv.Rows[i]["UTRNumber"].ToString();
                //ws.Cells["AB" + rw].Value = dtv.Rows[i]["DtReceipts"].ToString();
                //ws.Cells["AC" + rw].Value = dtv.Rows[i]["CntrNos"].ToString();


                ws.Cells["Q" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["R" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["S" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["T" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["U" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["V" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["W" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["X" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                PartyInvoiceNo = dtv.Rows[i]["FinalInvoice"].ToString();
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 29].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=FinanceInvoiceSummaryReport.xlsx");
            Response.End();

        }



        public DataTable GetInvoiceSummaryReport(string DtFrom, string DtTo, string PartyID)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_InvoiceCusBillingdtls.ID,  " +
                " Case when NVO_InvoiceCusBilling.BLTypes=1 then 'Export' when NVO_InvoiceCusBilling.BLTypes=2 then 'Import' end as ShipmentType,  " +
                            " (SELECT TOP(1) GeoLocation FROM NVO_GeoLocations WHERE Id = GeoLocID) as Location, " +
                            " (select top(1) TaxGSTNo from NVO_AgencyMaster where NVO_AgencyMaster.ID = AgentId)  as BranchGSTNo, " +
                            " case when InvTypes = 1 then 'DR' else 'CR' end as InvTypes,FinalInvoice, " +
                            " (select top(1) BLNumber from NVO_BOL where ID = NVO_InvoiceCusBilling.BLID) as BLNumber, " +
                            " (select top(1) SalesPerson from NVO_Booking where ID = BkgID) as SalesPerson, " +
                            " (select top(1) BLVesVoy from NVO_BOL where NVO_BOL.ID = NVO_InvoiceCusBilling.BLID) as VslVoy,  " +
                            " (select top(1) CompanyName from  NVO_NewCompnayDetails where Id = 1) as Princible,convert(varchar, Invdate, 103) as InvDate,PartyName,(select top (1) GSTIN from NVO_CusBranchLocation CB where CB.CID=PartyID ) as TaxNo,StateCode, " +
                            " (select top(1) CountryName from NVO_CountryMaster where Id = CountryID) as Country,NarrationDescription, " +
                            " (select top(1) SACCode from NVO_ChargeTB where Id = NVO_InvoiceCusBillingdtls.NarrationID) as SACCode, " +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where Id = NVO_InvoiceCusBillingdtls.CurrencyID) as Currency,Amount,LocalAmount,TaxAmt,InvTotal, " +
                            " GSTper,SGSTper,CGSTper,IGSTper, " +
                            " (select TOP(1) (select top(1) ReceiptNo from NVO_Receipts where NVO_Receipts.ID = NVO_ReceiptBL.ReceiptID) from NVO_ReceiptBL  where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBillingdtls.InvCusBillingID) as Receipts, " +
                            " (select TOP(1) (select top(1) Reference  from NVO_Receipts where NVO_Receipts.ID = NVO_ReceiptBL.ReceiptID) from NVO_ReceiptBL  where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBillingdtls.InvCusBillingID) as UTRNumber, " +
                            " (select TOP(1) (select top(1) convert(varchar, DtReceipt, 103)  from NVO_Receipts where NVO_Receipts.ID = NVO_ReceiptBL.ReceiptID) from NVO_ReceiptBL  where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBillingdtls.InvCusBillingID) as DtReceipts, " +
                            " case when isnull((select top(1) InvCusBillingId from NVO_ReceiptBL where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBilling.ID),0) = 0 then 'UNPAID' else 'PAID' end PatmntStatus, " +
                            " (select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID= NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID= 1 and NVO_InvoiceCusBillingTaxdtls.InvdtID= NVO_InvoiceCusBillingdtls.BLInvID) as SGSTAmt,  " +
                            " (select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 2 and NVO_InvoiceCusBillingTaxdtls.InvdtID = NVO_InvoiceCusBillingdtls.BLInvID) as CGSTAmt, " +
                            " (select Top(1) TaxAmount from NVO_InvoiceCusBillingTaxdtls where TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID and NVO_InvoiceCusBillingTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and TaxCodeID = 3 and NVO_InvoiceCusBillingTaxdtls.InvdtID = NVO_InvoiceCusBillingdtls.BLInvID) as IGSTAmt, " +
                            " (select top(1) CntrNo from NVO_view_CntrNoMulitpleDetails where BLID = NVO_InvoiceCusBilling.BLID) as CntrNos " +
                            " from NVO_InvoiceCusBilling " +
                            " inner join NVO_InvoiceCusBillingdtls on NVO_InvoiceCusBillingdtls.InvCusBillingID = NVO_InvoiceCusBilling.Id " +
                            " inner join NVO_V_CusInvoiceTaxdtls on NVO_V_CusInvoiceTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and " +
                            " NVO_V_CusInvoiceTaxdtls.TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CID=NVO_InvoiceCusBilling.PartyID " +
                            " WHERE NVO_InvoiceCusBilling.IsFinal = 1 ";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " AND CustomerID= " + PartyID.ToString();
                else
                    strWhere += " and CustomerID = " + PartyID.ToString();

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " AND InvDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and InvDate between '" + DtFrom + "' and '" + DtTo + "'";



            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }



        public void InvoiceCrCrSummaryReport(string DtFrom,string DtTo, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("DR/CRSummaryReport");

            ws.Cells["A2"].Value = "DR/CR SUMMARY OUTSTANDING REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:P2"];
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
            ws.Cells["B7"].Value = "DR/CR#";
            ws.Cells["C7"].Value = "DOC No#";
            ws.Cells["D7"].Value = "CREATED DATE";
            ws.Cells["E7"].Value = "BL No.";
            ws.Cells["F7"].Value = "VSL-VOY";
            ws.Cells["G7"].Value = "GATWAY PORT";
            ws.Cells["H7"].Value = "COMPANY";
            ws.Cells["I7"].Value = "TF CODE";
            ws.Cells["J7"].Value = "TOTAL AMOUNT";
            ws.Cells["K7"].Value = "CHARGE TYPE";
            ws.Cells["L7"].Value = "AMOUNT";
            ws.Cells["M7"].Value = "USD";
            ws.Cells["N7"].Value = "REMARKS";
            ws.Cells["O7"].Value = "CUSTOMER GSTIN";
            ws.Cells["P7"].Value = "AGENT GSTIN";
            ws.Cells["Q7"].Value = "SHIPPER";
            ws.Cells["R7"].Value = "POL";
            ws.Cells["S7"].Value = "POD";


            r = ws.Cells["A7:S7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetDrCrInvoiceSummaryReport(DtFrom, DtTo, PartyID);


            int rw = 8;
            int frowid = 0;
            string PartyInvoiceNo = "";
            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":S" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                if (dtv.Rows[i]["FinalInvoice"].ToString() != PartyInvoiceNo)
                {
                    ws.Cells["A" + rw].Value = sl++;
                }
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["InvTypes"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["FinalInvoice"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["InvDate"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["VslVoy"].ToString();
                    ws.Cells["G" + rw].Value = "";
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                    ws.Cells["I" + rw].Value = "";
               
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["NarrationDescription"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["ROE"].ToString();
                    ws.Cells["N" + rw].Value = "";
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["TaxNo"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["BranchGSTNo"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["POLName"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["PODName"].ToString();
                    ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    ws.Cells["J" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["L" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                //}
                //else
                //{
                //    ws.Cells["A" + rw].Value = "";
                //    ws.Cells["B" + rw].Value = "";
                //    ws.Cells["C" + rw].Value = "";
                //    ws.Cells["D" + rw].Value = "";
                //    ws.Cells["E" + rw].Value = "";
                //    ws.Cells["F" + rw].Value = "";
                //    ws.Cells["G" + rw].Value = "";
                //    ws.Cells["H" + rw].Value = "";
                //    ws.Cells["I" + rw].Value = "";
                //    ws.Cells["J" + rw].Value = "";
                //    ws.Cells["K" + rw].Value = dtv.Rows[i]["NarrationDescription"].ToString();
                //    ws.Cells["L" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                //    ws.Cells["M" + rw].Value = dtv.Rows[i]["ROE"].ToString();
                //    ws.Cells["N" + rw].Value = "";
                //    ws.Cells["O" + rw].Value = "";
                //    ws.Cells["P" + rw].Value = "";
                //    ws.Cells["Q" + rw].Value = "";
                //    ws.Cells["R" + rw].Value = "";
                //    ws.Cells["S" + rw].Value = "";
                //    ws.Cells["J" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    ws.Cells["L" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //    ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //}
                PartyInvoiceNo = dtv.Rows[i]["FinalInvoice"].ToString();
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:S" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:S" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:S" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:S" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 21].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=DRCRSummaryReport.xlsx");
            Response.End();


        }


        public DataTable GetDrCrInvoiceSummaryReport(string DtFrom, string DtTo, string PartyID)
        {
            string strWhere = "";
            string _Query = " Select " +
                            " (SELECT TOP(1) GeoLocation FROM NVO_GeoLocations WHERE Id = GeoLocID) as Location,  " +
                            " (select top(1) TaxGSTNo from NVO_AgencyMaster where NVO_AgencyMaster.ID = AgentId)  as BranchGSTNo,  " +
                            " case when InvTypes = 1 then 'DR' else 'CR' end as InvTypes,FinalInvoice, " +
                            " (select top(1) BLNumber from NVO_BOL where ID = NVO_InvoiceCusBilling.BLID) as BLNumber, " +
                            " (select top(1) SalesPerson from NVO_Booking where ID = BkgID) as SalesPerson,  " +
                            " (select top(1) BLVesVoy from NVO_BOL where NVO_BOL.ID = NVO_InvoiceCusBilling.BLID) as VslVoy, " +
                            " (select top(1) CompanyName from  NVO_NewCompnayDetails where Id = 1) as Princible,convert(varchar, Invdate, 103) as InvDate,PartyName,(select top (1) GSTIN from NVO_CusBranchLocation CB where CB.CID=PartyID ) as TaxNo,StateCode, NarrationDescription, " +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where Id = NVO_View_InvoiceDRCrSummary.CurrencyID) as Currency, " +
                           " case when InvTypes = 1 then Amount else -Amount end Amount, case when InvTypes = 1 then LocalAmount else -LocalAmount end LocalAmount, " +
                           " case when InvTypes = 1 then InvTotal else -InvTotal end InvTotal," +
                            " ROE,NVO_View_InvoiceDRCrSummary.sno, " +
                            " (SELECT TOP(1) PartyName FROM NVO_BOLCustomerDetails WHERE PartyTypeID= 1 AND NVO_BOLCustomerDetails.BLID=NVO_InvoiceCusBilling.BLID) AS Shipper, " +
                            " (select top(1)(select top(1) PortName from NVO_PortMaster where ID = NVO_BOL.POLID) from NVO_BOL where NVO_BOL.ID = NVO_InvoiceCusBilling.BLID) as POLName, " +
                            " (select top(1)(select top(1) PortName from NVO_PortMaster where ID = NVO_BOL.PODID) from NVO_BOL where NVO_BOL.ID = NVO_InvoiceCusBilling.BLID) as PODName" +
                            " from NVO_InvoiceCusBilling " +
                            " inner join NVO_View_InvoiceDRCrSummary on NVO_View_InvoiceDRCrSummary.InvCusBillingID = NVO_InvoiceCusBilling.Id " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CID = NVO_InvoiceCusBilling.PartyID " +
                            " WHERE NVO_InvoiceCusBilling.IsFinal = 1";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null && PartyID.ToString() != "null")

                if (strWhere == "")
                    strWhere += _Query + " AND CustomerID= " + PartyID.ToString();
                else
                    strWhere += " and CustomerID = " + PartyID.ToString();

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " AND InvDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and InvDate '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query + " order by FinalInvoice, Sno ";


            return Manag.GetViewData(strWhere, "");
        }


        public void FinancePerformaInvoiceSummaryReport(string DtFrom, string DtTo, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("PerformaInvoiceSummaryReport");

            ws.Cells["A2"].Value = "PERFORMA INVOICE SUMMARY REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:AA2"];
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
            ws.Cells["B7"].Value = "LOCATION";
            ws.Cells["C7"].Value = "BRANCH GST";
            ws.Cells["D7"].Value = "DOC TYPE";
            ws.Cells["E7"].Value = "DOC No#";
            ws.Cells["F7"].Value = "BL NUMBER";
            ws.Cells["G7"].Value = "VSL-VOY";
            ws.Cells["H7"].Value = "PRINCIPAL";
            ws.Cells["I7"].Value = "DOC DATE";
            ws.Cells["J7"].Value = "SUPPLY TO";
            ws.Cells["K7"].Value = "SUPPLY TO TAX NUMBER";
            ws.Cells["L7"].Value = "SUPPLY STATE";
            ws.Cells["M7"].Value = "SUPPLY COUNTRY";
            ws.Cells["N7"].Value = "CHARGE DESCRIPTION";
            ws.Cells["O7"].Value = "SAC CODE";
            ws.Cells["P7"].Value = "CURRENCY";
            ws.Cells["Q7"].Value = "DOC AMOUNT";
            ws.Cells["R7"].Value = "LOCAL AMOUNT";
            ws.Cells["S7"].Value = "GST%";
            ws.Cells["T7"].Value = "SGST";
            ws.Cells["U7"].Value = "CGST";
            ws.Cells["V7"].Value = "IGST";
            ws.Cells["W7"].Value = "TAX TOTAL";
            ws.Cells["X7"].Value = "GRAND TOTAL";
            ws.Cells["Y7"].Value = "PAYMENT STATUS";
            ws.Cells["Z7"].Value = "RECEIPT NUMBER";
            ws.Cells["AA7"].Value = "UTR REFERENCE";
            ws.Cells["AB7"].Value = "CLEARENCE DATE";


            r = ws.Cells["A7:AB7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetPerformaInvoiceSummaryReport(DtFrom, DtTo,PartyID);


            int rw = 8;
            int frowid = 0;
            string PartyInvoiceNo = "";

            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":AB" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                if (dtv.Rows[i]["InvoiceNo"].ToString() != PartyInvoiceNo)
                {
                    ws.Cells["A" + rw].Value = sl++;
                }
                   
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Location"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["BranchGSTNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["InvTypes"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["InvoiceNo"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["VslVoy"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["Princible"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["InvDate"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["TaxNo"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["StateCode"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["Country"].ToString();
                //}
                //else
                //{
                //    ws.Cells["A" + rw].Value = "";
                //    ws.Cells["B" + rw].Value = "";
                //    ws.Cells["C" + rw].Value = "";
                //    ws.Cells["D" + rw].Value = "";
                //    ws.Cells["E" + rw].Value = "";
                //    ws.Cells["F" + rw].Value = "";
                //    ws.Cells["G" + rw].Value = "";
                //    ws.Cells["H" + rw].Value = "";
                //    ws.Cells["I" + rw].Value = "";
                //    ws.Cells["J" + rw].Value = "";
                //    ws.Cells["K" + rw].Value = "";
                //    ws.Cells["L" + rw].Value = "";
                //    ws.Cells["M" + rw].Value = "";

                //}

                ws.Cells["N" + rw].Value = dtv.Rows[i]["NarrationDescription"].ToString();
                ws.Cells["O" + rw].Value = dtv.Rows[i]["SACCode"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["Amount"].ToString();
                ws.Cells["R" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                ws.Cells["S" + rw].Value = dtv.Rows[i]["GSTper"].ToString();
                ws.Cells["T" + rw].Value = dtv.Rows[i]["SGSTper"].ToString();
                ws.Cells["U" + rw].Value = dtv.Rows[i]["CGSTper"].ToString();
                ws.Cells["V" + rw].Value = dtv.Rows[i]["IGSTper"].ToString();
                ws.Cells["W" + rw].Value = dtv.Rows[i]["TaxAmt"].ToString();
                ws.Cells["X" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                ws.Cells["Y" + rw].Value = dtv.Rows[i]["PatmntStatus"].ToString();
                ws.Cells["Z" + rw].Value = dtv.Rows[i]["Receipts"].ToString();
                ws.Cells["AA" + rw].Value = dtv.Rows[i]["UTRNumber"].ToString();
                ws.Cells["AB" + rw].Value = dtv.Rows[i]["DtReceipts"].ToString();
                ws.Cells["Q" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["R" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["S" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["T" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["U" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["V" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["W" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["X" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                PartyInvoiceNo = dtv.Rows[i]["InvoiceNo"].ToString();
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:AB" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AB" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AB" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AB" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 28].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=FinancePerformaInvoiceSummaryReport.xlsx");
            Response.End();

        }


        public DataTable GetPerformaInvoiceSummaryReport(string DtFrom, string DtTo, string PartyID)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_InvoiceCusBillingdtls.ID, " +
                            " (SELECT TOP(1) GeoLocation FROM NVO_GeoLocations WHERE Id = GeoLocID) as Location, " +
                            " (select top(1) TaxGSTNo from NVO_AgencyMaster where NVO_AgencyMaster.ID = AgentId)  as BranchGSTNo, " +
                            " case when InvTypes = 1 then 'DR' else 'CR' end as InvTypes,InvoiceNo, " +
                            " (select top(1) BLNumber from NVO_BOL where ID = NVO_InvoiceCusBilling.BLID) as BLNumber, " +
                            " (select top(1) SalesPerson from NVO_Booking where ID = BkgID) as SalesPerson, " +
                            " (select top(1) BLVesVoy from NVO_BOL where NVO_BOL.ID = NVO_InvoiceCusBilling.BLID) as VslVoy,  " +
                            " (select top(1) CompanyName from  NVO_NewCompnayDetails where Id = 1) as Princible,convert(varchar, Invdate, 103) as InvDate,PartyName,(select top (1) GSTIN from NVO_CusBranchLocation CB where CB.CID=PartyID ) as TaxNo,StateCode, " +
                            " (select top(1) CountryName from NVO_CountryMaster where Id = CountryID) as Country,NarrationDescription, " +
                            " (select top(1) SACCode from NVO_ChargeTB where Id = NVO_InvoiceCusBillingdtls.NarrationID) as SACCode, " +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where Id = NVO_InvoiceCusBillingdtls.CurrencyID) as Currency,Amount,LocalAmount,TaxAmt,InvTotal, " +
                            " GSTper,SGSTper,CGSTper,IGSTper, " +
                            " (select TOP(1) (select top(1) ReceiptNo from NVO_Receipts where NVO_Receipts.ID = NVO_ReceiptBL.ReceiptID) from NVO_ReceiptBL  where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBillingdtls.InvCusBillingID) as Receipts, " +
                            " (select TOP(1) (select top(1) Reference  from NVO_Receipts where NVO_Receipts.ID = NVO_ReceiptBL.ReceiptID) from NVO_ReceiptBL  where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBillingdtls.InvCusBillingID) as UTRNumber, " +
                            " (select TOP(1) (select top(1) convert(varchar, DtReceipt, 103)  from NVO_Receipts where NVO_Receipts.ID = NVO_ReceiptBL.ReceiptID) from NVO_ReceiptBL  where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBillingdtls.InvCusBillingID) as DtReceipts, " +
                            " case when isnull((select top(1) InvCusBillingId from NVO_ReceiptBL where NVO_ReceiptBL.InvCusBillingId = NVO_InvoiceCusBilling.ID),0) = 0 then 'UNPAID' else 'PAID' end PatmntStatus " +
                            " from NVO_InvoiceCusBilling " +
                            " inner join NVO_InvoiceCusBillingdtls on NVO_InvoiceCusBillingdtls.InvCusBillingID = NVO_InvoiceCusBilling.Id " +
                            " inner join NVO_V_CusInvoiceTaxdtls on NVO_V_CusInvoiceTaxdtls.InvCusBillingID = NVO_InvoiceCusBillingdtls.InvCusBillingID and " +
                            " NVO_V_CusInvoiceTaxdtls.TaxNarrationID = NVO_InvoiceCusBillingdtls.NarrationID " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CID=NVO_InvoiceCusBilling.PartyID " +
                            " WHERE isnull(NVO_InvoiceCusBilling.IsFinal,0) != 1 ";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " AND CustomerID= " + PartyID.ToString();
                else
                    strWhere += " and CustomerID = " + PartyID.ToString();

           
            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " AND InvDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and InvDate between '" + DtFrom + "' and '" + DtTo + "'";


            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }


        public void ReceiptSummaryReport(string DtFrom, string DtTo, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("ReceiptSummaryReport");

            ws.Cells["A2"].Value = "RECEIPT SUMMARY REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:AA2"];
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
            ws.Cells["B7"].Value = "RECEIPT#";
            ws.Cells["C7"].Value = "CREATED DATE";
            ws.Cells["D7"].Value = "TOTAL AMOUNT";
            ws.Cells["E7"].Value = "PAY MODE";
            ws.Cells["F7"].Value = "CHEQUE/DD#";
            ws.Cells["G7"].Value = "CHEQUE/DD DATE";
            ws.Cells["H7"].Value = "BANK";
            ws.Cells["I7"].Value = "PAYER";
            ws.Cells["J7"].Value = "BLNUMBER";
            ws.Cells["K7"].Value = "DR NOTES#";
            ws.Cells["L7"].Value = "DR DATE";
            ws.Cells["M7"].Value = "COMPNAY";
            ws.Cells["N7"].Value = "RECEIPT TOTAL(USD)";
            ws.Cells["O7"].Value = "TDS AMOUNT";
            ws.Cells["P7"].Value = "WRITE OF AMOUNT";
            ws.Cells["Q7"].Value = "RECEIPT TOTAL(CY)";
            ws.Cells["R7"].Value = "INV TOTAL";
            ws.Cells["S7"].Value = "EXCESS AMOUNT";
            ws.Cells["T7"].Value = "REMARKS";
            ws.Cells["U7"].Value = "TAN#";
            ws.Cells["V7"].Value = "DEPOSIT BANK";
            ws.Cells["W7"].Value = "VESSEL";
            ws.Cells["X7"].Value = "IMP/EXP VOYAGE";
            ws.Cells["Y7"].Value = "SAILING DATE";
            r = ws.Cells["A7:Y7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetRceiptSummaryReport(DtFrom, DtTo,PartyID);


            int rw = 8;
            int frowid = 0;

            string PartyInvoiceNo = "";
            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":Y" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                if (dtv.Rows[i]["ReceiptNo"].ToString() != PartyInvoiceNo)
                {
                    ws.Cells["A" + rw].Value = sl++;
                }
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ReceiptNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["DtReceipt"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["PaymentMode"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Reference"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["PaymentDate"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["Bank"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["DrNotes"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["DrDate"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                //}
                //else
                //{
                //    ws.Cells["A" + rw].Value = "";
                //    ws.Cells["B" + rw].Value = "";
                //    ws.Cells["C" + rw].Value = "";
                //    ws.Cells["D" + rw].Value = "";
                //    ws.Cells["E" + rw].Value = "";
                //    ws.Cells["F" + rw].Value = "";
                //    ws.Cells["G" + rw].Value = "";
                //    ws.Cells["H" + rw].Value = "";
                //    ws.Cells["I" + rw].Value = "";
                //    ws.Cells["J" + rw].Value = "";
                //    ws.Cells["K" + rw].Value = "";
                //    ws.Cells["L" + rw].Value = "";
                //    ws.Cells["M" + rw].Value = "";
                //}
                   
                ws.Cells["N" + rw].Value = "";
                ws.Cells["O" + rw].Value = dtv.Rows[i]["TDSAmount"].ToString();
                ws.Cells["P" + rw].Value = "";
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["ReceiptAmt"].ToString();
                ws.Cells["R" + rw].Value = dtv.Rows[i]["InvAmount"].ToString();
                ws.Cells["S" + rw].Value = dtv.Rows[i]["ExcessLocalAmt"].ToString();
                ws.Cells["T" + rw].Value = dtv.Rows[i]["Remarks"].ToString();

                ws.Cells["U" + rw].Value = dtv.Rows[i]["TanNo"].ToString();
                ws.Cells["V" + rw].Value = dtv.Rows[i]["Bank"].ToString();
                ws.Cells["W" + rw].Value = dtv.Rows[i]["ExpVoyage"].ToString();
                ws.Cells["X" + rw].Value = dtv.Rows[i]["ImpVoyage"].ToString();
               
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["O" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["Q" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["R" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["S" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                PartyInvoiceNo = dtv.Rows[i]["ReceiptNo"].ToString();

                rw++;

            }

            rw -= 1;

            ws.Cells["A7:Y" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Y" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Y" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:Y" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 28].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=ReceiptSummaryReport.xlsx");
            Response.End();

        }



        public DataTable GetRceiptSummaryReport(string DtFrom, string DtTo, string PartyID)
        {
            string strWhere = "";
            string _Query = " select ReceiptNo,convert(varchar,DtReceipt, 103) as DtReceipt,LocalAmount, " +
                            " (select top(1)GeneralName from NVO_GeneralMaster  where Id = NVO_Receipts.PaymentTypes) as PaymentMode, Reference, " +
                            " convert(varchar, PaymentDate, 103) as PaymentDate,(select top(1) BankName from NVO_FinBankMaster where Id = Bank) as Bank,NVO_Receipts.PartyName, " +
                            " (select  top(1)(select top(1) BLNumber from NVO_BOL  where ID = NVO_InvoiceCusBilling.BLID)  from NVO_InvoiceCusBilling where Id = NVO_ReceiptBL.InvCusBillingId) as BLNumber, " +
                            " (select FinalInvoice from NVO_InvoiceCusBilling where Id = NVO_ReceiptBL.InvCusBillingId) as DrNotes, " +
                            " (select convert(varchar, Invdate, 103) from NVO_InvoiceCusBilling where Id = NVO_ReceiptBL.InvCusBillingId) as DrDate,NVO_Receipts.PartyName, " +
                            " isnull(TDSAmount, 0) as TDSAmount, NVO_ReceiptBL.Amount as ReceiptAmt, " +
                            " (select top(1) InvTotal from NVO_InvoiceCusBilling where Id = NVO_ReceiptBL.InvCusBillingId) as InvAmount,ExcessLocalAmt,NVO_Receipts.Remarks, " +
                            " (select top(1)TanNo from NVO_CusBranchLocation where CID = NVO_Receipts.PartyID) as TanNo, " +
                            " (select BLVesVoy from NVO_BOL where ID = NVO_InvoiceCusBilling.BLID) as ExpVoyage,  " +
                            " (select top(1)(select top(1) VesVoy from NVO_View_VoyageDetails where ID = NVO_BOLImpVoyageDetails.VesVoyID) from NVO_BOLImpVoyageDetails " +
                            " where BLID = NVO_InvoiceCusBilling.BLID) as ImpVoyage " +
                            " from NVO_Receipts " +
                            " inner join NVO_ReceiptBL on NVO_ReceiptBL.ReceiptID = NVO_Receipts.ID " +
                            " inner join NVO_InvoiceCusBilling on NVO_InvoiceCusBilling.Id = NVO_ReceiptBL.InvCusBillingId";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " WHERE NVO_Receipts.PartyID= " + PartyID.ToString();
                else
                    strWhere += " and NVO_Receipts.PartyID = " + PartyID.ToString();

    
            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE DtReceipt between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and DtReceipt between '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public void FinanceOutStaingProfomaReport(string DtFrom, string PartyID, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("FinanceOutStandingReport");

            ws.Cells["A2"].Value = "CUSTOMER OUTSTANDING PROFOMA REPORT";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:T2"];
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
            ws.Cells["B7"].Value = "GEOLOCATION";
            ws.Cells["C7"].Value = "AGENCY NAME";
            ws.Cells["D7"].Value = "INVOICE PARTY";
            ws.Cells["E7"].Value = "SHIPPER NAME";
            ws.Cells["F7"].Value = "BL/BOOKING No.";
            ws.Cells["G7"].Value = "DOC NO#";
            ws.Cells["H7"].Value = "SHIPMENT TYPE";
            ws.Cells["I7"].Value = "DATE";
            ws.Cells["J7"].Value = "VSL-VOY";
            ws.Cells["K7"].Value = "DOC AMOUNT";
            ws.Cells["L7"].Value = "RECIPT AMOUNT";
            ws.Cells["M7"].Value = "TDS AMOUNT";
            ws.Cells["N7"].Value = "WRITE OFF AMOUNT";
            ws.Cells["O7"].Value = "BALANCE";
            ws.Cells["P7"].Value = "AGE";
            ws.Cells["Q7"].Value = "CREDIT DAYS";
            ws.Cells["R7"].Value = "CREDIT AMOUNT";
            ws.Cells["S7"].Value = "DO NUMBER";
           // ws.Cells["T7"].Value = "LINE";


            r = ws.Cells["A7:T7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            DataTable dtv = GetOutStandingCustomerProfomaReport(DtFrom, PartyID);


            int rw = 8;
            int frowid = 0;

            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                frowid = rw;
                ExcelRange rng = ws.Cells["A" + frowid + ":S" + frowid];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = dtv.Rows[i]["GeoLoc"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["shipperName"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["FinalInvoice"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["InvDate"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["VslVoy"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["InvTotal"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["ReceiptAmount"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["TDSAmount"].ToString();
                ws.Cells["N" + rw].Value = 0.00;
                ws.Cells["O" + rw].Value = dtv.Rows[i]["OutStanding"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["Age"].ToString();
                ws.Cells["Q" + rw].Value = dtv.Rows[i]["CreditDays"].ToString();
                ws.Cells["R" + rw].Value = dtv.Rows[i]["CreditLimit"].ToString();
                ws.Cells["S" + rw].Value = "";
               // ws.Cells["T" + rw].Value = "";


                ws.Cells["K" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["L" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["M" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["N" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["O" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                rw++;

            }

            rw -= 1;

            ws.Cells["A7:T" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:T" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:T" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:T" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 24].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=FinanceOutStandingProformaReport.xlsx");
            Response.End();

        }
        public DataTable GetOutStandingCustomerProfomaReport(string DtFrom, string PartyID)
        {

            string strWhere = "";

            string _Query = "select *  from NVO_V_CustomerOutstandingProformaReport ";

            if (PartyID.ToString() != "" && PartyID.ToString() != "0" && PartyID.ToString() != "?" && PartyID.ToString() != "undefined" && PartyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where PartyID= " + PartyID.ToString();
                else
                    strWhere += " and PartyID = " + PartyID.ToString();

            if (DtFrom != "" && DtFrom != "0" && DtFrom != null && DtFrom != "?" && DtFrom != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where InvDatev <='" + DtFrom + "'";
                else
                    strWhere += " and InvDatev <='" + DtFrom + "'";


            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
    }
}