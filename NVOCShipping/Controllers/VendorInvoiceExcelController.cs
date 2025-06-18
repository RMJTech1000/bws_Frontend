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
    public class VendorInvoiceExcelController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: VendorInvoiceExcel
        public ActionResult Index()
        {
            return View();
        }
        public void PortBillingExcelTemplate(string ChargeCodeID, string VendorID, string VesVoyID)
        {
       
            DataTable dtv = GetPortBillingExcelTemplateValues(ChargeCodeID, VendorID, VesVoyID);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("PortBillingUpload");

                //Record Headers          
      
                ws.Cells["A1"].Value = "ID";
                ws.Cells["B1"].Value = "VesVoy";
                ws.Cells["C1"].Value = "BLNo";
                ws.Cells["D1"].Value = "ContainerNo";
                ws.Cells["E1"].Value = "Charges";
                ws.Cells["F1"].Value = "Currency";
                ws.Cells["G1"].Value = "Amount";
                ws.Cells["H1"].Value = "ExRate";
                ws.Cells["I1"].Value = "LocalAmount";
                ws.Cells["J1"].Value = "TaxPercentage";
                ws.Cells["K1"].Value = "TaxAmount";
                ws.Cells["L1"].Value = "NetAmount";
                ws.Cells["M1"].Value = "Result";
                ws.Cells["N1"].Value = "Status";
                ExcelRange r; 
                r = ws.Cells["A1:N1"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 2;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = "0";
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ChargeCode"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ROE"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["TaxPercentage"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["TaxAmt"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["NetAmount"].ToString();
                    ws.Cells["M" + rw].Value = "";
                    ws.Cells["N" + rw].Value = "";
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A1:L" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:L" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:L" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:L" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 14].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=PortBillingUpload.xlsx");
                Response.End();

            }

        }

        public DataTable GetPortBillingExcelTemplateValues(string ChargeCodeID, string VendorID, string VesVoyID)
        {
            string strWhere = "";

            string _Query = "Select NVO_Booking.ID as BkgID,NVO_Booking.RRID,BookingNo,NVO_BOL.ID as BLID,NVO_Booking.BkgPartyID, (select top 1  CurrencyCode from NVO_CurrencyMaster Where ID = NVO_BLVenCharges.CurrencyID) As Currency,(select top 1  ID from NVO_CurrencyMaster Where ID = NVO_BLVenCharges.CurrencyID) As CurrencyID , " +
          " C.CntrNo, " +
          " (case when(select ExemptFromTax from NVO_FinCustomerCreditControl  where PartyID = (select top 1 AgentID From NVO_BLVenChargesAgentName where BkgID = NVO_Booking.ID) ) = 1 then  0 else  isnull((select top(1) TaxPercentage from NVO_ChargeTaxEngineDtls join NVO_ChgTaxDeclaration on NVO_ChgTaxDeclaration.Id = " +
          " NVO_ChargeTaxEngineDtls.TaxPercentageID where NVO_ChargeTaxEngineDtls.ChargeTBID = NVO_BLVenCharges.ChargeCodeID),0) end) as TaxPercentage," +
        " (select top 1  Qty from NVO_BookingCntrTypes Where BKgID = NVO_Booking.ID) As Qty, NVO_Booking.VesVoy, NVO_Booking.VesVoyID,(select top 1  Size from NVO_BookingCntrTypes " +
        " inner join NVO_tblCntrTypes ct ON ct.ID = NVO_BookingCntrTypes.CntrTypes Where BKgID = NVO_Booking.ID) As Size,(select top 1  ct.ID from NVO_BookingCntrTypes " +
            " inner join NVO_tblCntrTypes ct ON ct.ID = NVO_BookingCntrTypes.CntrTypes Where BKgID = NVO_Booking.ID) As CntrTypeID, ISNULL(ReqRate, 0) AS LocalAmount, " +

       " CAST(ROUND(((((case when(select ExemptFromTax from NVO_FinCustomerCreditControl  where PartyID = 2733) = 1 then  0 else isnull((select top(1) TaxPercentage from NVO_ChargeTaxEngineDtls inner join NVO_ChgTaxDeclaration on NVO_ChgTaxDeclaration.Id = NVO_ChargeTaxEngineDtls.TaxPercentageID  where NVO_ChargeTaxEngineDtls.ChargeTBID = NVO_BLVenCharges.ChargeCodeID),0) end)) *((Case when cntrtype = 17 then 1 else (select top(1) Qty from NVO_BookingCntrTypes " +
       "  where BKgID = NVO_BLVenCharges.BkgID and CntrTypes = CntrType) end) *ReqRate * ((case when CurrencyID = 1 then 1 else (select top(1) Rate from NVO_ExRate where ToCurrency = 1 order by Date desc) end)))) / 100), 2) AS DECIMAL(10,2)) as TaxAmt, " +

       " (ISNULL(ReqRate, 0) + CAST(ROUND(((((case when(select ExemptFromTax from NVO_FinCustomerCreditControl  where PartyID = 2733) = 1 then  0 else isnull((select top(1) TaxPercentage from NVO_ChargeTaxEngineDtls inner join NVO_ChgTaxDeclaration on NVO_ChgTaxDeclaration.Id = NVO_ChargeTaxEngineDtls.TaxPercentageID  where " +

       " NVO_ChargeTaxEngineDtls.ChargeTBID = NVO_BLVenCharges.ChargeCodeID),0) end)) *((Case when cntrtype = 17 then 1 else (select top(1) Qty from NVO_BookingCntrTypes where BKgID = NVO_BLVenCharges.BkgID and CntrTypes = CntrType) end) *ReqRate * ((case when CurrencyID = 1 then 1 else (select top(1) Rate from NVO_ExRate where ToCurrency = 1 order by Date desc) end)))) / 100), 2) AS DECIMAL(10,2))) as NetAmount, " +
      " (select top 1  CurrencyCode from NVO_CurrencyMaster Where ID = NVO_BLVenCharges.CurrencyID) As Currency, case when CurrencyID = NVO_BLVenCharges.CurrencyID then 1 else (select top(1) Rate from NVO_ExRate where ToCurrency = NVO_BLVenCharges.CurrencyID order by Date desc) end as ROE, " +

        " ChargeCodeID, (select top 1  ChgDesc from NVO_ChargeTB Where ID = NVO_BLVenCharges.ChargeCodeID) As ChargeCode,ISNULL((select top 1  FileName from NVO_PORTBillAttachments Where BLID = NVO_BOL.ID and ChargeCodeID=NVO_BLVenCharges.ChargeCodeID AND CntrNO= C.CntrNo),'') As FileName from NVO_Booking inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID inner join NVO_BLVenCharges on NVO_BLVenCharges.BkgID = NVO_Booking.ID  inner join  NVO_BOLCntrDetails BC ON BC.BkgId =NVO_Booking.ID inner join  NVO_Containers C ON C.ID = BC.CNTRID where ISNULL(NVO_BLVenCharges.IsPortBill,0) != 1 ";


            if (ChargeCodeID != "" && ChargeCodeID != "0" && ChargeCodeID != "null" && ChargeCodeID != "?")
                if (strWhere == "")
                    strWhere += _Query + " and ChargeCodeID =" + ChargeCodeID;
                else
                    strWhere += " and ChargeCodeID =" + ChargeCodeID;
    

            if (VendorID != "" && VendorID != "0" && VendorID != "null" && VendorID != "?")
                if (strWhere == "")
                    strWhere += _Query + " and (select top 1 AgentID From NVO_BLVenChargesAgentName where BkgID = NVO_Booking.ID ) =" + VendorID;
                else
                    strWhere += " and (select top 1 AgentID From NVO_BLVenChargesAgentName where BkgID = NVO_Booking.ID ) =" + VendorID;

            if (VesVoyID != "" && VesVoyID != "0" && VesVoyID != "null" && VesVoyID != "?")
                if (strWhere == "")
                    strWhere += _Query + " and NVO_Booking.VesVoyID =" + VesVoyID;
                else
                    strWhere += " and NVO_Booking.VesVoyID =" + VesVoyID;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }


    }
}