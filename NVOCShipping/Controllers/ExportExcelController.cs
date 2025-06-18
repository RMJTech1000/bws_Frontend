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
    public class ExportExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: ExportExcel
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult BookingExcelReport()
        {
            return View();
        }

        public ActionResult ExportLoadingReport()
        {
            return View();

        }

        public ActionResult WeeklyExportPickupReport()
        {
            return View();
        }

        public ActionResult WeeklyCustomerReport()
        {
            return View();
        }
        public ActionResult FeederWiseLifting()
        {
            ViewBag.Message = "Your FeederWiseLifting.";

            return View();
        }
        public ActionResult SOAReport()
        {
            ViewBag.Message = "Your SOAReport.";

            return View();
        }
        public ActionResult SOAExportReport()
        {
            ViewBag.Message = "Your SOAExportReport.";

            return View();
        }
        public ActionResult SOAImportReport()
        {
            ViewBag.Message = "Your SOAImportReport.";

            return View();
        }

        public ActionResult WeeklyFeederWiseLifting()
        {
            ViewBag.Message = "Your FeederWiseLifting.";

            return View();
        }

        public ActionResult WeeklyLiftingReport()
        {
            return View();
        }

        public ActionResult BookingSummaryReport()
        {
            return View();
        }

        
        public void BookingReport(string DtFrom, string DtTo, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetBookingSearchValues(DtFrom, DtTo, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("BookingReport");

                ws.Cells["A2"].Value = "Booking Report";
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
                ws.Cells["B7"].Value = "Bkg.Loc";
                ws.Cells["C7"].Value = "Bkg.Ref No";
                ws.Cells["D7"].Value = "IsOnline";
                ws.Cells["E7"].Value = "Vessel/Voyage";
                ws.Cells["F7"].Value = "ETD";
                ws.Cells["G7"].Value = "Booking Creation Date";
                ws.Cells["H7"].Value = "Shipper";
                ws.Cells["I7"].Value = "Booking Party";
                ws.Cells["J7"].Value = "Invoicing Party";
                ws.Cells["K7"].Value = "POL";
                ws.Cells["L7"].Value = "Ts Ports";
                ws.Cells["M7"].Value = "POD";
                ws.Cells["N7"].Value = "FPOD";
                ws.Cells["O7"].Value = "VSL Operator";
                ws.Cells["P7"].Value = "Cntr type";
                ws.Cells["Q7"].Value = "Volume";
                ws.Cells["R7"].Value = "ALLOTED/PICKED UP";
                ws.Cells["S7"].Value = "Charge Code";
                ws.Cells["T7"].Value = "Customer Rate";
                ws.Cells["U7"].Value = "Manifest Rate";
                ws.Cells["V7"].Value = "Difference";
                ws.Cells["W7"].Value = "Approval RR No";
                ws.Cells["X7"].Value = "Payment Term";
                ws.Cells["Y7"].Value = "BOOKING NO";
                ws.Cells["Z7"].Value = "Cntr No";
                ws.Cells["AA7"].Value = "Cntr Type";
                ws.Cells["AB7"].Value = "Gr.Wt(Tons)";
                ws.Cells["AC7"].Value = "Booking remarks";
                ws.Cells["AD7"].Value = "Booking Status";
                ws.Cells["AE7"].Value = "Sale Pic";
                ws.Cells["AF7"].Value = "Commodity";
                ws.Cells["AG7"].Value = "BLNumber";
                ws.Cells["AH7"].Value = "Invoice No";
                ws.Cells["AI7"].Value = "Sailing Date";
                //ws.Cells["AJ7"].Value = "Last Modified By";
                //ws.Cells["AK7"].Value = "CRO Modified On(DD/MM/YYY)";

                r = ws.Cells["A7:AI7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;
                int rw_begin = 8, rw_end = 0; ;
                int rw = 8;
                int frowid = 0;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    frowid = rw;


                    ExcelRange rng = ws.Cells["A" + frowid + ":AI" + frowid];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Bkgloc"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["IsOnline"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["ETD"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["BkgDate"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["BkgParty"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["InvParty"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["PortofLoading"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["TranshimentPort"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["PlaceofDischarge"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["CntrType"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["Volume"];
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["PickUpQty"].ToString();
                    ws.Cells["W" + rw].Value = dtv.Rows[i]["RRNo"].ToString();
                    ws.Cells["X" + rw].Value = dtv.Rows[i]["FreightPayment"].ToString();
                    ws.Cells["Y" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["AC" + rw].Value = dtv.Rows[i]["Remarks"].ToString();
                    ws.Cells["AD" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["AE" + rw].Value = dtv.Rows[i]["SalesPerson"].ToString();
                    ws.Cells["AF" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();

                    //ws.Cells["AJ" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();
                    //ws.Cells["AK" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();
                    int stRow = rw;
                    int StRow1 = rw;
                    int StRow2 = rw;
                    DataTable dtR = GetRRDetails(dtv.Rows[i]["RRID"].ToString());
                    rw = stRow;
                    int chgRow = 1;
                    for (int k = 0; k < dtR.Rows.Count; k++)
                    {

                        ws.Cells["S" + rw].Value = dtR.Rows[k]["ChgCode"].ToString();
                        ws.Cells["T" + rw].Value = dtR.Rows[k]["CustomerRate"].ToString();
                        ws.Cells["U" + rw].Value = dtR.Rows[k]["ManifRate"].ToString();
                        ws.Cells["V" + rw].Value = dtR.Rows[k]["RateDiff"].ToString();

                        rw++;
                        chgRow++;

                    }
                    rw = StRow1;
                    int cntrRow = 1;
                    DataTable dtCn = GetCntrPickupDetails(dtv.Rows[i]["ID"].ToString());
                    for (int j = 0; j < dtCn.Rows.Count; j++)

                    {
                        //rw = StRow1;
                        ws.Cells["Z" + rw].Value = dtCn.Rows[j]["CntrNo"].ToString();
                        ws.Cells["AA" + rw].Value = dtCn.Rows[j]["Size"].ToString();
                        ws.Cells["AB" + rw].Value = dtCn.Rows[j]["GrossWt"].ToString();
                        rw++;
                        cntrRow++;

                    }
                    int totalcolumn = (chgRow < cntrRow) ? cntrRow : chgRow;

                    rw = StRow2;
                    int blRow = 1;
                    DataTable dtbldetails = GetBLDetails(dtv.Rows[i]["ID"].ToString());
                    for (int l = 0; l < dtbldetails.Rows.Count; l++)

                    {
                        //rw = StRow1;
                        ws.Cells["AG" + rw].Value = dtbldetails.Rows[l]["BLNumber"].ToString();
                        ws.Cells["AH" + rw].Value = dtbldetails.Rows[l]["InvoiceNo"].ToString();
                        ws.Cells["AI" + rw].Value = dtbldetails.Rows[l]["SOBDate"].ToString();
                        rw++;
                        blRow++;

                    }
                    sl++;
                    // rw += 1;
                    rw = (totalcolumn + rw) - 1;
                }

                rw -= 1;
                rw_end = rw;


                ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_begin, (rw_end - 1));
                ws.Cells["Q" + rw_end].Style.Numberformat.Format = "#";

                ws.Cells["A7:AI" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AI" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AI" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AI" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 40].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=BookingReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetBookingSearchValues(string DtFrom, string DtTo, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = "select NVO_Booking.ID,BookingNo,isnull(RRNo,'') as RRNo,isnull(RRID,0) as RRID,(select top 1 GeoLocation from NVO_GeoLocations Where id = NVO_AgencyMaster.GeoLocationID) AS Bkgloc,convert(varchar, (select top 1 ETD from NVO_VoyageRoute Where VoyageID = NVO_Booking.VesVoyID),103) As ETD,convert(varchar, BkgDate, 106) as BkgDate, " +
                            " (select top(1) CustomerName from NVO_CustomerMaster where Id = ShipperID ) Shipper, '' AS InvParty,'' AS IsOnline,isnull(ShipmentType, '') as ShipmentType ,BkgParty,POL,case when BkgStatus = 1 then 'DRAFT' else case when BKGStatus = 2  then 'CONFIRM' else case when BkgStatus = 3 then 'CANCELED' end end end as Status, " +
          "(select top(1) PortName from NVO_PortMaster where Id = POLID) PortofLoading, (select top(1) PortName from NVO_PortMaster where Id = PODID) PlaceofDischarge, (select top(1) FPOD from NVO_BLRelease where BKgID = NVO_Booking.ID ) FPOD,  (select top(1) PortName from NVO_PortMaster where Id = TSPORTID ) TranshimentPort,  " +
        " (select top 1(type + '-' + size) from NVO_tblCntrTypes where ID = NVO_BookingCntrTypes.CntrTypes) as CntrType,(select top(1) Qty from NVO_BookingCntrTypes where BKgID = NVO_Booking.ID ) Volume,  " +
     "(select top 1 ReqQty from NVO_CRODetails inner join NVO_CROmaster on NVO_CROmaster.id = NVO_CRODetails.CROID Where BKGID = NVO_Booking.ID) As PickUpQty, (select top(1) AgencyName from NVO_AgencyMaster where Id = AgentID) CreatedAgency, " +
   " (select top(1) AgencyName from NVO_AgencyMaster where Id = DestinationAgentID) DesAgency, (select top(1) GeneralName from NVO_GeneralMaster where Id = ShipmentTypeID) ShipmentType, (select top(1) Description from NVO_tblDLValues where Id = ServiceTypeID) ServiceType, " +
   "(select top(1) CustomerName from NVO_CustomerMaster where Id = ShipperID ) Shipper,   (select top(1) DepName from NVO_DepotMaster where Id = pickupdepotid ) PickupDepot,  (select top(1) Username from NVO_UserDetails where Id = PreparedBYID ) PreparedBy,(Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = VesVoyID )as VesVoy, " +
 " (select top(1) SlotRef from NVO_SLOTMaster where Id = SlotContractID ) SlotContract,  (select  upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster   inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where CID =SlotOperatorID ) SlotOperator,Remarks,isnull(SalesPerson,'') as SalesPerson,NVO_Booking.CommodityType," +
 " (select top(1) PartyName from NVO_InvoiceCusBilling where BKgID = NVO_Booking.ID ) InvoiceParty, " +
 " (select top(1) GeneralName  from NVO_RatesheetCharges rc INNER JOIN NVO_GeneralMaster ON NVO_GeneralMaster.ID = rc.PaymentModeID where rc.RRID = NVO_Booking.RRID and rc.ChargeCodeID in (1, 22)) FreightPayment   " +
 "  from NVO_Booking " +
  " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID = NVO_Booking.AgentID inner join NVO_BookingCntrTypes on NVO_BookingCntrTypes.BKgID = NVO_Booking.ID ";



            if (Status != "" && Status != "0" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BkgStatus=" + Status;
                else
                    strWhere += " and BkgStatus=" + Status;

            if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "2" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where NVO_Booking.Agentid = " + AgencyID.ToString();
                else
                    strWhere += " and NVO_Booking.Agentid = " + AgencyID.ToString();

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and (select top 1 ETD from NVO_VoyageRoute Where VoyageID = NVO_Booking.VesVoyID) between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and (select top 1 ETD from NVO_VoyageRoute Where VoyageID = NVO_Booking.VesVoyID) between '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
        public DataTable GetRRDetails(string RRID)
        {

            string _Query = " Select  DISTINCT CT.ID, RC.RRID,CT.ChgCode ,RC.ManifRate,RC.CustomerRate,RC.RateDiff,(select top(1) GeneralName from NVO_GeneralMaster where NVO_GeneralMaster.ID = PaymentModeID) as FreightTerm from NVO_RatesheetCharges  RC inner Join NVO_ChargeTB CT on CT.ID = RC.ChargeCodeID where RC.RRID = " + RRID + "  AND RC.ChargeTypeID = 1 and RC.TariffTypeID IN(135,136,137)";

            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetCntrPickupDetails(string BkgID)
        {

            string _Query = "   select BkgID,(select top 1 CntrNo from NVO_Containers where ID=CntrID) AS CntrNo,GrSWt as GrossWt, Size from NVO_BOLCntrDetails where BkgId =" + BkgID;

            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetBLDetails(string BkgID)
        {

            string _Query = " SELECT ID,BKGID,BLNumber,(select top(1) FinalInvoice from NVO_InvoiceCusBilling where BLID = NVO_BOL.ID ) InvoiceNo, " +
                        " convert(varchar, (SELECT TOP(1) SOBDate FROM NVO_BLRelease WHERE BLID = NVO_BOL.ID),103) As SOBDate FROM NVO_BOL WHERE BkgID=  " + BkgID;

            return Manag.GetViewData(_Query, "");
        }
        public void FeederWiseReport(string DtFrom, string DtTo, string Status, string User, string AgencyID)
        {
            //DataTable dtv = GetBookingSearchValues(DtFrom, DtTo, Status, AgencyID);
            //if (dtv.Rows.Count > 0)
            //{

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("FeederLifting");

            ws.Cells["A2"].Value = "Feeder Lifting Report";
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
            ws.Cells["B7"].Value = "POD/FEEDER";
            ws.Cells["C7"].Value = "XPF";
            ws.Cells["D7"].Value = "GFS";
            ws.Cells["E7"].Value = "SEA-LEAD";
            ws.Cells["F7"].Value = "SSL";
            ws.Cells["G7"].Value = "IWS";
            ws.Cells["H7"].Value = "MILAHA";
            ws.Cells["I7"].Value = "SARJAK";
            ws.Cells["J7"].Value = "SEABRIDGE";
            ws.Cells["K7"].Value = "TSS";
            ws.Cells["L7"].Value = "BTL";
            ws.Cells["M7"].Value = "FEEDER TECH";

            r = ws.Cells["A7:M7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            int sl = 1;

            int rw = 8;

            //for (int i = 0; i < dtv.Rows.Count; i++)
            //{
            ws.Cells["A" + rw].Value = sl;
            ws.Cells["B" + rw].Value = "";
            ws.Cells["C" + rw].Value = "";
            ws.Cells["D" + rw].Value = "";
            ws.Cells["E" + rw].Value = "";
            ws.Cells["F" + rw].Value = "";
            ws.Cells["G" + rw].Value = "";
            ws.Cells["H" + rw].Value = "";
            ws.Cells["I" + rw].Value = "";
            ws.Cells["J" + rw].Value = "";
            ws.Cells["K" + rw].Value = "";
            ws.Cells["L" + rw].Value = "";
            ws.Cells["M" + rw].Value = "";

            sl++;
            rw += 1;
            // }

            rw -= 1;

            ws.Cells["A7:M" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:M" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:M" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:M" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 20].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Feederwiselifting.xlsx");
            Response.End();

            // }

        }

        public void TDRReport(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID)
        {
            DataTable _dtv = GetTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);
            if (_dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TDR");

                ws.Cells["A2"].Value = "TERMINAL DEPARTURE REPORT";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:I2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);


                //Record Headers
                ws.Cells["A4"].Value = "PORT OF LOADING :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["A5"].Value = "NAME OF VESSEL :";
                ws.Cells["A5"].Style.Font.Bold = true;
                ws.Cells["A6"].Value = "ETA POL :";
                ws.Cells["A6"].Style.Font.Bold = true;
                ws.Cells["A7"].Value = "NEXT PORT :";
                ws.Cells["A7"].Style.Font.Bold = true;
                ws.Cells["A8"].Value = "CONTAINER OPERATOR :";
                ws.Cells["A8"].Style.Font.Bold = true;


                ws.Cells["B4"].Value = _dtv.Rows[0]["POL"].ToString();
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["B5"].Value = _dtv.Rows[0]["VesselName"].ToString();
                ws.Cells["B5"].Style.Font.Bold = true;
                ws.Cells["B6"].Value = _dtv.Rows[0]["ETA"].ToString();
                ws.Cells["B6"].Style.Font.Bold = true;
                ws.Cells["B7"].Value = _dtv.Rows[0]["NextPort"].ToString();
                ws.Cells["B7"].Style.Font.Bold = true;
                ws.Cells["B8"].Value = _dtv.Rows[0]["VesselOperator"].ToString();
                ws.Cells["B8"].Style.Font.Bold = true;

                ws.Cells["A4:B8"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A4:B8"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A4:B8"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A4:B8"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells["H4"].Value = "TERMINAL";
                ws.Cells["H4"].Style.Font.Bold = true;
                ws.Cells["H5"].Value = "VOYAGE NUMBER :";
                ws.Cells["H5"].Style.Font.Bold = true;
                ws.Cells["H6"].Value = "ETD POL :";
                ws.Cells["H6"].Style.Font.Bold = true;
                ws.Cells["H7"].Value = "ETA NEXT PORT :";
                ws.Cells["H7"].Style.Font.Bold = true;



                ws.Cells["I4"].Value = _dtv.Rows[0]["TerminalName"].ToString();
                ws.Cells["I4"].Style.Font.Bold = true;
                ws.Cells["I5"].Value = _dtv.Rows[0]["VoyageNo"].ToString();
                ws.Cells["I5"].Style.Font.Bold = true;
                ws.Cells["I6"].Value = _dtv.Rows[0]["ETD"].ToString();
                ws.Cells["I6"].Style.Font.Bold = true;
                ws.Cells["I7"].Value = _dtv.Rows[0]["NextPortETA"].ToString();
                ws.Cells["I7"].Style.Font.Bold = true;

                ws.Cells["H4:I8"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["H4:I8"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["H4:I8"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["H4:I8"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["A10"].Value = "S. No.";
                ws.Cells["B10"].Value = "CONTAINER#";
                ws.Cells["C10"].Value = "SIZE TYPE";
                ws.Cells["D10"].Value = "COMMODITY";
                ws.Cells["E10"].Value = "SERVICE";
                ws.Cells["F10"].Value = "TSPORT";
                ws.Cells["G10"].Value = "POD";
                ws.Cells["H10"].Value = "FINAL DESTINATION";
                ws.Cells["I10"].Value = "PICK UP DATE";
                ws.Cells["J10"].Value = "GW.KGS";
                ws.Cells["K10"].Value = "BLNUMBER";
                ws.Cells["L10"].Value = "POD AGENT";
                ws.Cells["M10"].Value = "VESSEL OPERATOR";
                ws.Cells["N10"].Value = "SLOT TERM";

                r = ws.Cells["A10:O10"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 11;

                DataTable _dtC = GetCntrDtlsTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);

                for (int k = 0; k < _dtC.Rows.Count; k++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = _dtC.Rows[k]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = _dtC.Rows[k]["Size"].ToString();
                    ws.Cells["D" + rw].Value = _dtC.Rows[k]["Commodity"].ToString();
                    ws.Cells["E" + rw].Value = _dtC.Rows[k]["ServiceType"].ToString();
                    ws.Cells["F" + rw].Value = _dtC.Rows[k]["TSPORT"].ToString();
                    ws.Cells["G" + rw].Value = _dtC.Rows[k]["POD"].ToString();
                    ws.Cells["H" + rw].Value = _dtC.Rows[k]["FPOD"].ToString();
                    ws.Cells["I" + rw].Value = _dtC.Rows[k]["PickUpDate"].ToString();
                    ws.Cells["J" + rw].Value = _dtC.Rows[k]["GrsWt"].ToString();
                    ws.Cells["K" + rw].Value = _dtC.Rows[k]["BLNumber"].ToString();
                    ws.Cells["L" + rw].Value = _dtC.Rows[k]["DestinationAgent"].ToString();
                    ws.Cells["M" + rw].Value = _dtC.Rows[k]["Operator"].ToString();
                    ws.Cells["N" + rw].Value = _dtC.Rows[k]["SlotTerm"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A10:O" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A10:O" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A10:O" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A10:O" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TDReport.xlsx");
                Response.End();

            }

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

            return Manag.GetViewData(strWhere, "");
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

            return Manag.GetViewData(strWhere, "");
        }


        public void ExportLoadingReportView(string DtFrom, string DtTo, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetExportLoadingView(DtFrom, DtTo, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ExportLoadingReport");

                ws.Cells["A2"].Value = "Export Loading Report";
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
                ws.Cells["B7"].Value = "Bkg.Loc";
                ws.Cells["C7"].Value = "Agency";
                ws.Cells["D7"].Value = "BL Number";
                ws.Cells["E7"].Value = "Vessel/Voyage";
                ws.Cells["F7"].Value = "ETD";
                ws.Cells["G7"].Value = "Terminal";
                ws.Cells["H7"].Value = "Booking Creation Date";
                ws.Cells["I7"].Value = "Shipper";
                ws.Cells["J7"].Value = "Booking Party";
                ws.Cells["K7"].Value = "POO";
                ws.Cells["L7"].Value = "Sale Pic";
                ws.Cells["M7"].Value = "POL";
                ws.Cells["N7"].Value = "POD";
                ws.Cells["O7"].Value = "FPOD";
                ws.Cells["P7"].Value = "Ts Port";
                ws.Cells["Q7"].Value = "20";
                ws.Cells["R7"].Value = "40";
                ws.Cells["S7"].Value = "Cntr Type";
                ws.Cells["T7"].Value = "Cntr No";
                ws.Cells["U7"].Value = "Commodity";
                ws.Cells["V7"].Value = "Charge Code";
                ws.Cells["W7"].Value = "Customer Rate";
                ws.Cells["X7"].Value = "Manifest Rate";
                ws.Cells["Y7"].Value = "Difference";
                ws.Cells["Z7"].Value = "Freight Term";
                ws.Cells["AA7"].Value = "VSL Operator";
                ws.Cells["AB7"].Value = "Slot";
                ws.Cells["AC7"].Value = "Booking Remarks";

                r = ws.Cells["A7:AC7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;
                int frowid = 0;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    frowid = rw;


                    ExcelRange rng = ws.Cells["A" + frowid + ":AC" + frowid];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Bkgloc"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["ETD"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Terminal"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["BkgDate"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["BkgParty"].ToString();

                    ws.Cells["K" + rw].Value = dtv.Rows[i]["POO"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["SalesPerson"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["PortofLoading"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["PlaceofDischarge"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["TranshimentPort"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["Size20"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["Size40"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["CntrType"].ToString();
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["U" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();
                    int stRow = rw;
                    int StRow1 = rw;
                    DataTable dtR = GetRRDetails(dtv.Rows[i]["RRID"].ToString());
                    rw = stRow;
                    int chgRow = 1;
                    for (int k = 0; k < dtR.Rows.Count; k++)
                    {

                        ws.Cells["V" + StRow1].Value = dtR.Rows[k]["ChgCode"].ToString();
                        ws.Cells["W" + StRow1].Value = dtR.Rows[k]["CustomerRate"].ToString();
                        ws.Cells["X" + StRow1].Value = dtR.Rows[k]["ManifRate"].ToString();
                        ws.Cells["Y" + StRow1].Value = dtR.Rows[k]["RateDiff"].ToString();
                        ws.Cells["Z" + StRow1].Value = dtR.Rows[k]["FreightTerm"].ToString();

                        StRow1++;
                        chgRow++;

                    }

                    ws.Cells["AA" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                    ws.Cells["AB" + rw].Value = dtv.Rows[i]["SlotCost"].ToString();
                 
                    ws.Cells["AC" + rw].Value = dtv.Rows[i]["Remarks"].ToString();
                    rw = StRow1;
                }

                rw -= 1;

                ws.Cells["A7:AC" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AC" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AC" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AC" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 40].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ExportLoadingReports.xlsx");
                Response.End();

            }

        }

        public DataTable GetExportLoadingView(string DtFrom, string DtTo, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_Booking.ID,BookingNo,isnull(RRNo,'') as RRNo,NVO_AgencyMaster.AgencyName,NVO_Booking.Remarks," +
                            " isnull(RRID, 0) as RRID,(select top 1 GeoLocation from NVO_GeoLocations Where id = NVO_AgencyMaster.GeoLocationID) AS Bkgloc, " +
                            " convert(varchar, (select top 1 ETD from NVO_VoyageRoute Where VoyageID = NVO_Booking.VesVoyID),103) As ETD, " +
                            "  (select top(1)(select top(1) TerminalName from NVO_TerminalMaster where ID = TerminalID) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID=NVO_Booking.VesVoyID) as Terminal, " +
                            " convert(varchar, BkgDate, 106) as BkgDate,  (select top(1) CustomerName from NVO_CustomerMaster where Id = ShipperID ) Shipper,  " +
                            " '' AS InvParty,'' AS IsOnline, isnull(ShipmentType, '') as ShipmentType ,BkgParty,POL, " +
                            " case when BkgStatus = 1 then 'DRAFT' else case when BKGStatus = 2  then 'CONFIRM' else case when BkgStatus = 3 then 'CANCELED' end end end as Status, " +
                            " (select top(1) PortName from NVO_PortMaster where Id = NVO_Booking.POLID) PortofLoading,  " +
                            " (select top(1) PortName from NVO_PortMaster where Id = NVO_Booking.PODID) PlaceofDischarge,   " +
                            "  (select top(1) PortName from NVO_PortMaster where Id = NVO_Booking.POOID ) POO," +
                            " (select top(1) PortName from NVO_PortMaster where Id = NVO_Booking.FPODID ) FPOD,  (select top(1) PortName from NVO_PortMaster where Id = TSPORTID ) TranshimentPort,   " +
                            " (select(select top 1(type + '-' + size) from NVO_tblCntrTypes where  Id = NVO_Containers.TypeID) from NVO_Containers where NVO_Containers.ID = NVO_BOLCntrDetails.CntrID) as CntrType, " +
                            "  case when CntrTypes in (1) then SlotAmt20 else case when CntrTypes in (2,3) then SlotAmt40 end end as SlotCost, " +
                            " (select top(1) CntrNo from NVO_Containers where NVO_Containers.Id = NVO_BOLCntrDetails.CntrID and NVO_BOLCntrDetails.BkgId = NVO_BookingCntrTypes.BKgID and NVO_Containers.TypeID = NVO_BookingCntrTypes.CntrTypes) as CntrNo,  " +
                            " isnull((select(select count(SizeID) from NVO_tblCntrTypes where SizeID = 1 and EQTypeID = 1 and Id = NVO_Containers.TypeID) from NVO_Containers where NVO_Containers.ID = NVO_BOLCntrDetails.CntrID),0) as Size20,   " +
                            " isnull((select(select count(SizeID) from NVO_tblCntrTypes where SizeID in (2, 3) and EQTypeID = 1 and Id = NVO_Containers.TypeID) from NVO_Containers where NVO_Containers.Id = NVO_BOLCntrDetails.CntrID),0) as Size40, " +
                            " (select top(1) ReleaseOrderNo from NVO_CROMaster where NVO_CROMaster.BkgID = NVO_Booking.Id) as CRONo, " +
                            " (select top(1) Qty from NVO_BookingCntrTypes where BKgID = NVO_Booking.ID ) Volume,   " +
                            " (select top 1 ReqQty from NVO_CRODetails inner join NVO_CROmaster on NVO_CROmaster.id = NVO_CRODetails.CROID Where BKGID = NVO_Booking.ID) As PickUpQty, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where Id = AgentID) CreatedAgency,  " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where Id = DestinationAgentID) DesAgency,  " +
                            " (select top(1) GeneralName from NVO_GeneralMaster where Id = NVO_Booking.ShipmentTypeID) ShipmentType,  " +
                            " (select top(1) Description from NVO_tblDLValues where Id = ServiceTypeID) ServiceType,  " +
                            " (select top(1) CustomerName from NVO_CustomerMaster where Id = ShipperID ) Shipper,    " +
                            " (select top(1) DepName from NVO_DepotMaster where Id = pickupdepotid ) PickupDepot,   " +
                            " (select top(1) Username from NVO_UserDetails where Id = PreparedBYID ) PreparedBy, " +
                            " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute " +
                            " where VoyageID = V.ID) from NVO_Voyage V where V.ID = VesVoyID )as VesVoy,   " +
                            " (select top(1) SlotRef from NVO_SLOTMaster where Id = SlotContractID ) SlotContract,   " +
                            " (select  upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where CID = SlotOperatorID ) SlotOperator, " +
                            " NVO_Booking.Remarks,isnull(SalesPerson, '') as SalesPerson,NVO_Booking.CommodityType " +
                            " from NVO_Booking " +
                            " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID = NVO_Booking.AgentID " +
                            " inner join NVO_BookingCntrTypes on NVO_BookingCntrTypes.BKgID = NVO_Booking.ID " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                            " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOL.ID";


            if (Status != "" && Status != "0" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BkgStatus=" + Status;
                else
                    strWhere += " and BkgStatus=" + Status;

            //if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "2" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

            //    if (strWhere == "")
            //        strWhere += _Query + " where NVO_Booking.Agentid = " + AgencyID.ToString();
            //    else
            //        strWhere += " and NVO_Booking.Agentid = " + AgencyID.ToString();

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and NVO_Booking.BkgDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and NVO_Booking.BkgDate between '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public void ImportLoadingReportView(string DtFrom, string DtTo, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetImportLandingView(DtFrom, DtTo, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ImportLandingReport");

                ws.Cells["A2"].Value = "Import Landing Report";
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
                ws.Cells["B7"].Value = "Geo Location";
                ws.Cells["C7"].Value = "Agency Name";
                ws.Cells["D7"].Value = "BL Number";
                ws.Cells["E7"].Value = "Load - Vessel/Voyage";
                ws.Cells["F7"].Value = "POL";
                ws.Cells["G7"].Value = "Sailing Date";
                ws.Cells["H7"].Value = "T/S Port";
                ws.Cells["I7"].Value = "POD";

                ws.Cells["J7"].Value = "FPOD";
                ws.Cells["K7"].Value = "Discharge - Vessel/Voyage";
                ws.Cells["L7"].Value = "ETA Date";
                ws.Cells["M7"].Value = "Shipper ";
                ws.Cells["N7"].Value = "Consignee";
                ws.Cells["O7"].Value = "Sales Person";
                ws.Cells["P7"].Value = "20'";
                ws.Cells["Q7"].Value = "40'";
                ws.Cells["R7"].Value = "Cntr Type";
                ws.Cells["S7"].Value = "Cntr No";
                ws.Cells["T7"].Value = "Commodity";
                ws.Cells["U7"].Value = "Charge Code";
                ws.Cells["V7"].Value = "Manifest Rate";
                ws.Cells["W7"].Value = "Freight Term";


                r = ws.Cells["A7:W7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 0;

                int rw = 8;
                int frowid = 0;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    frowid = rw;


                    ExcelRange rng = ws.Cells["A" + frowid + ":W" + frowid];
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    sl++;
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Bkgloc"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["BLVesVoy"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["SOBDate"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["TSPORT"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["ImpVesVoy"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["ETA"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Consignee"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["SalesPerson"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["Size20"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["Size40"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["CntrType"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();
                    int stRow = rw;
                    int StRow1 = rw;
                    if (dtv.Rows[i]["RRID"].ToString() != "")
                    {
                        DataTable dtR = GetRRDetails(dtv.Rows[i]["RRID"].ToString());
                        rw = stRow;
                        int chgRow = 1;
                        for (int k = 0; k < dtR.Rows.Count; k++)
                        {

                            ws.Cells["U" + StRow1].Value = dtR.Rows[k]["ChgCode"].ToString();
                            ws.Cells["V" + StRow1].Value = dtR.Rows[k]["ManifRate"].ToString();
                            ws.Cells["W" + StRow1].Value = dtR.Rows[k]["FreightTerm"].ToString();
                            StRow1++;
                            chgRow++;

                        }
                    }
                    else
                    {
                        int chgRow = 1;
                        ws.Cells["U" + StRow1].Value = "";
                        ws.Cells["V" + StRow1].Value = "";
                        ws.Cells["W" + StRow1].Value = "";

                        StRow1++;
                        chgRow++;
                    }

                    rw = StRow1;


                }

                rw -= 1;

                ws.Cells["A7:W" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:W" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:W" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:W" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 40].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ImportLandingReports.xlsx");
                Response.End();

            }

        }

        public DataTable GetImportLandingView(string DtFrom, string DtTo, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = " SELECT * FROM View_ImpLoadingReport ";




            //if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "2" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

            //    if (strWhere == "")
            //        strWhere += _Query + " where NVO_v_ImportBLView.AgencyID = " + AgencyID.ToString();
            //    else
            //        strWhere += " and NVO_v_ImportBLView.AgencyID = " + AgencyID.ToString();

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE ETA between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and ETAbetween '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public void ExportLoadListReport(string VesVoyID, string User, string AgencyID)
        {
            DataTable _dtv = GetExportVslVoyDetails(VesVoyID);
            if (_dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("LoadList");

                ws.Cells["A5"].Value = "LOAD LIST REPORT";
                ws.Cells["A5"].Style.Font.Bold = true;
                ws.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A5:K5"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);


                //Record Headers
                ws.Cells["A7"].Value = "DATE:";
                ws.Cells["A7"].Style.Font.Bold = true;
                ws.Cells["E7"].Value = "USER NAME :";
                ws.Cells["E7"].Style.Font.Bold = true;


                ws.Cells["B7"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["B7"].Style.Font.Bold = true;

                ws.Cells["F7"].Value = User;
                ws.Cells["F7"].Style.Font.Bold = true;

                ws.Cells["A7:F7"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F7"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F7"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F7"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells["A9"].Value = "VESSEL/VOYAGE :";
                ws.Cells["A9"].Style.Font.Bold = true;
               

                ws.Cells["B9"].Value = _dtv.Rows[0]["VesselName"].ToString();
               // ws.Cells["B10"].Value = _dtv.Rows[0]["LoadPort"].ToString();
               // ws.Cells["B11"].Value = _dtv.Rows[0]["NextPort"].ToString();



                ws.Cells["C9"].Value = "ETA :";
                ws.Cells["C9"].Style.Font.Bold = true;
                ws.Cells["E9"].Value = "ETD :";
                ws.Cells["E9"].Style.Font.Bold = true;

                int Rvsl = 10;
                DataTable _dtx = GetExportVslVoyPortDetails(VesVoyID);
                for( int z=0; z<_dtx.Rows.Count; z++)
                {   if (Rvsl == 10)
                    {
                        ws.Cells["A10"].Value = "LOAD PORT:";
                        ws.Cells["A10"].Style.Font.Bold = true;
                    }
                    else if (Rvsl == 11)
                    {
                        ws.Cells["A11"].Value = "NEXT PORT:";
                        ws.Cells["A11"].Style.Font.Bold = true;
                    }
                    ws.Cells["B"+ Rvsl].Value = _dtx.Rows[z]["PortName"].ToString();
                    ws.Cells["c" + Rvsl].Value = _dtx.Rows[z]["ETA"].ToString();
                    ws.Cells["E" + Rvsl].Value = _dtx.Rows[z]["ETD"].ToString();
                    Rvsl++;
                }

                ws.Cells["f9"].Value = "Operator Code :";
                ws.Cells["f9"].Style.Font.Bold = true;
                ws.Cells["g9"].Value = "Ship Call No :";
                ws.Cells["g9"].Style.Font.Bold = true;
                ws.Cells["h9"].Value = "Vessel ID:";
                ws.Cells["h9"].Style.Font.Bold = true;
                if (AgencyID == "4")
                    ws.Cells["f10"].Value = "GN";
                else
                    ws.Cells["f10"].Value = "GNV";

                ws.Cells["f10"].Style.Font.Bold = true;
                ws.Cells["g10"].Value = _dtv.Rows[0]["ShipCallno"].ToString();
                ws.Cells["g10"].Style.Font.Bold = true;
                ws.Cells["h10"].Value = _dtv.Rows[0]["VesselID"].ToString();
                ws.Cells["h10"].Style.Font.Bold = true;

                // ws.Cells["D10"].Value = _dtv.Rows[0]["LoadPortETA"].ToString();
                // ws.Cells["D11"].Value = _dtv.Rows[0]["NextPortETA"].ToString();

                //ws.Cells["E10"].Value = "ETD:";
                //ws.Cells["E10"].Style.Font.Bold = true;
                //ws.Cells["E11"].Value = "ETD:";
                //ws.Cells["E11"].Style.Font.Bold = true;

                // ws.Cells["F10"].Value = _dtv.Rows[0]["LoadPortETD"].ToString();
                // ws.Cells["F11"].Value = _dtv.Rows[0]["NextPortETD"].ToString();



                //ws.Cells["A13"].Value = "SLOT OPERATOR";
                //ws.Cells["A13"].Style.Font.Bold = true;
                //ws.Cells["B13"].Value = "";

                ws.Cells["A9:h14"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A9:h14"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A9:h14"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A9:h14"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells["A15"].Value = "S.No.";
                ws.Cells["B15"].Value = "CONTAINER#";
                ws.Cells["C15"].Value = "SIZE TYPE";
                ws.Cells["D15"].Value = "ISOCODE";
                ws.Cells["E15"].Value = "COMMODITY";
                ws.Cells["F15"].Value = "GR.WEIGHT";
                ws.Cells["G15"].Value = "BL NUMBER";
                ws.Cells["H15"].Value = "POL";
                ws.Cells["I15"].Value = "POD";
                //ws.Cells["J15"].Value = "MANIFEST RATE";
                //ws.Cells["K15"].Value = "SELLING RATE";
                ws.Cells["J15"].Value = "DESTINATION AGENT";
                ws.Cells["K15"].Value = "DESTINATION ADDRESS";


                r = ws.Cells["A15:K15"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                int sl = 1;

                int rw = 16;
                DataTable _dtC = GetCntrDtlsVesvoyValues(VesVoyID);

                for (int k = 0; k < _dtC.Rows.Count; k++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = _dtC.Rows[k]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = _dtC.Rows[k]["Size"].ToString();
                    ws.Cells["D" + rw].Value = _dtC.Rows[k]["ISOCode"].ToString();
                    ws.Cells["E" + rw].Value = _dtC.Rows[k]["Commodity"].ToString();
                    ws.Cells["F" + rw].Value = _dtC.Rows[k]["Grwt"];
                    ws.Cells["G" + rw].Value = _dtC.Rows[k]["BookingNo"].ToString();
                    ws.Cells["H" + rw].Value = _dtC.Rows[k]["POL"].ToString();
                    ws.Cells["I" + rw].Value = _dtC.Rows[k]["POD"].ToString();
                    //ws.Cells["J" + rw].Value = _dtC.Rows[k]["ManifRate"];
                    //ws.Cells["K" + rw].Value = _dtC.Rows[k]["CustomerRate"];

                    ws.Cells["J" + rw].Value = _dtC.Rows[k]["DestCustomerName"];
                    ws.Cells["K" + rw].Value = _dtC.Rows[k]["DestAddress"];
                    

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A15:K" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A15:K" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A15:K" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A15:K" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=LoadListReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetExportVslVoyDetails(string VesVoy)
        {
            string strWhere = "";

            string _Query = "select VOY.ID, (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + '-' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID)  " +
              " from NVO_Voyage V where V.ID = VOY.ID)as VesselName ,(select top(1) PortName from NVO_VoyageRoute  inner join NVO_PortMaster On NVO_PortMaster.ID = NVO_VoyageRoute.PortID " +
              " where NVO_VoyageRoute.VoyageID = VOY.ID order by NVO_VoyageRoute.RID desc)as LoadPort, (select top(1) PortName from NVO_VoyageRoute " +
              " inner join NVO_PortMaster On NVO_PortMaster.ID = NVO_VoyageRoute.PortID where NVO_VoyageRoute.VoyageID = VOY.ID order by NVO_VoyageRoute.RID asc)as NextPort, " +
              " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VOY.ID order by NVO_VoyageRoute.RID desc),105) as LoadPortETA, " +
              " convert(varchar, (select top(1)  ETD from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VOY.ID order by NVO_VoyageRoute.RID desc),105) as LoadPortETD, " +
              " convert(varchar, (select top(1)  ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VOY.ID order by NVO_VoyageRoute.RID asc),105) as NextPortETA, " +
              " convert(varchar, (select top(1)  ETD from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = VOY.ID order by NVO_VoyageRoute.RID asc),105) as NextPortETD, " +
              " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID=282 and NVO_VoyageNotesDtls.VoyageID=VOY.ID) as ShipCallno, " +
              " (select top(1) VesselID from NVO_VesselMaster where NVO_VesselMaster.ID=VOY.VesselID) VesselID " +
              " from NVO_Voyage VOY ";

            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  VOY.ID=" + VesVoy;
                else

                    strWhere += " AND VOY.ID=" + VesVoy;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetCntrDtlsVesvoyValues(string VesVoyID)
        {

            //string _Query = "select NVO_Booking.ID,(Select top 1 CntrNo from NVO_Containers WHERE NVO_Containers.ID=NVO_BOLCntrDetails.CntrID) As CntrNo, " +
            //" (select top 1 Size FROM NVO_tblCntrTypes WHERE NVO_tblCntrTypes.ID = NVO_Containers.TypeID) As Size, (select top 1 ISOCode FROM NVO_tblCntrTypes WHERE NVO_tblCntrTypes.ID = NVO_Containers.ISOCodeID) As ISOCode, " +
            //" (select top 1 GeneralName FROM NVO_GeneralMaster WHERE NVO_Booking.CommodityTypeID = NVO_GeneralMaster.ID) As Commodity,  NVO_BOLCntrDetails.GrsWt as Grwt,NVO_Booking.BookingNo,NVO_Booking.POL,NVO_Booking.POD, " +
            //" isnull((Select  DISTINCT  top 1 SUM (RC.ManifRate) As ManifRate from NVO_RatesheetCharges  RC inner Join NVO_ChargeTB CT on CT.ID = RC.ChargeCodeID  where RC.RRID = NVO_Booking.RRID AND RC.TariffTypeID IN(135) ),0) As ManifRate, " +
            //" isnull((Select  DISTINCT  top 1 SUM(RC.CustomerRate) As ManifRate from NVO_RatesheetCharges RC inner Join NVO_ChargeTB CT on CT.ID = RC.ChargeCodeID where RC.RRID = NVO_Booking.RRID AND RC.TariffTypeID IN(135) ),0) As CustomerRate, " +
            //" (select top(1) AgencyName from NVO_AgencyMaster where ID= NVO_BOL.AgencyID) as DestCustomerName,(select top(1) Address from NVO_AgencyMaster where Id = NVO_BOL.AgencyID) as DestAddress " +
            //" from NVO_Booking inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOL.ID " +
            //" inner join NVO_Containers on NVO_Containers.ID = NVO_BOLCntrDetails.CntrID where NVO_BOL.BLVesVoyID = " + VesVoyID + "";

            string _Query = " select distinct NVO_Booking.ID,CntrNo,(select top 1 Size FROM NVO_tblCntrTypes WHERE NVO_tblCntrTypes.ID = NVO_Containers.TypeID) As Size, " +
                            " (select top 1 ISOCode FROM NVO_tblCntrTypes WHERE NVO_tblCntrTypes.ID = NVO_Containers.ISOCodeID) As ISOCode, " +
                            " (select top 1 GeneralName FROM NVO_GeneralMaster WHERE NVO_Booking.CommodityTypeID = NVO_GeneralMaster.ID) As Commodity,  '' as Grwt, " +
                            " NVO_Booking.BookingNo,NVO_Booking.POL,NVO_Booking.POD, " +
                            " isnull((Select  DISTINCT  top 1 SUM(RC.ManifRate) As ManifRate from NVO_RatesheetCharges RC inner Join NVO_ChargeTB CT on CT.ID = RC.ChargeCodeID  where RC.RRID = NVO_Booking.RRID AND RC.TariffTypeID IN(135)),0) As ManifRate, " +
                            " isnull((Select  DISTINCT  top 1 SUM(RC.CustomerRate) As ManifRate from NVO_RatesheetCharges RC inner Join NVO_ChargeTB CT on CT.ID = RC.ChargeCodeID where RC.RRID = NVO_Booking.RRID AND RC.TariffTypeID IN(135) ),0) As CustomerRate, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_Booking.DestinationAgentID) as DestCustomerName,  " +
                            " (select top(1) Address from NVO_AgencyMaster where Id = NVO_Booking.DestinationAgentID) as DestAddress " +
                            " from NVO_Booking " +
                            " inner join NVO_ContainerTxns on NVO_ContainerTxns.BLNumber = NVO_Booking.ID " +
                            " inner join NVO_Containers on NVO_Containers.Id = NVO_ContainerTxns.ContainerID " +
                            " where NVO_Booking.VesVoyID = " + VesVoyID + "";
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetExportVslVoyPortDetails(string VesVoy)
        {
            string strWhere = "";

            string _Query = " select(select top(1) PortName from NVO_PortMaster where NVO_PortMaster.Id = NVO_VoyageRoute.PortID) as PortName, " +
                            " Convert(varchar, ETA, 103) as ETA,Convert(varchar, ETD, 103) as ETD " +
                            " from NVO_VoyageRoute";

            if (VesVoy != "" && VesVoy != "0" && VesVoy != "null" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " where  VoyageID=" + VesVoy;
                else

                    strWhere += " AND VoyageID=" + VesVoy;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public void WeeklyExportPickupSummaryReport(string DtFrom, string DtTo, string User, string AgencyID)
        {
            //DataTable _dtv = GetExportVslVoyDetails(VesVoyID);
            //if (_dtv.Rows.Count > 0)
            //{

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("LoadList");

            //ExcelWorksheet ws = pck.Workbook.Worksheets[1];
            Color colCost;
            ws.Name = "REPORT"; //Setting Sheet's name
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
            int MaxCol = 18;
            int RowIndex = 2;
            ws.Cells[RowIndex, 1].Value = "PICK UP DETAILS FOR WEEK # 39";
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
            //ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            RowIndex++;

            ws.Cells[RowIndex, 1, RowIndex, 6].Value = "BASIC DETAILS";
            ws.Cells[RowIndex, 1, RowIndex, 6].Merge = true;
            ws.Cells[RowIndex, 1, RowIndex, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 12.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;


            colCost = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
            ws.Cells[RowIndex, 7, RowIndex, 10].Value = "BOOKING ISSUED";
            ws.Cells[RowIndex, 7, RowIndex, 10].Merge = true;
            ws.Cells[RowIndex, 7, RowIndex, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 7, RowIndex, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 7, RowIndex, 10].Style.Fill.BackgroundColor.SetColor(colCost);

            colCost = System.Drawing.ColorTranslator.FromHtml("#FFDAE1");
            ws.Cells[RowIndex, 11, RowIndex, 14].Value = "CANCELLED";
            ws.Cells[RowIndex, 11, RowIndex, 14].Merge = true;
            ws.Cells[RowIndex, 11, RowIndex, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 11, RowIndex, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 11, RowIndex, 14].Style.Fill.BackgroundColor.SetColor(colCost);

            colCost = System.Drawing.ColorTranslator.FromHtml("#D3FAFA");
            ws.Cells[RowIndex, 15, RowIndex, 18].Value = "CONFIRMED";
            ws.Cells[RowIndex, 15, RowIndex, 18].Merge = true;
            ws.Cells[RowIndex, 15, RowIndex, 18].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 15, RowIndex, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 15, RowIndex, 18].Style.Fill.BackgroundColor.SetColor(colCost);

            RowIndex++;
            ws.Cells[RowIndex, 1].Value = "S.No.";
            ws.Cells[RowIndex, 2].Value = "VESSEL.";
            ws.Cells[RowIndex, 3].Value = "ETD";
            ws.Cells[RowIndex, 4].Value = "CUSTOMER";
            ws.Cells[RowIndex, 5].Value = "BOOKING/BL NUMBER";

            ws.Cells[RowIndex, 6].Value = "POD";
            ws.Cells[RowIndex, 7].Value = "20";
            ws.Cells[RowIndex, 8].Value = "40";
            ws.Cells[RowIndex, 9].Value = "20RF";
            ws.Cells[RowIndex, 10].Value = "40RF";

            ws.Cells[RowIndex, 11].Value = "20";
            ws.Cells[RowIndex, 12].Value = "40";
            ws.Cells[RowIndex, 13].Value = "20RF";
            ws.Cells[RowIndex, 14].Value = "40RF";

            ws.Cells[RowIndex, 15].Value = "20";
            ws.Cells[RowIndex, 16].Value = "40";
            ws.Cells[RowIndex, 17].Value = "20RF";
            ws.Cells[RowIndex, 18].Value = "40RF";
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.WrapText = true;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
            RowIndex++; int InvCount = 0;
            int rw_bgn = 0;
            rw_bgn = RowIndex;
            DataTable _dt = GetWeelyExportBookingView(DtFrom, DtTo, AgencyID);
            for (int x = 0; x < _dt.Rows.Count; x++)
            {
                InvCount++;
                ws.Cells[RowIndex, 1].Value = InvCount;
                ws.Cells[RowIndex, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 2].Value = _dt.Rows[x]["VesVoy"].ToString().ToUpper();
                ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 3].Value = _dt.Rows[x]["ETDDate"].ToString().ToUpper();
                ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 4].Value = _dt.Rows[x]["BkgParty"].ToString().ToUpper();
                ws.Cells[RowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 5].Value = _dt.Rows[x]["BookingNo"].ToString().ToUpper();
                ws.Cells[RowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 6].Value = _dt.Rows[x]["POD"].ToString().ToUpper();
                ws.Cells[RowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                ws.Cells[RowIndex, 7].Value = _dt.Rows[x]["DGP20"];
                ws.Cells[RowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 8].Value = _dt.Rows[x]["DGP40"];
                ws.Cells[RowIndex, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 9].Value = _dt.Rows[x]["DFR20"];
                ws.Cells[RowIndex, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 10].Value = _dt.Rows[x]["DFR40"];
                ws.Cells[RowIndex, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 11].Value = _dt.Rows[x]["FGP20"];
                ws.Cells[RowIndex, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 12].Value = _dt.Rows[x]["FGP40"];
                ws.Cells[RowIndex, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 13].Value = _dt.Rows[x]["FFR20"];
                ws.Cells[RowIndex, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 14].Value = _dt.Rows[x]["FFR40"];
                ws.Cells[RowIndex, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 15].Value = _dt.Rows[x]["CGP20"];
                ws.Cells[RowIndex, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 16].Value = _dt.Rows[x]["CGP40"];
                ws.Cells[RowIndex, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 17].Value = _dt.Rows[x]["CFR20"];
                ws.Cells[RowIndex, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 18].Value = _dt.Rows[x]["CFR40"];
                ws.Cells[RowIndex, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                RowIndex++;
            }

            string styleRev = "#FFDAE1";
            string styleCost = "#DAFFED";
            string styleprofit = "#E1E199";
            rw_bgn = 7;

            ws.Cells["G" + RowIndex].Formula = "SUM(G5:G" + (RowIndex - 1) + ")";
            ws.Cells["h" + RowIndex].Formula = "SUM(H5:H" + (RowIndex - 1) + ")";
            ws.Cells["i" + RowIndex].Formula = "SUM(I5:I" + (RowIndex - 1) + ")";
            ws.Cells["j" + RowIndex].Formula = "SUM(J5:J" + (RowIndex - 1) + ")";
            ws.Cells["K" + RowIndex].Formula = "SUM(K5:K" + (RowIndex - 1) + ")";
            ws.Cells["L" + RowIndex].Formula = "SUM(L5:L" + (RowIndex - 1) + ")";
            ws.Cells["M" + RowIndex].Formula = "SUM(M5:M" + (RowIndex - 1) + ")";
            ws.Cells["N" + RowIndex].Formula = "SUM(N5:N" + (RowIndex - 1) + ")";
            ws.Cells["O" + RowIndex].Formula = "SUM(O5:O" + (RowIndex - 1) + ")";
            ws.Cells["P" + RowIndex].Formula = "SUM(P5:P" + (RowIndex - 1) + ")";
            ws.Cells["Q" + RowIndex].Formula = "SUM(Q5:Q" + (RowIndex - 1) + ")";
            ws.Cells["R" + RowIndex].Formula = "SUM(R5:R" + (RowIndex - 1) + ")";


            rw_bgn = 7;
            colCost = System.Drawing.ColorTranslator.FromHtml(styleCost);
            ws.Cells["G" + (rw_bgn - 2) + ":J" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["G" + (rw_bgn - 2) + ":J" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

            colCost = System.Drawing.ColorTranslator.FromHtml(styleRev);
            ws.Cells["K" + (rw_bgn - 2) + ":N" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["K" + (rw_bgn - 2) + ":N" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

            colCost = System.Drawing.ColorTranslator.FromHtml(styleprofit);
            ws.Cells["O" + (rw_bgn - 2) + ":R" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["O" + (rw_bgn - 2) + ":R" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);


            ws.Cells["A3:R" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:R" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:R" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:R" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells.AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=WeeklyExportReport.xlsx");
            Response.End();

        }

        //}

        public DataTable GetWeelyExportBookingView(string DtFrom, string DtTo, string AgencyID)
        {
            string strWhere = "";
            string _Query = " SELECT BookingNo,VesVoy,BkgParty,POD,BkgDate,(select top(1) convert(varchar, ETD, 103) from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID) as ETDDate, BkgStatus, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (1, 6, 11)  and BkgStatus = 1 and Bkg.Id = NVO_Booking.Id) as DGP20, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (2, 3, 7, 12)  and BkgStatus = 1 and Bkg.Id = NVO_Booking.Id) as DGP40, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (8)  and BkgStatus = 1 and Bkg.Id = NVO_Booking.Id) as DFR20, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (9)  and BkgStatus = 1 and Bkg.Id = NVO_Booking.Id) as DFR40, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (1, 6, 11)  and BkgStatus = 2 and Bkg.Id = NVO_Booking.Id) as FGP20, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (2, 3, 7, 12)  and BkgStatus = 2 and Bkg.Id = NVO_Booking.Id) as FGP40, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (8)  and BkgStatus = 2 and Bkg.Id = NVO_Booking.Id) as FFR20, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (9)  and BkgStatus = 2 and Bkg.Id = NVO_Booking.Id) as FFR40, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (1, 6, 11)  and BkgStatus = 3 and Bkg.Id = NVO_Booking.Id) as CGP20, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (2, 3, 7, 12)  and BkgStatus = 3 and Bkg.Id = NVO_Booking.Id) as CGP40, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (8)  and BkgStatus = 3 and Bkg.Id = NVO_Booking.Id) as CFR20, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes inner join NVO_Booking Bkg on Bkg.ID = NVO_BookingCntrTypes.BKgID where CntrTypes in (9)  and BkgStatus = 3 and Bkg.Id = NVO_Booking.Id) as CFR40 " +
                            " from NVO_Booking ";

            strWhere += _Query + " WHERE AgentID=" + AgencyID;

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE BkgDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and BkgDate between '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere + " order by BkgDate", "");
        }


        public void WeeklyCustomerTuesSummaryReport(string DtFrom, string DtTo, string User, string AgencyID)
        {
            //DataTable _dtv = GetExportVslVoyDetails(VesVoyID);
            //if (_dtv.Rows.Count > 0)
            //{

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("LoadList");

            //ExcelWorksheet ws = pck.Workbook.Worksheets[1];
            Color colCost;
            ws.Name = "REPORT"; //Setting Sheet's name
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
            int MaxCol = 7;
            int RowIndex = 2;
            ws.Cells[RowIndex, 1].Value = "TOP CUSTOMER WEEK";
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
            //ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            RowIndex++;

            ws.Cells[RowIndex, 1, RowIndex, 2].Value = "BASIC DETAILS";
            ws.Cells[RowIndex, 1, RowIndex, 2].Merge = true;
            ws.Cells[RowIndex, 1, RowIndex, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 12.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;


            colCost = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
            ws.Cells[RowIndex, 3, RowIndex, 7].Value = "";
            ws.Cells[RowIndex, 3, RowIndex, 7].Merge = true;
            ws.Cells[RowIndex, 3, RowIndex, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 3, RowIndex, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 3, RowIndex, 7].Style.Fill.BackgroundColor.SetColor(colCost);



            RowIndex++;
            ws.Cells[RowIndex, 1].Value = "S.No.";
            ws.Cells[RowIndex, 2].Value = "CUSTOMER";
            ws.Cells[RowIndex, 3].Value = "20";
            ws.Cells[RowIndex, 4].Value = "40";
            ws.Cells[RowIndex, 5].Value = "20RF";
            ws.Cells[RowIndex, 6].Value = "40RF";
            ws.Cells[RowIndex, 7].Value = "SALES PIC";

            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.WrapText = true;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
            ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
            RowIndex++; int InvCount = 0;
            int rw_bgn = 0;
            rw_bgn = RowIndex;
            int CountBkgNo = 0;
            DataTable _dt = GetWeelyExportCustomerView(DtFrom, DtTo, AgencyID);
            for (int x = 0; x < _dt.Rows.Count; x++)
            {
                InvCount++;
                ws.Cells[RowIndex, 1].Value = InvCount;
                ws.Cells[RowIndex, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 2].Value = _dt.Rows[x]["BkgParty"].ToString().ToUpper();
                ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                ws.Cells[RowIndex, 3].Value = _dt.Rows[x]["FGP20"];
                ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 4].Value = _dt.Rows[x]["FGP40"];
                ws.Cells[RowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 5].Value = _dt.Rows[x]["FFR20"];
                ws.Cells[RowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 6].Value = _dt.Rows[x]["FFR40"];
                ws.Cells[RowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[RowIndex, 7].Value = _dt.Rows[x]["SalesPerson"].ToString().ToUpper();
                ws.Cells[RowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                CountBkgNo += Int32.Parse(_dt.Rows[x]["BkgNo"].ToString().ToUpper());
                RowIndex++;
            }

            string styleRev = "#FFDAE1";
            string styleCost = "#DAFFED";
            string styleprofit = "#E1E199";
            rw_bgn = 7;

            ws.Cells[RowIndex, 2].Value = "TOTAL";
            ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 2].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 2].Style.Font.Bold = true;
            ws.Cells[RowIndex, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[RowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);

            ws.Cells["C" + RowIndex].Formula = "SUM(C5:C" + (RowIndex - 1) + ")";
            ws.Cells["D" + RowIndex].Formula = "SUM(D5:F" + (RowIndex - 1) + ")";
            ws.Cells["E" + RowIndex].Formula = "SUM(E5:E" + (RowIndex - 1) + ")";
            ws.Cells["F" + RowIndex].Formula = "SUM(F5:F" + (RowIndex - 1) + ")";


            rw_bgn = 7;
            colCost = System.Drawing.ColorTranslator.FromHtml(styleRev);
            ws.Cells["C" + (rw_bgn - 2) + ":F" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C" + (rw_bgn - 2) + ":F" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);
            ws.Cells["A3:g" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            RowIndex++;
            RowIndex++;
            RowIndex++;

            ws.Cells[RowIndex, 2].Value = "NO OF TEUS";
            ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 2].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 2].Style.Font.Bold = true;
            ws.Cells[RowIndex, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


            ws.Cells["c" + RowIndex].Formula = "=SUM(C" + (RowIndex - 3) + "+ D" + (RowIndex - 3) + "*2 +E" + (RowIndex - 3) + "+F" + (RowIndex - 3) + "*2)";
            ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 3].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 3].Style.Font.Bold = true;
            ws.Cells[RowIndex, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            RowIndex++;
            ws.Cells[RowIndex, 2].Value = "NO. OF CONTAINERS";
            ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 2].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 2].Style.Font.Bold = true;
            ws.Cells[RowIndex, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            ws.Cells["c" + RowIndex].Formula = "=SUM(C" + (RowIndex - 4) + "+ D" + (RowIndex - 4) + " +E" + (RowIndex - 4) + "+F" + (RowIndex - 4) + ")";
            ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 3].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 3].Style.Font.Bold = true;
            ws.Cells[RowIndex, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            RowIndex++;
            ws.Cells[RowIndex, 2].Value = "NO. OF CUSTOMERS";

            ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 2].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 2].Style.Font.Bold = true;
            ws.Cells[RowIndex, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            ws.Cells["c" + RowIndex].Value = (RowIndex - 10);
            ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 3].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 3].Style.Font.Bold = true;
            ws.Cells[RowIndex, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            RowIndex++;
            ws.Cells[RowIndex, 2].Value = "NO. OF BLS";
            ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 2].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 2].Style.Font.Bold = true;
            ws.Cells[RowIndex, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            ws.Cells["c" + RowIndex].Value = CountBkgNo;
            ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            ws.Cells[RowIndex, 3].Style.Font.Size = 14.0F;
            ws.Cells[RowIndex, 3].Style.Font.Bold = true;
            ws.Cells[RowIndex, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            RowIndex++;
            ws.Cells["A3:g" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;



            ws.Cells["A3:g" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A3:g" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells.AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=WeeklyExportReport.xlsx");
            Response.End();

        }


        public DataTable GetWeelyExportCustomerView(string DtFrom, string DtTo, string AgencyID)
        {
            string strWhere = "";
            string _Query = " select BkgParty,SalesPerson,SUM(FGP20) as FGP20,SUM(FGP40) as FGP40,SUM(FFR20) as FFR20,SUM(FFR40) as FFR40,count(BookingNo) as BkgNo " +
                            " from v_NVO_CustomerBookingTues ";


            strWhere += _Query + " WHERE AgentID=" + AgencyID;

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE BkgDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and BkgDate between '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere + " group by BkgParty,SalesPerson", "");
        }

        public ActionResult WeeklyFeederLifitingReport(string DtFrom, string DtTo, string User, string AgencyID)
        {
            WeeklyFeederLifitingReportvalues(DtFrom, DtTo, User, AgencyID);
            return View();
        }
        public void WeeklyFeederLifitingReportvalues(string DtFrom, string DtTo, string User, string AgencyID)
        {
            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("WeeklyFeederLifitingReport");
            ws.Name = "WeeklyFeederLifitingReport"; //Setting Sheet's name
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
            int RowIndex = 2;
            int ColIndex = 3;
            ws.Cells[RowIndex, 1].Value = "SR No.";
            ws.Cells[RowIndex, 2].Value = "FEEDER OPERATOR";

            DataTable _dtP = GetWeeklyFeederLifingPortValues(DtFrom, DtTo, AgencyID);
            for (int j = 0; j < _dtP.Rows.Count; j++)
            {
                ws.Cells[RowIndex, ColIndex].Value = _dtP.Rows[j]["POD"].ToString();
                ColIndex++;
            }

            //DataTable dtx = GetWeeklyFeederLifingSlotValues(DtFrom, DtTo);
            //for (int j = 0; j < dtx.Rows.Count; j++)
            //{
            //    ws.Cells[RowIndex, ColIndex].Value = dtx.Rows[j]["Slot"].ToString();
            //    ColIndex++;
            //}


            ws.Cells[RowIndex, ColIndex].Value = "TOTAL";
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Font.Size = 10.0F;
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Font.Bold = true;
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.WrapText = true;
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Font.Size = 10.0F;
            ws.Cells[RowIndex, 1, RowIndex, ColIndex].Style.Font.Bold = true;

            RowIndex++;
            DataTable dtx = GetWeeklyFeederLifingSlotValues(DtFrom, DtTo, AgencyID);
            //DataTable _dtP = GetWeeklyFeederLifingPortValues(DtFrom, DtTo);
            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                ws.Cells[RowIndex, 1].Value = (RowIndex - 1);
                ws.Cells[RowIndex, 2].Value = dtx.Rows[i]["Slot"].ToString();
                RowIndex++;
            }
            int ColIndexz = 3;

            for (int z = 0; z < _dtP.Rows.Count; z++)
            {
                int RowIndexz = 3;
                for (int y = 0; y < dtx.Rows.Count; y++)
                {
                    DataTable _dtz = GetWeeklyFeederLifingPortSlotValues(DtFrom, DtTo, _dtP.Rows[z]["PODID"].ToString(), dtx.Rows[y]["Slot"].ToString(), AgencyID);
                    if (_dtz.Rows.Count > 0)
                    {
                        if (_dtz.Rows[0]["TotalCntr"].ToString() != "")
                            ws.Cells[RowIndexz, ColIndexz].Value = _dtz.Rows[0]["TotalCntr"];
                        else
                            ws.Cells[RowIndexz, ColIndexz].Value = "";
                        RowIndexz++;
                    }
                }
                ColIndexz++;
            }

            //for (int y = 0; y < dtx.Rows.Count; y++)
            //{
            //    int RowIndexz = 3;
            //    for (int z = 0; z < _dtP.Rows.Count; z++)
            //    {
            //        DataTable _dtz = GetWeeklyFeederLifingPortSlotValues(DtFrom, DtTo, _dtP.Rows[z]["POLID"].ToString(), dtx.Rows[y]["Slot"].ToString());
            //        if (_dtz.Rows.Count > 0)
            //        {
            //            if (_dtz.Rows[0]["TotalCntr"].ToString() != "")
            //                ws.Cells[RowIndexz, ColIndexz].Value = _dtz.Rows[0]["TotalCntr"];
            //            else
            //                ws.Cells[RowIndexz, ColIndexz].Value = "";
            //            RowIndexz++;
            //        }
            //    }
            //    ColIndexz++;
            //}
            int Rows = 3;
            string styleColor = "#DAFFED";
            Color colCost;
            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                ws.Cells[Rows, ColIndex].Formula = "SUM(C" + Rows + ":" + GetColumnName(ColIndex - 2) + Rows + ")";

                colCost = System.Drawing.ColorTranslator.FromHtml(styleColor);
                ws.Cells[Rows - 1, ColIndex, Rows, ColIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                ws.Cells[Rows - 1, ColIndex, Rows, ColIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[Rows - 1, ColIndex, Rows, ColIndex].Style.Fill.BackgroundColor.SetColor(colCost);
                Rows++;
            }


            ws.Cells["A1:" + Rows].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:" + Rows].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:" + Rows].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A1:" + Rows].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells.AutoFitColumns();

            #region Add2ndSheet

            ws = pck.Workbook.Worksheets.Add("FeederLiftingSummary");
            ws.Cells["B2"].Value = "S. No.";
            ws.Cells["C2"].Value = "BOOKING/BL NO";
            ws.Cells["D2"].Value = "BOOKING DATE";
            ws.Cells["E2"].Value = "CONFIRM DATE";
            ws.Cells["F2"].Value = "BOOKING STATUS";
            ws.Cells["G2"].Value = "POL";
            ws.Cells["H2"].Value = "POD";
            ws.Cells["I2"].Value = "FEEDER NAME";
            ws.Cells["J2"].Value = "VESSEL/VOYAGE";
            //ws.Cells["K2"].Value = "NO.OF.20";
            //ws.Cells["L2"].Value = "NO.OF.40";
            ws.Cells["K2"].Value = "CntrNo";
            ws.Cells["L2"].Value = "Container Types";

            ExcelRange r = ws.Cells["b2:L2"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            int SRowIndex = 3;
            int MaxRow2 = 1;
            DataTable _dty = GetWeeklyFeederReportSummaryValues(DtFrom, DtTo, AgencyID);
            for (int f = 0; f < _dty.Rows.Count; f++)
            {
                ws.Cells["B" + SRowIndex].Value = MaxRow2++;
                ws.Cells["C" + SRowIndex].Value = _dty.Rows[f]["BookingNo"].ToString();
                ws.Cells["D" + SRowIndex].Value = _dty.Rows[f]["BkgDate"].ToString();
                ws.Cells["E" + SRowIndex].Value = _dty.Rows[f]["ConfirmDate"].ToString();
                ws.Cells["F" + SRowIndex].Value = _dty.Rows[f]["BkgStatusV"].ToString();
                ws.Cells["G" + SRowIndex].Value = _dty.Rows[f]["POL"].ToString();
                ws.Cells["H" + SRowIndex].Value = _dty.Rows[f]["POD"].ToString();
                ws.Cells["I" + SRowIndex].Value = _dty.Rows[f]["Slot"].ToString();
                ws.Cells["J" + SRowIndex].Value = _dty.Rows[f]["VesVoy"].ToString();
                //ws.Cells["K" + SRowIndex].Value = _dty.Rows[f]["GP20"];
                //ws.Cells["L" + SRowIndex].Value = _dty.Rows[f]["GP40"];
                ws.Cells["K" + SRowIndex].Value = _dty.Rows[f]["CntrNo"].ToString();
                ws.Cells["L" + SRowIndex].Value = _dty.Rows[f]["size"].ToString();
                SRowIndex++;
            }
            ws.Cells["B2:N" + SRowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["B2:N" + SRowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["B2:N" + SRowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["B2:N" + SRowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells.AutoFitColumns();
            #endregion



            string _fileName = "WeeklyFeederLiftingReport " + Session.SessionID + ".xlsx";
            string _filePath = Server.MapPath("~/PreXL/");
            Byte[] bin = pck.GetAsByteArray();
            System.IO.File.WriteAllBytes(_filePath + "\\" + _fileName, bin);
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + _fileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(_filePath + "\\" + _fileName);
            Response.Flush();
            Response.End();
        }


        public DataTable GetWeeklyFeederLifingPortValues(string FromDate, string ToDate, string AgentID)
        {

            string strWhere = "";
            string _Query = "select distinct (select top(1) PortName  from NVO_PortMaster where ID =PODID) as POD,PODID from NVO_Booking " +
                            " inner join NVO_ContainerTxns on NVO_ContainerTxns.BLNumber=NVO_Booking.ID " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID";
            strWhere += _Query + "  where BkgStatus = 2 and NVO_ContainerTxns.StatusCode='FB' and AgentID=" + AgentID;
            if (FromDate != "" && FromDate != "undefined" || ToDate != "" && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");

        }

        public DataTable GetWeeklyFeederLifingSlotValues(string FromDate, string ToDate, string AgentID)
        {
            //     string _Query = " select distinct(select top(1) CustomerName from view_CustomerFeederName where CID = SlotOperatorID) as Slot, " +
            // " (select top(1) ID from view_CustomerFeederName where CID = SlotOperatorID) as SlotID " +
            string strWhere = "";
            string _Query = " select distinct slot from view_FeederNamemultilpevalues";

            strWhere += _Query + " where StatusCode='FB' and AgentID = " + AgentID;

            if (FromDate != "" && FromDate != "undefined" || ToDate != "" && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");

        }

        public DataTable GetWeeklyFeederLifingPortSlotValues(string FromDate, string ToDate, string PortID, string FeederName, string AgentID)
        {

            string strWhere = "";
            string _Query = " select sum(isnull(GP20,0)) + (sum(isnull(GP40,0))*2) as TotalCntr  from View_FeederLifingReport";
            strWhere += _Query + " where  PODID=" + PortID + " and Slot='" + FeederName + "' and StatusCode = 'FB' and AgentID = " + AgentID;

            if (FromDate != "" && FromDate != "undefined" || ToDate != "" && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");

        }
        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }


        public ActionResult WeeklyLifitingMISReportAllExcel(string DtFrom, string DtTo, string AgencyID)
        {
            BindWeeklyLifitingMISALLReportGetValue(AgencyID, DtFrom, DtTo);
            return View();
        }

        public void BindWeeklyLifitingMISALLReportGetValue(string AgentID, string FromDate, string ToDate)
        {
            using (ExcelPackage _xlPV = new ExcelPackage())
            {
                _xlPV.Workbook.Properties.Author = "WEEKLY LIFTING REPORT";
                _xlPV.Workbook.Properties.Title = "WEEKLY LIFTING REPORT";
                _xlPV.Workbook.Worksheets.Add("WEEKLY LIFTING REPORT");
                ExcelWorksheet ws = _xlPV.Workbook.Worksheets[1];
                Color colCost;
                ws.Name = "REPORT"; //Setting Sheet's name
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
                int MaxCol = 30;
                int RowIndex = 2;
                ws.Cells[RowIndex, 1].Value = "WEEKLY LIFTING REPORT";
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
                //ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 14.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                RowIndex++;

                ws.Cells[RowIndex, 1, RowIndex, 14].Value = "";
                ws.Cells[RowIndex, 1, RowIndex, 14].Merge = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 12.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;


                colCost = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
                ws.Cells[RowIndex, 20, RowIndex, 25].Value = "REVENUE";
                ws.Cells[RowIndex, 20, RowIndex, 25].Merge = true;
                ws.Cells[RowIndex, 20, RowIndex, 25].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 20, RowIndex, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 20, RowIndex, 25].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#FFDAE1");
                ws.Cells[RowIndex, 26, RowIndex, 28].Value = "COST";
                ws.Cells[RowIndex, 26, RowIndex, 28].Merge = true;
                ws.Cells[RowIndex, 26, RowIndex, 28].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 26, RowIndex, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 26, RowIndex, 28].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#E1E199");
                ws.Cells[RowIndex, 29, RowIndex, 29].Value = "PROFIT";
                ws.Cells[RowIndex, 29, RowIndex, 29].Merge = true;
                ws.Cells[RowIndex, 29, RowIndex, 29].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 29, RowIndex, 29].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 29, RowIndex, 29].Style.Fill.BackgroundColor.SetColor(colCost);



                //ws.Cells[RowIndex, 34].Value = "";



                RowIndex++;
                ws.Cells[RowIndex, 1].Value = "SR No.";
                ws.Cells[RowIndex, 2].Value = "Location.";
                ws.Cells[RowIndex, 3].Value = "Agent Name.";
                ws.Cells[RowIndex, 4].Value = "Rate Request No.";
                ws.Cells[RowIndex, 5].Value = "BLNo";
                ws.Cells[RowIndex, 6].Value = "Container No";
                ws.Cells[RowIndex, 7].Value = "POL";
                ws.Cells[RowIndex, 8].Value = "POD";

                ws.Cells[RowIndex, 9].Value = "T/S";

                ws.Cells[RowIndex, 10].Value = "FPOD";
                ws.Cells[RowIndex, 11].Value = "Vessel /Voy ";
                ws.Cells[RowIndex, 12].Value = "Terminal Name";
                ws.Cells[RowIndex, 13].Value = "Slot Operator";
                ws.Cells[RowIndex, 14].Value = "ETD'";
                ws.Cells[RowIndex, 15].Value = "Commodity Type ";
                ws.Cells[RowIndex, 16].Value = "20'";
                ws.Cells[RowIndex, 17].Value = "40'";
                ws.Cells[RowIndex, 18].Value = "Container Types";
                ws.Cells[RowIndex, 19].Value = "Lease Type";

                ws.Cells[RowIndex, 20].Value = "Ocen Freight ($)";
                ws.Cells[RowIndex, 21].Value = "BAF ($)";
                ws.Cells[RowIndex, 22].Value = "DGS ($)";
                ws.Cells[RowIndex, 23].Value = "Surcharges ($)";
                ws.Cells[RowIndex, 24].Value = "Load Port THC ($)";
                // ws.Cells[RowIndex, 25].Value = "Discharge Port THC ($)";
                // ws.Cells[RowIndex, 26].Value = "HDL/TPT ($)";
                //  ws.Cells[RowIndex, 27].Value = "MISC ($)";
                ws.Cells[RowIndex, 25].Value = "Total Revenue ($)";

                ws.Cells[RowIndex, 26].Value = "SLOT ($)";
                //  ws.Cells[RowIndex, 30].Value = "T/S THC COST ($)";
                // ws.Cells[RowIndex, 31].Value = "T/S COMMISSION ($)";
                // ws.Cells[RowIndex, 32].Value = "T/S SLOT COST ($)";
                ws.Cells[RowIndex, 27].Value = "LOAD PORT PHC ($)";
                //ws.Cells[RowIndex, 34].Value = "DISCHARGE PORT PHC ($)";
                //ws.Cells[RowIndex, 35].Value = "LOAD COMMISSION ($)";
                //ws.Cells[RowIndex, 36].Value = "DISCHARGE COMMISSION ($)";
                //ws.Cells[RowIndex, 37].Value = "MISC COST ($)";
                //ws.Cells[RowIndex, 38].Value = "LOG. FEE ($)";
                ws.Cells[RowIndex, 28].Value = "TOTAL COST ($)";
                ws.Cells[RowIndex, 29].Value = "PROFIT / LOSS ($)";
                ws.Cells[RowIndex, 30].Value = "FRIGHT TERMS";


                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.WrapText = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                RowIndex++; int InvCount = 0;
                int rw_bgn = 0;
                rw_bgn = RowIndex;
                DataTable dtx = GetALLCntrDtlsTerminalDepPdfValues(AgentID, FromDate, ToDate);
                for (int i = 0; i < dtx.Rows.Count; i++)
                {
                    decimal RevenusAmt = decimal.Parse(dtx.Rows[i]["OFTt"].ToString()) + decimal.Parse(dtx.Rows[i]["DGSt"].ToString()) + decimal.Parse(dtx.Rows[i]["BAFt"].ToString()) + decimal.Parse(dtx.Rows[i]["SURCHARGESt"].ToString()) + decimal.Parse(dtx.Rows[i]["LOADTHCt"].ToString());
                    decimal CostAmt = decimal.Parse(dtx.Rows[i]["SlotAmtt"].ToString()) + decimal.Parse(dtx.Rows[i]["LOADPHCt"].ToString());


                    InvCount++;
                    ws.Cells[RowIndex, 1].Value = InvCount;
                    ws.Cells[RowIndex, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 2].Value = dtx.Rows[i]["Locations"].ToString();
                    ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 3].Value = dtx.Rows[i]["AgentName"].ToString();
                    ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 4].Value = dtx.Rows[i]["RRNO"].ToString();
                    ws.Cells[RowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 5].Value = dtx.Rows[i]["BookingNo"].ToString();
                    ws.Cells[RowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 6].Value = dtx.Rows[i]["CntrNo"].ToString();
                    ws.Cells[RowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 7].Value = dtx.Rows[i]["POL"].ToString();
                    ws.Cells[RowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 8].Value = dtx.Rows[i]["POD"].ToString();
                    ws.Cells[RowIndex, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 9].Value = dtx.Rows[i]["TSPORT"].ToString();
                    ws.Cells[RowIndex, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 10].Value = dtx.Rows[i]["FPOD"].ToString();
                    ws.Cells[RowIndex, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 11].Value = dtx.Rows[i]["VesVoy"].ToString();
                    ws.Cells[RowIndex, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 12].Value = dtx.Rows[i]["Terminal"].ToString();
                    ws.Cells[RowIndex, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 13].Value = dtx.Rows[i]["SlotOperator"].ToString();
                    ws.Cells[RowIndex, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 14].Value = dtx.Rows[i]["ETDDate"].ToString();
                    ws.Cells[RowIndex, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 15].Value = dtx.Rows[i]["Commodity"].ToString();
                    ws.Cells[RowIndex, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 16].Value = dtx.Rows[i]["Size20"].ToString();
                    ws.Cells[RowIndex, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 17].Value = dtx.Rows[i]["Size40"].ToString();
                    ws.Cells[RowIndex, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 18].Value = dtx.Rows[i]["TypeSize"].ToString();
                    ws.Cells[RowIndex, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 19].Value = dtx.Rows[i]["LeasTerms"].ToString();
                    ws.Cells[RowIndex, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 20].Value = dtx.Rows[i]["OFTt"];
                    ws.Cells[RowIndex, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 21].Value = dtx.Rows[i]["BAFt"];
                    ws.Cells[RowIndex, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 22].Value = dtx.Rows[i]["DGSt"];
                    ws.Cells[RowIndex, 22].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 23].Value = dtx.Rows[i]["SURCHARGESt"];
                    ws.Cells[RowIndex, 23].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 24].Value = dtx.Rows[i]["LOADTHCt"];
                    ws.Cells[RowIndex, 24].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 25].Value = dtx.Rows[i]["DESTTHCt"];
                    //ws.Cells[RowIndex, 25].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 26].Value = dtx.Rows[i]["HLDTPTt"];
                    //ws.Cells[RowIndex, 26].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["MISCt"];
                    //ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //  ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["MISRevAmt"];
                    // ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);




                    ws.Cells[RowIndex, 25].Value = RevenusAmt;
                    ws.Cells[RowIndex, 25].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 26].Value = dtx.Rows[i]["SlotAmtt"];
                    ws.Cells[RowIndex, 26].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //   ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["TSCOSTt"];
                    //      ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ////ws.Cells[RowIndex, 30].Value = // dtx.Rows[i]["TSCOMMt"];
                    //if (dtx.Rows[i]["TSPORT"].ToString() == "")
                    //    ws.Cells[RowIndex, 31].Value = 0.00;
                    //else
                    //    ws.Cells[RowIndex, 31].Value = 10.00;
                    //ws.Cells[RowIndex, 31].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 32].Value = dtx.Rows[i]["TSSlotCost"];
                    //ws.Cells[RowIndex, 32].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["LOADPHCt"];
                    ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 34].Value = dtx.Rows[i]["DESTPHCt"];
                    //ws.Cells[RowIndex, 34].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    //decimal ExCommt = decimal.Parse(dtx.Rows[i]["EXCOMMt"].ToString());
                    //if (ExCommt <= 10)
                    //    ws.Cells[RowIndex, 35].Value = 10;
                    //else
                    //    ws.Cells[RowIndex, 35].Value = dtx.Rows[i]["EXCOMMt"];


                    //ws.Cells[RowIndex, 35].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    ////ws.Cells[RowIndex, 35].Value = dtx.Rows[i]["EXCOMMt"];
                    ////ws.Cells[RowIndex, 35].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //decimal ExICommt = decimal.Parse(dtx.Rows[i]["ICOMMt"].ToString());

                    //if (ExICommt <= 10)
                    //    ws.Cells[RowIndex, 36].Value = 10;
                    //else
                    //    ws.Cells[RowIndex, 36].Value = dtx.Rows[i]["ICOMMt"];


                    //ws.Cells[RowIndex, 36].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ////ws.Cells[RowIndex, 36].Value = dtx.Rows[i]["ICOMMt"];
                    ////ws.Cells[RowIndex, 36].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 37].Value = dtx.Rows[i]["MISCCOSTt"];
                    //ws.Cells[RowIndex, 37].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 38].Value = 5.00;
                    //ws.Cells[RowIndex, 38].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 28].Value = CostAmt;
                    ws.Cells[RowIndex, 28].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 29].Value = (RevenusAmt - CostAmt);
                    ws.Cells[RowIndex, 29].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 30].Value = dtx.Rows[i]["FreightTerms"].ToString();
                    ws.Cells[RowIndex, 30].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    RowIndex++;

                }


                string styleRev = "#FFDAE1";
                string styleCost = "#DAFFED";
                string styleprofit = "#E1E199";
                rw_bgn = 7;
                colCost = System.Drawing.ColorTranslator.FromHtml(styleCost);
                ws.Cells["T" + (rw_bgn - 2) + ":y" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["T" + (rw_bgn - 2) + ":y" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(styleRev);
                ws.Cells["z" + (rw_bgn - 2) + ":Ab" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["z" + (rw_bgn - 2) + ":Ab" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(styleprofit);
                ws.Cells["Ac" + (rw_bgn - 2) + ":Ac" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["Ac" + (rw_bgn - 2) + ":Ac" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                ws.Cells["A3:Ac" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:Ac" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:Ac" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:Ac" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells[1, 1, 23, 30].AutoFitColumns();

                string _filePath = Server.MapPath("~/PreXL/");
                string _fileName = "WeeklyLiftingReport " + Session.SessionID + ".xlsx";
                Byte[] bin = _xlPV.GetAsByteArray();
                System.IO.File.WriteAllBytes(_filePath + "\\" + _fileName, bin);
                Response.Clear();
                Response.AppendHeader("content-disposition", "attachment; filename=" + _fileName);
                Response.ContentType = "application/octet-stream";
                Response.WriteFile(_filePath + "\\" + _fileName);
                Response.Flush();
                Response.End();


            }

        }

        public DataTable GetALLCntrDtlsTerminalDepPdfValues(string AgentID, string FromDate, string ToDate)
        {
            string strWhere = "";

            string _Query = " select AgentID,Id,RRID,RRNO,BookingNo,POL,POD,FPOD,TSPORT,CntrNo,VesVoy,Terminal,SlotOperator,TypeSize,ETDDate,ETDDatev,FreightTerms,	 " +
                            " (select top(1)(select top(1) GeoLocation from NVO_GeoLocations  where NVO_GeoLocations.ID=NVO_AgencyMaster.GeoLocationID) from NVO_AgencyMaster where NVO_AgencyMaster.Id = NVO_V_MiSViewALLDataReport.AgentID) as Locations, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where NVO_AgencyMaster.ID = NVO_V_MiSViewALLDataReport.AgentID) as AgentName, " +
                          " Size20,Size40,SizeRF40,SizeOT40,LeasTerms,Commodity," +
                          " cast(isnull((OFT / (case when OFTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.OFTCurrId)  end)),0) as decimal(10,2)) as OFTt," +
                          " cast(isnull((BAF / (case when BAFCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.BAFCurrId)  end)),0) as decimal(10,2)) as BAFt, " +
                          " cast(isnull((SURCHARGES / (case when SURCHARGESCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SURCHARGESCurrId)  end)),0) as decimal(10,2))  as SURCHARGESt, " +
                          " cast(isnull((DGS / (case when DGS_CCURRId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DGS_CCURRId)  end)),0) as decimal(10,2))  as DGSt, " +
                          " cast(isnull((LOADTHC/ (case when LOADTHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOADTHCCurrId)  end)),0) as decimal(10,2))  as LOADTHCt, " +
                          " cast(isnull((DESTTHC / (case when DESTTHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DESTTHCCurrId)  end)),0) as decimal(10,2))  as DESTTHCt, " +
                          " cast(isnull((HLDTPT / (case when HLDTPTCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.HLDTPTCCurrId)  end)),0) as decimal(10,2))  as HLDTPTt, " +
                          " cast(isnull((MISC / (case when MISCCcurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCcurrId)  end)),0) as decimal(10,2)) as MISCt1, " +

                          " cast(isnull((SUV_MISC / (case when SUV_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SUV_MISCCCurrId)  end)),0)  + " +
                          " isnull((LOLO_MISC / (case when LOLO_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOLO_MISCCCurrId)  end)),0)  + " +
                          " isnull((WASH_MISC / (case when WASH_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.WASH_MISCCCurrId)  end)),0)  + " +
                          " isnull((CMC_MISC / (case when CMC_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.CMC_MISCCCurrId)  end)),0) as decimal(10,2)) as MISCt, " +
                          " SlotAmt  as SlotAmtt," +
                          //" cast(isnull((SlotAmt / (case when SlotCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SlotCurrId)  end)),0) as decimal(10,2)) as SlotAmtt," +
                          " cast(isnull((TSCOST / (case when TSCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOSTCurrId)  end)),0)  as decimal(10,2)) as TSCOSTt, " +
                          " TSSlotCost, " +
                          // " cast(isnull((TSSlotCost / (case when TSCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOSTCurrId)  end)),0)  as decimal(10,2)) as TSSlotCost,"+

                          " cast(isnull((EXCOMM / (case when EXCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.EXCOMMCurrId)  end)),0) as decimal(10,2)) as EXCOMMt, " +
                          " cast(isnull((ICOMM / (case when ICOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.ICOMMCurrId)  end)),0) as decimal(10,2)) as ICOMMt, " +

                          " cast(isnull((TSCOMM / (case when TSCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOMMCurrId)  end)),0) as decimal(10,2)) as TSCOMMt, " +
                          " cast(isnull((LOADPHC / (case when LOADPHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOADPHCCurrId)  end)),0)  as decimal(10,2)) as LOADPHCt, " +
                          " cast(isnull((DESTPHC / (case when DESTPHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DESTPHCCurrId)  end)),0)  as decimal(10,2)) as DESTPHCt, " +
                          " cast(isnull((MISCCOST / (case when MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCOSTCurrId)  end)),0)  as decimal(10,2)) as MISCCOSTt1, " +

                          " cast(isnull((SUV_MISCCOST / (case when SUV_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SUV_MISCCOSTCurrId)  end)),0)  + " +
                          " isnull((LOLO_MISCCOST / (case when LOLO_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOLO_MISCCOSTCurrId)  end)),0)  + " +
                          " isnull((WASH_MISCCOST / (case when WASH_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.WASH_MISCCOSTCurrId)  end)),0) as decimal(10, 2)) as MISCCOSTt, " +
                          " MISRevAmt, " +
                          " cast(isnull((MISCostAmt / (case when MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCOSTCurrId)  end)),0)  as decimal(10,2)) as MISCostAmtv " +

                          " from NVO_V_MiSViewALLDataReport_New_All NVO_V_MiSViewALLDataReport";

            strWhere += _Query + " Where AgentID=" + AgentID;

            if (FromDate != "" && FromDate != "undefined" || ToDate != "" && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,ETDDatev,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,ETDDatev,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }


        public DataTable GetWeeklyFeederReportSummaryValues(string FromDate, string ToDate, string AgentID)
        {

            string strWhere = "";
            string _Query = " select BookingNo,BkgStatus,convert(varchar,BkgDate, 103) as BkgDate,convert(varchar,DtMovement, 103) as ConfirmDate,POL,POD, " +
                            " VesVoy,case when BkgStatus = 1 then 'DRAFT' else 'CONFIRM' end as BkgStatusV, " +
                            " (select top(1) CustomerName from view_CustomerFeederName where CID = SlotOperatorID) as Slot, " +
                            " (select sum(Qty) from NVO_BookingCntrTypes where CntrTypes in(1, 4, 6, 8, 10, 11) and BKgID = NVO_Booking.ID) as GP20, " +
                            " (select sum(Qty) from NVO_BookingCntrTypes where CntrTypes in(2, 3, 5, 7, 9, 12) and BKgID = NVO_Booking.ID) as GP40, " +
                            " NVO_Containers.CntrNo, (select top(1) Size from NVO_tblCntrTypes where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as Size " +
                            " from NVO_Booking " +
                            " inner join NVO_ContainerTxns on NVO_ContainerTxns.BLNumber=NVO_Booking.ID " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID";

            strWhere += _Query + "  where BkgStatus = 2 and NVO_ContainerTxns.StatusCode='FB'  and AgentID=" + AgentID;

            if (FromDate != "" && FromDate != "undefined" || ToDate != "" && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");

        }
    }

}
