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
    public class BookingExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }
        public void CreateBookingExcel(string POL, string RRNumber, string BookingNo, string BkgParty, string POD, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetBookingSearchValues(POL, RRNumber, BkgParty, BookingNo, POD, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("BookingSearch");

                ws.Cells["A2"].Value = "Booking Search List";
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
                ws.Cells["B7"].Value = "BookingNo";
                ws.Cells["C7"].Value = "RR Number";
                ws.Cells["D7"].Value = "Booking Date";
                ws.Cells["E7"].Value = "Booking Party";
                ws.Cells["F7"].Value = "Shipment Type";
                ws.Cells["G7"].Value = "POL";
                ws.Cells["H7"].Value = "POD";
                ws.Cells["I7"].Value = "FPOD";
                ws.Cells["J7"].Value = "Transhipment";
                ws.Cells["K7"].Value = "Loading Agency";
                ws.Cells["L7"].Value = "Destination Agency";
                ws.Cells["M7"].Value = "Service Type";
                ws.Cells["N7"].Value = "Shipper";
                ws.Cells["O7"].Value = "Pickup Depot";
                ws.Cells["P7"].Value = "PreparedBy";
                ws.Cells["Q7"].Value = "VesVoy";
                ws.Cells["R7"].Value = "Slot Contract";
                ws.Cells["S7"].Value = "Slot Operator";
                r = ws.Cells["A7:S7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["RRNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["BkgDate"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["BkgParty"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["PortofLoading"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["PlaceofDischarge"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["TranshimentPort"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["CreatedAgency"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["DesAgency"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["ServiceType"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["PickupDepot"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["PreparedBy"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["SlotContract"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:S" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:S" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:S" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:S" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=BookingSearchList.xlsx");
                Response.End();

            }

        }

        public DataTable GetBookingSearchValues(string POL, string RRNumber, string BookingNo, string BkgParty, string POD, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = "select ID,BookingNo,RRNo,convert(varchar, BkgDate,106) as BkgDate,ShipmentType,BkgParty,POL,case when BkgStatus =1 then 'DRAFT' else case when BKGStatus= 2  then 'CONFIRM' else case when BkgStatus = 3 then 'CANCELED' end end end as Status,(select top(1) PortName from NVO_PortMaster where Id = POLID) PortofLoading, " +
             " (select top(1) PortName from NVO_PortMaster where Id = PODID) PlaceofDischarge, " +
             " (select top(1) PortName from NVO_PortMaster where Id = FPODID ) FPOD, " +
            " (select top(1) PortName from NVO_PortMaster where Id = TSPORTID ) TranshimentPort, " +
            " (select top(1) AgencyName from NVO_AgencyMaster where Id = AgentID) CreatedAgency, " +
         " (select top(1) AgencyName from NVO_AgencyMaster where Id = DestinationAgentID) DesAgency, " +
        " (select top(1) GeneralName from NVO_GeneralMaster where Id = ShipmentTypeID) ShipmentType, " +
  " (select top(1) Description from NVO_tblDLValues where Id = ServiceTypeID) ServiceType, " +
   " (select top(1) CustomerName from NVO_CustomerMaster where Id = ShipperID ) Shipper, " +
   " (select top(1) DepName from NVO_DepotMaster where Id = pickupdepotid ) PickupDepot, " +
 " (select top(1) Username from NVO_UserDetails where Id = PreparedBYID ) PreparedBy,(Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID)  + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID)from NVO_Voyage V where V.ID = VesVoyID )as VesVoy, " +
" (select top(1) SlotRef from NVO_SLOTMaster where Id = SlotContractID ) SlotContract, (select top(1) CustomerName from NVO_CustomerMaster where Id = SlotOperatorID )SlotOperator from NVO_Booking ";

            if (BookingNo != "" && BookingNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BookingNo like '%" + BookingNo + "%'";
                else
                    strWhere += " and BookingNo like '%" + BookingNo + "%'";

            if (RRNumber != "" && RRNumber != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where RRNO like '%" + RRNumber + "%'";
                else
                    strWhere += " and RRNO like '%" + RRNumber + "%'";

            if (BkgParty != "" && BkgParty != "0" && BkgParty != "null" && BkgParty != "?" && BkgParty != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BkgPartyID = " + BkgParty;
                else
                    strWhere += " and BkgPartyID = " + BkgParty;

            if (POL != "" && POL != "0" && POL != "null" && POL != "?" && POL != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where POLID= " + POL;
                else
                    strWhere += " and POLID= " + POL;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?" && POD != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PODID= " + POD;
                else
                    strWhere += " and PODID= " + POD;

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


            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }


        public void CreateBOLExcel(string POL, string RRNumber, string BookingNo, string BkgParty, string POD, string Status, string User,string AgencyID)
        {
            DataTable dtv = GetBOLSearchValues(POL, RRNumber, BkgParty, BookingNo, POD, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("BOLSearch");

                ws.Cells["A2"].Value = "BOL Search List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:I2"];
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
                ws.Cells["B7"].Value = "BL Number";
                ws.Cells["C7"].Value = "RR Number";
                ws.Cells["D7"].Value = "Booking Party";
                ws.Cells["E7"].Value = "Shipment Type";
                ws.Cells["F7"].Value = "POL";
                ws.Cells["G7"].Value = "POD";
                ws.Cells["H7"].Value = "FPOD";
                ws.Cells["I7"].Value = "Transhipment Port";
                //ws.Cells["J7"].Value = "Transhipment";
                //ws.Cells["K7"].Value = "Loading Agency";
                //ws.Cells["L7"].Value = "Destination Agency";
                //ws.Cells["M7"].Value = "Service Type";
                //ws.Cells["N7"].Value = "Shipper";
                //ws.Cells["O7"].Value = "Pickup Depot";
                //ws.Cells["P7"].Value = "PreparedBy";
                //ws.Cells["Q7"].Value = "VesVoy";
                //ws.Cells["R7"].Value = "Slot Contract";
                //ws.Cells["S7"].Value = "Slot Operator";
                r = ws.Cells["A7:S7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["RRNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["BkgParty"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["TSPORT"].ToString();
                   // ws.Cells["I" + rw].Value = dtv.Rows[i]["TSPORT"].ToString();
                    //ws.Cells["J" + rw].Value = dtv.Rows[i]["TranshimentPort"].ToString();
                    //ws.Cells["K" + rw].Value = dtv.Rows[i]["CreatedAgency"].ToString();
                    //ws.Cells["L" + rw].Value = dtv.Rows[i]["DesAgency"].ToString();
                    //ws.Cells["M" + rw].Value = dtv.Rows[i]["ServiceType"].ToString();
                    //ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    //ws.Cells["O" + rw].Value = dtv.Rows[i]["PickupDepot"].ToString();
                    //ws.Cells["P" + rw].Value = dtv.Rows[i]["PreparedBy"].ToString();
                    //ws.Cells["Q" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    //ws.Cells["R" + rw].Value = dtv.Rows[i]["SlotContract"].ToString();
                    //ws.Cells["S" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:I" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=BOLSearchList.xlsx");
                Response.End();

            }

        }

        public DataTable GetBOLSearchValues(string POL, string RRNumber, string BookingNo, string BkgParty, string POD, string Status, string AgentId)
        {
            string strWhere = "";
            string _Query = " Select case when NVO_BOL.ID is null  then 0 else NVO_BOL.ID end as ID,NVO_Booking.Id as BkgID,BookingNo as BLNumber,case when isnull(NVO_BOL.Status,0)= 0 then 'DRAFT'  when NVO_BOL.Status = 2 then 'CONFIRMED' when BkgStatus = 3 then 'CANCELLED'  end as Status , BookingNo,RRNo,BkgParty,ShipmentType,POL,POD,FPOD,TSPORT from NVO_Booking left outer join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID where BkgStatus=2 ";

            if (BookingNo != "" && BookingNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and BookingNo like '%" + BookingNo + "%'";
                else
                    strWhere += " and BookingNo like '%" + BookingNo + "%'";

            if (RRNumber != "" && RRNumber != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and RRNO like '%" + RRNumber + "%'";
                else
                    strWhere += " and RRNO like '%" + RRNumber + "%'";

            if (BkgParty != "" && BkgParty != "0" && BkgParty != "null" && BkgParty != "?" && BkgParty != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and BkgPartyID = " + BkgParty;
                else
                    strWhere += " and BkgPartyID = " + BkgParty;

            if (POL != "" && POL != "0" && POL != "null" && POL != "?" && POL != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and POLID= " + POL;
                else
                    strWhere += " and POLID= " + POL;

            if (POD != "" && POD != "0" && POD != "null" && POD != "?" && POD != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and PODID= " + POD;
                else
                    strWhere += " and PODID= " + POD;

           if ( AgentId != "" && AgentId != "0" && AgentId != "2" && AgentId != "undefined" && AgentId != null)

                if (strWhere == "")
                    strWhere += _Query + " and NVO_Booking.AgentID = " + AgentId;
                else
                    strWhere += " and NVO_Booking.AgentID = " + AgentId;
            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
    }
}