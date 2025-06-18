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
    public class RRExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }
        public void CreateRRExcel(string POL, string RRNumber, string POD, string BookingParty, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetRRSearchValues(POL, RRNumber, POD, BookingParty, Status, AgencyID);


            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("RRSearch");

            ws.Cells["A2"].Value = "Rate Request Search List";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:AI2"];
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
            ws.Cells["B7"].Value = "RR Number";
            ws.Cells["C7"].Value = "Created On";
            ws.Cells["D7"].Value = "Agency Name";
            ws.Cells["E7"].Value = "Shipment Type";
            ws.Cells["F7"].Value = "POO";
            ws.Cells["G7"].Value = "POL";
            ws.Cells["H7"].Value = "POD";
            ws.Cells["I7"].Value = "FPOD";
            ws.Cells["J7"].Value = "Route Type";
            ws.Cells["K7"].Value = "Transhipment";
            ws.Cells["L7"].Value = "Service Type";
            ws.Cells["M7"].Value = "Booking Party";
            ws.Cells["N7"].Value = "Shipper";
            ws.Cells["O7"].Value = "Status";
            ws.Cells["P7"].Value = "Created By";
            ws.Cells["Q7"].Value = "Approved By";
            ws.Cells["R7"].Value = "Approved On";
            ws.Cells["S7"].Value = "Rejected On";
            ws.Cells["T7"].Value = "Rejected By";
            ws.Cells["U7"].Value = "Rejected Reason";
            ws.Cells["V7"].Value = "Requoted On";
            ws.Cells["W7"].Value = "Requoted By";
            ws.Cells["X7"].Value = "Requoted Reason";
            ws.Cells["Y7"].Value = "Special Instruction-slot Details& Services";
            ws.Cells["Z7"].Value = " ExpFreeDays";
            ws.Cells["AA7"].Value = " ImpFreeDays";
            ws.Cells["AB7"].Value = " Equip Type";
            ws.Cells["AC7"].Value = "Charge Description";
            ws.Cells["AD7"].Value = "Currency";
            ws.Cells["AE7"].Value = "Payment Mode";
            ws.Cells["AF7"].Value = "Requested Rate";
            ws.Cells["AG7"].Value = "Manifest Rate";
            ws.Cells["AH7"].Value = "Customer Rate";
            ws.Cells["AI7"].Value = "Rate Different";
            r = ws.Cells["A7:AI7"];
            r.Style.Font.Bold = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            int sl = 1;

            int rw = 8;
            int frowid = 0;
            if (dtv.Rows.Count > 0)
            {
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    frowid = rw;
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["RatesheetNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Datev"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["PortofOrgin"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["PortofLoading"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["PlaceofDischarge"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["RouteTypes"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["TranshimentPort"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["ServiceType"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["BkgParty"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["CreatedBy"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["ApprovedBy"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["DtAppr"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["RejectedBy"].ToString();
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["DtRej"].ToString();
                    ws.Cells["U" + rw].Value = dtv.Rows[i]["RjRemarks"].ToString();
                    ws.Cells["V" + rw].Value = dtv.Rows[i]["DtRequote"].ToString();
                    ws.Cells["W" + rw].Value = dtv.Rows[i]["RequotedBy"].ToString();
                    ws.Cells["X" + rw].Value = dtv.Rows[i]["ReRemarks"].ToString();
                    ws.Cells["Y" + rw].Value = dtv.Rows[i]["Remarks"].ToString();
                    ws.Cells["Z" + rw].Value = dtv.Rows[i]["ExpFreeDays"];
                    ws.Cells["AA" + rw].Value = dtv.Rows[i]["ImpFreeDays"];

                    sl++;
                    int stRow = rw;
                    int StRow1 = rw;

                    DataTable dtR = GetRRRateValues(dtv.Rows[i]["ID"].ToString());
                    rw = stRow;
                    int chgRow = 1;
                    for (int k = 0; k < dtR.Rows.Count; k++)
                    {

                        ws.Cells["AB" + StRow1].Value = dtR.Rows[k]["CntrType"].ToString();
                        ws.Cells["AC" + StRow1].Value = dtR.Rows[k]["ChgCode"].ToString();
                        ws.Cells["AD" + StRow1].Value = dtR.Rows[k]["Currency"].ToString();
                        ws.Cells["AE" + StRow1].Value = dtR.Rows[k]["Paymode"].ToString();
                        ws.Cells["AF" + StRow1].Value = dtR.Rows[k]["ReqRate"];
                        ws.Cells["AG" + StRow1].Value = dtR.Rows[k]["ManifRate"];
                        ws.Cells["AH" + StRow1].Value = dtR.Rows[k]["CustomerRate"];
                        ws.Cells["AI" + StRow1].Value = dtR.Rows[k]["RateDiff"];
                        StRow1++;
                        chgRow++;

                    }

                    rw = StRow1;
                }

                rw -= 1;
            }
            ws.Cells["A7:AI" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AI" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AI" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AI" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 40].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=RRSearchlist.xlsx");
            Response.End();



        }

        public DataTable GetRRSearchValues(string PortofLoading, string RRNumber, string PortofDischarge, string BookingParty, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = " select ID, RatesheetNo,(select top(1) CustomerName from NVO_view_CustomerDetails where cid = BookingPartyID) as BookingParty, " +
                            " convert(varchar, date, 101) as Datev,(select top(1) CityName from NVO_CityMaster where Id = PortOfOrgin) PortofOrgin, " +
                            " (select top(1) PortName from NVO_PortMaster where Id = PortofLoading) PortofLoading," +
                            " (select top(1) Status from NVO_RRStatusMaster where ID =RSStatus) as Status ,(select top(1) PortName from NVO_PortMaster where Id = PlaceofDischargeId) PlaceofDischarge," +
                      " (select top(1) CityName from NVO_CityMaster where Id =finalpodid ) FPOD, (select top(1) PortName from NVO_PortMaster where Id = TranshimentPortID ) TranshimentPort," +
                   " (select top(1) AgencyName from NVO_AgencyMaster where Id = AgentId) Agency, (select top(1) GeneralName from NVO_GeneralMaster where Id = ShipmentID) ShipmentType," +
                " (select top(1) Description from NVO_tblDLValues where Id = ServiceTypeID) ServiceType, case when RouteType = 1 then 'DIRECT' when RouteType = 2 then 'TRANSHIPMENT' else '' end as RouteTypes," +
                " (select top(1) CustomerName from NVO_view_CustomerDetails where cid =BookingPartyID ) BkgParty,(select top(1) CustomerName from NVO_view_CustomerDetails where cid = ShipperID ) Shipper," +
                " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.ApprovedBy) as ApprovedBy, " +
                  " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.UserID) as CreatedBy, " +
                   " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.RejectedBy) as RejectedBy, " +
                  " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.RequotedBy) as RequotedBy, " +
                " convert(varchar, DtApproved, 101) as DtAppr, convert(varchar, DtRejected, 101) as DtRej,convert(varchar, DtRequoted, 101) as DtRequote,RjRemarks,ReRemarks,Remarks,(select top 1 ExpFreeDays from NVO_RatesheetMode  where RRID = NVO_Ratesheet.ID) AS ExpFreeDays, (select top 1 ImpFreeDays from NVO_RatesheetMode  where RRID = NVO_Ratesheet.ID) AS ImpFreeDays   from NVO_Ratesheet ";

            if (RRNumber != "" && RRNumber != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where RatesheetNo like'%" + RRNumber + "%'";
                else
                    strWhere += " and RatesheetNo like'%" + RRNumber + "%'";

            if (PortofLoading != "" && PortofLoading != "null" && PortofLoading != "?" && PortofLoading != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PortOfLoading=" + PortofLoading;
                else
                    strWhere += " and PortOfLoading=" + PortofLoading;

            if (PortofDischarge != "" && PortofDischarge != "null" && PortofDischarge != "?" && PortofDischarge != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PlaceofDischargeId=" + PortofDischarge;
                else
                    strWhere += " and PlaceofDischargeId=" + PortofDischarge;


            if (BookingParty != "" && BookingParty != "null" && BookingParty != "?" && BookingParty != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BookingPartyID=" + BookingParty;
                else
                    strWhere += " and BookingPartyID=" + BookingParty;

            if (Status != "" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where RSStatus=" + Status;
                else
                    strWhere += " and RSStatus=" + Status;

            if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "2" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where NVO_Ratesheet.Agentid = " + AgencyID.ToString();
                else
                    strWhere += " and NVO_Ratesheet.Agentid = " + AgencyID.ToString();

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetRRRateValues(string RRID)

        {
            string _Query = "select CID, RRID, (SELECT TOP  1 SIZE FROM NVO_tblCntrTypes WHERE ID = RC.CntrType) as CntrType , " +
                " (SELECT TOP  1 CurrencyCode FROM NVO_CurrencyMaster WHERE ID = RC.CurrencyID) as Currency, " +
                "   (SELECT TOP  1 (ChgCode +'-'+ ChgDesc) FROM NVO_ChargeTB WHERE ID=RC.ChargeCodeID) as ChgCode, " +
                " (SELECT TOP  1 GeneralName FROM NVO_GeneralMaster WHERE ID = RC.PaymentModeID) as Paymode ,ReqRate,ManifRate,CustomerRate,RateDiff from NVO_RatesheetCharges RC  WHERE RC.TariffTypeID in (135, 136, 137) and RC.ChargeTypeID = 1 and RC.RRID = " + RRID;
            return Manag.GetViewData(_Query, "");
        }


        public void CreateRRSummary(string Status, string User, string AgencyID, string FromDt, string ToDt)
        {
            DataTable dtv = GetRRSummarySearchValues(Status, AgencyID, FromDt, ToDt);
           

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("RRSummary");

                ws.Cells["A2"].Value = "Rate Request Search List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:AI2"];
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
                ws.Cells["B7"].Value = "RR Number";
                ws.Cells["C7"].Value = "Created On";
                ws.Cells["D7"].Value = "Agency Name";
                ws.Cells["E7"].Value = "Shipment Type";
                ws.Cells["F7"].Value = "POO";
                ws.Cells["G7"].Value = "POL";
                ws.Cells["H7"].Value = "POD";
                ws.Cells["I7"].Value = "FPOD";
                ws.Cells["J7"].Value = "Route Type";
                ws.Cells["K7"].Value = "Transhipment";
                ws.Cells["L7"].Value = "Service Type";
                ws.Cells["M7"].Value = "Booking Party";
                ws.Cells["N7"].Value = "Shipper";
                ws.Cells["O7"].Value = "Status";
                ws.Cells["P7"].Value = "Created By";
                ws.Cells["Q7"].Value = "Approved By";
                ws.Cells["R7"].Value = "Approved On";
                ws.Cells["S7"].Value = "Rejected On";
                ws.Cells["T7"].Value = "Rejected By";
                ws.Cells["U7"].Value = "Rejected Reason";
                ws.Cells["V7"].Value = "Requoted On";
                ws.Cells["W7"].Value = "Requoted By";
                ws.Cells["X7"].Value = "Requoted Reason";
                ws.Cells["Y7"].Value = "Special Instruction-slot Details& Services";
                ws.Cells["Z7"].Value = " FreeDays Mode";
                ws.Cells["AA7"].Value = " ExpFreeDays";
                ws.Cells["AB7"].Value = " ImpFreeDays";
                ws.Cells["AC7"].Value = " Equip Type";
                ws.Cells["AD7"].Value = "Charge Description";
                ws.Cells["AE7"].Value = "Currency";
                ws.Cells["AF7"].Value = "Payment Mode";
                ws.Cells["AG7"].Value = "Requested Rate";
                ws.Cells["AH7"].Value = "Manifest Rate";
                ws.Cells["AI7"].Value = "Customer Rate";
                ws.Cells["AJ7"].Value = "Rate Different";
                ws.Cells["AK7"].Value = "Slot Operator";
                ws.Cells["AL7"].Value = "Slot Term";
                ws.Cells["AM7"].Value = "Slot Cost for 20'";
                ws.Cells["AN7"].Value = "Slot Cost for 40'";

                r = ws.Cells["A7:AN7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;
                int frowid = 0;
            if (dtv.Rows.Count > 0)
            {
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    frowid = rw;
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["RatesheetNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Datev"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["PortofOrgin"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["PortofLoading"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["PlaceofDischarge"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["RouteTypes"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["TranshimentPort"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["ServiceType"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["BkgParty"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["CreatedBy"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["ApprovedBy"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["DtAppr"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["RejectedBy"].ToString();
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["DtRej"].ToString();
                    ws.Cells["U" + rw].Value = dtv.Rows[i]["RjRemarks"].ToString();
                    ws.Cells["V" + rw].Value = dtv.Rows[i]["DtRequote"].ToString();
                    ws.Cells["W" + rw].Value = dtv.Rows[i]["RequotedBy"].ToString();
                    ws.Cells["X" + rw].Value = dtv.Rows[i]["ReRemarks"].ToString();
                    ws.Cells["Y" + rw].Value = dtv.Rows[i]["Remarks"].ToString();


                    sl++;
                    int stRow = rw;
                    int StRow1 = rw;
                    DataTable dtRC = GetRRSummaryRateSheetMode(dtv.Rows[i]["ID"].ToString());
                    rw = stRow;
                    int chgRow = 1;
                    for (int k = 0; k < dtRC.Rows.Count; k++)
                    {
                        ws.Cells["Z" + StRow1].Value = dtRC.Rows[k]["Mode"];
                        ws.Cells["AA" + StRow1].Value = dtRC.Rows[k]["ExpFreeDays"];
                        ws.Cells["AB" + StRow1].Value = dtRC.Rows[k]["ImpFreeDays"];

                        StRow1++;
                        chgRow++;

                    }

                    // rw = StRow1;
                    //sl++;
                    int rowid = 0;
                    DataTable dtR = GetRRSummaryRateValues(dtv.Rows[i]["ID"].ToString());

                    for (int j = 0; j < dtR.Rows.Count; j++)
                    {


                        ws.Cells["AC" + (StRow1 - 1)].Value = dtR.Rows[j]["CntrType"].ToString();
                        ws.Cells["AD" + (StRow1 - 1)].Value = dtR.Rows[j]["ChgCode"].ToString();
                        ws.Cells["AE" + (StRow1 - 1)].Value = dtR.Rows[j]["Currency"].ToString();
                        ws.Cells["AF" + (StRow1 - 1)].Value = dtR.Rows[j]["Paymode"].ToString();
                        ws.Cells["AG" + (StRow1 - 1)].Value = dtR.Rows[j]["ReqRate"];
                        ws.Cells["AH" + (StRow1 - 1)].Value = dtR.Rows[j]["ManifRate"];
                        ws.Cells["AI" + (StRow1 - 1)].Value = dtR.Rows[j]["CustomerRate"];
                        ws.Cells["AJ" + (StRow1 - 1)].Value = dtR.Rows[j]["RateDiff"];

                        StRow1++;
                        chgRow++;
                        rowid++;
                    }
                    // rw = StRow1;
                    // sl++;

                    rowid++;
                    DataTable dtS = GetRRSummaryRateSheetSlot(dtv.Rows[i]["ID"].ToString());
                    for (int l = 0; l < dtS.Rows.Count; l++)
                    {

                        ws.Cells["AK" + (StRow1 - rowid)].Value = dtS.Rows[l]["SlotOpr"].ToString();
                        ws.Cells["AL" + (StRow1 - rowid)].Value = dtS.Rows[l]["SlotTerm"].ToString();
                        ws.Cells["AM" + (StRow1 - rowid)].Value = dtS.Rows[l]["SlotAmt20"].ToString();
                        ws.Cells["AN" + (StRow1 - rowid)].Value = dtS.Rows[l]["SlotAmt40"].ToString();

                        StRow1++;
                        chgRow++;

                    }
                    rw = StRow1 - 1;
                    // sl++;
                }

                rw -= 1;
            }
            ws.Cells["A7:AN" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AN" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AN" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AN" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 40].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=RRSummarylist.xlsx");
                Response.End();

           

        }

        public DataTable GetRRSummarySearchValues(string Status, string AgencyID, string FromDt, string ToDt)
        {
            string fromdate = ""; string ToDate = "";
            if (FromDt != "" && FromDt != null && FromDt != "undefined")
                fromdate = (DateTime.Parse(FromDt)).ToString("MM/dd/yyyy");
            if (ToDt != "" && ToDt != null && ToDt != "undefined")
                ToDate = (DateTime.Parse(ToDt)).ToString("MM/dd/yyyy");

            string strWhere = "";
            string _Query = " select ID, RatesheetNo,(select top(1) CustomerName from NVO_view_CustomerDetails where cid = BookingPartyID) as BookingParty, " +
                            " convert(varchar, date, 101) as Datev,(select top(1) CityName from NVO_CityMaster where Id = PortOfOrgin) PortofOrgin, " +
                            " (select top(1) PortName from NVO_PortMaster where Id = PortofLoading) PortofLoading," +
                            " (select top(1) Status from NVO_RRStatusMaster where ID =RSStatus) as Status ,(select top(1) PortName from NVO_PortMaster where Id = PlaceofDischargeId) PlaceofDischarge," +
                      " (select top(1) CityName from NVO_CityMaster where Id =finalpodid ) FPOD, (select top(1) PortName from NVO_PortMaster where Id = TranshimentPortID ) TranshimentPort," +
                   " (select top(1) AgencyName from NVO_AgencyMaster where Id = AgentId) Agency, (select top(1) GeneralName from NVO_GeneralMaster where Id = ShipmentID) ShipmentType," +
                " (select top(1) Description from NVO_tblDLValues where Id = ServiceTypeID) ServiceType, case when RouteType = 1 then 'DIRECT' when RouteType = 2 then 'TRANSHIPMENT' else '' end as RouteTypes," +
                " (select top(1) CustomerName from NVO_view_CustomerDetails where cid =BookingPartyID ) BkgParty,(select top(1) CustomerName from NVO_view_CustomerDetails where cid = ShipperID ) Shipper," +
                " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.ApprovedBy) as ApprovedBy, " +
                  " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.UserID) as CreatedBy, " +
                   " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.RejectedBy) as RejectedBy, " +
                  " (select top(1) UserName from NVO_UserDetails where ID = NVO_Ratesheet.RequotedBy) as RequotedBy, " +
                " convert(varchar, DtApproved, 101) as DtAppr, convert(varchar, DtRejected, 101) as DtRej,convert(varchar, DtRequoted, 101) as DtRequote,RjRemarks,ReRemarks,Remarks,(select top 1 ExpFreeDays from NVO_RatesheetMode  where RRID = NVO_Ratesheet.ID) AS ExpFreeDays, (select top 1 ImpFreeDays from NVO_RatesheetMode  where RRID = NVO_Ratesheet.ID) AS ImpFreeDays   from NVO_Ratesheet ";



            if (FromDt != "" && FromDt != "undefined" && FromDt != null || ToDt != "" && ToDt != "undefined" && FromDt != ToDt)
                if (strWhere == "")
                    strWhere += _Query + " where  NVO_Ratesheet.date   between '" + fromdate + "' and '" + ToDate + "' ";
                else
                    strWhere += "  and NVO_Ratesheet.date between '" + fromdate + "' and '" + ToDate + "'  ";


            if (Status != "" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where RSStatus=" + Status;
                else
                    strWhere += " and RSStatus=" + Status;

            if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "?" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where NVO_Ratesheet.Agentid = " + AgencyID.ToString();
                else
                    strWhere += " and NVO_Ratesheet.Agentid = " + AgencyID.ToString();

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetRRSummaryRateValues(string RRID)

        {
            string _Query = "select CID, RRID, (SELECT TOP  1 SIZE FROM NVO_tblCntrTypes WHERE ID = RC.CntrType) as CntrType , " +
                " (SELECT TOP  1 CurrencyCode FROM NVO_CurrencyMaster WHERE ID = RC.CurrencyID) as Currency, " +
                "   (SELECT TOP  1 (ChgCode +'-'+ ChgDesc) FROM NVO_ChargeTB WHERE ID=RC.ChargeCodeID) as ChgCode, " +
                " (SELECT TOP  1 GeneralName FROM NVO_GeneralMaster WHERE ID = RC.PaymentModeID) as Paymode ,ReqRate,ManifRate,CustomerRate,RateDiff from NVO_RatesheetCharges RC  WHERE RC.TariffTypeID in (135, 136, 137) and RC.ChargeTypeID = 1 and RC.RRID = " + RRID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetRRSummaryRateSheetMode(string RRID)

        {
            string _Query = "select case when ModeID=1 then 'Combined' when ModeID=2 then 'Detention' " +
                "  when ModeID=3 then 'Demmurage' else '' end as Mode,* from NVO_RatesheetMode where RRID = " + RRID;
            return Manag.GetViewData(_Query, "");
        }
        public DataTable GetRRSummaryRateSheetSlot(string RRID)

        {
            string _Query = "select (SELECT TOP  1 Description FROM NVO_tblDLValues WHERE ID = isnull(RC.SlotTermID,0)) as SlotTerm," +
                " (SELECT TOP  1 CustomerName FROM NVO_view_CustomerDetails WHERE CID = isnull(RC.SlotOperatorID,0)) as SlotOpr," +
                " * from NVO_RatesheetSlotCharges RC where RC.RRID = " + RRID;
            return Manag.GetViewData(_Query, "");
        }

    }
}