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
    public class VoyageExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }
        public void CreateVoyageExcel(string AgencyID, string VesselName, string VoyageNo, string User)
        {
            DataTable dtv = GetVoyageSearchValues(AgencyID, VesselName, VoyageNo);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("VoyageList");

                ws.Cells["A2"].Value = "Voyage Search List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:J2"];
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
                ws.Cells["B7"].Value = "Vessel";
                ws.Cells["C7"].Value = "Service Name";
                ws.Cells["D7"].Value = "RIN";
                ws.Cells["E7"].Value = "Export Voyage Code";
                ws.Cells["F7"].Value = "Import Voyage Code";
                ws.Cells["G7"].Value = "Port";
                ws.Cells["H7"].Value = "Terminal";
                ws.Cells["I7"].Value = "ETA";
                ws.Cells["J7"].Value = "ETD";
                r = ws.Cells["A7:J7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Vessel"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Service"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["RIN"].ToString();

                    DataTable dtR = GetVoyageRoute(dtv.Rows[i]["ID"].ToString());
                    for (int k = 0; k < dtR.Rows.Count; k++)
                    {
                        ws.Cells["E" + rw].Value = dtR.Rows[k]["ExportVoyageCd"].ToString();
                        ws.Cells["F" + rw].Value = dtR.Rows[k]["ImportVoyageCd"].ToString();
                        ws.Cells["G" + rw].Value = dtR.Rows[k]["Port"].ToString();
                        ws.Cells["H" + rw].Value = dtR.Rows[k]["Terminal"].ToString();
                        ws.Cells["I" + rw].Value = dtR.Rows[k]["ETA"].ToString();
                        ws.Cells["J" + rw].Value = dtR.Rows[k]["ETD"].ToString();
                        rw++;
                    }

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:J" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=Voyagelist.xlsx");
                Response.End();

            }

        }

        public DataTable GetVoyageSearchValues(string AgencyID, string VesselName, string VoyageNo)
        {
            string strWhere = "";
            string _Query = "Select V.ID,(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) as Vessel,(select top(1) ServiceName from NVO_Services where ID = V.ServiceID) as Service,(select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) as OBVoyNo,RIN from  NVO_Voyage V  where V.IsTransferred <> 1 and  ";

            if (VesselName != "" && VesselName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID )  like '%" + VesselName + "%'";
                else
                    strWhere += " and (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID ) like '%" + VesselName + "%'";

            if (VoyageNo != "" && VoyageNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) like '%" + VoyageNo + "%'";
                else
                    strWhere += " and (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) like '%" + VoyageNo + "%'";


            if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != "null")

                if (strWhere == "")
                    strWhere += _Query + "   V.AgencyID = " + AgencyID.ToString() + "";
                else
                    strWhere += " and  V.AgencyID = " + AgencyID.ToString() + "";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetVoyageRoute(string VoyageID)
        {


            string _Query = "Select RID,ExportVoyageCd,ImportVoyageCd,(select TOP 1 PortName From NVO_PortMainMaster  WHERE ID = PortID) as Port, " +
                " (select TOP 1 TerminalName From NVO_TerminalMaster  WHERE ID=TerminalID) as Terminal, convert(varchar, NVo_VoyageRoute.ETA, 103) as ETA,convert(varchar, NVo_VoyageRoute.ETD, 103) as ETD from NVO_VoyageRoute where VoyageID = " + VoyageID;

            return Manag.GetViewData(_Query, "");
        }

        public void CreateVoyAllocationExcel(string AgencyID, string VesVoyID, string User)
        {
            DataTable dtv = GetVoyAllocationValues(AgencyID, VesVoyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("VoyageAllocationList");

                ws.Cells["A2"].Value = "Voyage Allocation List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:J2"];
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
                ws.Cells["B7"].Value = "Voyage";
                ws.Cells["C7"].Value = "Voyage Types";
                ws.Cells["D7"].Value = "Leg Information";
                ws.Cells["E7"].Value = "BLNumber";
                ws.Cells["F7"].Value = "POL";
                ws.Cells["G7"].Value = "POD";
                ws.Cells["H7"].Value = "TSPORT";
                ws.Cells["I7"].Value = "1st Leg Vessel/Voyage";
                ws.Cells["J7"].Value = "2nd Leg Vessel/Voyage";
                r = ws.Cells["A7:J7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["VoyageTypes"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["LegInformation"].ToString();

                    DataTable dtBL = GetVoyAllocationBLValues(dtv.Rows[i]["ID"].ToString());
                    for (int k = 0; k < dtBL.Rows.Count; k++)
                    {
                        ws.Cells["E" + rw].Value = dtBL.Rows[i]["BLNumber"].ToString();
                        ws.Cells["F" + rw].Value = dtBL.Rows[k]["POL"].ToString();
                        ws.Cells["G" + rw].Value = dtBL.Rows[k]["POD"].ToString();
                        ws.Cells["H" + rw].Value = dtBL.Rows[k]["TSPORT"].ToString();
                        ws.Cells["I" + rw].Value = dtBL.Rows[k]["VesVoy1"].ToString();
                        ws.Cells["J" + rw].Value = dtBL.Rows[k]["VesVoyAlloc"].ToString();
                        rw++;
                    }

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:J" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=VoyageAllocationlist.xlsx");
                Response.End();

            }

        }

        public DataTable GetVoyAllocationValues(string AgencyID, string VesVoyID)
        {
            string strWhere = "";
            string _Query = "select  VA.ID,  (Select top 1 (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = VA.VesVoyID )as VesVoy, GM.GeneralName as VoyageTypes,GM1.GeneralName as LegInformation  from NVO_VoyageAllocation VA " +
           " inner join NVO_GeneralMaster GM ON GM.ID = VA.VoyageTypes and GM.SeqNo = 27 " +
         " inner join NVO_GeneralMaster GM1 ON GM1.ID = VA.leginformation and GM1.SeqNo = 26 ";

            if (VesVoyID.ToString() != "" && VesVoyID.ToString() != "0" && VesVoyID.ToString() != "null" && VesVoyID.ToString() != "?")

                if (strWhere == "")
                    strWhere += _Query + " where VA.VesvoyID = " + VesVoyID.ToString();
                else
                    strWhere += " and VA.VesvoyID = " + VesVoyID.ToString();

            if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != null && AgencyID.ToString() != "?")

                if (strWhere == "")
                    strWhere += _Query + " where VA.AgencyID = " + AgencyID.ToString();
                else
                    strWhere += " and VA.AgencyID = " + AgencyID.ToString();

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetVoyAllocationBLValues(string VoyAllocID)
        {

            string _Query = "select DISTINCT VAD.VID, BL.ID AS BLID,BL.bkgid,BL.BLNumber,(select top(1) PortName from NVO_PortMaster where ID = BL.POLID) As POL,(select top(1) PortName from NVO_PortMaster where ID = BL.PODID) As POD,(select top(1) PortName from NVO_PortMaster where ID = B.TSPORTID) As TSPORT,(Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = NVO_BOLVoyageDetails.VesVoyID )as VesVoy1," +
                "  (Select top 1 (select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID)  + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V where V.ID = NVO_VoyageAllocation.VesVoyID )as VesVoyAlloc from NVO_BOL BL  left outer join NVO_BOLVoyageDetails ON NVO_BOLVoyageDetails.BLID = BL.ID  left outer join NVO_Booking B ON B.ID = BL.BKGID  left outer join NVO_VoyageAllocationDtls VAD ON VAD.BLID = BL.ID " +
                " left outer join NVO_VoyageAllocation on NVO_VoyageAllocation.ID =VAD.VoyAllocID WHERE VAD.VoyAllocID  =" + VoyAllocID;

            return Manag.GetViewData(_Query, "");
        }

        public void CreateVesselMasterExcel(string VesselName, string Status, string User)
        {
            DataTable dtv = GetVesselMasterValues(VesselName, Status);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("VesselMasterList");
                ws.Cells["A2"].Value = "Vessel Master List";
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
                ws.Cells["B7"].Value = "Vessel Name";
                ws.Cells["C7"].Value = "Vessel Call Sign";
                ws.Cells["D7"].Value = "IMO Number";
                ws.Cells["E7"].Value = "MMSI";
                ws.Cells["F7"].Value = "Flag";
                ws.Cells["G7"].Value = "Vessel ID";
                ws.Cells["H7"].Value = "Vessel Owner";
                ws.Cells["I7"].Value = "Status";

                r = ws.Cells["A7:I7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesselName"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["VesselCallSign"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["IMONumber"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["MMSI"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Flag"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["VesselID"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["VesselOwner"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();


                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:J" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=VesselMasterlist.xlsx");
                Response.End();

            }

        }

        public DataTable GetVesselMasterValues(string VesselName, string Status)
        {
            string strWhere = "";

            string _Query = " select case when status = 1 then 'Active' when status = 0 then 'Inactive' ELSE '' END as StatusResult,* from NVO_VesselMaster ";



            if (VesselName != "" && VesselName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where VesselName like '%" + VesselName + "%'";
                else
                    strWhere += " and VesselName like '%" + VesselName + "%'";

            if (Status == "1")
                if (strWhere == "")
                    strWhere += _Query + " where Status =" + Status;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public void CreateServiceMasterExcel(string ServiceName, string GeoLocID, string User)
        {
            DataTable dtv = GetServiceMasterValues(ServiceName, GeoLocID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ServiceSetupList");
                ws.Cells["A2"].Value = "Service Setup List";
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
                ws.Cells["B7"].Value = "Service Name";
                ws.Cells["C7"].Value = "Dt Effective";
                ws.Cells["D7"].Value = "Port";
                ws.Cells["E7"].Value = "TmTransitTime";
                ws.Cells["F7"].Value = "Day Window Commmence";
                ws.Cells["G7"].Value = "TmWindowCommence";
                ws.Cells["H7"].Value = "Day Window Close";
                ws.Cells["I7"].Value = "TmWindowClose";
                ws.Cells["J7"].Value = "Day Port Stay";
                ws.Cells["K7"].Value = "TmPortStay";
                ws.Cells["L7"].Value = "Operators";
                ws.Cells["M7"].Value = "Slot Reference";
                r = ws.Cells["A7:M7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;
                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ServiceName"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["DtEffective"].ToString();
                    DataTable dtR = GetServiceRouteValues(dtv.Rows[i]["ID"].ToString());
                    for (int k = 0; k < dtR.Rows.Count; k++)
                    {
                        ws.Cells["D" + rw].Value = dtR.Rows[k]["Port"].ToString();
                        ws.Cells["E" + rw].Value = dtR.Rows[k]["TmTransitTime"].ToString();
                        ws.Cells["F" + rw].Value = dtR.Rows[k]["DayWinComm"].ToString();
                        ws.Cells["G" + rw].Value = dtR.Rows[k]["TmWindowCommence"].ToString();
                        ws.Cells["H" + rw].Value = dtR.Rows[k]["DayWinClose"].ToString();
                        ws.Cells["I" + rw].Value = dtR.Rows[k]["TmWindowClose"].ToString();
                        ws.Cells["J" + rw].Value = dtR.Rows[k]["DayPortStay"].ToString();
                        ws.Cells["K" + rw].Value = dtR.Rows[k]["TmPortStay"].ToString();

                        DataTable dtOP = GetServiceOperators(dtv.Rows[i]["ID"].ToString());
                        for (int j = 0; j < dtOP.Rows.Count; j++)
                        {
                            ws.Cells["L" + rw].Value = dtOP.Rows[j]["Operator"].ToString();
                            ws.Cells["M" + rw].Value = dtOP.Rows[j]["SlotRef"].ToString();
                        }
                        rw++;
                    }


                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:M" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ServiceSetuplist.xlsx");
                Response.End();

            }

        }

        public DataTable GetServiceMasterValues(string ServiceName, string GeoLocID)
        {
            string strWhere = "";

            string _Query = " select Id,ServiceName,convert(varchar,DtEffective, 103) As DtEffective  from NVO_Services ";

            if (ServiceName != "" && ServiceName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where ServiceName like '%" + ServiceName + "%'";
                else
                    strWhere += " and ServiceName like '%" + ServiceName + "%'";

            if (GeoLocID.ToString() != "" && GeoLocID.ToString() != "0" && GeoLocID.ToString() != "null" && GeoLocID.ToString() != "?")

                if (strWhere == "")
                    strWhere += _Query + " where GeoLocID = " + GeoLocID.ToString();
                else
                    strWhere += " and GeoLocID = " + GeoLocID.ToString();
            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetServiceRouteValues(string ServicesID)
        {

            string _Query = "Select (select TOP 1 PortName From NVO_PortMainMaster  WHERE ID = PortID ) as Port, Case when DayWindowClose = 0 then '' when DayWindowClose = 1 then 'SUN' WHEN DayWindowClose = 2 THEN 'MON' when DayWindowClose = 3 then 'TUE' WHEN DayWindowClose = 4 THEN 'WED' when DayWindowClose = 5 then 'THU' WHEN DayWindowClose = 6 THEN 'FRI' WHEN DayWindowClose = 7 THEN 'SAT' END AS DayWinClose,Case when DayWindowCommence = 0 then '' when DayWindowCommence = 1 then 'SUN' WHEN DayWindowCommence = 2 THEN 'MON' when DayWindowCommence = 3 then 'TUE' WHEN DayWindowCommence = 4 THEN 'WED' when DayWindowCommence = 5 then 'THU' WHEN DayWindowCommence = 6 THEN 'FRI' WHEN DayWindowCommence = 7 THEN 'SAT' END AS DayWinComm,*  from NVO_SERVICEROUTE where ServicesID = " + ServicesID;

            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetServiceOperators(string ServicesID)
        {

            string _Query = "select * from NVO_ServiceOpertaors where ServicesID = " + ServicesID;

            return Manag.GetViewData(_Query, "");
        }

    }
}