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
    public class ContainersExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }

        public void CntrMasterExcel(string CntrNo, string TypeID, string StatusID, string LeasingPartnerID, string BoxOwnerID, string User)
        {
            DataTable dtv = GetCntrMasterValues(CntrNo, TypeID, StatusID, LeasingPartnerID, BoxOwnerID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ContainerMaster");

                ws.Cells["A2"].Value = "Container Master List";
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
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Container Type";
                ws.Cells["D7"].Value = "ISOCODE";
                ws.Cells["E7"].Value = "Grade";
                ws.Cells["F7"].Value = "Cubic Capacity";
                ws.Cells["G7"].Value = "Gross Weight";
                ws.Cells["H7"].Value = "Net Weight";
                ws.Cells["I7"].Value = "Tare Weight";
                ws.Cells["J7"].Value = "DtManufacture";
                ws.Cells["K7"].Value = "Box Owner";
                ws.Cells["L7"].Value = "Lease Partner";
                ws.Cells["M7"].Value = "Lease Term";
                ws.Cells["N7"].Value = "Reference";
                ws.Cells["O7"].Value = "Container Status";
                //ws.Cells["P7"].Value = "PreparedBy";
                //ws.Cells["Q7"].Value = "VesVoy";
                //ws.Cells["R7"].Value = "Slot Contract";
                //ws.Cells["S7"].Value = "Slot Operator";
                r = ws.Cells["A7:O7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ISOCODE"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Grade"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["CubicCapacity"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["GrWtx100"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["NtWtx100"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["TareWtx100"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["DtManuf"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["BoxOwner"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["LeasingPartner"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["LeaseTerm"].ToString();

                    if (dtv.Rows[i]["PickUpRefID"].ToString() == "0")
                    {
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["Reference"].ToString();
                    }
                    else
                    {
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["PickUpRef"].ToString();
                    }
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    //ws.Cells["P" + rw].Value = dtv.Rows[i]["PreparedBy"].ToString();
                    //ws.Cells["Q" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    //ws.Cells["R" + rw].Value = dtv.Rows[i]["SlotContract"].ToString();
                    //ws.Cells["S" + rw].Value = dtv.Rows[i]["SlotOperator"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:O" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CntrMasterList.xlsx");
                Response.End();

            }

        }

        public DataTable GetCntrMasterValues(string CntrNo, string TypeID, string StatusID, string LeasingPartnerID, string BoxOwnerID)
        {
            string _Query = "Select CN.ID,CN.CntrNo,PickUpRefID,(select top 1 SIZE from NVO_tblCntrTypes where ID = TypeID) as Size,(select top 1 ISOCODE from NVO_tblCntrTypes where ID = ISOCODEID) as ISOCODE," +
                " (select top 1 GeneralName from NVO_generalMaster where ID = GradeID) as Grade,CubicCapacity,GrWtx100,NtWtx100,TareWtx100,Reference,(Select top 1 ContractRefNO from nvo_leasecontract Where ID=PickUpRefID) as PickUpRef, convert(varchar, DtManufacture, 103) as DtManuf,( select top 1 GeneralName from NVO_generalMaster where ID = StatusID) as Status,(select top 1 CustomerName from NVO_CustomerMaster where ID = BoxOwnerID) as BoxOwner,(select top 1 upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where NVO_CusBranchLocation.CID=LeasingPartnerID) as LeasingPartner,   ( select top 1 GeneralName from NVO_generalMaster where ID = LeaseTermID ) LeaseTerm from NVO_Containers CN ";

            string strWhere = "";

            if (CntrNo != "" && CntrNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CN.CntrNo like '%" + CntrNo + "%'";
                else
                    strWhere += " and CN.CntrNo like '%" + CntrNo + "%'";

            if (TypeID != "" && TypeID != "0" && TypeID != "null" && TypeID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where CN.TypeID=" + TypeID;
                else
                    strWhere += " and CN.TypeID =" + TypeID;

            if (BoxOwnerID != "" && BoxOwnerID != "0" && BoxOwnerID != "null" && BoxOwnerID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where CN.BoxOwnerID=" + BoxOwnerID;
                else
                    strWhere += " and CN.BoxOwnerID =" + BoxOwnerID;

            if (LeasingPartnerID.ToString() != "" && LeasingPartnerID != "0" && LeasingPartnerID != "null" && LeasingPartnerID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where CN.LeasingPartnerID=" + LeasingPartnerID;
                else
                    strWhere += " and CN.LeasingPartnerID =" + LeasingPartnerID;

            if (StatusID != "" && StatusID != "0" && StatusID != "null" && StatusID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where CN.StatusID=" + StatusID;
                else
                    strWhere += " and CN.StatusID =" + StatusID;

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public void InventoryTrackingExcel(string CntrID, string TypeID, string User)
        {
            DataTable dtv = GetInventoryTrackingExcelValues(CntrID, TypeID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("InventoryTracking");

                ws.Cells["A2"].Value = "Inventory TrackingList";
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
                ws.Cells["C4"].Value = "Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "Container No";
                ws.Cells["C7"].Value = "Container Type";
                ws.Cells["D7"].Value = "StatusCode";
                ws.Cells["E7"].Value = "Movement Date";
                ws.Cells["F7"].Value = "Agency";
                ws.Cells["G7"].Value = "Location";
                ws.Cells["H7"].Value = "Vessel / Voyage";
                ws.Cells["I7"].Value = "Transit Mode";
                ws.Cells["J7"].Value = "Depot";
                ws.Cells["K7"].Value = "BL Number";
                ws.Cells["L7"].Value = "Customer";
                ws.Cells["M7"].Value = "UserName";
                ws.Cells["N7"].Value = "Lease Reference";
                ws.Cells["O7"].Value = "Lease Partner";
                r = ws.Cells["A7:O7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["SIZE"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Statuscode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DtMovement"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["NextPort"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["VESVOY"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["TransitMode"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Depot"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["CustomerName"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["UserName"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["LeaseReference"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["LeasePartner"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:O" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:O" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=InventoryTracking.xlsx");
                Response.End();

            }

        }

        public DataTable GetInventoryTrackingExcelValues(string CntrID, string TypeID)
        {
            string _Query = " select DISTINCT CT.ID As TxnsID,C.ID As CntrID,C.CntrNo, " +
    " (select top(1) AgencyName from NVO_AgencyMaster where ID = CT.AgencyID) AS AgencyName, " +
    " ISNULL((select top(1) NextPortID   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc),0) AS CurrentPortID ," +
      " (select top(1) Size   from NVO_tblCntrTypes where ID = C.TypeID) AS SIZE, CT.Statuscode,CT.DtMovement, " +
      " (select top(1) PortName from NVO_PortMaster where ID = CT.LocationID order by CT.DtMovement desc ) as FromPort, " +
         " ( select top(1) PortName from NVO_PortMaster where ID = CT.NextPortID order by CT.DtMovement desc ) as NextPort, " +
      " (select top(1) GeneralName from NVO_GeneralMaster where ID = CT.ModeOfTransportID AND SeqNo = 31 order by CT.DtMovement desc ) as TransitMode," +
       " (select top(1) DepName from NVO_DepotMaster where ID = CT.DepotID order by CT.DtMovement desc ) as Depot," +
        "  ( select top(1) U.UserName   from NVO_UserDetails U WHERE ID =CT.UserID)  AS UserName , " +
        " (select top(1) BookingNo from NVO_Booking where ID = CT.BLNumber order by CT.DtMovement desc ) as BLNumber," +
      " (select top(1) CustomerName from NVO_view_CustomerDetails where CID = CT.CustomerID order by CT.DtMovement desc ) as CustomerName," +
      " (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute " +
      " where VoyageID = V.ID) from NVO_Voyage V where CT.ContainerID = C.ID and CT.VesVoyID = V.ID Order by DtMovement desc)  As VESVOY," +
      "  (select top(1) ContractRefNo from NVO_LeaseContract where ID = C.PickUpRefID ) as LeaseReference, "+
       " (select top(1) CustomerName from NVO_view_CustomerDetails where cID = C.LeasingPartnerID ) as LeasePartner from NVO_ContainerTxns CT " +
    " INNER join NVO_Containers C ON C.ID = CT.ContainerID where CT.STATUSCODE NOT IN ('PENDING') ";

            string strWhere = "";

            if (CntrID != "" && CntrID != "0" && CntrID != "null" && CntrID != "?")
                if (strWhere == "")
                    strWhere += _Query + " AND C.ID=" + CntrID;
                else
                    strWhere += " and C.ID =" + CntrID;


            if (TypeID != "" && TypeID != "0" && TypeID != "null" && TypeID != "?")
                if (strWhere == "")
                    strWhere += _Query + " AND C.TypeID=" + TypeID;
                else
                    strWhere += " and C.TypeID =" + TypeID;

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere + "Order by CT.DtMovement desc ", "");
        }

        public void InventoryMovementExcel(string CntrID, string TypeID, string LocationID, string StatusCode, string VesVoyID, string User)
        {
            DataTable dtv = GetInventoryMovementExcelValues(CntrID, TypeID, LocationID, StatusCode, VesVoyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("InventoryMovement");

                ws.Cells["A2"].Value = "Inventory Movement List";
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
                ws.Cells["C7"].Value = "Container Type";
                ws.Cells["D7"].Value = "StatusCode";
                ws.Cells["E7"].Value = "Movement Date";
                ws.Cells["F7"].Value = "Agency";
                ws.Cells["G7"].Value = "Location";
                ws.Cells["H7"].Value = "Vessel / Voyage";
                ws.Cells["I7"].Value = "Transit Mode";
                ws.Cells["J7"].Value = "Depot";
                ws.Cells["K7"].Value = "BL Number";
                ws.Cells["L7"].Value = "Customer";
                ws.Cells["M7"].Value = "UserName";
                r = ws.Cells["A7:M7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Statuscode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DtMovement"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["FromPort"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["VESVOY"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["TransitMode"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Depot"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["CustomerName"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["UserName"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:M" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=InventoryMovement.xlsx");
                Response.End();

            }

        }

        public DataTable GetInventoryMovementExcelValues(string CntrID, string TypeID, string LocationID, string StatusCode, string VesVoyID)
        {
            string _Query = "Select DISTINCT C.ID As CntrID, '0' as ChkFlag, C.CntrNo,( select top(1) StatusCode   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) as StatusCode, ISNULL((select top(1) NextPortID   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc),0) AS CurrentPortID, " +
              " ( select top(1) A.AgencyName from NVO_AgencyMaster A  Inner join NVO_ContainerTxns CT ON CT.ContainerID =  C.ID where A.ID = CT.AgencyID ORDER BY CT.DtMovement DESC ) AS AgencyName ," +
              " (select top(1) DtMovement   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) AS DtMovement , " +
               " (select top(1) BL.BLNumber from NVO_BOL BL Inner join NVO_ContainerTxns CT ON CT.BLNumber = BL.ID  WHERE CT.ContainerID = C.ID ORDER BY CT.DtMovement DESC ) AS BLNumber , " +

            " (select top(1) Size   from NVO_tblCntrTypes where ID = C.TypeID) AS Size, " +
           " (select top(1) PortName   from NVO_PortMaster " +
           " Inner Join NVO_ContainerTxns on NVO_ContainerTxns.ContainerID= C.ID " +
           " where NVO_ContainerTxns.NextPortID =NVO_PortMaster.ID order by DtMovement desc) AS FromPort," +
           " (select top(1) G.GeneralName from NVO_GeneralMaster G  Inner join NVO_ContainerTxns CT ON CT.ContainerID = C.ID where G.ID = CT.ModeOfTransportID ORDER BY CT.DtMovement DESC ) AS TransitMode , " +
        " ( select top(1) D.DepName   from NVO_DepotMaster D Inner join NVO_ContainerTxns CT ON CT.ContainerID = C.ID where D.ID = CT.DepotID ORDER BY CT.DtMovement DESC ) AS Depot , " +
          " ( select top(1) U.UserName   from NVO_UserDetails U Inner join NVO_ContainerTxns CT ON CT.ContainerID = C.ID where U.ID = CT.UserID ORDER BY CT.DtMovement DESC ) AS UserName , " +
       " (select top(1) CustomerName   from NVO_CustomerMaster where ID = C.CustomerID) AS CustomerName, (Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) from NVO_Voyage V " +
      " Inner JOIN NVO_ContainerTxns ON NVO_ContainerTxns.VesVoyID = V.ID where ContainerID = C.ID  Order by DtMovement desc)  As VESVOY from NVO_Containers C " +
       "";
            string strWhere = "";


            if (StatusCode != "" && StatusCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where (select top(1) StatusCode   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) like '%" + StatusCode + "%'";
                else
                    strWhere += " and (select top(1) StatusCode   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) like '%" + StatusCode + "%'";

            if (CntrID != "" && CntrID != "0" && CntrID != "null" && CntrID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where C.ID=" + CntrID;
                else
                    strWhere += " and C.ID =" + CntrID;


            if (TypeID != "" && TypeID != "0" && TypeID != "null" && TypeID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where C.TypeID=" + TypeID;
                else
                    strWhere += " and C.TypeID =" + TypeID;

            if (LocationID != "" && LocationID != "0" && LocationID != "null" && LocationID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where (select top(1) NextPortID   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) =" + LocationID;
                else
                    strWhere += " and (select top(1) NextPortID   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) =" + LocationID;

            if (VesVoyID != "" && VesVoyID != "0" && VesVoyID != "null" && VesVoyID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where C.VesVoyID=" + VesVoyID;
                else
                    strWhere += " and C.VesVoyID =" + VesVoyID;


            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public void CntrCurrentReport(string FromDt, string ToDt, string User)
        {
            DataTable dtv = GetCntrCurrentReport( FromDt,  ToDt);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();



                #region 1st SHEET
                var ws = pck.Workbook.Worksheets.Add("CurrentList");
                ws.Cells["A2"].Value = "Container Current Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:H2"];
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
                ws.Cells["D7"].Value = "Status Code";
                ws.Cells["E7"].Value = "Last MovementDate";
                ws.Cells["F7"].Value = "Agency";
                ws.Cells["G7"].Value = "Current Port";
                ws.Cells["H7"].Value = "Current Depot";


                r = ws.Cells["A7:H7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int slno = 1;
                int row = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + row].Value = slno;
                    ws.Cells["B" + row].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + row].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["D" + row].Value = dtv.Rows[i]["StatusCode"].ToString();
                    ws.Cells["E" + row].Value = dtv.Rows[i]["LastDtMovement"].ToString();
                    ws.Cells["F" + row].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["G" + row].Value = dtv.Rows[i]["CurrentPort"].ToString();
                    ws.Cells["H" + row].Value = dtv.Rows[i]["CurrentDepot"].ToString();

                    slno++;
                    row += 1;
                }

                row -= 1;

                ws.Cells["A7:H" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, row, 10].AutoFitColumns();
                #endregion



                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ContainerCurrentList.xlsx");
                Response.End();

            }

        }


        public DataTable GetCntrCurrentReport(string FromDt, string ToDt)
        {


            string _Query = "Select ID,CntrNo,StatusCode,(select top 1 DtMovement from NVO_ContainerTxns where ContainerID = NVO_Containers.ID and " +
                " NVO_ContainerTxns.StatusCode!='PENDING' ORDER BY ID DESC) as LastDtMovement,(select top 1 AgencyName from NVO_AgencyMaster where ID = AgencyID) as Agency," +
                " (select top 1 Type + '-' + Size from NVO_tblCntrTypes where ID = TypeID) Size,(select top 1 PortName from NVO_PortMaster where ID = CurrentPortID) as CurrentPort, " +
                " isnull((select top 1 DepName from NVO_DepotMaster where ID = DepotID),'') as CurrentDepot from NVO_Containers  ";

            string strWhere = "";

            if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                if (strWhere == "") 
                    strWhere += _Query + " WHERE convert(varchar, (select top 1 DtMovement from NVO_ContainerTxns where ContainerID = NVO_Containers.ID and NVO_ContainerTxns.StatusCode != 'PENDING' ORDER BY ID DESC), 23)  between '" + FromDt + "' and '" + ToDt + "'";
                else
                    strWhere += "  and  convert(varchar, (select top 1 DtMovement from NVO_ContainerTxns where ContainerID = NVO_Containers.ID and NVO_ContainerTxns.StatusCode != 'PENDING' ORDER BY ID DESC), 23)  between '" + FromDt + "' and '" + ToDt + "'";

            if (strWhere == "")
                strWhere = _Query;

          


            return Manag.GetViewData(strWhere + " order By NVO_Containers.ID ", "");
        }


        public void CntrHistoryReport(string FromDt, string ToDt,string CntrID, string User)
        {
            DataTable dtv = GetCntrHistoryReport(FromDt, ToDt,CntrID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();



                #region 1st SHEET
                var ws = pck.Workbook.Worksheets.Add("Container History");
                ws.Cells["A2"].Value = "Container History Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r;
                r = ws.Cells["A2:M2"];
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
                ws.Cells["C7"].Value = "Type";
                ws.Cells["D7"].Value = "StatusCode";
                ws.Cells["E7"].Value = "DtMovement";
                ws.Cells["F7"].Value = "Agency";
                ws.Cells["G7"].Value = "Location";
                ws.Cells["H7"].Value = "Depot";
                ws.Cells["I7"].Value = "Vessel/Voyage";
                ws.Cells["J7"].Value = "Transit Mode";
                ws.Cells["K7"].Value = "BookingNo";
                ws.Cells["L7"].Value = "UserName";
                ws.Cells["M7"].Value = "Created On";


                r = ws.Cells["A7:M7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int slno1 = 1;
                int rw = 8;
                DataTable dtc = GetCntrHistoryReport(FromDt, ToDt,CntrID);
                string cntrno = "";

                for (int i = 0; i < dtc.Rows.Count; i++)
                {
                    //if (dtc.Rows[i]["CntrNo"].ToString() != cntrno)
                    //{

                    ws.Cells["A" + rw].Value = slno1;
                    ws.Cells["B" + rw].Value = dtc.Rows[i]["CntrNo"].ToString();
                    ws.Cells["C" + rw].Value = dtc.Rows[i]["Size"].ToString();
                    ws.Cells["D" + rw].Value = dtc.Rows[i]["StatusCode"].ToString();
                    ws.Cells["E" + rw].Value = dtc.Rows[i]["DtMovement"].ToString();
                    ws.Cells["F" + rw].Value = dtc.Rows[i]["Agency"].ToString();
                    ws.Cells["G" + rw].Value = dtc.Rows[i]["PortCode"].ToString();
                    ws.Cells["H" + rw].Value = dtc.Rows[i]["Depot"].ToString();
                    ws.Cells["I" + rw].Value = dtc.Rows[i]["VesVoy"].ToString();
                    ws.Cells["J" + rw].Value = dtc.Rows[i]["Transit"].ToString();
                    ws.Cells["K" + rw].Value = dtc.Rows[i]["BookingNo"].ToString();
                    ws.Cells["L" + rw].Value = dtc.Rows[i]["UserName"].ToString();
                    ws.Cells["M" + rw].Value = dtc.Rows[i]["DTCreated"].ToString();
                    // cntrno = dtc.Rows[i]["CntrNo"].ToString();
                    //}
                    slno1++;
                    rw += 1;


                }

                rw -= 1;

                ws.Cells["A7:M" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:M" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, rw, 24].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ContainerHistory.xlsx");
                Response.End();

            }

        }

        public DataTable GetCntrHistoryReport(string FromDt, string ToDt, string CntrID)
        {
            string _Query = "select C.ID,C.CntrNo,CT.StatusCode,CT.DtMovement, " +
                             " (select top 1 Type + '-' + Size from NVO_tblCntrTypes where ID = C.TypeID) Size, (select top 1 AgencyCode from NVO_AgencyMaster where ID = CT.AgencyID) Agency, " +
                       " (select top 1 PortCode from NVO_PortMaster where ID = CT.LocationID) PortCode, isnull((select top 1 DepName from NVO_DepotMaster where ID = CT.LocationID),'') Depot, " +
                      " isnull((select top 1 UserName from NVO_UserDetails where ID = CT.UserID),'') UserName, " +
                    " isnull((select top 1 GeneralName from NVO_GeneralMaster where ID = CT.ModeOfTransportID),'') Transit, " +
                      " isnull((select top 1 VesVoy from NVO_View_VoyageDetails where ID = CT.VesVoyID),'') VesVoy, " +
                    " isnull((select top 1 BookingNo from NVO_Booking where ID = CT.BLNumber),'') BookingNo, " +
                   " isnull((select top 1 UserName from NVO_UserDetails where ID = CT.UserID),'') UserName,CT.DTCreated from NVO_Containers C inner join NVO_ContainerTxns CT on CT.ContainerID = C.ID where ct.StatusCode not in ('PENDING')  ";
            string strWhere = "";

            if (CntrID != "" && CntrID != "0" && CntrID != "null" && CntrID != "?")
                if (strWhere == "")
                    strWhere += _Query + " and C.ID=" + CntrID;
                else
                    strWhere += " and C.ID =" + CntrID;

            if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " WHERE convert(varchar, CT.DtMovement, 23)   between '" + FromDt + "' and '" + ToDt + "'";
                else
                    strWhere += "  and convert(varchar, CT.DtMovement, 23)  between '" + FromDt + "' and '" + ToDt + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere + " Order By  C.ID", "");
        }

    }
}