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
    public class TDRMultiPortExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: TDRMultiPortExcel
        public ActionResult Index()
        {
            return View();
        }

        public void TDRExcelReport(string VesVoyID, string PODID, string TSPORTID, string AgentID, string VesselOpID, string AgencyID, string NextPortID)
        {
            DataTable _dtv = GetTerminalDepPdfValues(VesVoyID, AgencyID, NextPortID);
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

                //ws.Cells["A7"].Value = "NEXT PORT :";
                //ws.Cells["A7"].Style.Font.Bold = true;



                ws.Cells["A4:B7"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A4:B7"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A4:B7"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A4:B7"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells["B4"].Value = _dtv.Rows[0]["POL"].ToString();
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["B5"].Value = _dtv.Rows[0]["VesselName"].ToString();
                ws.Cells["B5"].Style.Font.Bold = true;
                ws.Cells["B6"].Value = _dtv.Rows[0]["ETA"].ToString();
                ws.Cells["B6"].Style.Font.Bold = true;
                //ws.Cells["B7"].Value = _dtv.Rows[0]["NextPort"].ToString();
                //ws.Cells["B7"].Style.Font.Bold = true;

                //ws.Cells["B8"].Value = " OCEANUS CONTAINER LINES PTE LTD";
                //ws.Cells["B8"].Style.Font.Bold = true;

                ws.Cells["H4:I7"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["H4:I7"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["H4:I7"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["H4:I7"].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                ws.Cells["H4"].Value = "TERMINAL";
                ws.Cells["H4"].Style.Font.Bold = true;
                ws.Cells["H5"].Value = "VOYAGE NUMBER :";
                ws.Cells["H5"].Style.Font.Bold = true;
                ws.Cells["H6"].Value = "ETD POL :";
                ws.Cells["H6"].Style.Font.Bold = true;

                //ws.Cells["H7"].Value = "ETA NEXT PORT :";
                //ws.Cells["H7"].Style.Font.Bold = true;


                ws.Cells["I4"].Value = _dtv.Rows[0]["TerminalName"].ToString();
                ws.Cells["I4"].Style.Font.Bold = true;
                ws.Cells["I5"].Value = _dtv.Rows[0]["VoyageNo"].ToString();
                ws.Cells["I5"].Style.Font.Bold = true;
                ws.Cells["I6"].Value = _dtv.Rows[0]["ETD"].ToString();
                ws.Cells["I6"].Style.Font.Bold = true;
                //ws.Cells["I7"].Value = _dtv.Rows[0]["NextPortETA"].ToString();
                //ws.Cells["I7"].Style.Font.Bold = true;


                int row = 7;
                for (int i = 1; i < _dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + row].Value = _dtv.Rows[i]["NextPortCol"].ToString();
                    ws.Cells["A" + row].Style.Font.Bold = true;
                    ws.Cells["B" + row].Value = _dtv.Rows[i]["POL"].ToString();
                    ws.Cells["B" + row].Style.Font.Bold = true;
                    ws.Cells["H" + row].Value = _dtv.Rows[i]["NextPortETACol"].ToString();
                    ws.Cells["H" + row].Style.Font.Bold = true;
                    ws.Cells["I" + row].Value = _dtv.Rows[i]["ETA"].ToString();
                    ws.Cells["I" + row].Style.Font.Bold = true;

                    ws.Cells["A" + row + ":" + "B" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A" + row + ":" + "B" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A" + row + ":" + "B" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A" + row + ":" + "B" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    ws.Cells["H" + row + ":" + "I" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["H" + row + ":" + "I" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["H" + row + ":" + "I" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["H" + row + ":" + "I" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    row++;
                }

                ws.Cells["A" + row].Value = "CONTAINER OPERATOR :";
                ws.Cells["A" + row].Style.Font.Bold = true;

                ws.Cells["B" + row].Value = " GLOBAL NETWORK LINES PTE LTD";
                ws.Cells["B" + row].Style.Font.Bold = true;
                ws.Cells["A" + row + ":" + "B" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "B" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "B" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "B" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                row++;
                // int rw = row + 2;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["A" + row].Value = "S. No.";
                ws.Cells["B" + row].Value = "CONTAINER#";
                ws.Cells["C" + row].Value = "SIZE TYPE";
                ws.Cells["D" + row].Value = "COMMODITY";
                ws.Cells["E" + row].Value = "SERVICE";
                ws.Cells["F" + row].Value = "TSPORT";
                ws.Cells["G" + row].Value = "POD";
                ws.Cells["H" + row].Value = "FINAL DESTINATION";
                ws.Cells["I" + row].Value = "PICK UP DATE";
                ws.Cells["J" + row].Value = "GW.KGS";
                ws.Cells["K" + row].Value = "BLNUMBER";
                ws.Cells["L" + row].Value = "POD AGENT";
                ws.Cells["M" + row].Value = "VESSEL OPERATOR";
                ws.Cells["N" + row].Value = "SLOT TERM";

                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":" + "N" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                r = ws.Cells["A" + row + ":" + "N" + row];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;
                row++;

                DataTable _dtC = GetCntrDtlsTerminalDepPdfValues(VesVoyID, PODID, TSPORTID, AgentID, VesselOpID);

                for (int k = 0; k < _dtC.Rows.Count; k++)
                {
                    ws.Cells["A" + row].Value = sl;
                    ws.Cells["B" + row].Value = _dtC.Rows[k]["CntrNo"].ToString();
                    ws.Cells["C" + row].Value = _dtC.Rows[k]["Size"].ToString();
                    ws.Cells["D" + row].Value = _dtC.Rows[k]["Commodity"].ToString();
                    ws.Cells["E" + row].Value = _dtC.Rows[k]["ServiceType"].ToString();
                    ws.Cells["F" + row].Value = _dtC.Rows[k]["TSPORT"].ToString();
                    ws.Cells["G" + row].Value = _dtC.Rows[k]["POD"].ToString();
                    ws.Cells["H" + row].Value = _dtC.Rows[k]["FPOD"].ToString();
                    ws.Cells["I" + row].Value = _dtC.Rows[k]["PickUpDate"].ToString();
                    ws.Cells["J" + row].Value = _dtC.Rows[k]["GrsWt"].ToString();
                    ws.Cells["K" + row].Value = _dtC.Rows[k]["BLNumber"].ToString();
                    ws.Cells["L" + row].Value = _dtC.Rows[k]["DestinationAgent"].ToString();
                    ws.Cells["M" + row].Value = _dtC.Rows[k]["Operator"].ToString();
                    ws.Cells["N" + row].Value = _dtC.Rows[k]["SlotTerm"].ToString();

                    ws.Cells["A" + row + ":" + "N" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A" + row + ":" + "N" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A" + row + ":" + "N" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A" + row + ":" + "N" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sl++;
                    row += 1;
                }

                row -= 1;



                ws.Cells[1, 1, row, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TDReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetTerminalDepPdfValues(string VesVoy, string AgencyID, string NextPortID)
        {
            string strWhere = "";


            DataTable dtvPr = GetFirstPort(VesVoy);
            int port1 = 0;
            port1 = Int32.Parse(dtvPr.Rows[0]["PortID"].ToString());

            string _Query = " select V.ID,VR.RID,(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) as VesselName,ExportVoyageCd AS VoyageNo,  Convert(varchar, VR.ETA, 106) AS ETA, Convert(varchar, VR.ETD, 106) AS ETD, " +
            " (select top(1) PortName from NVO_PortMainMaster where ID = VR.PortID) as POL, (select top(1) TerminalName from NVO_VoyageRoute inner join NVO_TerminalMaster On NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID where NVO_VoyageRoute.VoyageID = V.ID " +
             " order by NVO_VoyageRoute.RID asc) as TerminalName, CASE WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 1 THEN 'NextPort'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 2 THEN 'NextPort1' " +
              "  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 3 THEN 'NextPort2'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 4 THEN 'NextPort3' " +
              "  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 5 THEN 'NextPort4'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 6 THEN 'NextPort5' WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 7 THEN 'NextPort6' END NextPortCol, CASE WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 1 THEN 'NextPortETA' WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 2 THEN 'NextPortETA1' " +
              " WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 3 THEN 'NextPortETA2'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 4 THEN 'NextPortETA3' WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 5 THEN 'NextPortETA4' " +
              " WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 6 THEN 'NextPortETA5'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 7 THEN 'NextPortETA6'  END NextPortETACol from NVO_Voyage V INNER JOIN NVO_VoyageRoute VR ON VR.VoyageID = V.ID WHERE VR.VoyageID = " + VesVoy + " and  V.AgencyID=" + AgencyID;

            if (NextPortID != "" && NextPortID != "0" && NextPortID != null && NextPortID != "?")

                if (strWhere == "")
                    strWhere += _Query + " and VR.PortID in (" + port1 + "," + NextPortID + ")";
                else
                    strWhere += " and VR.PortID  in (" + port1 + "," + NextPortID + ")";
           


            if (strWhere == "")
                strWhere = _Query;

            strWhere += " UNION  " +
                       " select V.ID,VR.RID,(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) as VesVoy,ExportVoyageCd AS VoyageNo,  Convert(varchar, VR.ETA, 106) AS ETA, Convert(varchar, VR.ETD, 106) AS ETD, " +
          " (select top(1) PortName from NVO_PortMainMaster where ID = VR.PortID) as POL, (select top(1) TerminalName from NVO_VoyageRoute inner join NVO_TerminalMaster On NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID where NVO_VoyageRoute.VoyageID = V.ID " +
           " order by NVO_VoyageRoute.RID asc) as TerminalName, CASE WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 1 THEN 'NextPort'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 2 THEN 'NextPort1' " +
            "  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 3 THEN 'NextPort2'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 4 THEN 'NextPort3' " +
            "  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 5 THEN 'NextPort4' " +
            "  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 6 THEN 'NextPort5' WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 7 THEN 'NextPort6' END NextPortCol, CASE WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 1 THEN 'NextPortETA' WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 2 THEN 'NextPortETA1' " +
            " WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 3 THEN 'NextPortETA2'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 4 THEN 'NextPortETA3' WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 5 THEN 'NextPortETA4'" +
            " WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 6 THEN 'NextPortETA5'  WHEN ROW_NUMBER() OVER(ORDER BY VR.RID) = 7 THEN 'NextPortETA6'  END NextPortETACol from NVO_Voyage V INNER JOIN NVO_VoyageRoute VR ON VR.VoyageID = V.ID  inner join NVO_Gateway_VoyageBL GT ON GT.VoyageID =V.ID  where VR.VoyageID=" + VesVoy + " and  GT.AgencyID=" + AgencyID;


            if (NextPortID != "" && NextPortID != "0" && NextPortID != null && NextPortID != "?")

                if (strWhere == "")
                    strWhere += _Query + " and VR.PortID in (" + port1 + "," + NextPortID + ")";
                else
                    strWhere += " and VR.PortID  in (" + port1 + "," + NextPortID + ")";


            return Manag.GetViewData(strWhere + " ORDER BY VR.RID ", "");
        }

        public DataTable GetFirstPort(string VesVoy)
        {

            string _Query = "select top (1) PortID from NVO_VoyageRoute WHERE VoyageID=" + VesVoy;
            return Manag.GetViewData(_Query, "");
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
    }
}