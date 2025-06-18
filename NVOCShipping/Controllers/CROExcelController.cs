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
    public class CROExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }
        public void CreateCROExcel(string BkgID, string CusID, string ReleaseOrderNo, string User, string AgencyID)
        {
            DataTable dtv = GetCROSearchValues(BkgID, CusID, ReleaseOrderNo , AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CROList");

                ws.Cells["A2"].Value = "CRO List";
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
                ws.Cells["B7"].Value = "ReleaseOrderNo";
                ws.Cells["C7"].Value = "BookingNo";
                ws.Cells["D7"].Value = "ValidFrom";
                ws.Cells["E7"].Value = "ValidTill";
                ws.Cells["F7"].Value = "Customer";
                ws.Cells["G7"].Value = "Vessel/Voyage";
                ws.Cells["H7"].Value = "ETA";
                ws.Cells["I7"].Value = "ETD";
                ws.Cells["J7"].Value = "CutOff Date";
                ws.Cells["K7"].Value = "Pick Up Depot";
                ws.Cells["L7"].Value = "Surveyor";
                ws.Cells["M7"].Value = "LineCode";
                ws.Cells["N7"].Value = "Remarks";
                ws.Cells["O7"].Value = "Status";
                ws.Cells["P7"].Value = "Cntr Types";
                ws.Cells["Q7"].Value = "Required Qty";
                r = ws.Cells["A7:Q7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ReleaseOrderNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ValidFrom"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ValidTill"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Customer"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ETA"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["ETD"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["CutOffDt"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["pickUpDepot"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["Surveyor"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["LineCode"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Remarks"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Status"].ToString();

                    DataTable dtQ = GetQTYValues(dtv.Rows[i]["ID"].ToString());
                    for (int k = 0; k < dtQ.Rows.Count; k++)
                    {
                    ws.Cells["P" + rw].Value = dtQ.Rows[k]["CntrType"].ToString();
                    ws.Cells["Q" + rw].Value = dtQ.Rows[k]["ReqQty"].ToString();
                    rw++;

                    }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:Q" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:Q" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:Q" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:Q" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CROList.xlsx");
                Response.End();

            }

        }

        public DataTable GetCROSearchValues(string BkgID, string CusID, string ReleaseOrderNo, string AgencyID)
        {
            string strWhere = "";
            string _Query = "Select ID, ReleaseOrderNo, (select top(1) BookingNo from NVO_Booking where ID = BkgID) as BookingNo, (select top(1) CustomerName from NVO_CustomerMaster where ID = CusID) as Customer,(Select top 1(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID)  + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID)from NVO_Voyage V where V.ID = NVO_CROMaster.VesselID )as VesVoy, "+
          " convert(varchar, ValidTill, 106) as ValidTill,convert(varchar, Date, 106) as ValidFrom,convert(varchar, ETADate, 106) as ETA,convert(varchar, ETDDate, 106) as ETD,convert(varchar, CutDate, 106) as CutoffDt,(select top(1) DepName from NVO_DepotMaster where ID = PickDepoID) as pickUpDepot,(select top(1) CustomerName from NVO_CustomerMaster where ID = Surveyor) as Surveyor,Linecode,Remarks, case when CROStatus = 1 then 'CANCELLED' ELSE 'ACTIVE' END AS Status from NVO_CROMaster";

           if (ReleaseOrderNo != "" && ReleaseOrderNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where ReleaseOrderNo like '%" + ReleaseOrderNo + "%'";
                else
                    strWhere += " and ReleaseOrderNo like '%" + ReleaseOrderNo + "%'";

            if (BkgID != "" && BkgID != "0" && BkgID != "null" && BkgID != "?")
                if (strWhere == "")
                    strWhere += _Query + " where BkgID= " + BkgID;
                else
                    strWhere += " and BkgID= " + BkgID;

            if (CusID != "" && CusID != "0" && CusID != "null" && CusID != "?")
                if (strWhere == "")
                    strWhere += _Query + " where CusID=" + CusID;
                else
                    strWhere += " and CusID=" + CusID;

            if (AgencyID != "" && AgencyID != "0" && AgencyID != "2" && AgencyID != "undefined" && AgencyID != "null")

                if (strWhere == "")
                    strWhere += _Query + " where NVO_CROMaster.AgencyID = " + AgencyID.ToString();
                else
                    strWhere += " and NVO_CROMaster.AgencyID = " + AgencyID.ToString();

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
        public DataTable GetQTYValues(string CROID)
        {


            string _Query = "select CID,ReqQty,(Select top 1 Size from NVO_tblCntrTypes WHERE ID = CntrTypeID) as CntrType from NVO_CRODetails where CROID=" + CROID;

            return Manag.GetViewData(_Query, "");
        }


    }
}