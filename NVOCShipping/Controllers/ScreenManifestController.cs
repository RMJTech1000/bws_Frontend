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
    public class ScreenManifestController : Controller
    {
        // GET: ScreenManifest
        MasterManager Manag = new MasterManager();
        public ActionResult Index()
        {
            return View();
        }

        public void ScreenManifestReport(string VesVoyID, string User, string AgencyID)
        {
            DataTable dtv = GetScreenManifestReport(VesVoyID);
      
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ScreeningManifest");

                ws.Cells["A2"].Value = "Screening Manifest ";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:K2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "User :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = User;
                ws.Cells["B4"].Style.Font.Bold = true;
                ws.Cells["C4"].Value = "Downloaded Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;

                //ws.Cells["A5"].Value = "From Date :";
                //ws.Cells["A5"].Style.Font.Bold = true;
                //ws.Cells["B5"].Value = FromDt;
                //ws.Cells["B5"].Style.Font.Bold = true;
                //ws.Cells["C5"].Value = "To Date :";
                //ws.Cells["C5"].Style.Font.Bold = true;
                //ws.Cells["D5"].Value = ToDt;
                //ws.Cells["D5"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "BLNumber";
                ws.Cells["C7"].Value = "Container No";
                ws.Cells["D7"].Value = "Type-Size";
                ws.Cells["E7"].Value = "Shipper Name";
                ws.Cells["F7"].Value = "Shipper Address/Country";
                ws.Cells["G7"].Value = "Consignee Name";
                ws.Cells["H7"].Value = "Consignee Address/Country";
                ws.Cells["I7"].Value = "Notifier Name";
                ws.Cells["J7"].Value = "Notifier Address/Country";
                ws.Cells["K7"].Value = "Origin";
                ws.Cells["L7"].Value = "POL";
                ws.Cells["M7"].Value = "POD";
                ws.Cells["N7"].Value = "DELIVERY";
                ws.Cells["O7"].Value = "HS Codes";
                ws.Cells["P7"].Value = "Commodity Full Description";
          
                r = ws.Cells["A7:P7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 1;

                int rw = 8;
            if (dtv.Rows.Count > 0)
            {

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["ShipperAddress"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Consignee"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ConsigneeAddress"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Notify"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["NotifyAddress"].ToString();
                    if (dtv.Rows[i]["Origin"].ToString()!="")
                    {
                        ws.Cells["K" + rw].Value = dtv.Rows[i]["Origin"].ToString();
                    }
                    else
                    {
                        ws.Cells["K" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    }
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    if (dtv.Rows[i]["FPOD"].ToString() != "")
                    {
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    }
                    else
                    {
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    }
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["HSCode"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["Cargo"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:P" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:P" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:P" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:P" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ScreeningManifestReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetScreenManifestReport(string VesVoyID)
        {


            string _Query = "select BK.BKGID,bb.VesVoyID,BK.BLNumber,(select TOP 1 CntrNo from NVO_containers C WHERE C.ID=BC.CntrID ) CntrNo,Size,"+
                            " (select TOP 1 AgencyName from NVO_AgencyMaster A WHERE A.ID = BK.AgencyID) AS Shipper,"+
                           " ((select TOP 1 Address from NVO_AgencyMaster A WHERE A.ID = BK.AgencyID ) +' / ' + (select TOP 1 CountryName from NVO_AgencyMaster A inner " +
                           "join NVO_CountryMaster CM ON CM.ID =A.CountryID WHERE A.ID = BK.AgencyID)) AS ShipperAddress,"+
                           " (select TOP 1 AgencyName from NVO_AgencyMaster A WHERE A.ID = BK.DesAgentID ) AS Consignee,"+
                          " ((select TOP 1 Address from NVO_AgencyMaster A WHERE A.ID = BK.DesAgentID ) +' / ' + (select TOP 1 CountryName from NVO_AgencyMaster" +
                          " A inner join NVO_CountryMaster CM ON CM.ID = A.CountryID WHERE A.ID = BK.DesAgentID)) AS ConsigneeAddress, 'SAME AS CONSIGNEE' as Notify, " +
                          " 'SAME AS CONSIGNEE' as NotifyAddress,(select TOP 1 PortName from NVO_PortMaster A WHERE A.ID = bb.PODID ) AS POD,"+
                          " (select TOP 1 PortName from NVO_PortMaster A WHERE A.ID = bb.POLID ) AS POL,"+
                          " (select TOP 1 CityName from NVO_CityMaster A WHERE A.ID = bb.FPODID ) AS FPOD," +
                         " (select TOP 1 CityName from NVO_CityMaster A WHERE A.ID = bb.POOID ) AS Origin, BC.HSCode,BK.CagoDescription,BC.Cargo" +
                        " from NVO_BOL BK inner join NVO_BOLCntrDetails BC ON BC.BkgId = BK.BkgID inner join NVO_Booking bb ON bb.ID = BK.BkgID "+
                        " WHERE BLTypes = 40  and bb.VesVoyID="+ VesVoyID;

            string strWhere = "";

            //if (LocID != "undefined" && LocID != "" && LocID != "0" && LocID != "null" && LocID != "?")

            //    if (strWhere == "")
            //        strWhere += _Query + " WHERE cmd1.locationid=" + LocID;
            //    else
            //        strWhere += " and cmd1.locationid = " + LocID;


            //if (StatusCode != "" && StatusCode != "undefined")
            //    if (strWhere == "")
            //        strWhere += _Query + " WHERE cmd1.StatusCode ='" + StatusCode + "'";
            //    else
            //        strWhere += " and cmd1.StatusCode ='" + StatusCode + "'";



            //if (FromDt != "" && FromDt != "undefined" || ToDt != "" && ToDt != "undefined")
            //    if (strWhere == "")
            //        strWhere += _Query + " WHERE  convert(varchar, cmd1.DtMovement , 23) between '" + FromDt + "' and '" + ToDt + "'";
            //    else
            //        strWhere += "  and convert(varchar, cmd1.DtMovement , 23)  between '" + FromDt + "' and '" + ToDt + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
    }

}