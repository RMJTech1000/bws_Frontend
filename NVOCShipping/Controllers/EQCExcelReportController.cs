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
using System.Text;

namespace NVOCShipping.Controllers
{
    public class EQCExcelReportController : Controller
    {
        // GET: EQCExcelReport
        MasterManager Manag = new MasterManager();
        public ActionResult Index()
        {
            return View();
        }
        public int? TryParseNullable(string val)
        {
            int outValue;
            int.TryParse(val, out outValue);
            //return int.TryParse(val, out outValue) ? (int?)outValue : null;
            if (outValue > 0)
                return outValue;
            else
                return null;
        }
        public void EQCIdlingReportValues(string Agency, string Type, string User)
        {
            DataTable dtv = GetEQCIdlingReport(Agency, Type);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                #region 1st page
                var ws = pck.Workbook.Worksheets.Add("Summary");

                ws.Cells["A2"].Value = "EQCIdling Report List";
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
                ws.Cells["C4"].Value = "Downloaded Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D4"].Style.Font.Bold = true;


                //Record Headers
                ws.Cells["A7:E7"].Value = "";
                ws.Cells["A7:E7"].Merge = true;
                ws.Cells["A7:E7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["F7:AC7"].Value = "Dry";
                ws.Cells["F7:AC7"].Merge = true;
                ws.Cells["F7:AC7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AD7:AU7"].Value = "Reefer";
                ws.Cells["AD7:AU7"].Merge = true;
                ws.Cells["AD7:AU7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["F8:I8"].Value = "< 7 Days";
                ws.Cells["F8:I8"].Merge = true;
                ws.Cells["F8:I8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["J8:M8"].Value = "8 - 14 Days";
                ws.Cells["J8:M8"].Merge = true;
                ws.Cells["J8:M8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["N8:Q8"].Value = "15 - 29 Days";
                ws.Cells["N8:Q8"].Merge = true;
                ws.Cells["N8:Q8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["R8:U8"].Value = "30 - 59 Days";
                ws.Cells["R8:U8"].Merge = true;
                ws.Cells["R8:U8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["V8:Y8"].Value = "> 59 Days";
                ws.Cells["V8:Y8"].Merge = true;
                ws.Cells["V8:Y8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["Z8:AC8"].Value = "TOTAL";
                ws.Cells["Z8:AC8"].Merge = true;
                ws.Cells["Z8:AC8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AD8:AF8"].Value = "< 7 Days";
                ws.Cells["AD8:AF8"].Merge = true;
                ws.Cells["AD8:AF8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AG8:AI8"].Value = "8 - 14 Days";
                ws.Cells["AG8:AI8"].Merge = true;
                ws.Cells["AG8:AI8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AJ8:AL8"].Value = "15 - 29 Days";
                ws.Cells["AJ8:AL8"].Merge = true;
                ws.Cells["AJ8:AL8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AM8:AO8"].Value = "30 - 59 Days";
                ws.Cells["AM8:AO8"].Merge = true;
                ws.Cells["AM8:AO8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AP8:AR8"].Value = "> 59 Days";
                ws.Cells["AP8:AR8"].Merge = true;
                ws.Cells["AP8:AR8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["AS8:AU8"].Value = "TOTAL";
                ws.Cells["AS8:AU8"].Merge = true;
                ws.Cells["AS8:AU8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                r = ws.Cells["F7:AC7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightSalmon);
                r = ws.Cells["AD7:AU7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightPink);

                ws.Cells["A9"].Value = "Country Name";
                ws.Cells["B9"].Value = "Agency";
                ws.Cells["C9"].Value = "Geolocation";
                ws.Cells["D9"].Value = "Type";
                ws.Cells["E9"].Value = "StatusCode";

                ws.Cells["F9"].Value = "20GP";
                ws.Cells["G9"].Value = "20HQ";
                ws.Cells["H9"].Value = "40GP";
                ws.Cells["I9"].Value = "40HQ";

                ws.Cells["J9"].Value = "20GP";
                ws.Cells["K9"].Value = "20HQ";
                ws.Cells["L9"].Value = "40GP";
                ws.Cells["M9"].Value = "40HQ";

                ws.Cells["N9"].Value = "20GP";
                ws.Cells["O9"].Value = "20HQ";
                ws.Cells["P9"].Value = "40GP";
                ws.Cells["Q9"].Value = "40HQ";

                ws.Cells["R9"].Value = "20GP";
                ws.Cells["S9"].Value = "20HQ";
                ws.Cells["T9"].Value = "40GP";
                ws.Cells["U9"].Value = "40HQ";

                ws.Cells["V9"].Value = "20GP";
                ws.Cells["W9"].Value = "20HQ";
                ws.Cells["X9"].Value = "40GP";
                ws.Cells["Y9"].Value = "40HQ";

                ws.Cells["Z9"].Value = "20GP";
                ws.Cells["AA9"].Value = "20HQ";
                ws.Cells["AB9"].Value = "40GP";
                ws.Cells["AC9"].Value = "40HQ";

                ws.Cells["AD9"].Value = "20'RF";
                ws.Cells["AE9"].Value = "40'RF";
                ws.Cells["AF9"].Value = "HQ'RF";

                ws.Cells["AG9"].Value = "20'RF";
                ws.Cells["AH9"].Value = "40'RF";
                ws.Cells["AI9"].Value = "HQ'RF";

                ws.Cells["AJ9"].Value = "20'RF";
                ws.Cells["AK9"].Value = "40'RF";
                ws.Cells["AL9"].Value = "HQ'RF";

                ws.Cells["AM9"].Value = "20'RF";
                ws.Cells["AN9"].Value = "40'RF";
                ws.Cells["AO9"].Value = "HQ'RF";

                ws.Cells["AP9"].Value = "20'RF";
                ws.Cells["AQ9"].Value = "40'RF";
                ws.Cells["AR9"].Value = "HQ'RF";

                ws.Cells["AS9"].Value = "20'RF";
                ws.Cells["AT9"].Value = "40'RF";
                ws.Cells["AU9"].Value = "HQ'RF";

                int sl = 1;
                int rw = 10;

                string input = Type;  

                switch (input)
                {
                    case "1":
                        ws.Cells["D10:D14"].Value = "EXPORT";
                        ws.Cells["D10:D14"].Merge = true;
                        ws.Cells["D10:D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        //ws.Cells["D10:D14"].AutoFitColumns();

                        ws.Cells["E10"].Value = "MS";
                        ws.Cells["E11"].Value = "FC";
                        ws.Cells["E12"].Value = "FI";
                        ws.Cells["E13"].Value = "FL";
                        ws.Cells["E14"].Value = "FB";

                        ws.Cells["A7:AU" + 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 14].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[1, 1, 14, 50].AutoFitColumns();

                        //ws.Cells["F8:I8" + 14].Style.Font.Bold = true;
                        //ws.Cells["F8:I8" + 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //ws.Cells["F8:I8" + 14].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                        //ws.Cells["J8:M8" + 14].Style.Font.Bold = true;
                        //ws.Cells["J8:M8" + 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //ws.Cells["J8:M8" + 14].Style.Fill.BackgroundColor.SetColor(Color.MediumPurple);

                        break;
                    case "3":
                        ws.Cells["D10:D12"].Value = "TRANSHIPMENT";
                        ws.Cells["D10:D12"].Merge = true;
                        ws.Cells["D10:D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                       // ws.Cells["D10:D12"].AutoFitColumns();

                        ws.Cells["E10"].Value = "TZ";
                        ws.Cells["E11"].Value = "TZFB";
                        ws.Cells["A7:AU" + 12].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 12].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 12].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                       // ws.Cells[1, 1, 12, 50].AutoFitColumns();
                        break;
                     
                    case "2":
                        ws.Cells["D10:D14"].Value = "IMPORT";
                        ws.Cells["D10:D14"].Merge = true;
                        ws.Cells["D10:D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells["D10:D14"].AutoFitColumns();

                        ws.Cells["E10"].Value = "FV";
                        ws.Cells["E11"].Value = "FI";
                        ws.Cells["E12"].Value = "FVICD";
                        ws.Cells["E13"].Value = "DV";
                        ws.Cells["E14"].Value = "MA";

                       ws.Cells["A7:AU" + 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 14].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[1, 1, 14, 50].AutoFitColumns();
                        break;
                    case "4":
                        ws.Cells["D10:D12"].Value = "DAMAGE";
                        ws.Cells["D10:D12"].Merge = true;
                        ws.Cells["D10:D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells["D10:D12"].AutoFitColumns();

                        ws.Cells["E10"].Value = "DL";
                        ws.Cells["E11"].Value = "UR";
                        ws.Cells["A7:AU" + 12].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 12].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 12].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[1, 1, 12, 50].AutoFitColumns();
                        break;
                    case "5":
                        ws.Cells["D10"].Value = "AVAILABLE";
                        ws.Cells["E10"].Value = "AV";
                        ws.Cells["A7:AU" + 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[1, 1, 11, 50].AutoFitColumns();
                        break;
                    default:
                        ws.Cells["D10:D14"].Value = "EXPORT";
                        ws.Cells["D10:D14"].Merge = true;
                        ws.Cells["D10:D14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        //ws.Cells["D10:D14"].AutoFitColumns();

                        ws.Cells["E10"].Value = "MS";
                        ws.Cells["E11"].Value = "FC";
                        ws.Cells["E12"].Value = "FI";
                        ws.Cells["E13"].Value = "FL";
                        ws.Cells["E14"].Value = "FB";

                        ws.Cells["D15:D16"].Value = "TRANSHIPMENT";
                        ws.Cells["D15:D16"].Merge = true;
                        ws.Cells["D15:D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                      

                        ws.Cells["E15"].Value = "TZ";
                        ws.Cells["E16"].Value = "TZFB";

                        ws.Cells["D17:D21"].Value = "IMPORT";
                        ws.Cells["D17:D21"].Merge = true;
                        ws.Cells["D17:D21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                       

                        ws.Cells["E17"].Value = "FV";
                        ws.Cells["E18"].Value = "FI";
                        ws.Cells["E19"].Value = "FVICD";
                        ws.Cells["E20"].Value = "DV";
                        ws.Cells["E21"].Value = "MA";

                        ws.Cells["D22:D23"].Value = "DAMAGE";
                        ws.Cells["D22:D23"].Merge = true;
                        ws.Cells["D22:D23"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        

                        ws.Cells["E22"].Value = "DL";
                        ws.Cells["E23"].Value = "UR";

                        ws.Cells["D24"].Value = "AVAILABLE";
                        ws.Cells["E24"].Value = "AV";
                        int Rowv = 24;

                        //ws.Cells["F8:I8" + Rowv].Style.Font.Bold = true;
                        //ws.Cells["F8:I8" + Rowv].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //ws.Cells["F8:I8" + Rowv].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                        //ws.Cells["J8:M8" + Rowv].Style.Font.Bold = true;
                        //ws.Cells["J8:M8" + Rowv].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //ws.Cells["J8:M8" + Rowv].Style.Fill.BackgroundColor.SetColor(Color.MediumPurple);

                        ws.Cells["A7:AU" + Rowv].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + Rowv].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + Rowv].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A7:AU" + Rowv].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[1, 1, Rowv, 50].AutoFitColumns();

                        break;

                }
            

                //for (int i = 0; i < dtv.Rows.Count; i++)
                //{
                    ws.Cells["A" + rw].Value = "";
                    ws.Cells["B" + rw].Value = "";
                    ws.Cells["C" + rw].Value = "";
                    ws.Cells["D" + rw].Value = "";
                    ws.Cells["E" + rw].Value = "";
                    //ws.Cells["F" + rw].Value = "";
                    //ws.Cells["G" + rw].Value = "";
                    ws.Cells["H" + rw].Value = "";
                    ws.Cells["I" + rw].Value = "";
                    ws.Cells["J" + rw].Value = "";
                    ws.Cells["K" + rw].Value = "";
                    ws.Cells["L" + rw].Value = "";
                    ws.Cells["M" + rw].Value = "";
                    ws.Cells["N" + rw].Value = "";
                    ws.Cells["O" + rw].Value = "";
                    ws.Cells["P" + rw].Value = "";
                    ws.Cells["Q" + rw].Value = "";
                    ws.Cells["R" + rw].Value = "";
                    ws.Cells["S" + rw].Value = "";
                    ws.Cells["T" + rw].Value = "";
                    ws.Cells["U" + rw].Value = "";
                    ws.Cells["V" + rw].Value = "";
                    ws.Cells["W" + rw].Value = "";
                    ws.Cells["X" + rw].Value = "";
                    ws.Cells["Y" + rw].Value = "";
                    ws.Cells["Z" + rw].Value = "";
                    ws.Cells["AA" + rw].Value = "";
                    ws.Cells["AB" + rw].Value = "";
                    ws.Cells["AC" + rw].Value = "";
                    ws.Cells["AD" + rw].Value = "";
                    ws.Cells["AE" + rw].Value = "";
                    ws.Cells["AF" + rw].Value = "";
                    ws.Cells["AG" + rw].Value = "";
                    ws.Cells["AH" + rw].Value = "";
                    ws.Cells["AI" + rw].Value = "";
                    ws.Cells["AJ" + rw].Value = "";
                    ws.Cells["AK" + rw].Value = "";
                    ws.Cells["AL" + rw].Value = "";
                    ws.Cells["AM" + rw].Value = "";
                    ws.Cells["AN" + rw].Value = "";
                    ws.Cells["AO" + rw].Value = "";
                    ws.Cells["AP" + rw].Value = "";
                    ws.Cells["AQ" + rw].Value = "";
                    ws.Cells["AR" + rw].Value = "";
                    ws.Cells["AS" + rw].Value = "";
                    ws.Cells["AT" + rw].Value = "";
                    ws.Cells["AU" + rw].Value = "";


                    sl++;
                    rw += 1;
               // }

                rw -= 1;






                #endregion

                #region 2ND SHEET
                ws = pck.Workbook.Worksheets.Add("IdlingBreakUp");
                ws.Cells["A2"].Value = "Idling Detail Report List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r = ws.Cells["A2:X2"];
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
                ws.Cells["B7"].Value = "Type";
                ws.Cells["C7"].Value = "Country";
                ws.Cells["D7"].Value = "Geo Location";
                ws.Cells["E7"].Value = "Agency";
                ws.Cells["F7"].Value = "Container No";
                ws.Cells["G7"].Value = "Container Type";
                ws.Cells["H7"].Value = "StatusCode";
                ws.Cells["I7"].Value = "Date Movement";
                ws.Cells["J7"].Value = "Ageing";
                ws.Cells["K7"].Value = "Location";
                ws.Cells["L7"].Value = "Vessel/Voyage";
                ws.Cells["M7"].Value = "Transit";
                ws.Cells["N7"].Value = "Depot";
                ws.Cells["O7"].Value = "BLNumber";
                ws.Cells["P7"].Value = "Orgin";
                ws.Cells["Q7"].Value = "POL";
                ws.Cells["R7"].Value = "POD";
                ws.Cells["S7"].Value = "FPOD";
                ws.Cells["T7"].Value = "Customer";
                ws.Cells["U7"].Value = "Vendor";
                ws.Cells["V7"].Value = "Status";
                ws.Cells["W7"].Value = "Created By";
                ws.Cells["X7"].Value = "Created On";
                ws.Cells["Y7"].Value = "Container Owner";
                ws.Cells["Z7"].Value = "Leasing Partner";
                ws.Cells["AA7"].Value = "Leasing Term";
                r = ws.Cells["A7:AA7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int slno = 1;
                int row = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + row].Value = slno;
                    ws.Cells["B" + row].Value = "";
                    ws.Cells["C" + row].Value = "";
                    ws.Cells["C" + row].Value = "";
                    ws.Cells["D" + row].Value = "";
                    ws.Cells["E" + row].Value = "";
                    ws.Cells["F" + row].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["G" + row].Value = dtv.Rows[i]["CntrType"].ToString();
                    ws.Cells["H" + row].Value = dtv.Rows[i]["StatusCode"].ToString();
                    ws.Cells["I" + row].Value = dtv.Rows[i]["DtMovement"].ToString();
                    ws.Cells["J" + row].Value = "";
                    ws.Cells["K" + row].Value = "";
                    ws.Cells["L" + row].Value = "";
                    ws.Cells["M" + row].Value = "";
                    ws.Cells["N" + row].Value = "";
                    ws.Cells["O" + row].Value = "";
                    ws.Cells["P" + row].Value = "";
                    ws.Cells["Q" + row].Value = "";
                    ws.Cells["R" + row].Value = "";
                    ws.Cells["S" + row].Value = "";
                    ws.Cells["T" + row].Value = "";
                    ws.Cells["U" + row].Value = "";
                    ws.Cells["V" + row].Value = "";
                    ws.Cells["W" + row].Value = "";
                    ws.Cells["X" + row].Value = "";
                    ws.Cells["Y" + row].Value = "";
                    ws.Cells["Z" + row].Value = "";
                    ws.Cells["AA" + row].Value = "";
                    slno++;
                    row += 1;
                }

                row -= 1;

                ws.Cells["A7:AA" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AA" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AA" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AA" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[1, 1, row, 24].AutoFitColumns();
                #endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=EQCIdlingReport.xlsx");
                Response.End();

            }

        }


        public DataTable GetEQCIdlingReport(string Agency, string Type)
        {


            string _Query = " Select DISTINCT C.ID,C.CntrNo,CT.Type +'-'+ CT.Size as CntrType, Datediff(d, (select top(1) DtMovement   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc),GETDATE()) Days, " +
                " (select top(1) DtMovement   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) DtMovement," +
                 " (select top(1) A.AgencyName   from NVO_AgencyMaster A inner join NVO_ContainerTxns ct on ct.containerID =C.ID where A.ID = ct.AgencyID order by ct.DtMovement desc) AgencyName ," +
                  " (select top(1) P.PortName   from NVO_PortMaster P inner join NVO_ContainerTxns ct on ct.containerID =C.ID where P.ID = ct.NextPortID order by ct.DtMovement desc) LOCATION ," +
              " (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) StatusCode FROM NVO_Containers C INNER JOIN NVO_tblCntrTypes CT ON CT.ID = C.TypeID " +
             " WHERE (select top(1) StatusCode   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc)  NOT IN('PENDING')  and  (select top(1) StatusCode   from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc)  IN('MS','FC','FI','FL','FB','TZ','TZFB','FV','FVICD','DV','MA','DL','UR','AV')  ";

            string strWhere = "";

            //if (PortID != "" && PortID != "0" && PortID != "null" && PortID != "?")

            //    if (strWhere == "")
            //        strWhere += _Query + " and (select top(1) NextPortID from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc)=" + PortID;
            //    else
            //        strWhere += " and (select top(1) NextPortID from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) = " + PortID;


            //if (StatusCode != "" && StatusCode != "undefined")
            //    if (strWhere == "")
            //        strWhere += _Query + " and (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) ='" + StatusCode + "'";
            //    else
            //        strWhere += " and (select top(1) StatusCode from NVO_ContainerTxns where ContainerID = C.ID order by DtMovement desc) ='" + StatusCode + "'";



          

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }



        public void EQCAgeWiseReportNew(string cntrStatus, string Country, string GeoLoc, string LeaseTerm, string Grade, string CntrType, string User, string GeoLocName, string CountryName, string GradeName, string CntrTypeName, string LeaseTermName)
        {
            string criteria = "0";
            int rw_bgn = 0;
            int rw_end = 0;

            int rw = 0;
            //  int mrg_bgn = 0, mrg_end = 0;

            string styleDry = "#FFCC99";
            string styleReefer = "#FFCCFF";
            string styleFlatRack = "#99FF99";
            string styleOpenTop = "#99FFFF";
            string styleTank = "#99FF99";

            DataTable dtNew = null;

            Server.ScriptTimeout = 3600;

            #region Excel
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("EQCAgeWise");

            Color colFromHex;


            ws.Cells["A1"].Value = "EQC AGEWISE REPORT";
            ExcelRange r = ws.Cells["A1:AB1"];
            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            r.Merge = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            r.Style.Font.Color.SetColor(Color.Black);
            r.Style.Font.Bold = true;
            r.Style.Font.SetFromFont(new Font("Arial", 12));

            var allCells = ws.Cells[2, 1, 500, 50];
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Arial", 8));

            ExcelRange rng = ws.Cells["A2:E2"];

            //ws.Cells["A3"].Value = "Line :";
            //ws.Cells["B3"].Value = "OCEANUS LINES";

            ws.Cells["A4"].Value = "User :";
            ws.Cells["B4"].Value = User;
            ws.Cells["A4"].Style.Font.Bold = true;
            ws.Cells["B4"].Style.Font.Bold = true;

            ws.Cells["A5"].Value = "Date :";
            ws.Cells["B5"].Value = System.DateTime.Today.Date.ToShortDateString();
            ws.Cells["A5"].Style.Font.Bold = true;
            ws.Cells["B5"].Style.Font.Bold = true;

            ws.Cells["A6"].Value = "Status :";
            ws.Cells["B6"].Value = cntrStatus;
            ws.Cells["A6"].Style.Font.Bold = true;
            ws.Cells["B6"].Style.Font.Bold = true;
            // e_msg += "<br>" + cntrStatus;

            if (Session["UserName"] == null)
                Session["UserName1"] = "SYS GEN";
            else
                Session["UserName1"] = Session["UserName"];

            rw = 8;

            var Cntrs = cntrStatus.Split(',');
            DataTable _dtValue = null;
            string country1 = "", GeoLocation = "";
            string mainsector = ""; string mainsectorfr = ""; string mainsectorOt = ""; string mainsectortnk = ""; string mainsectorrf = "";
            int country_start = 0, country_end = 0, GeoLoc_start = 0, GeoLoc_end = 0;
            int mainsector_start = 0, mainsector_end = 0;

            for (int y = 0; y < Cntrs.Length; y++)
            {
                _dtValue = _dtAgewiseRepValue(cntrStatus, Country, GeoLoc, Grade, LeaseTerm, CntrType, criteria);
            }


            #region Dry Cntr Type
            rw = 8;
            rng = ws.Cells["E" + rw + ":AB" + rw];
            rng.Merge = true;
            rng.Style.Font.Bold = true;
            rng.Value = "DRY";
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw = 9;
            rng = ws.Cells["A" + rw + ":AD" + rw];
            rng.Style.Font.Bold = true;


            rng = ws.Cells["G" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";

            rng = ws.Cells["K" + rw + ":N" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";

            rng = ws.Cells["O" + rw + ":R" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";

            rng = ws.Cells["S" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";

            rng = ws.Cells["W" + rw + ":Z" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";

            rng = ws.Cells["AA" + rw + ":AD" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;

            rng = ws.Cells["A" + rw + ":AD" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            ws.Cells["E" + rw].Value = "Lease Term";
            ws.Cells["F" + rw].Value = "StatusCode";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["G" + rw].Value = "20GP";
            ws.Cells["H" + rw].Value = "20HQ";
            ws.Cells["I" + rw].Value = "40GP";
            ws.Cells["J" + rw].Value = "40HQ";

            //8 - 14 Days
            ws.Cells["K" + rw].Value = "20GP";
            ws.Cells["L" + rw].Value = "20HQ";
            ws.Cells["M" + rw].Value = "40GP";
            ws.Cells["N" + rw].Value = "40HQ";

            //15 - 29 Days
            ws.Cells["O" + rw].Value = "20GP";
            ws.Cells["P" + rw].Value = "20HQ";
            ws.Cells["Q" + rw].Value = "40GP";
            ws.Cells["R" + rw].Value = "40HQ";

            //>30 - 59 Days
            ws.Cells["S" + rw].Value = "20GP";
            ws.Cells["T" + rw].Value = "20HQ";
            ws.Cells["U" + rw].Value = "40GP";
            ws.Cells["V" + rw].Value = "40HQ";

            //> 59 Days
            ws.Cells["W" + rw].Value = "20GP";
            ws.Cells["X" + rw].Value = "20HQ";
            ws.Cells["Y" + rw].Value = "40GP";
            ws.Cells["Z" + rw].Value = "40HQ";

            //TOTAL
            ws.Cells["AA" + rw].Value = "20GP";
            ws.Cells["AB" + rw].Value = "20HQ";
            ws.Cells["AC" + rw].Value = "40GP";
            ws.Cells["AD" + rw].Value = "40HQ";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;


            DataTable dt = _dtValue;
            DataView dv1 = new DataView(dt);

            dtNew = dv1.ToTable();

            rw++;
            int flag1 = 0;
            object val_this = null, val_prev = null, val_this_GeoLoc = null, val_prev_GeoLoc = null;
            object val_this1 = null, val_prev1 = null;

            int dtrows = dtNew.Rows.Count - 1;
            int rw_bgn_sub = rw;
            StringBuilder sbSub = new StringBuilder();
            StringBuilder sbSub1 = new StringBuilder();
            StringBuilder sbSub2 = new StringBuilder();
            StringBuilder sbSub3 = new StringBuilder();
            StringBuilder sbSub4 = new StringBuilder();
            ExcelRange rngtemp;
            // main sector
            for (int i = 0; i < dtNew.Rows.Count; i++)
            {


                if (mainsector.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {

                    if (mainsector != "" && i > 0)
                    {
                        if (dtNew.Rows.Count > 0)
                        {
                            rw++;
                            rw_end = rw - 1;
                            sbSub.AppendLine(rw_end.ToString());

                            rngtemp = ws.Cells["A" + rw_end + ":D" + rw_end];
                            rngtemp.Value = "Sub Total - " + mainsector;
                            rngtemp.Merge = true;
                            rngtemp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // rngtemp.Style.Font.Bold = true;

                            //Region merge 13-sep-2019  rgs
                            ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["W" + rw_end].Formula = string.Format("=SUM(W{0}:W{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["X" + rw_end].Formula = string.Format("=SUM(X{0}:X{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["Y" + rw_end].Formula = string.Format("=SUM(Y{0}:Y{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["Z" + rw_end].Formula = string.Format("=SUM(Z{0}:Z{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AA" + rw_end].Formula = string.Format("=SUM(AA{0}:AA{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AB" + rw_end].Formula = string.Format("=SUM(AB{0}:AB{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AC" + rw_end].Formula = string.Format("=SUM(AC{0}:AC{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AD" + rw_end].Formula = string.Format("=SUM(AD{0}:AD{1})", rw_bgn_sub, (rw_end - 1));
                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }
                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsector = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsector;
                    rw_bgn_sub = rw;

                }
                country1 = dtNew.Rows[i]["CountryName"].ToString();
                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    // country_end = rw - 2;
                    country_end = rw - 1;
                    ws.Cells[country_start, 2, country_end, 2].Merge = true;
                    ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 2, country_end, 2].Merge = true;
                    ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                // For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {

                    GeoLoc_end = rw - 1;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["E" + rw].Value = dtNew.Rows[i]["LeaseTerm"].ToString();
                ws.Cells["F" + rw].Value = dtNew.Rows[i]["StatusCode"].ToString();
                ws.Cells["F" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry20L7"].ToString());
                ws.Cells["H" + rw].Value = TryParseNullable("");
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry40L7"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQL7"].ToString());



                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry208to14"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable("");
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry408to14"].ToString());
                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQ8to14"].ToString());

                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry2015to29"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable("");
                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry4015to29"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQ15to29"].ToString());

                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry2030to59"].ToString());
                ws.Cells["T" + rw].Value = TryParseNullable("");
                ws.Cells["U" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry4030to59"].ToString());
                ws.Cells["V" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQ30to59"].ToString());

                ws.Cells["W" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry20G59"].ToString());
                ws.Cells["X" + rw].Value = TryParseNullable("");
                ws.Cells["Y" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry40G59"].ToString());
                ws.Cells["Z" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQG59"].ToString());

                ws.Cells["AA" + rw].Formula = "=G" + rw + "+K" + rw + "+O" + rw + "+S" + rw + "+W" + rw;//"=D" + rw + "+H" + rw + "+L" + rw + "+P" + rw + "+T" + rw;
                ws.Cells["AB" + rw].Formula = "=H" + rw + "+L" + rw + "+P" + rw + "+T" + rw + "+X" + rw;//"=E" + rw + "+I" + rw + "+M" + rw + "+R" + rw + "+U" + rw;
                ws.Cells["AC" + rw].Formula = "=I" + rw + "+M" + rw + "+Q" + rw + "+U" + rw + "+Y" + rw;//"=F" + rw + "+J" + rw + "+N" + rw + "+R" + rw + "+V" + rw;
                ws.Cells["AD" + rw].Formula = "=J" + rw + "+N" + rw + "+R" + rw + "+V" + rw + "+Z" + rw;//"=G" + rw + "+K" + rw + "+O" + rw + "+S" + rw + "+W" + rw;

                //13-sept -2019
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;

            }


            #region Final sub-total

            if (mainsector.Trim() != "")
            {
                if (mainsector != "")
                {
                    rw++;
                    rw_end = rw - 1;
                    sbSub.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        rngtemp = ws.Cells["A" + rw_end + ":F" + rw_end];
                        rngtemp.Value = "Sub Total - " + mainsector;
                        rngtemp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        rngtemp.Merge = true;
                        // rngtemp.Style.Font.Bold = true;

                        //Region merge 13-sep-2019
                        ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["W" + rw_end].Formula = string.Format("=SUM(W{0}:W{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["X" + rw_end].Formula = string.Format("=SUM(X{0}:X{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["Y" + rw_end].Formula = string.Format("=SUM(Y{0}:Y{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["Z" + rw_end].Formula = string.Format("=SUM(Z{0}:Z{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["AA" + rw_end].Formula = string.Format("=SUM(AA{0}:AA{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["AB" + rw_end].Formula = string.Format("=SUM(AB{0}:AB{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["AC" + rw_end].Formula = string.Format("=SUM(AC{0}:AC{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["AD" + rw_end].Formula = string.Format("=SUM(AD{0}:AD{1})", rw_bgn_sub, (rw_end - 1));

                        if (country_start > 0)
                        {
                            country_end = rw - 2;
                            ws.Cells[country_start, 2, country_end, 2].Merge = true;
                            ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            country_start = 0;
                        }
                        if (GeoLoc_start > 0)
                        {
                            GeoLoc_end = rw - 2;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            GeoLoc_start = 0;
                        }
                    }

                }


            }


            #endregion

            rw_end = rw - 1;

            #region foooter
            rng = ws.Cells["A" + rw + ":AD" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            string row_nos = "";
            string[] sub_rows = sbSub.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows)
            {
                if (ln != "")
                    row_nos += "+G" + ln;
            }

            //Total 13-sep-2019 rgs
            rngtemp = ws.Cells["A" + rw + ":F" + rw];
            rngtemp.Value = "GRAND TOTAL";
            rngtemp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rngtemp.Merge = true;

            ws.Cells["A" + rw + ":AD" + rw].Style.Font.Bold = true;
            ws.Cells["A" + rw + ":AD" + rw].Style.Font.Size = 9;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // ws.Cells["C" + rw].Value = "TOTAL";

            ws.Cells["G" + rw].Formula = "=" + row_nos;
            ws.Cells["H" + rw].Formula = "=" + row_nos.Replace("G", "H");
            ws.Cells["I" + rw].Formula = "=" + row_nos.Replace("G", "I");
            ws.Cells["J" + rw].Formula = "=" + row_nos.Replace("G", "J");
            ws.Cells["K" + rw].Formula = "=" + row_nos.Replace("G", "K");
            ws.Cells["L" + rw].Formula = "=" + row_nos.Replace("G", "L");
            ws.Cells["M" + rw].Formula = "=" + row_nos.Replace("G", "M");
            ws.Cells["N" + rw].Formula = "=" + row_nos.Replace("G", "N");
            ws.Cells["O" + rw].Formula = "=" + row_nos.Replace("G", "O");
            ws.Cells["P" + rw].Formula = "=" + row_nos.Replace("G", "P");
            ws.Cells["Q" + rw].Formula = "=" + row_nos.Replace("G", "Q");
            ws.Cells["R" + rw].Formula = "=" + row_nos.Replace("G", "R");
            ws.Cells["S" + rw].Formula = "=" + row_nos.Replace("G", "S");
            ws.Cells["T" + rw].Formula = "=" + row_nos.Replace("G", "T");
            ws.Cells["U" + rw].Formula = "=" + row_nos.Replace("G", "U");
            ws.Cells["V" + rw].Formula = "=" + row_nos.Replace("G", "V");
            ws.Cells["W" + rw].Formula = "=" + row_nos.Replace("G", "W");
            ws.Cells["X" + rw].Formula = "=" + row_nos.Replace("G", "X");
            ws.Cells["Y" + rw].Formula = "=" + row_nos.Replace("G", "Y");
            ws.Cells["Z" + rw].Formula = "=" + row_nos.Replace("G", "Z");
            ws.Cells["AA" + rw].Formula = "=" + row_nos.Replace("G", "AA");
            ws.Cells["AB" + rw].Formula = "=" + row_nos.Replace("G", "AB");
            ws.Cells["AC" + rw].Formula = "=" + row_nos.Replace("G", "AC");
            ws.Cells["AD" + rw].Formula = "=" + row_nos.Replace("G", "AD");
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            #endregion

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleDry);
            rng = ws.Cells["G" + (rw_bgn - 2) + ":I" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["G" + (rw_bgn - 1) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFDF80");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":N" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["O" + (rw_bgn - 2) + ":R" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#CCFFFF");
            rng = ws.Cells["S" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9999");
            rng = ws.Cells["W" + (rw_bgn - 2) + ":Z" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#CCFF99");
            rng = ws.Cells["AA" + (rw_bgn - 2) + ":AD" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A"+rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Flat Rack Cntr Type
            rw++;
            #region header
            rng = ws.Cells["A" + rw + ":V" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "FLAT RACK";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20FR";
            ws.Cells["F" + rw].Value = "40FR";
            ws.Cells["G" + rw].Value = "HQFR";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20FR";
            ws.Cells["I" + rw].Value = "40FR";
            ws.Cells["J" + rw].Value = "HQFR";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20FR";
            ws.Cells["L" + rw].Value = "40FR";
            ws.Cells["M" + rw].Value = "HQFR";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20FR";
            ws.Cells["O" + rw].Value = "40FR";
            ws.Cells["P" + rw].Value = "HQFR";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20FR";
            ws.Cells["R" + rw].Value = "40FR";
            ws.Cells["S" + rw].Value = "HQFR";

            //TOTAL
            ws.Cells["T" + rw].Value = "20FR";
            ws.Cells["U" + rw].Value = "40FR";
            ws.Cells["V" + rw].Value = "HQFR";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            #endregion

            rw_bgn = rw;
            rw_end = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();
            int rw_bgn_sub1 = rw;
            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " FlRack20L7 > 0 OR FlRack40L7 > 0 OR FlRackHQL7 > 0 OR " +
                " FlRack208to14 > 0 OR FlRack408to14 > 0 OR FlRackHQ8to14 > 0 OR " +
                " FlRack2015to29 > 0 OR FlRack4015to29 > 0 OR FlRackHQ15to29 > 0 OR " +
                " FlRack2030to59 > 0 OR FlRack4030to59 > 0 OR FlRackHQ30to59 > 0 OR " +
                " FlRack20G59 > 0 OR FlRack40G59 > 0 OR FlRackHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;
            rw++;

            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                // rng = ws.Cells["D" + rw + ":AA" + rw];
                //rng.Style.Font.Bold = true;
                // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (mainsector.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {
                    if (mainsector != "" && i > 0)
                    {
                        rw++;
                        rw_end = rw - 1;

                        sbSub1.AppendLine(rw_end.ToString());
                        if (dtNew.Rows.Count > 0)
                        {

                            //#region REGION wise sub total
                            ExcelRange range1 = ws.Cells["A" + rw_end + ":D" + rw_end];
                            //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                            range1.Value = "Sub Total - " + mainsector;
                            //  range1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range1.Merge = true;
                            //range1.Style.Font.Bold = true;
                            //range1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //range1.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                            //Region merge 13-sep-2019 rgs
                            ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //Subtotal 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            //End Subtotal

                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.BackgroundColor.SetColor(Constants.ReportSubtotalBgColor);
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Bold = true;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Color.SetColor(Constants.ReportSubtotalFontColor);

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(G{0}:H{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub1, (rw_end - 1));

                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }

                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsector = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsector;
                    rw_bgn_sub1 = rw;
                }

                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                // For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;
                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQG59"].ToString());

                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;
            }

            #region Final sub-total  flatrack

            if (mainsector.Trim() != "") //nothing
            {
                if (mainsector != "")//nothing
                {
                    rw++;
                    rw_end = rw - 1;
                    sbSub1.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        ExcelRange range = ws.Cells["A" + rw_end + ":D" + rw_end];
                        range.Value = "Sub Total - " + mainsector;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Merge = true;
                        //  range.Style.Font.Bold = true;

                        //Region merge 13-sep-2019 rgs
                        ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin 13-sept-2019 rgs
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //End Subtotal

                        ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub1, (rw_end - 1));

                        if (country_start > 0)
                        {
                            country_end = rw - 2;
                            //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                            ws.Cells[country_start, 2, country_end, 2].Merge = true;
                            ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            country_start = 0;
                        }

                        if (GeoLoc_start > 0)
                        {
                            GeoLoc_end = rw - 2;
                            //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            GeoLoc_start = 0;
                        }
                    }
                }

                //mainsector = dtNew.Rows[i]["mainsector"].ToString();
                //ws.Cells["A" + rw].Value = mainsector;
                //ws.Cells["A" + rw].Merge = true;
                //rw_bgn_sub = rw;
            }

            #endregion


            #region foooter total flatrack
            // rng = ws.Cells["A" + rw + ":AA" + rw];
            // rng.Style.Font.Bold = true;
            // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            string row_nos1 = "";
            string[] sub_rows1 = sbSub1.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows1)
            {
                if (ln != "")
                    row_nos1 += "+E" + ln;
            }
            if (dtNew.Rows.Count > 0)
            {
                // ws.Cells["C" + rw].Value = "TOTAL";
                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=" + row_nos1;
                ws.Cells["F" + rw].Formula = "=" + row_nos1.Replace("E", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos1.Replace("E", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos1.Replace("E", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos1.Replace("E", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos1.Replace("E", "J");
                ws.Cells["K" + rw].Formula = "=" + row_nos1.Replace("E", "K");
                ws.Cells["L" + rw].Formula = "=" + row_nos1.Replace("E", "L");
                ws.Cells["M" + rw].Formula = "=" + row_nos1.Replace("E", "M");
                ws.Cells["N" + rw].Formula = "=" + row_nos1.Replace("E", "N");
                ws.Cells["O" + rw].Formula = "=" + row_nos1.Replace("E", "O");
                ws.Cells["P" + rw].Formula = "=" + row_nos1.Replace("E", "P");
                ws.Cells["Q" + rw].Formula = "=" + row_nos1.Replace("E", "Q");
                ws.Cells["R" + rw].Formula = "=" + row_nos1.Replace("E", "R");
                ws.Cells["S" + rw].Formula = "=" + row_nos1.Replace("E", "S");
                ws.Cells["T" + rw].Formula = "=" + row_nos1.Replace("E", "T");
                ws.Cells["U" + rw].Formula = "=" + row_nos1.Replace("E", "U");
                ws.Cells["V" + rw].Formula = "=" + row_nos1.Replace("E", "V");
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            #endregion

            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end-1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end - 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end - 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end - 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleFlatRack);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999 ");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Open Top Cntr Type

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "OPEN TOP";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":V" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20'OT";
            ws.Cells["F" + rw].Value = "40'OT";
            ws.Cells["G" + rw].Value = "HQ'OT";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20'OT";
            ws.Cells["I" + rw].Value = "40'OT";
            ws.Cells["J" + rw].Value = "HQ'OT";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20'OT";
            ws.Cells["L" + rw].Value = "40'OT";
            ws.Cells["M" + rw].Value = "HQ'OT";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20'OT";
            ws.Cells["O" + rw].Value = "40'OT";
            ws.Cells["P" + rw].Value = "HQ'OT";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20'OT";
            ws.Cells["R" + rw].Value = "40'OT";
            ws.Cells["S" + rw].Value = "HQ'OT";

            //TOTAL
            ws.Cells["T" + rw].Value = "20'OT";
            ws.Cells["U" + rw].Value = "40'OT";
            ws.Cells["V" + rw].Value = "HQ'OT";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;
            int rw_bgn_sub2 = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " OpTop20L7 > 0 OR OpTop40L7 > 0 OR OpTopHQL7 > 0 OR " +
                " OpTop208to14 > 0 OR OpTop408to14 > 0 OR OpTopHQ8to14 > 0 OR " +
                " OpTop2015to29 > 0 OR OpTop4015to29 > 0 OR OpTopHQ15to29 > 0 OR " +
                " OpTop2030to59 > 0 OR OpTop4030to59 > 0 OR OpTopHQ30to59 > 0 OR " +
                " OpTop20G59 > 0 OR OpTop40G59 > 0 OR OpTopHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;

            rw++;
            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                rng = ws.Cells["A" + rw + ":AB" + rw];
                //rng.Style.Font.Bold = true;
                //  rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (mainsectorOt.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {
                    if (mainsectorOt != "" && i > 0)
                    {
                        rw++;
                        rw_end = rw - 1;

                        sbSub2.AppendLine(rw_end.ToString());
                        //#region REGION wise sub total
                        if (dtNew.Rows.Count > 0)
                        {
                            ExcelRange range2 = ws.Cells["A" + rw_end + ":D" + rw_end];
                            range2.Value = "Sub Total - " + mainsectorOt;
                            range2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range2.Merge = true;
                            // range2.Style.Font.Bold = true;

                            //Region merge 13-sep-2019 rgs
                            ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.BackgroundColor.SetColor(Constants.ReportSubtotalBgColor);
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Bold = true;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Color.SetColor(Constants.ReportSubtotalFontColor);

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            //End Subtotal

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub2, (rw_end - 1));

                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }

                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsectorOt = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsectorOt;
                    rw_bgn_sub2 = rw;
                }



                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                //For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQG59"].ToString());

                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;
            }

            #region Final sub-total open top

            if (mainsector.Trim() != "") //nothing
            {
                if (mainsector != "")//nothing
                {

                    rw++;
                    rw_end = rw - 1;
                    sbSub2.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        ExcelRange range = ws.Cells["A" + rw_end + ":D" + rw_end];
                        //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                        range.Value = "Sub Total - " + mainsector;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Merge = true;
                        range.Style.Font.Bold = true;

                        //Region merge 13-sep-2019 rgs
                        ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin 13-sept-2019 rgs
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //End Subtotal

                        ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub2, (rw_end - 1));
                    }

                }

                //mainsector = dtNew.Rows[i]["mainsector"].ToString();
                //ws.Cells["A" + rw].Value = mainsector;
                //ws.Cells["A" + rw].Merge = true;
                //rw_bgn_sub = rw;
            }

            #endregion

            #region foooter total open top
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            string row_nos2 = "";
            string[] sub_rows2 = sbSub2.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows2)
            {
                if (ln != "")
                    row_nos2 += "+E" + ln;
            }
            if (dtNew.Rows.Count > 0)
            {
                //ws.Cells["C" + rw].Value = "TOTAL";
                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=" + row_nos2;
                ws.Cells["F" + rw].Formula = "=" + row_nos2.Replace("E", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos2.Replace("E", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos2.Replace("E", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos2.Replace("E", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos2.Replace("E", "J");
                ws.Cells["K" + rw].Formula = "=" + row_nos2.Replace("E", "K");
                ws.Cells["L" + rw].Formula = "=" + row_nos2.Replace("E", "L");
                ws.Cells["M" + rw].Formula = "=" + row_nos2.Replace("E", "M");
                ws.Cells["N" + rw].Formula = "=" + row_nos2.Replace("E", "N");
                ws.Cells["O" + rw].Formula = "=" + row_nos2.Replace("E", "O");
                ws.Cells["P" + rw].Formula = "=" + row_nos2.Replace("E", "P");
                ws.Cells["Q" + rw].Formula = "=" + row_nos2.Replace("E", "Q");
                ws.Cells["R" + rw].Formula = "=" + row_nos2.Replace("E", "R");
                ws.Cells["S" + rw].Formula = "=" + row_nos2.Replace("E", "S");
                ws.Cells["T" + rw].Formula = "=" + row_nos2.Replace("E", "T");
                ws.Cells["U" + rw].Formula = "=" + row_nos2.Replace("E", "U");
                ws.Cells["V" + rw].Formula = "=" + row_nos2.Replace("E", "V");
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            #endregion total open top total open top


            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleOpenTop);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Reefer Type

            rw++;
            rng = ws.Cells["E" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "REEFER";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["E" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20'RF";
            ws.Cells["F" + rw].Value = "40'RF";
            ws.Cells["G" + rw].Value = "HQ'RF";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20'RF";
            ws.Cells["I" + rw].Value = "40'RF";
            ws.Cells["J" + rw].Value = "HQ'RF";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20'RF";
            ws.Cells["L" + rw].Value = "40'RF";
            ws.Cells["M" + rw].Value = "HQ'RF";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20'RF";
            ws.Cells["O" + rw].Value = "40'RF";
            ws.Cells["P" + rw].Value = "HQ'RF";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20'RF";
            ws.Cells["R" + rw].Value = "40'RF";
            ws.Cells["S" + rw].Value = "HQ'RF";

            //TOTAL
            ws.Cells["T" + rw].Value = "20'RF";
            ws.Cells["U" + rw].Value = "40'RF";
            ws.Cells["V" + rw].Value = "HQ'RF";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;
            int rw_bgn_sub3 = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " RF20L7 > 0 OR RF40L7 > 0 OR RFHQL7 > 0 OR " +
                " RF208to14 > 0 OR RF408to14 > 0 OR RFHQ8to14 > 0 OR " +
                " RF2015to29 > 0 OR RF4015to29 > 0 OR RFHQ15to29 > 0 OR " +
                " RF2030to59 > 0 OR RF4030to59 > 0 OR RFHQ30to59 > 0 OR " +
                " RF20G59 > 0 OR RF40G59 > 0 OR RFHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;
            rw++;

            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                rng = ws.Cells["A" + rw + ":AB" + rw];
                //rng.Style.Font.Bold = true;
                // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (mainsectorrf.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {
                    if (mainsectorrf != "" && i > 0)
                    {
                        rw++;
                        rw_end = rw - 1;

                        sbSub3.AppendLine(rw_end.ToString());
                        //#region REGION wise sub total
                        if (dtNew.Rows.Count > 0)
                        {
                            ExcelRange range2 = ws.Cells["A" + rw_end + ":D" + rw_end];
                            //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                            range2.Value = "Sub Total - " + mainsectorrf;
                            range2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range2.Merge = true;
                            //  range2.Style.Font.Bold = true;

                            //range2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //range2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.BackgroundColor.SetColor(Constants.ReportSubtotalBgColor);
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Bold = true;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Color.SetColor(Constants.ReportSubtotalFontColor);

                            //Region merge 13-sep-2019 rgs
                            ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            //End Subtotal

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub3, (rw_end - 1));

                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }

                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsectorrf = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsectorrf;
                    rw_bgn_sub3 = rw;
                }


                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                //For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQG59"].ToString());

                //ws.Cells["S" + rw].Formula = "=C" + rw + "+F" + rw + "+I" + rw + "+L" + rw + "+O" + rw;
                //ws.Cells["T" + rw].Formula = "=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                //ws.Cells["U" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;

                //13-Sept-2019 RGS
                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;
            }

            #region Final sub-total reefer

            if (mainsectorrf.Trim() != "") //nothing
            {
                if (mainsectorrf != "")//nothing
                {
                    rw++;
                    rw_end = rw - 1;
                    sbSub3.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        ExcelRange range = ws.Cells["A" + rw_end + ":D" + rw_end];
                        //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                        range.Value = "Sub Total - " + mainsectorrf;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Merge = true;
                        //range.Style.Font.Bold = true;

                        //Region merge 13-sep-2019 rgs
                        ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin 13-sept-2019 rgs
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //End Subtotal

                        ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub2, (rw_end - 1));
                    }

                }

                //mainsector = dtNew.Rows[i]["mainsector"].ToString();
                //ws.Cells["A" + rw].Value = mainsector;
                //ws.Cells["A" + rw].Merge = true;
                //rw_bgn_sub = rw;
            }

            #endregion

            #region foooter total reefer
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            string row_nos3 = "";
            string[] sub_rows3 = sbSub3.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows3)
            {
                if (ln != "")
                    row_nos3 += "+E" + ln;
            }

            if (dtNew.Rows.Count > 0)
            {
                // ws.Cells["C" + rw].Value = "TOTAL";

                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=" + row_nos3;
                ws.Cells["F" + rw].Formula = "=" + row_nos3.Replace("E", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos3.Replace("E", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos3.Replace("E", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos3.Replace("E", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos3.Replace("E", "J");
                ws.Cells["K" + rw].Formula = "=" + row_nos3.Replace("E", "K");
                ws.Cells["L" + rw].Formula = "=" + row_nos3.Replace("E", "L");
                ws.Cells["M" + rw].Formula = "=" + row_nos3.Replace("E", "M");
                ws.Cells["N" + rw].Formula = "=" + row_nos3.Replace("E", "N");
                ws.Cells["O" + rw].Formula = "=" + row_nos3.Replace("E", "O");
                ws.Cells["P" + rw].Formula = "=" + row_nos3.Replace("E", "P");
                ws.Cells["Q" + rw].Formula = "=" + row_nos3.Replace("E", "Q");
                ws.Cells["R" + rw].Formula = "=" + row_nos3.Replace("E", "R");
                ws.Cells["S" + rw].Formula = "=" + row_nos3.Replace("E", "S");
                ws.Cells["T" + rw].Formula = "=" + row_nos3.Replace("E", "T");
                ws.Cells["U" + rw].Formula = "=" + row_nos3.Replace("E", "U");
                ws.Cells["V" + rw].Formula = "=" + row_nos3.Replace("E", "V");
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            #endregion total open top total open top

            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleReefer);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Tank Cntr Type

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TANK";
            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20'TK";
            ws.Cells["F" + rw].Value = "40'TK";
            ws.Cells["G" + rw].Value = "HQ'TK";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20'TK";
            ws.Cells["I" + rw].Value = "40'TK";
            ws.Cells["J" + rw].Value = "HQ'TK";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20'TK";
            ws.Cells["L" + rw].Value = "40'TK";
            ws.Cells["M" + rw].Value = "HQ'TK";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20'TK";
            ws.Cells["O" + rw].Value = "40'TK";
            ws.Cells["P" + rw].Value = "HQ'TK";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20'TK";
            ws.Cells["R" + rw].Value = "40'TK";
            ws.Cells["S" + rw].Value = "HQ'TK";

            //TOTAL
            ws.Cells["T" + rw].Value = "20'TK";
            ws.Cells["U" + rw].Value = "40'TK";
            ws.Cells["V" + rw].Value = "HQ'TK";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " Tank20L7 > 0 OR Tank40L7 > 0 OR TankHQL7 > 0 OR " +
                " Tank208to14 > 0 OR Tank408to14 > 0 OR TankHQ8to14 > 0 OR " +
                " Tank2015to29 > 0 OR Tank4015to29 > 0 OR TankHQ15to29 > 0 OR " +
                " Tank2030to59 > 0 OR Tank4030to59 > 0 OR TankHQ30to59 > 0 OR " +
                " Tank20G59 > 0 OR Tank40G59 > 0 OR TankHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;
            rw++;

            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                // rng = ws.Cells["A" + rw + ":Z" + rw];
                //rng.Style.Font.Bold = true;
                // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                //For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["B" + rw].Value;
                val_prev_GeoLoc = ws.Cells["B" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQG59"].ToString());

                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

                rw++;
            }

            if (dtNew.Rows.Count > 0)
            {
                rw_end = rw - 1;

                //rng = ws.Cells["A" + rw + ":Z" + rw];
                //rng.Style.Font.Bold = true;
                //rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //ws.Cells["C" + rw].Value = "TOTAL";
                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=SUM(E" + rw_bgn + ":E" + rw_end + ")";
                ws.Cells["F" + rw].Formula = "=SUM(F" + rw_bgn + ":F" + rw_end + ")";
                ws.Cells["G" + rw].Formula = "=SUM(G" + rw_bgn + ":G" + rw_end + ")";
                ws.Cells["H" + rw].Formula = "=SUM(H" + rw_bgn + ":H" + rw_end + ")";
                ws.Cells["I" + rw].Formula = "=SUM(I" + rw_bgn + ":I" + rw_end + ")";
                ws.Cells["J" + rw].Formula = "=SUM(J" + rw_bgn + ":J" + rw_end + ")";
                ws.Cells["K" + rw].Formula = "=SUM(K" + rw_bgn + ":K" + rw_end + ")";
                ws.Cells["L" + rw].Formula = "=SUM(L" + rw_bgn + ":L" + rw_end + ")";
                ws.Cells["M" + rw].Formula = "=SUM(M" + rw_bgn + ":M" + rw_end + ")";
                ws.Cells["N" + rw].Formula = "=SUM(N" + rw_bgn + ":N" + rw_end + ")";
                ws.Cells["O" + rw].Formula = "=SUM(O" + rw_bgn + ":O" + rw_end + ")";
                ws.Cells["P" + rw].Formula = "=SUM(P" + rw_bgn + ":P" + rw_end + ")";
                ws.Cells["Q" + rw].Formula = "=SUM(Q" + rw_bgn + ":Q" + rw_end + ")";
                ws.Cells["R" + rw].Formula = "=SUM(R" + rw_bgn + ":R" + rw_end + ")";
                ws.Cells["S" + rw].Formula = "=SUM(S" + rw_bgn + ":S" + rw_end + ")";
                ws.Cells["T" + rw].Formula = "=SUM(T" + rw_bgn + ":T" + rw_end + ")";
                ws.Cells["U" + rw].Formula = "=SUM(U" + rw_bgn + ":U" + rw_end + ")";
                ws.Cells["V" + rw].Formula = "=SUM(V" + rw_bgn + ":V" + rw_end + ")";
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleTank);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion


            ws.Cells[7, 7, ws.Dimension.End.Row, ws.Dimension.End.Column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //  ws.Cells[ws.Dimension.Address].AutoFitColumns();

            #region Column Width to 5

            //ws.Column(1).AutoFit();
            ws.Column(1).Width = 15;
            ws.Column(1).Style.WrapText = true;
            // ws.Column(1).BestFit = true;
            ws.Column(2).AutoFit();
            ws.Column(2).Style.WrapText = true;

            ws.Column(3).AutoFit();
            ws.Column(3).Style.WrapText = true;

            ws.Column(4).AutoFit();
            ws.Column(4).Style.WrapText = true;
            ws.Column(5).AutoFit();
            ws.Column(5).Style.WrapText = true;
            ws.Column(6).AutoFit();
            ws.Column(6).Style.WrapText = true;


            ws.Column(7).Width = 5.71;
            ws.Column(8).Width = 5.71;
            ws.Column(9).Width = 5.71;
            ws.Column(10).Width = 5.71;
            ws.Column(11).Width = 5.71;
            ws.Column(12).Width = 5.71;
            ws.Column(13).Width = 5.71;
            ws.Column(14).Width = 5.71;
            ws.Column(15).Width = 5.71;
            ws.Column(16).Width = 5.71;
            ws.Column(17).Width = 5.71;
            ws.Column(18).Width = 5.71;
            ws.Column(19).Width = 5.71;
            ws.Column(20).Width = 5.71;
            ws.Column(21).Width = 5.71;
            ws.Column(22).Width = 5.71;
            ws.Column(23).Width = 5.71;
            ws.Column(24).Width = 5.71;
            ws.Column(25).Width = 5.71;
            ws.Column(26).Width = 5.71;
            ws.Column(27).Width = 5.71;
            ws.Column(28).Width = 5.71;
            ws.Column(29).Width = 5.71;
            ws.Column(30).Width = 5.71;
            #endregion

            #endregion

            #region sheet 2  summary
           // ExcelPackage pck = new ExcelPackage();
             ws = pck.Workbook.Worksheets.Add("Summary");

        


            ws.Cells["A1"].Value = "EQC AGEWISE REPORT";
            r = ws.Cells["A1:AB1"];
            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            r.Merge = true;
            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            r.Style.Font.Color.SetColor(Color.Black);
            r.Style.Font.Bold = true;
            r.Style.Font.SetFromFont(new Font("Arial", 12));

            allCells = ws.Cells[2, 1, 500, 50];
            cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Arial", 8));

             rng = ws.Cells["A2:E2"];

            //ws.Cells["A3"].Value = "Line :";
            //ws.Cells["B3"].Value = "OCEANUS LINES";

            ws.Cells["A4"].Value = "User :";
            ws.Cells["B4"].Value = User;
            ws.Cells["A4"].Style.Font.Bold = true;
            ws.Cells["B4"].Style.Font.Bold = true;

            ws.Cells["A5"].Value = "Date :";
            ws.Cells["B5"].Value = System.DateTime.Today.Date.ToShortDateString();
            ws.Cells["A5"].Style.Font.Bold = true;
            ws.Cells["B5"].Style.Font.Bold = true;

            ws.Cells["A6"].Value = "Status :";
            ws.Cells["B6"].Value = cntrStatus;
            ws.Cells["A6"].Style.Font.Bold = true;
            ws.Cells["B6"].Style.Font.Bold = true;
            // e_msg += "<br>" + cntrStatus;

            if (Session["UserName"] == null)
                Session["UserName1"] = "SYS GEN";
            else
                Session["UserName1"] = Session["UserName"];

            rw = 8;

             Cntrs = cntrStatus.Split(',');
             _dtValue = null;
            //string country1 = "", GeoLocation = "";
            //string mainsector = ""; string mainsectorfr = ""; string mainsectorOt = ""; string mainsectortnk = ""; string mainsectorrf = "";
            //int country_start = 0, country_end = 0, GeoLoc_start = 0, GeoLoc_end = 0;
            //int mainsector_start = 0, mainsector_end = 0;

            for (int y = 0; y < Cntrs.Length; y++)
            {
                _dtValue = _dtAgewiseRepValueSummary(cntrStatus, Country, GeoLoc, Grade, LeaseTerm, CntrType, criteria);
            }


            #region Dry Cntr Type
            rw = 8;
            rng = ws.Cells["E" + rw + ":AB" + rw];
            rng.Merge = true;
            rng.Style.Font.Bold = true;
            rng.Value = "DRY";
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw = 9;
            rng = ws.Cells["A" + rw + ":AD" + rw];
            rng.Style.Font.Bold = true;


            rng = ws.Cells["G" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";

            rng = ws.Cells["K" + rw + ":N" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";

            rng = ws.Cells["O" + rw + ":R" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";

            rng = ws.Cells["S" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";

            rng = ws.Cells["W" + rw + ":Z" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";

            rng = ws.Cells["AA" + rw + ":AD" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;

            rng = ws.Cells["A" + rw + ":AD" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw + ":E" + rw].Value = "Location";
            ws.Cells["D" + rw + ":E" + rw].Merge = true;
            ws.Cells["D" + rw + ":E" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            // ws.Cells["E" + rw].Value = "";
            ws.Cells["F" + rw].Value = "StatusCode";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["G" + rw].Value = "20GP";
            ws.Cells["H" + rw].Value = "20HQ";
            ws.Cells["I" + rw].Value = "40GP";
            ws.Cells["J" + rw].Value = "40HQ";

            //8 - 14 Days
            ws.Cells["K" + rw].Value = "20GP";
            ws.Cells["L" + rw].Value = "20HQ";
            ws.Cells["M" + rw].Value = "40GP";
            ws.Cells["N" + rw].Value = "40HQ";

            //15 - 29 Days
            ws.Cells["O" + rw].Value = "20GP";
            ws.Cells["P" + rw].Value = "20HQ";
            ws.Cells["Q" + rw].Value = "40GP";
            ws.Cells["R" + rw].Value = "40HQ";

            //>30 - 59 Days
            ws.Cells["S" + rw].Value = "20GP";
            ws.Cells["T" + rw].Value = "20HQ";
            ws.Cells["U" + rw].Value = "40GP";
            ws.Cells["V" + rw].Value = "40HQ";

            //> 59 Days
            ws.Cells["W" + rw].Value = "20GP";
            ws.Cells["X" + rw].Value = "20HQ";
            ws.Cells["Y" + rw].Value = "40GP";
            ws.Cells["Z" + rw].Value = "40HQ";

            //TOTAL
            ws.Cells["AA" + rw].Value = "20GP";
            ws.Cells["AB" + rw].Value = "20HQ";
            ws.Cells["AC" + rw].Value = "40GP";
            ws.Cells["AD" + rw].Value = "40HQ";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;


            dt = _dtValue;
            dv1 = new DataView(dt);

            dtNew = dv1.ToTable();

            rw++;
            //int flag1 = 0;
            //object val_this = null, val_prev = null, val_this_GeoLoc = null, val_prev_GeoLoc = null;
            //object val_this1 = null, val_prev1 = null;

             dtrows = dtNew.Rows.Count - 1;
             rw_bgn_sub = rw;
             sbSub = new StringBuilder();
             sbSub1 = new StringBuilder();
             sbSub2 = new StringBuilder();
             sbSub3 = new StringBuilder();
             sbSub4 = new StringBuilder();
            //ExcelRange rngtemp;
            // main sector
            for (int i = 0; i < dtNew.Rows.Count; i++)
            {


                if (mainsector.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {

                    if (mainsector != "" && i > 0)
                    {
                        if (dtNew.Rows.Count > 0)
                        {
                            rw++;
                            rw_end = rw - 1;
                            sbSub.AppendLine(rw_end.ToString());

                            rngtemp = ws.Cells["A" + rw_end + ":D" + rw_end];
                            rngtemp.Value = "Sub Total - " + mainsector;
                            rngtemp.Merge = true;
                            rngtemp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // rngtemp.Style.Font.Bold = true;

                            //Region merge 13-sep-2019  rgs
                            ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["W" + rw_end].Formula = string.Format("=SUM(W{0}:W{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["X" + rw_end].Formula = string.Format("=SUM(X{0}:X{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["Y" + rw_end].Formula = string.Format("=SUM(Y{0}:Y{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["Z" + rw_end].Formula = string.Format("=SUM(Z{0}:Z{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AA" + rw_end].Formula = string.Format("=SUM(AA{0}:AA{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AB" + rw_end].Formula = string.Format("=SUM(AB{0}:AB{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AC" + rw_end].Formula = string.Format("=SUM(AC{0}:AC{1})", rw_bgn_sub, (rw_end - 1));
                            ws.Cells["AD" + rw_end].Formula = string.Format("=SUM(AD{0}:AD{1})", rw_bgn_sub, (rw_end - 1));
                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }
                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsector = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsector;
                    rw_bgn_sub = rw;

                }
                country1 = dtNew.Rows[i]["CountryName"].ToString();
                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    // country_end = rw - 2;
                    country_end = rw - 1;
                    ws.Cells[country_start, 2, country_end, 2].Merge = true;
                    ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 2, country_end, 2].Merge = true;
                    ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                // For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {

                    GeoLoc_end = rw - 1;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                    ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
             
                ws.Cells["D" + rw + ":E" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw + ":E" + rw].Merge = true;
                ws.Cells["D" + rw + ":E" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                //ws.Cells["E" + rw].Value = dtNew.Rows[i]["LeaseTerm"].ToString();
                ws.Cells["F" + rw].Value = dtNew.Rows[i]["StatusCode"].ToString();
                ws.Cells["F" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry20L7"].ToString());
                ws.Cells["H" + rw].Value = TryParseNullable("");
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry40L7"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQL7"].ToString());



                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry208to14"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable("");
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry408to14"].ToString());
                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQ8to14"].ToString());

                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry2015to29"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable("");
                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry4015to29"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQ15to29"].ToString());

                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry2030to59"].ToString());
                ws.Cells["T" + rw].Value = TryParseNullable("");
                ws.Cells["U" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry4030to59"].ToString());
                ws.Cells["V" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQ30to59"].ToString());

                ws.Cells["W" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry20G59"].ToString());
                ws.Cells["X" + rw].Value = TryParseNullable("");
                ws.Cells["Y" + rw].Value = TryParseNullable(dtNew.Rows[i]["Dry40G59"].ToString());
                ws.Cells["Z" + rw].Value = TryParseNullable(dtNew.Rows[i]["DryHQG59"].ToString());

                ws.Cells["AA" + rw].Formula = "=G" + rw + "+K" + rw + "+O" + rw + "+S" + rw + "+W" + rw;//"=D" + rw + "+H" + rw + "+L" + rw + "+P" + rw + "+T" + rw;
                ws.Cells["AB" + rw].Formula = "=H" + rw + "+L" + rw + "+P" + rw + "+T" + rw + "+X" + rw;//"=E" + rw + "+I" + rw + "+M" + rw + "+R" + rw + "+U" + rw;
                ws.Cells["AC" + rw].Formula = "=I" + rw + "+M" + rw + "+Q" + rw + "+U" + rw + "+Y" + rw;//"=F" + rw + "+J" + rw + "+N" + rw + "+R" + rw + "+V" + rw;
                ws.Cells["AD" + rw].Formula = "=J" + rw + "+N" + rw + "+R" + rw + "+V" + rw + "+Z" + rw;//"=G" + rw + "+K" + rw + "+O" + rw + "+S" + rw + "+W" + rw;

                //13-sept -2019
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;

            }


            #region Final sub-total

            if (mainsector.Trim() != "")
            {
                if (mainsector != "")
                {
                    rw++;
                    rw_end = rw - 1;
                    sbSub.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        rngtemp = ws.Cells["A" + rw_end + ":F" + rw_end];
                        rngtemp.Value = "Sub Total - " + mainsector;
                        rngtemp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        rngtemp.Merge = true;
                        // rngtemp.Style.Font.Bold = true;

                        //Region merge 13-sep-2019
                        ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":AD" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["W" + rw_end].Formula = string.Format("=SUM(W{0}:W{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["X" + rw_end].Formula = string.Format("=SUM(X{0}:X{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["Y" + rw_end].Formula = string.Format("=SUM(Y{0}:Y{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["Z" + rw_end].Formula = string.Format("=SUM(Z{0}:Z{1})", rw_bgn_sub, (rw_end - 1));

                        ws.Cells["AA" + rw_end].Formula = string.Format("=SUM(AA{0}:AA{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["AB" + rw_end].Formula = string.Format("=SUM(AB{0}:AB{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["AC" + rw_end].Formula = string.Format("=SUM(AC{0}:AC{1})", rw_bgn_sub, (rw_end - 1));
                        ws.Cells["AD" + rw_end].Formula = string.Format("=SUM(AD{0}:AD{1})", rw_bgn_sub, (rw_end - 1));

                        if (country_start > 0)
                        {
                            country_end = rw - 2;
                            ws.Cells[country_start, 2, country_end, 2].Merge = true;
                            ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            country_start = 0;
                        }
                        if (GeoLoc_start > 0)
                        {
                            GeoLoc_end = rw - 2;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            GeoLoc_start = 0;
                        }
                    }

                }


            }


            #endregion

            rw_end = rw - 1;

            #region foooter
            rng = ws.Cells["A" + rw + ":AD" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

             row_nos = "";
             sub_rows = sbSub.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows)
            {
                if (ln != "")
                    row_nos += "+G" + ln;
            }

            //Total 13-sep-2019 rgs
            rngtemp = ws.Cells["A" + rw + ":F" + rw];
            rngtemp.Value = "GRAND TOTAL";
            rngtemp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rngtemp.Merge = true;

            ws.Cells["A" + rw + ":AD" + rw].Style.Font.Bold = true;
            ws.Cells["A" + rw + ":AD" + rw].Style.Font.Size = 9;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":AD" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // ws.Cells["C" + rw].Value = "TOTAL";

            ws.Cells["G" + rw].Formula = "=" + row_nos;
            ws.Cells["H" + rw].Formula = "=" + row_nos.Replace("G", "H");
            ws.Cells["I" + rw].Formula = "=" + row_nos.Replace("G", "I");
            ws.Cells["J" + rw].Formula = "=" + row_nos.Replace("G", "J");
            ws.Cells["K" + rw].Formula = "=" + row_nos.Replace("G", "K");
            ws.Cells["L" + rw].Formula = "=" + row_nos.Replace("G", "L");
            ws.Cells["M" + rw].Formula = "=" + row_nos.Replace("G", "M");
            ws.Cells["N" + rw].Formula = "=" + row_nos.Replace("G", "N");
            ws.Cells["O" + rw].Formula = "=" + row_nos.Replace("G", "O");
            ws.Cells["P" + rw].Formula = "=" + row_nos.Replace("G", "P");
            ws.Cells["Q" + rw].Formula = "=" + row_nos.Replace("G", "Q");
            ws.Cells["R" + rw].Formula = "=" + row_nos.Replace("G", "R");
            ws.Cells["S" + rw].Formula = "=" + row_nos.Replace("G", "S");
            ws.Cells["T" + rw].Formula = "=" + row_nos.Replace("G", "T");
            ws.Cells["U" + rw].Formula = "=" + row_nos.Replace("G", "U");
            ws.Cells["V" + rw].Formula = "=" + row_nos.Replace("G", "V");
            ws.Cells["W" + rw].Formula = "=" + row_nos.Replace("G", "W");
            ws.Cells["X" + rw].Formula = "=" + row_nos.Replace("G", "X");
            ws.Cells["Y" + rw].Formula = "=" + row_nos.Replace("G", "Y");
            ws.Cells["Z" + rw].Formula = "=" + row_nos.Replace("G", "Z");
            ws.Cells["AA" + rw].Formula = "=" + row_nos.Replace("G", "AA");
            ws.Cells["AB" + rw].Formula = "=" + row_nos.Replace("G", "AB");
            ws.Cells["AC" + rw].Formula = "=" + row_nos.Replace("G", "AC");
            ws.Cells["AD" + rw].Formula = "=" + row_nos.Replace("G", "AD");
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":AA" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            #endregion

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleDry);
            rng = ws.Cells["G" + (rw_bgn - 2) + ":I" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["G" + (rw_bgn - 1) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFDF80");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":N" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["O" + (rw_bgn - 2) + ":R" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#CCFFFF");
            rng = ws.Cells["S" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9999");
            rng = ws.Cells["W" + (rw_bgn - 2) + ":Z" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#CCFF99");
            rng = ws.Cells["AA" + (rw_bgn - 2) + ":AD" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A"+rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Flat Rack Cntr Type
            rw++;
            #region header
            rng = ws.Cells["A" + rw + ":V" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "FLAT RACK";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20FR";
            ws.Cells["F" + rw].Value = "40FR";
            ws.Cells["G" + rw].Value = "HQFR";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20FR";
            ws.Cells["I" + rw].Value = "40FR";
            ws.Cells["J" + rw].Value = "HQFR";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20FR";
            ws.Cells["L" + rw].Value = "40FR";
            ws.Cells["M" + rw].Value = "HQFR";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20FR";
            ws.Cells["O" + rw].Value = "40FR";
            ws.Cells["P" + rw].Value = "HQFR";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20FR";
            ws.Cells["R" + rw].Value = "40FR";
            ws.Cells["S" + rw].Value = "HQFR";

            //TOTAL
            ws.Cells["T" + rw].Value = "20FR";
            ws.Cells["U" + rw].Value = "40FR";
            ws.Cells["V" + rw].Value = "HQFR";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            #endregion

            rw_bgn = rw;
            rw_end = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();
             rw_bgn_sub1 = rw;
            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " FlRack20L7 > 0 OR FlRack40L7 > 0 OR FlRackHQL7 > 0 OR " +
                " FlRack208to14 > 0 OR FlRack408to14 > 0 OR FlRackHQ8to14 > 0 OR " +
                " FlRack2015to29 > 0 OR FlRack4015to29 > 0 OR FlRackHQ15to29 > 0 OR " +
                " FlRack2030to59 > 0 OR FlRack4030to59 > 0 OR FlRackHQ30to59 > 0 OR " +
                " FlRack20G59 > 0 OR FlRack40G59 > 0 OR FlRackHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;
            rw++;

            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                // rng = ws.Cells["D" + rw + ":AA" + rw];
                //rng.Style.Font.Bold = true;
                // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (mainsector.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {
                    if (mainsector != "" && i > 0)
                    {
                        rw++;
                        rw_end = rw - 1;

                        sbSub1.AppendLine(rw_end.ToString());
                        if (dtNew.Rows.Count > 0)
                        {

                            //#region REGION wise sub total
                            ExcelRange range1 = ws.Cells["A" + rw_end + ":D" + rw_end];
                            //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                            range1.Value = "Sub Total - " + mainsector;
                            //  range1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range1.Merge = true;
                            //range1.Style.Font.Bold = true;
                            //range1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //range1.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                            //Region merge 13-sep-2019 rgs
                            ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //Subtotal 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            //End Subtotal

                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.BackgroundColor.SetColor(Constants.ReportSubtotalBgColor);
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Bold = true;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Color.SetColor(Constants.ReportSubtotalFontColor);

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(G{0}:H{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub1, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub1, (rw_end - 1));

                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }

                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsector = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsector;
                    rw_bgn_sub1 = rw;
                }

                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                // For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;
                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRack40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["FlRackHQG59"].ToString());

                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;
            }

            #region Final sub-total  flatrack

            if (mainsector.Trim() != "") //nothing
            {
                if (mainsector != "")//nothing
                {
                    rw++;
                    rw_end = rw - 1;
                    sbSub1.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        ExcelRange range = ws.Cells["A" + rw_end + ":D" + rw_end];
                        range.Value = "Sub Total - " + mainsector;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Merge = true;
                        //  range.Style.Font.Bold = true;

                        //Region merge 13-sep-2019 rgs
                        ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub1, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin 13-sept-2019 rgs
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //End Subtotal

                        ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub1, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub1, (rw_end - 1));

                        if (country_start > 0)
                        {
                            country_end = rw - 2;
                            //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                            ws.Cells[country_start, 2, country_end, 2].Merge = true;
                            ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            country_start = 0;
                        }

                        if (GeoLoc_start > 0)
                        {
                            GeoLoc_end = rw - 2;
                            //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                            ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            GeoLoc_start = 0;
                        }
                    }
                }

                //mainsector = dtNew.Rows[i]["mainsector"].ToString();
                //ws.Cells["A" + rw].Value = mainsector;
                //ws.Cells["A" + rw].Merge = true;
                //rw_bgn_sub = rw;
            }

            #endregion


            #region foooter total flatrack
            // rng = ws.Cells["A" + rw + ":AA" + rw];
            // rng.Style.Font.Bold = true;
            // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

             row_nos1 = "";
             sub_rows1 = sbSub1.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows1)
            {
                if (ln != "")
                    row_nos1 += "+E" + ln;
            }
            if (dtNew.Rows.Count > 0)
            {
                // ws.Cells["C" + rw].Value = "TOTAL";
                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=" + row_nos1;
                ws.Cells["F" + rw].Formula = "=" + row_nos1.Replace("E", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos1.Replace("E", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos1.Replace("E", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos1.Replace("E", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos1.Replace("E", "J");
                ws.Cells["K" + rw].Formula = "=" + row_nos1.Replace("E", "K");
                ws.Cells["L" + rw].Formula = "=" + row_nos1.Replace("E", "L");
                ws.Cells["M" + rw].Formula = "=" + row_nos1.Replace("E", "M");
                ws.Cells["N" + rw].Formula = "=" + row_nos1.Replace("E", "N");
                ws.Cells["O" + rw].Formula = "=" + row_nos1.Replace("E", "O");
                ws.Cells["P" + rw].Formula = "=" + row_nos1.Replace("E", "P");
                ws.Cells["Q" + rw].Formula = "=" + row_nos1.Replace("E", "Q");
                ws.Cells["R" + rw].Formula = "=" + row_nos1.Replace("E", "R");
                ws.Cells["S" + rw].Formula = "=" + row_nos1.Replace("E", "S");
                ws.Cells["T" + rw].Formula = "=" + row_nos1.Replace("E", "T");
                ws.Cells["U" + rw].Formula = "=" + row_nos1.Replace("E", "U");
                ws.Cells["V" + rw].Formula = "=" + row_nos1.Replace("E", "V");
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            #endregion

            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end-1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end - 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end - 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 3) + ":U" + (rw_end - 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleFlatRack);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999 ");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Open Top Cntr Type

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "OPEN TOP";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":V" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20'OT";
            ws.Cells["F" + rw].Value = "40'OT";
            ws.Cells["G" + rw].Value = "HQ'OT";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20'OT";
            ws.Cells["I" + rw].Value = "40'OT";
            ws.Cells["J" + rw].Value = "HQ'OT";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20'OT";
            ws.Cells["L" + rw].Value = "40'OT";
            ws.Cells["M" + rw].Value = "HQ'OT";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20'OT";
            ws.Cells["O" + rw].Value = "40'OT";
            ws.Cells["P" + rw].Value = "HQ'OT";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20'OT";
            ws.Cells["R" + rw].Value = "40'OT";
            ws.Cells["S" + rw].Value = "HQ'OT";

            //TOTAL
            ws.Cells["T" + rw].Value = "20'OT";
            ws.Cells["U" + rw].Value = "40'OT";
            ws.Cells["V" + rw].Value = "HQ'OT";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;
            rw_bgn_sub2 = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " OpTop20L7 > 0 OR OpTop40L7 > 0 OR OpTopHQL7 > 0 OR " +
                " OpTop208to14 > 0 OR OpTop408to14 > 0 OR OpTopHQ8to14 > 0 OR " +
                " OpTop2015to29 > 0 OR OpTop4015to29 > 0 OR OpTopHQ15to29 > 0 OR " +
                " OpTop2030to59 > 0 OR OpTop4030to59 > 0 OR OpTopHQ30to59 > 0 OR " +
                " OpTop20G59 > 0 OR OpTop40G59 > 0 OR OpTopHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;

            rw++;
            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                rng = ws.Cells["A" + rw + ":AB" + rw];
                //rng.Style.Font.Bold = true;
                //  rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (mainsectorOt.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {
                    if (mainsectorOt != "" && i > 0)
                    {
                        rw++;
                        rw_end = rw - 1;

                        sbSub2.AppendLine(rw_end.ToString());
                        //#region REGION wise sub total
                        if (dtNew.Rows.Count > 0)
                        {
                            ExcelRange range2 = ws.Cells["A" + rw_end + ":D" + rw_end];
                            range2.Value = "Sub Total - " + mainsectorOt;
                            range2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range2.Merge = true;
                            // range2.Style.Font.Bold = true;

                            //Region merge 13-sep-2019 rgs
                            ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.BackgroundColor.SetColor(Constants.ReportSubtotalBgColor);
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Bold = true;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Color.SetColor(Constants.ReportSubtotalFontColor);

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            //End Subtotal

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub2, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub2, (rw_end - 1));

                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }

                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsectorOt = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsectorOt;
                    rw_bgn_sub2 = rw;
                }



                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                //For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTop40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["OpTopHQG59"].ToString());

                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;
            }

            #region Final sub-total open top

            if (mainsector.Trim() != "") //nothing
            {
                if (mainsector != "")//nothing
                {

                    rw++;
                    rw_end = rw - 1;
                    sbSub2.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        ExcelRange range = ws.Cells["A" + rw_end + ":D" + rw_end];
                        //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                        range.Value = "Sub Total - " + mainsector;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Merge = true;
                        range.Style.Font.Bold = true;

                        //Region merge 13-sep-2019 rgs
                        ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub2, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin 13-sept-2019 rgs
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //End Subtotal

                        ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub2, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub2, (rw_end - 1));
                    }

                }

                //mainsector = dtNew.Rows[i]["mainsector"].ToString();
                //ws.Cells["A" + rw].Value = mainsector;
                //ws.Cells["A" + rw].Merge = true;
                //rw_bgn_sub = rw;
            }

            #endregion

            #region foooter total open top
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

             row_nos2 = "";
             sub_rows2 = sbSub2.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows2)
            {
                if (ln != "")
                    row_nos2 += "+E" + ln;
            }
            if (dtNew.Rows.Count > 0)
            {
                //ws.Cells["C" + rw].Value = "TOTAL";
                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=" + row_nos2;
                ws.Cells["F" + rw].Formula = "=" + row_nos2.Replace("E", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos2.Replace("E", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos2.Replace("E", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos2.Replace("E", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos2.Replace("E", "J");
                ws.Cells["K" + rw].Formula = "=" + row_nos2.Replace("E", "K");
                ws.Cells["L" + rw].Formula = "=" + row_nos2.Replace("E", "L");
                ws.Cells["M" + rw].Formula = "=" + row_nos2.Replace("E", "M");
                ws.Cells["N" + rw].Formula = "=" + row_nos2.Replace("E", "N");
                ws.Cells["O" + rw].Formula = "=" + row_nos2.Replace("E", "O");
                ws.Cells["P" + rw].Formula = "=" + row_nos2.Replace("E", "P");
                ws.Cells["Q" + rw].Formula = "=" + row_nos2.Replace("E", "Q");
                ws.Cells["R" + rw].Formula = "=" + row_nos2.Replace("E", "R");
                ws.Cells["S" + rw].Formula = "=" + row_nos2.Replace("E", "S");
                ws.Cells["T" + rw].Formula = "=" + row_nos2.Replace("E", "T");
                ws.Cells["U" + rw].Formula = "=" + row_nos2.Replace("E", "U");
                ws.Cells["V" + rw].Formula = "=" + row_nos2.Replace("E", "V");
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            #endregion total open top total open top


            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleOpenTop);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Reefer Type

            rw++;
            rng = ws.Cells["E" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "REEFER";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["E" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20'RF";
            ws.Cells["F" + rw].Value = "40'RF";
            ws.Cells["G" + rw].Value = "HQ'RF";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20'RF";
            ws.Cells["I" + rw].Value = "40'RF";
            ws.Cells["J" + rw].Value = "HQ'RF";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20'RF";
            ws.Cells["L" + rw].Value = "40'RF";
            ws.Cells["M" + rw].Value = "HQ'RF";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20'RF";
            ws.Cells["O" + rw].Value = "40'RF";
            ws.Cells["P" + rw].Value = "HQ'RF";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20'RF";
            ws.Cells["R" + rw].Value = "40'RF";
            ws.Cells["S" + rw].Value = "HQ'RF";

            //TOTAL
            ws.Cells["T" + rw].Value = "20'RF";
            ws.Cells["U" + rw].Value = "40'RF";
            ws.Cells["V" + rw].Value = "HQ'RF";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;
            rw_bgn_sub3 = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " RF20L7 > 0 OR RF40L7 > 0 OR RFHQL7 > 0 OR " +
                " RF208to14 > 0 OR RF408to14 > 0 OR RFHQ8to14 > 0 OR " +
                " RF2015to29 > 0 OR RF4015to29 > 0 OR RFHQ15to29 > 0 OR " +
                " RF2030to59 > 0 OR RF4030to59 > 0 OR RFHQ30to59 > 0 OR " +
                " RF20G59 > 0 OR RF40G59 > 0 OR RFHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;
            rw++;

            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                rng = ws.Cells["A" + rw + ":AB" + rw];
                //rng.Style.Font.Bold = true;
                // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (mainsectorrf.Trim() != dtNew.Rows[i]["CountryName"].ToString().Trim())
                {
                    if (mainsectorrf != "" && i > 0)
                    {
                        rw++;
                        rw_end = rw - 1;

                        sbSub3.AppendLine(rw_end.ToString());
                        //#region REGION wise sub total
                        if (dtNew.Rows.Count > 0)
                        {
                            ExcelRange range2 = ws.Cells["A" + rw_end + ":D" + rw_end];
                            //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                            range2.Value = "Sub Total - " + mainsectorrf;
                            range2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range2.Merge = true;
                            //  range2.Style.Font.Bold = true;

                            //range2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //range2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Fill.BackgroundColor.SetColor(Constants.ReportSubtotalBgColor);
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Bold = true;
                            //ws.Cells["B" + rw_end + ":U" + rw_end].Style.Font.Color.SetColor(Constants.ReportSubtotalFontColor);

                            //Region merge 13-sep-2019 rgs
                            ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Merge = true;
                            ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            //End Region merge

                            //Subtotal Begin 13-sept-2019 rgs
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            //End Subtotal

                            ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub3, (rw_end - 1));
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub3, (rw_end - 1));

                            if (country_start > 0)
                            {
                                country_end = rw - 2;
                                //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                                ws.Cells[country_start, 2, country_end, 2].Merge = true;
                                ws.Cells[country_start, 2, country_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                country_start = 0;
                            }

                            if (GeoLoc_start > 0)
                            {
                                GeoLoc_end = rw - 2;
                                //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Merge = true;
                                ws.Cells[GeoLoc_start, 3, GeoLoc_end, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                GeoLoc_start = 0;
                            }
                        }

                    }
                    mainsectorrf = dtNew.Rows[i]["CountryName"].ToString();
                    ws.Cells["A" + rw].Value = mainsectorrf;
                    rw_bgn_sub3 = rw;
                }


                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                //For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["C" + rw].Value;
                val_prev_GeoLoc = ws.Cells["C" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["RF40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["RFHQG59"].ToString());

                //ws.Cells["S" + rw].Formula = "=C" + rw + "+F" + rw + "+I" + rw + "+L" + rw + "+O" + rw;
                //ws.Cells["T" + rw].Formula = "=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                //ws.Cells["U" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;

                //13-Sept-2019 RGS
                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                rw++;
            }

            #region Final sub-total reefer

            if (mainsectorrf.Trim() != "") //nothing
            {
                if (mainsectorrf != "")//nothing
                {
                    rw++;
                    rw_end = rw - 1;
                    sbSub3.AppendLine(rw_end.ToString());
                    if (dtNew.Rows.Count > 0)
                    {
                        //#region REGION wise sub total
                        ExcelRange range = ws.Cells["A" + rw_end + ":D" + rw_end];
                        //dtNew.Rows[i]["mainsector"].ToString() + " - " +
                        range.Value = "Sub Total - " + mainsectorrf;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Merge = true;
                        //range.Style.Font.Bold = true;

                        //Region merge 13-sep-2019 rgs
                        ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Merge = true;
                        ws.Cells[rw_bgn_sub3, 1, (rw_end - 1), 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        //End Region merge

                        //Subtotal Begin 13-sept-2019 rgs
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Bold = true;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Size = 9;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Font.Color.SetColor(Color.Black);
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells["A" + rw_end + ":V" + rw_end].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //End Subtotal

                        ws.Cells["E" + rw_end].Formula = string.Format("=SUM(E{0}:E{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["F" + rw_end].Formula = string.Format("=SUM(F{0}:F{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["G" + rw_end].Formula = string.Format("=SUM(G{0}:G{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["H" + rw_end].Formula = string.Format("=SUM(H{0}:H{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["I" + rw_end].Formula = string.Format("=SUM(I{0}:I{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["J" + rw_end].Formula = string.Format("=SUM(J{0}:J{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["U" + rw_end].Formula = string.Format("=SUM(U{0}:U{1})", rw_bgn_sub3, (rw_end - 1));
                        ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub2, (rw_end - 1));
                    }

                }

                //mainsector = dtNew.Rows[i]["mainsector"].ToString();
                //ws.Cells["A" + rw].Value = mainsector;
                //ws.Cells["A" + rw].Merge = true;
                //rw_bgn_sub = rw;
            }

            #endregion

            #region foooter total reefer
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

             row_nos3 = "";
              sub_rows3 = sbSub3.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var ln in sub_rows3)
            {
                if (ln != "")
                    row_nos3 += "+E" + ln;
            }

            if (dtNew.Rows.Count > 0)
            {
                // ws.Cells["C" + rw].Value = "TOTAL";

                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=" + row_nos3;
                ws.Cells["F" + rw].Formula = "=" + row_nos3.Replace("E", "F");
                ws.Cells["G" + rw].Formula = "=" + row_nos3.Replace("E", "G");
                ws.Cells["H" + rw].Formula = "=" + row_nos3.Replace("E", "H");
                ws.Cells["I" + rw].Formula = "=" + row_nos3.Replace("E", "I");
                ws.Cells["J" + rw].Formula = "=" + row_nos3.Replace("E", "J");
                ws.Cells["K" + rw].Formula = "=" + row_nos3.Replace("E", "K");
                ws.Cells["L" + rw].Formula = "=" + row_nos3.Replace("E", "L");
                ws.Cells["M" + rw].Formula = "=" + row_nos3.Replace("E", "M");
                ws.Cells["N" + rw].Formula = "=" + row_nos3.Replace("E", "N");
                ws.Cells["O" + rw].Formula = "=" + row_nos3.Replace("E", "O");
                ws.Cells["P" + rw].Formula = "=" + row_nos3.Replace("E", "P");
                ws.Cells["Q" + rw].Formula = "=" + row_nos3.Replace("E", "Q");
                ws.Cells["R" + rw].Formula = "=" + row_nos3.Replace("E", "R");
                ws.Cells["S" + rw].Formula = "=" + row_nos3.Replace("E", "S");
                ws.Cells["T" + rw].Formula = "=" + row_nos3.Replace("E", "T");
                ws.Cells["U" + rw].Formula = "=" + row_nos3.Replace("E", "U");
                ws.Cells["V" + rw].Formula = "=" + row_nos3.Replace("E", "V");
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            #endregion total open top total open top

            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleReefer);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion

            #region Tank Cntr Type

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TANK";
            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            rng = ws.Cells["E" + rw + ":G" + rw];
            rng.Merge = true;
            rng.Value = "< 7 Days";
            rng = ws.Cells["H" + rw + ":J" + rw];
            rng.Merge = true;
            rng.Value = "8 - 14 Days";
            rng = ws.Cells["K" + rw + ":M" + rw];
            rng.Merge = true;
            rng.Value = "15 - 29 Days";
            rng = ws.Cells["N" + rw + ":P" + rw];
            rng.Merge = true;
            rng.Value = ">30 - 59 Days";
            rng = ws.Cells["Q" + rw + ":S" + rw];
            rng.Merge = true;
            rng.Value = "> 59 Days";
            rng = ws.Cells["T" + rw + ":V" + rw];
            rng.Merge = true;
            rng.Value = "TOTAL";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw++;
            rng = ws.Cells["A" + rw + ":AB" + rw];
            rng.Style.Font.Bold = true;
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["A" + rw].Value = "Region";
            ws.Cells["B" + rw].Value = "Country";
            ws.Cells["C" + rw].Value = "Geo Location";
            ws.Cells["D" + rw].Value = "Location";
            //ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //< 7 Days
            ws.Cells["E" + rw].Value = "20'TK";
            ws.Cells["F" + rw].Value = "40'TK";
            ws.Cells["G" + rw].Value = "HQ'TK";

            //8 - 14 Days
            ws.Cells["H" + rw].Value = "20'TK";
            ws.Cells["I" + rw].Value = "40'TK";
            ws.Cells["J" + rw].Value = "HQ'TK";

            //15 - 29 Days
            ws.Cells["K" + rw].Value = "20'TK";
            ws.Cells["L" + rw].Value = "40'TK";
            ws.Cells["M" + rw].Value = "HQ'TK";

            //>30 - 59 Days
            ws.Cells["N" + rw].Value = "20'TK";
            ws.Cells["O" + rw].Value = "40'TK";
            ws.Cells["P" + rw].Value = "HQ'TK";

            //> 59 Days
            ws.Cells["Q" + rw].Value = "20'TK";
            ws.Cells["R" + rw].Value = "40'TK";
            ws.Cells["S" + rw].Value = "HQ'TK";

            //TOTAL
            ws.Cells["T" + rw].Value = "20'TK";
            ws.Cells["U" + rw].Value = "40'TK";
            ws.Cells["V" + rw].Value = "HQ'TK";

            //13-sept -2019 rgs
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            rw_bgn = rw;
            rw_end = rw;
            //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

            dt = _dtValue;
            dv1 = new DataView(dt);
            dv1.RowFilter = "(" +
                " Tank20L7 > 0 OR Tank40L7 > 0 OR TankHQL7 > 0 OR " +
                " Tank208to14 > 0 OR Tank408to14 > 0 OR TankHQ8to14 > 0 OR " +
                " Tank2015to29 > 0 OR Tank4015to29 > 0 OR TankHQ15to29 > 0 OR " +
                " Tank2030to59 > 0 OR Tank4030to59 > 0 OR TankHQ30to59 > 0 OR " +
                " Tank20G59 > 0 OR Tank40G59 > 0 OR TankHQG59 > 0 " +
                ")";
            dtNew = dv1.ToTable();
            dtrows = dtNew.Rows.Count - 1;
            rw++;

            for (int i = 0; i < dtNew.Rows.Count; i++)
            {
                // rng = ws.Cells["A" + rw + ":Z" + rw];
                //rng.Style.Font.Bold = true;
                // rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                country1 = dtNew.Rows[i]["CountryName"].ToString();

                ws.Cells["B" + rw].Value = country1;

                val_this = ws.Cells["B" + rw].Value;
                val_prev = ws.Cells["B" + (rw - 1)].Value;

                if (country_start == 0 && val_this.ToString() == val_prev.ToString())
                {
                    country_start = rw - 1;
                }

                if (country_start > 0 && val_this.ToString() != val_prev.ToString())
                {
                    country_end = rw - 1;
                    //e_msg += "<br>country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                if (i == dtrows && val_this.ToString() == val_prev.ToString())
                {
                    country_end = rw;
                    //e_msg += "<br>LastRow:country_start=" + country_start + ", country_end=" + country_end;
                    ws.Cells[country_start, 1, country_end, 1].Merge = true;
                    ws.Cells[country_start, 1, country_end, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    country_start = 0;
                }

                //For GeoLoc
                GeoLocation = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["C" + rw].Value = GeoLocation;

                val_this_GeoLoc = ws.Cells["B" + rw].Value;
                val_prev_GeoLoc = ws.Cells["B" + (rw - 1)].Value;

                if (GeoLoc_start == 0 && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_start = rw - 1;
                }

                if (GeoLoc_start > 0 && val_this_GeoLoc.ToString() != val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw - 1;
                    //e_msg += "<br>GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                if (i == dtrows && val_this_GeoLoc.ToString() == val_prev_GeoLoc.ToString())
                {
                    GeoLoc_end = rw;
                    //e_msg += "<br>LastRow:GeoLoc_start=" + GeoLoc_start + ", GeoLoc_end=" + GeoLoc_end;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Merge = true;
                    ws.Cells[GeoLoc_start, 2, GeoLoc_end, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    GeoLoc_start = 0;
                }

                ws.Cells["C" + rw].Value = dtNew.Rows[i]["GeoLocName"].ToString();
                ws.Cells["D" + rw].Value = dtNew.Rows[i]["CurrentPort"].ToString();
                ws.Cells["D" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ws.Cells["E" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank20L7"].ToString());
                ws.Cells["F" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank40L7"].ToString());
                ws.Cells["G" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQL7"].ToString());

                ws.Cells["H" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank208to14"].ToString());
                ws.Cells["I" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank408to14"].ToString());
                ws.Cells["J" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQ8to14"].ToString());

                ws.Cells["K" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank2015to29"].ToString());
                ws.Cells["L" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank4015to29"].ToString());
                ws.Cells["M" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQ15to29"].ToString());

                ws.Cells["N" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank2030to59"].ToString());
                ws.Cells["O" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank4030to59"].ToString());
                ws.Cells["P" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQ30to59"].ToString());

                ws.Cells["Q" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank20G59"].ToString());
                ws.Cells["R" + rw].Value = TryParseNullable(dtNew.Rows[i]["Tank40G59"].ToString());
                ws.Cells["S" + rw].Value = TryParseNullable(dtNew.Rows[i]["TankHQG59"].ToString());

                ws.Cells["T" + rw].Formula = "=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;//"=D" + rw + "+G" + rw + "+J" + rw + "+M" + rw + "+P" + rw;
                ws.Cells["U" + rw].Formula = "=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;//"=E" + rw + "+H" + rw + "+K" + rw + "+N" + rw + "+Q" + rw;
                ws.Cells["V" + rw].Formula = "=G" + rw + "+J" + rw + "+M" + rw + "+P" + rw + "+S" + rw;//"=F" + rw + "+I" + rw + "+L" + rw + "+O" + rw + "+R" + rw;

                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                //ws.Cells["" + rw].Value = dtNew.Rows[i][""].ToString();

                rw++;
            }

            if (dtNew.Rows.Count > 0)
            {
                rw_end = rw - 1;

                //rng = ws.Cells["A" + rw + ":Z" + rw];
                //rng.Style.Font.Bold = true;
                //rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //ws.Cells["C" + rw].Value = "TOTAL";
                //Total 13-sep-2019 rgs
                rngtemp = ws.Cells["A" + rw + ":D" + rw];
                rngtemp.Value = "GRAND TOTAL";
                rngtemp.Merge = true;

                //Total 13-sep-2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Bold = true;
                ws.Cells["A" + rw + ":V" + rw].Style.Font.Size = 9;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells["E" + rw].Formula = "=SUM(E" + rw_bgn + ":E" + rw_end + ")";
                ws.Cells["F" + rw].Formula = "=SUM(F" + rw_bgn + ":F" + rw_end + ")";
                ws.Cells["G" + rw].Formula = "=SUM(G" + rw_bgn + ":G" + rw_end + ")";
                ws.Cells["H" + rw].Formula = "=SUM(H" + rw_bgn + ":H" + rw_end + ")";
                ws.Cells["I" + rw].Formula = "=SUM(I" + rw_bgn + ":I" + rw_end + ")";
                ws.Cells["J" + rw].Formula = "=SUM(J" + rw_bgn + ":J" + rw_end + ")";
                ws.Cells["K" + rw].Formula = "=SUM(K" + rw_bgn + ":K" + rw_end + ")";
                ws.Cells["L" + rw].Formula = "=SUM(L" + rw_bgn + ":L" + rw_end + ")";
                ws.Cells["M" + rw].Formula = "=SUM(M" + rw_bgn + ":M" + rw_end + ")";
                ws.Cells["N" + rw].Formula = "=SUM(N" + rw_bgn + ":N" + rw_end + ")";
                ws.Cells["O" + rw].Formula = "=SUM(O" + rw_bgn + ":O" + rw_end + ")";
                ws.Cells["P" + rw].Formula = "=SUM(P" + rw_bgn + ":P" + rw_end + ")";
                ws.Cells["Q" + rw].Formula = "=SUM(Q" + rw_bgn + ":Q" + rw_end + ")";
                ws.Cells["R" + rw].Formula = "=SUM(R" + rw_bgn + ":R" + rw_end + ")";
                ws.Cells["S" + rw].Formula = "=SUM(S" + rw_bgn + ":S" + rw_end + ")";
                ws.Cells["T" + rw].Formula = "=SUM(T" + rw_bgn + ":T" + rw_end + ")";
                ws.Cells["U" + rw].Formula = "=SUM(U" + rw_bgn + ":U" + rw_end + ")";
                ws.Cells["V" + rw].Formula = "=SUM(V" + rw_bgn + ":V" + rw_end + ")";
            }
            else
            {
                //13-sept -2019 rgs
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + rw + ":V" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //ws.Cells["A" + (rw_bgn - 2) + ":U" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            //Header
            colFromHex = System.Drawing.ColorTranslator.FromHtml(styleTank);
            rng = ws.Cells["E" + (rw_bgn - 2) + ":G" + (rw_bgn - 2)];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#D8E4BC");
            rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E4DFEC");
            rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#E1E199");
            rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#FF9966");
            rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#ff9999");
            rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            colFromHex = System.Drawing.ColorTranslator.FromHtml("#BDBDBD");
            rng = ws.Cells["T" + (rw_bgn - 2) + ":V" + rw];
            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

            //ws.Cells["A" + rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
            //ws.Cells[8, 15, rw, 39].Style.Numberformat.Format = "#,##0.00";
            rw++;

            #endregion


            ws.Cells[7, 7, ws.Dimension.End.Row, ws.Dimension.End.Column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //  ws.Cells[ws.Dimension.Address].AutoFitColumns();

            #region Column Width to 5

            //ws.Column(1).AutoFit();
            ws.Column(1).Width = 15;
            ws.Column(1).Style.WrapText = true;
            // ws.Column(1).BestFit = true;
            ws.Column(2).AutoFit();
            ws.Column(2).Style.WrapText = true;

            ws.Column(3).AutoFit();
            ws.Column(3).Style.WrapText = true;

            ws.Column(4).AutoFit();
            ws.Column(4).Style.WrapText = true;
            ws.Column(5).AutoFit();
            ws.Column(5).Style.WrapText = true;
            ws.Column(6).AutoFit();
            ws.Column(6).Style.WrapText = true;


            ws.Column(7).Width = 5.71;
            ws.Column(8).Width = 5.71;
            ws.Column(9).Width = 5.71;
            ws.Column(10).Width = 5.71;
            ws.Column(11).Width = 5.71;
            ws.Column(12).Width = 5.71;
            ws.Column(13).Width = 5.71;
            ws.Column(14).Width = 5.71;
            ws.Column(15).Width = 5.71;
            ws.Column(16).Width = 5.71;
            ws.Column(17).Width = 5.71;
            ws.Column(18).Width = 5.71;
            ws.Column(19).Width = 5.71;
            ws.Column(20).Width = 5.71;
            ws.Column(21).Width = 5.71;
            ws.Column(22).Width = 5.71;
            ws.Column(23).Width = 5.71;
            ws.Column(24).Width = 5.71;
            ws.Column(25).Width = 5.71;
            ws.Column(26).Width = 5.71;
            ws.Column(27).Width = 5.71;
            ws.Column(28).Width = 5.71;
            ws.Column(29).Width = 5.71;
            ws.Column(30).Width = 5.71;
            #endregion

            #endregion

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=EQCAgeWiseReport.xlsx");
            Response.End();
            //return pck.GetAsByteArray();
        }

        public DataTable _dtAgewiseRepValue(string CntrStCode, string CntryId, string GeoLocId, string Grade, string LeaseTerm, string CntrTypes, string criteria)
        {
            string strWhere = "";

            string _Query = "SELECT LeaseTerm,CountryName,GeoLocName,StatusCode,CurrentPort ";

            if (CntrTypes == "1" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(Dry20L7,0)) Dry20L7, SUM(ISNULL(Dry40L7,0)) Dry40L7, SUM(ISNULL(DryHQL7,0)) DryHQL7,  " +
                                          " SUM(ISNULL(Dry208To14,0)) Dry208To14, SUM(ISNULL(Dry408To14,0)) Dry408To14, SUM(ISNULL(DryHQ8To14,0)) DryHQ8To14,  " +
                                          " SUM(ISNULL(Dry2015To29,0)) Dry2015To29, SUM(ISNULL(Dry4015To29,0)) Dry4015To29, " +
                                          " SUM(ISNULL(DryHQ15To29,0)) DryHQ15To29, SUM(ISNULL(Dry2030To59,0)) Dry2030To59, SUM(ISNULL(Dry4030To59,0)) Dry4030To59, " +
                                          " SUM(ISNULL(DryHQ30To59,0)) DryHQ30To59, SUM(ISNULL(Dry20G59,0)) Dry20G59,  " +
                                          " SUM(ISNULL(Dry40G59,0)) Dry40G59,SUM(ISNULL(DryHQG59,0)) DryHQG59, " +
                                          " SUM(ISNULL(Dry20HQL7,0)) Dry20HQL7,  " +
                                          " SUM(ISNULL(Dry20HQ8To14,0)) Dry20HQ8To14,SUM(ISNULL(Dry20HQ15To29,0)) Dry20HQ15To29, " +
                                          " SUM(ISNULL(Dry20HQ30To59,0)) Dry20HQ30To59, SUM(ISNULL(Dry20HQG59,0)) Dry20HQG59 ";



            if (CntrTypes == "2" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(FlRack20L7,0)) FlRack20L7, SUM(ISNULL(FlRack40L7,0)) FlRack40L7,SUM(ISNULL(FlRackHQL7,0)) FlRackHQL7,SUM(ISNULL(FlRack208To14,0)) FlRack208To14, SUM(ISNULL(FlRack408To14,0)) FlRack408To14, " +
                                       " SUM(ISNULL(FlRackHQ8To14,0)) FlRackHQ8To14,SUM(ISNULL(FlRack2015To29,0)) FlRack2015To29, SUM(ISNULL(FlRack4015To29,0)) FlRack4015To29,SUM(ISNULL(FlRackHQ15To29,0)) FlRackHQ15To29,SUM(ISNULL(FlRack2030To59,0)) FlRack2030To59, " +
                                       " SUM(ISNULL(FlRack4030To59,0)) FlRack4030To59,SUM(ISNULL(FlRackHQ30To59,0)) FlRackHQ30To59,SUM(ISNULL(FlRack20G59,0)) FlRack20G59,SUM(ISNULL(FlRack40G59,0)) FlRack40G59,SUM(ISNULL(FlRackHQG59,0)) FlRackHQG59 ";

            if (CntrTypes == "3" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(OpTop20L7,0)) OpTop20L7, SUM(ISNULL(OpTop40L7,0)) OpTop40L7, SUM(ISNULL(OpTopHQL7,0)) OpTopHQL7,SUM(ISNULL(OpTop208To14,0)) OpTop208To14, SUM(ISNULL(OpTop408To14,0)) OpTop408To14, " +
                                     " SUM(ISNULL(OpTopHQ8To14,0)) OpTopHQ8To14,SUM(ISNULL(OpTop2015To29,0)) OpTop2015To29,SUM(ISNULL(OpTop4015To29,0)) OpTop4015To29,SUM(ISNULL(OpTopHQ15To29,0)) OpTopHQ15To29, SUM(ISNULL(OpTop2030To59,0))OpTop2030To59,  " +
                                     " SUM(ISNULL(OpTop4030To59,0)) OpTop4030To59, SUM(ISNULL(OpTopHQ30To59,0)) OpTopHQ30To59,SUM(ISNULL(OpTop20G59,0)) OpTop20G59,SUM(ISNULL(OpTop40G59,0)) OpTop40G59, SUM(ISNULL(OpTopHQG59,0)) OpTopHQG59 ";

            if (CntrTypes == "4" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(RF20L7,0)) RF20L7,SUM(ISNULL(RF40L7,0)) RF40L7,SUM(ISNULL(RFHQL7,0)) RFHQL7,SUM(ISNULL(RF208To14,0)) RF208To14,SUM(ISNULL(RF408To14,0)) RF408To14," +
                                      " SUM(ISNULL(RFHQ8To14,0)) RFHQ8To14,SUM(ISNULL(RF2015To29,0)) RF2015To29,SUM(ISNULL(RF4015To29,0)) RF4015To29,SUM(ISNULL(RFHQ15To29,0)) RFHQ15To29, SUM(ISNULL(RF2030To59,0)) RF2030To59, " +
                                      " SUM(ISNULL(RF4030To59,0)) RF4030To59, SUM(ISNULL(RFHQ30To59,0)) RFHQ30To59, SUM(ISNULL(RF20G59,0)) RF20G59,SUM(ISNULL(RF40G59,0)) RF40G59, SUM(ISNULL(RFHQG59 ,0)) RFHQG59 ";


            if (CntrTypes == "5" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(Tank20L7,0)) Tank20L7, SUM(ISNULL(Tank40L7,0)) Tank40L7, SUM(ISNULL(TankHQL7,0)) TankHQL7,  " +
                                    " SUM(ISNULL(Tank208To14,0)) Tank208To14, SUM(ISNULL(Tank408To14,0)) Tank408To14, SUM(ISNULL(TankHQ8To14,0)) TankHQ8To14, " +
                                    " SUM(ISNULL(Tank2015To29,0)) Tank2015To29, SUM(ISNULL(Tank4015To29,0)) Tank4015To29,  " +
                                    " SUM(ISNULL(TankHQ15To29,0)) TankHQ15To29, SUM(ISNULL(Tank2030To59,0)) Tank2030To59, SUM(ISNULL(Tank4030To59,0)) Tank4030To59, " +
                                    " SUM(ISNULL(TankHQ30To59,0)) TankHQ30To59, SUM(ISNULL(Tank20G59,0)) Tank20G59, SUM(ISNULL(Tank40G59,0)) Tank40G59, SUM(ISNULL(TankHQG59,0)) TankHQG59 ";

            if (CntryId != "" && CntryId != "0" && CntryId != "null" && CntryId != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal WHERE CtryID=" + CntryId;
                else
                    strWhere += " and CtryID =" + CntryId;

            if (GeoLocId != "" && GeoLocId != "0" && GeoLocId != "null" && GeoLocId != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal WHERE GeoLocID=" + GeoLocId;
                else
                    strWhere += " and GeoLocID =" + GeoLocId;


            if (Grade != "" && Grade != "0" && Grade != "null" && Grade != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal WHERE GradeID=" + Grade;
                else
                    strWhere += " and GradeID =" + Grade;

            if (LeaseTerm != "" && LeaseTerm != "0" && LeaseTerm != null && LeaseTerm != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal WHERE TermID=" + LeaseTerm;
                else
                    strWhere += " and TermID =" + LeaseTerm;

            if (CntrStCode != "" && CntrStCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal  where StatusCode ='" + CntrStCode + "'";
                else
                    strWhere += " and StatusCode ='" + CntrStCode + "'";

            if (strWhere == "")
                strWhere += _Query + " from VInvSnShotFinal GROUP BY  LeaseTerm, CountryName, GeoLocName, StatusCode,CurrentPort ORDER BY  CountryName, GeoLocName,LeaseTerm,StatusCode,CurrentPort ";

            else
                strWhere += "  GROUP BY  CountryName, GeoLocName, StatusCode,CurrentPort,LeaseTerm ORDER BY CountryName, GeoLocName , LeaseTerm,StatusCode,CurrentPort ";


            return Manag.GetViewData(strWhere, "");
        }


        public DataTable _dtAgewiseRepValueSummary(string CntrStCode, string CntryId, string GeoLocId, string Grade, string LeaseTerm, string CntrTypes, string criteria)
        {
            string strWhere = "";

            string _Query = "SELECT CountryName,GeoLocName,StatusCode,CurrentPort ";

            if (CntrTypes == "1" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(Dry20L7,0)) Dry20L7, SUM(ISNULL(Dry40L7,0)) Dry40L7, SUM(ISNULL(DryHQL7,0)) DryHQL7,  " +
                                          " SUM(ISNULL(Dry208To14,0)) Dry208To14, SUM(ISNULL(Dry408To14,0)) Dry408To14, SUM(ISNULL(DryHQ8To14,0)) DryHQ8To14,  " +
                                          " SUM(ISNULL(Dry2015To29,0)) Dry2015To29, SUM(ISNULL(Dry4015To29,0)) Dry4015To29, " +
                                          " SUM(ISNULL(DryHQ15To29,0)) DryHQ15To29, SUM(ISNULL(Dry2030To59,0)) Dry2030To59, SUM(ISNULL(Dry4030To59,0)) Dry4030To59, " +
                                          " SUM(ISNULL(DryHQ30To59,0)) DryHQ30To59, SUM(ISNULL(Dry20G59,0)) Dry20G59,  " +
                                          " SUM(ISNULL(Dry40G59,0)) Dry40G59,SUM(ISNULL(DryHQG59,0)) DryHQG59, " +
                                          " SUM(ISNULL(Dry20HQL7,0)) Dry20HQL7,  " +
                                          " SUM(ISNULL(Dry20HQ8To14,0)) Dry20HQ8To14,SUM(ISNULL(Dry20HQ15To29,0)) Dry20HQ15To29, " +
                                          " SUM(ISNULL(Dry20HQ30To59,0)) Dry20HQ30To59, SUM(ISNULL(Dry20HQG59,0)) Dry20HQG59 ";



            if (CntrTypes == "2" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(FlRack20L7,0)) FlRack20L7, SUM(ISNULL(FlRack40L7,0)) FlRack40L7,SUM(ISNULL(FlRackHQL7,0)) FlRackHQL7,SUM(ISNULL(FlRack208To14,0)) FlRack208To14, SUM(ISNULL(FlRack408To14,0)) FlRack408To14, " +
                                       " SUM(ISNULL(FlRackHQ8To14,0)) FlRackHQ8To14,SUM(ISNULL(FlRack2015To29,0)) FlRack2015To29, SUM(ISNULL(FlRack4015To29,0)) FlRack4015To29,SUM(ISNULL(FlRackHQ15To29,0)) FlRackHQ15To29,SUM(ISNULL(FlRack2030To59,0)) FlRack2030To59, " +
                                       " SUM(ISNULL(FlRack4030To59,0)) FlRack4030To59,SUM(ISNULL(FlRackHQ30To59,0)) FlRackHQ30To59,SUM(ISNULL(FlRack20G59,0)) FlRack20G59,SUM(ISNULL(FlRack40G59,0)) FlRack40G59,SUM(ISNULL(FlRackHQG59,0)) FlRackHQG59 ";

            if (CntrTypes == "3" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(OpTop20L7,0)) OpTop20L7, SUM(ISNULL(OpTop40L7,0)) OpTop40L7, SUM(ISNULL(OpTopHQL7,0)) OpTopHQL7,SUM(ISNULL(OpTop208To14,0)) OpTop208To14, SUM(ISNULL(OpTop408To14,0)) OpTop408To14, " +
                                     " SUM(ISNULL(OpTopHQ8To14,0)) OpTopHQ8To14,SUM(ISNULL(OpTop2015To29,0)) OpTop2015To29,SUM(ISNULL(OpTop4015To29,0)) OpTop4015To29,SUM(ISNULL(OpTopHQ15To29,0)) OpTopHQ15To29, SUM(ISNULL(OpTop2030To59,0))OpTop2030To59,  " +
                                     " SUM(ISNULL(OpTop4030To59,0)) OpTop4030To59, SUM(ISNULL(OpTopHQ30To59,0)) OpTopHQ30To59,SUM(ISNULL(OpTop20G59,0)) OpTop20G59,SUM(ISNULL(OpTop40G59,0)) OpTop40G59, SUM(ISNULL(OpTopHQG59,0)) OpTopHQG59 ";

            if (CntrTypes == "4" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(RF20L7,0)) RF20L7,SUM(ISNULL(RF40L7,0)) RF40L7,SUM(ISNULL(RFHQL7,0)) RFHQL7,SUM(ISNULL(RF208To14,0)) RF208To14,SUM(ISNULL(RF408To14,0)) RF408To14," +
                                      " SUM(ISNULL(RFHQ8To14,0)) RFHQ8To14,SUM(ISNULL(RF2015To29,0)) RF2015To29,SUM(ISNULL(RF4015To29,0)) RF4015To29,SUM(ISNULL(RFHQ15To29,0)) RFHQ15To29, SUM(ISNULL(RF2030To59,0)) RF2030To59, " +
                                      " SUM(ISNULL(RF4030To59,0)) RF4030To59, SUM(ISNULL(RFHQ30To59,0)) RFHQ30To59, SUM(ISNULL(RF20G59,0)) RF20G59,SUM(ISNULL(RF40G59,0)) RF40G59, SUM(ISNULL(RFHQG59 ,0)) RFHQG59 ";


            if (CntrTypes == "5" || CntrTypes == "undefined")
                _Query += ",SUM(ISNULL(Tank20L7,0)) Tank20L7, SUM(ISNULL(Tank40L7,0)) Tank40L7, SUM(ISNULL(TankHQL7,0)) TankHQL7,  " +
                                    " SUM(ISNULL(Tank208To14,0)) Tank208To14, SUM(ISNULL(Tank408To14,0)) Tank408To14, SUM(ISNULL(TankHQ8To14,0)) TankHQ8To14, " +
                                    " SUM(ISNULL(Tank2015To29,0)) Tank2015To29, SUM(ISNULL(Tank4015To29,0)) Tank4015To29,  " +
                                    " SUM(ISNULL(TankHQ15To29,0)) TankHQ15To29, SUM(ISNULL(Tank2030To59,0)) Tank2030To59, SUM(ISNULL(Tank4030To59,0)) Tank4030To59, " +
                                    " SUM(ISNULL(TankHQ30To59,0)) TankHQ30To59, SUM(ISNULL(Tank20G59,0)) Tank20G59, SUM(ISNULL(Tank40G59,0)) Tank40G59, SUM(ISNULL(TankHQG59,0)) TankHQG59 ";

            if (CntryId != "" && CntryId != "0" && CntryId != "null" && CntryId != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal_Summary WHERE CtryID=" + CntryId;
                else
                    strWhere += " and CtryID =" + CntryId;

            if (GeoLocId != "" && GeoLocId != "0" && GeoLocId != "null" && GeoLocId != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal_Summary WHERE GeoLocID=" + GeoLocId;
                else
                    strWhere += " and GeoLocID =" + GeoLocId;


            //if (Grade != "" && Grade != "0" && Grade != "null" && Grade != "?")
            //    if (strWhere == "")
            //        strWhere += _Query + " from VInvSnShotFinal_Summary WHERE GradeID=" + Grade;
            //    else
            //        strWhere += " and GradeID =" + Grade;

            //if (LeaseTerm != "" && LeaseTerm != "0" && LeaseTerm != null && LeaseTerm != "?")
            //    if (strWhere == "")
            //        strWhere += _Query + " from VInvSnShotFinal_Summary WHERE TermID=" + LeaseTerm;
            //    else
            //        strWhere += " and TermID =" + LeaseTerm;

            if (CntrStCode != "" && CntrStCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinal_Summary  where StatusCode ='" + CntrStCode + "'";
                else
                    strWhere += " and StatusCode ='" + CntrStCode + "'";

            if (strWhere == "")
                strWhere += _Query + " from VInvSnShotFinal_Summary GROUP BY   CountryName, GeoLocName, StatusCode,CurrentPort ORDER BY  CountryName, GeoLocName,StatusCode,CurrentPort ";

            else
                strWhere += "  GROUP BY  CountryName, GeoLocName, StatusCode,CurrentPort ORDER BY CountryName, GeoLocName ,StatusCode,CurrentPort ";


            return Manag.GetViewData(strWhere, "");
        }
        public DataTable _dtAgewiseCntrsize()
        {
            string _Query = "select * from NVO_tblcntrTypes where  EQTypeId = 1";
            return Manag.GetViewData(_Query, "");
        }
        public static String getColumnNameIndex(int column)
        {
            column--;
            String col = Convert.ToString((char)('A' + (column % 26)));
            while (column >= 26)
            {
                column = (column / 26) - 1;
                col = Convert.ToString((char)('A' + (column % 26))) + col;
            }
            return col;
        }
    }
}