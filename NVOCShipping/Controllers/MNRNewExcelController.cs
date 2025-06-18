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
    public class MNRNewExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: MNRNewExcel
        public ActionResult Index()
        {
            return View();
        }
        public void MNRNewReport(string DtFrom, string DtTo, string User, string Status, string LocID, string DepotID, string AgencyID)
        {

            try
            {
                string DtFromV = "", DtToV = "";
                DateTime outDate = new DateTime();
                if (DateTime.TryParse(DtFrom, out outDate))
                    DtFromV = outDate.ToString("MM/dd/yyyy");
                if (DateTime.TryParse(DtTo, out outDate))
                    DtToV = outDate.ToString("MM/dd/yyyy");



                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MNRReport");

                ws.Cells["A2"].Value = "MNR Details List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:M2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A4"].Value = "From Date :";
                ws.Cells["A4"].Style.Font.Bold = true;
                ws.Cells["B4"].Value = DtFrom;
                ws.Cells["B4"].Style.Font.Bold = true;

                ws.Cells["C4"].Value = " To Date :";
                ws.Cells["C4"].Style.Font.Bold = true;
                ws.Cells["D4"].Value = DtTo;
                ws.Cells["D4"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A7"].Value = "S. No.";
                ws.Cells["B7"].Value = "MNR Reference";
                ws.Cells["C7"].Value = "Container No";
                ws.Cells["D7"].Value = "EmptyGate In";
                ws.Cells["E7"].Value = "Container Type";
                ws.Cells["F7"].Value = "Agency";
                ws.Cells["G7"].Value = "Location";
                ws.Cells["H7"].Value = "Depot";
                ws.Cells["I7"].Value = "BLNumber";
                ws.Cells["J7"].Value = "Consignee";
                //ws.Cells["K7"].Value = "Vessel";
                //ws.Cells["L7"].Value = "Voyage";
                ws.Cells["K7"].Value = "Estimate Amount";
                ws.Cells["L7"].Value = "Currency";
                ws.Cells["M7"].Value = "Estimate Date";
                ws.Cells["N7"].Value = "Approved Amount";
                ws.Cells["O7"].Value = "Approved Date";
                ws.Cells["P7"].Value = "Accountability";
                ws.Cells["Q7"].Value = "Line Item";
                ws.Cells["R7"].Value = "Amount";
                //ws.Cells["S7"].Value = "Invoice No";
                //ws.Cells["T7"].Value = "Collected Amount";
                ws.Cells["S7"].Value = "Recovery Remarks";
                ws.Cells["T7"].Value = "MNR Status";
                r = ws.Cells["A7:T7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;
                int frowid = 0;
                int rw = 8;
                DataTable dtv = GetMNRDetails(DtFrom, DtTo, Status, LocID, DepotID, AgencyID);
                if (dtv.Rows.Count > 0)
                {
                    for (int i = 0; i < dtv.Rows.Count; i++)
                    {
                        frowid = rw;


                        //ExcelRange rng = ws.Cells["A" + frowid + ":S" + frowid];
                        //rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                        ws.Cells["A" + rw].Value = sl++;
                        ws.Cells["B" + rw].Value = dtv.Rows[i]["MNRRefNo"].ToString();
                        ws.Cells["C" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                        ws.Cells["D" + rw].Value = dtv.Rows[i]["EmptyGateINv"].ToString();
                        ws.Cells["E" + rw].Value = dtv.Rows[i]["CntrType"].ToString();
                        ws.Cells["F" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                        ws.Cells["G" + rw].Value = dtv.Rows[i]["Location"].ToString();
                        ws.Cells["H" + rw].Value = dtv.Rows[i]["Depot"].ToString();
                        ws.Cells["I" + rw].Value = dtv.Rows[i]["BLNUmber"].ToString();
                        ws.Cells["J" + rw].Value = dtv.Rows[i]["Consignee"].ToString();
                        //ws.Cells["K" + rw].Value = dtv.Rows[i]["Vessel"].ToString();
                        //ws.Cells["L" + rw].Value = dtv.Rows[i]["Voyage"].ToString();
                        ws.Cells["K" + rw].Value = dtv.Rows[i]["EstAmt"].ToString();
                        ws.Cells["L" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["EstimateDate"].ToString();
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["ApprovedAmt"].ToString(); ;
                        ws.Cells["O" + rw].Value = dtv.Rows[i]["ApprovedDate"].ToString();
                        ws.Cells["T" + rw].Value = dtv.Rows[i]["MNRStatus"].ToString();
                        //sl++;
                        int stRow = rw;
                        int StRow1 = rw;
                        DataTable dtR = GetMNRRecoveryDetails(dtv.Rows[i]["ID"].ToString());
                        rw = stRow;
                        int chgRow = 1;


                        for (int k = 0; k < dtR.Rows.Count; k++)
                        {

                            ws.Cells["P" + StRow1].Value = dtR.Rows[k]["Accountability"].ToString();
                            ws.Cells["Q" + StRow1].Value = dtR.Rows[k]["ItemNo"].ToString();
                            ws.Cells["R" + StRow1].Value = dtR.Rows[k]["RecoveryAmount"].ToString();
                            ws.Cells["S" + StRow1].Value = dtR.Rows[k]["Remarks"].ToString();
                            StRow1++;
                            chgRow++;

                        }
                        rw = StRow1;
                        rw += 1;

                        //rw -= 1;
                    }
                }
                ws.Cells["A7:T" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:T" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:T" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:T" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();


                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MNRReport.xlsx");
                Response.End();



            }
            catch (Exception ex)
            {

            }
        }

        public DataTable GetMNRDetails(string DtFrom, string DtTo, string Status, string LocID, string DepotID, string AgencyID)
        {
            string strWhere = "";
            string _Query = "select * from NVO_ViewMNRDetailsReport ";

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " Where convert(varchar,EmptyGateIN, 23)  between '" + DtFrom + "' and '" + DtTo + "'";

            if (LocID != "" && LocID != "null" && LocID != "?" && LocID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where LocationID=" + LocID;
                else
                    strWhere += " and LocationID =" + LocID;


            if (Status != "" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where MNRStatusID=" + Status;
                else
                    strWhere += " and MNRStatusID =" + Status;


            if (DepotID != "" && DepotID != "null" && DepotID != "?" && DepotID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where DepotID=" + DepotID;
                else
                    strWhere += " and DepotID =" + DepotID;

            if (AgencyID != "" && AgencyID != "null" && AgencyID != "?" && AgencyID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where AgencyID=" + AgencyID;
                else
                    strWhere += " and AgencyID =" + AgencyID;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere + " Order by ID Desc ", "");
        }

        public DataTable GetMNRRecoveryDetails(string MNRID)
        {

            string _Query = "select * from NVO_MNRNewRecoveryDtls WHERE MNRID=" + MNRID;

            return Manag.GetViewData(_Query, "");
        }
    }
}