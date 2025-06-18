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
    public class MNRExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }
        public void CreateMNRSummaryExcel(string DepotID, string ReqRefNo, string CntrID, string DtFrom, string DtTo, string Status, string User)
        {
            DataTable dtv = GetMNRDetailSearchValues(DepotID, ReqRefNo, CntrID, DtFrom, DtTo, Status);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MNRRepair");

                ws.Cells["A2"].Value = "MNR Repair List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:L2"];
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
                ws.Cells["B7"].Value = "Reference#";
                ws.Cells["C7"].Value = "Requested Date";
                ws.Cells["D7"].Value = " Approved Date";
                ws.Cells["E7"].Value = "Container#";
                ws.Cells["F7"].Value = "Size & Type";
                ws.Cells["G7"].Value = "Requested By";
                ws.Cells["H7"].Value = "Currency";
                ws.Cells["I7"].Value = "Status";
                ws.Cells["J7"].Value = "Estimated Cost";
                ws.Cells["K7"].Value = "Approved Cost";
                ws.Cells["L7"].Value = "Approved + 50%";


                r = ws.Cells["A7:L7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;
                double amt = 0.00;
                double amt1 = 0.00;
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ReqRefNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["DtReq"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["DtApprv"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["TypeSize"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["RequestedBy"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["EstTotalCost"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["ApprTotalCost"].ToString();

                    double.TryParse(dtv.Rows[i]["EstTotalCost"].ToString(), out amt);
                    ws.Cells["J" + rw].Value = amt;
                    ws.Cells["J" + rw].Style.Numberformat.Format = "#,##0.00";

                    double.TryParse(dtv.Rows[i]["ApprTotalCost"].ToString(), out amt1);
                    ws.Cells["K" + rw].Value = amt1;
                    ws.Cells["K" + rw].Style.Numberformat.Format = "#,##0.00";

                    // Approved Cost *50 PERCENT / 100

                    ws.Cells["L" + rw].Formula = " = " + ws.Cells["K" + rw].Address + "  + (" + ws.Cells["K" + rw].Address + " * 50 / 100  )";
                    ws.Cells["K" + rw].Style.Numberformat.Format = "#,##0.00";

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:L" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:L" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:L" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:L" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MNRSummary.xlsx");
                Response.End();

            }

        }


        public void CreateMNRDetailExcel(string DepotID, string ReqRefNo, string CntrID, string DtFrom, string DtTo, string Status, string User)
        {
        
            try
            {
                string DtFromV = "", DtToV = "";
                DateTime outDate = new DateTime();
                if (DateTime.TryParse(DtFrom, out outDate))
                    DtFromV = outDate.ToString("MM/dd/yyyy");
                if (DateTime.TryParse(DtTo, out outDate))
                    DtToV = outDate.ToString("MM/dd/yyyy");

                DataTable dtv = GetMNRDetailSearchValues(DepotID, ReqRefNo, CntrID, DtFrom, DtTo, Status);
                if (dtv != null && dtv.Rows.Count > 0)
                {
                    ExcelPackage pck = new ExcelPackage();
                    var ws = pck.Workbook.Worksheets.Add("MNR REPORT");
                    ws.Cells["A1"].Value = "Container Repair Request";
                    ws.Cells["A1"].Style.Font.Bold = true;
                    ws.Cells["A1:Q1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ExcelRange r = ws.Cells["A1:Q1"];
                    r.Merge = true;

                    r = ws.Cells["A1:V1"];
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    r.Style.Font.Color.SetColor(Color.Black);
                    
                    ws.Cells["A2"].Value = "Date :";
                    ws.Cells["B2"].Value = System.DateTime.Now.ToShortDateString();

                    ws.Cells["K4"].Value = "Estimated";
                    ws.Cells["K4"].Style.Font.Bold = true;
                    ws.Cells["K4:O4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ExcelRange r2 = ws.Cells["K4:O4"];
                    r2.Merge = true;

                    ws.Cells["P4"].Value = "Approved";
                    ws.Cells["P4"].Style.Font.Bold = true;
                    ws.Cells["P4:T4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ExcelRange r3 = ws.Cells["P4:T4"];
                    r3.Merge = true;

                    #region header
                    ws.Cells["A5"].Value = "Ref No.";
                    ws.Cells["B5"].Value = "Container No";
                    ws.Cells["C5"].Value = "Category";
                    ws.Cells["D5"].Value = "Comp Cd.";
                    ws.Cells["E5"].Value = "Loc Cd";
                    ws.Cells["F5"].Value = "Dm Cd";
                    ws.Cells["G5"].Value = "Rep Type.";
                    ws.Cells["H5"].Value = "Mat Type";
                    ws.Cells["I5"].Value = "Measurement W*H";
                    ws.Cells["J5"].Value = "Unit";
                    ws.Cells["K5"].Value = "Lab Hr";
                    ws.Cells["L5"].Value = "Lab Cost";
                    ws.Cells["M5"].Value = "Total Lab Cost";
                    ws.Cells["N5"].Value = "Mat Cost";
                    ws.Cells["O5"].Value = "Total Cost";
                    ws.Cells["P5"].Value = "Lab Hr";
                    ws.Cells["Q5"].Value = "Lab Cost";
                    ws.Cells["R5"].Value = "Total Lab Cost";
                    ws.Cells["S5"].Value = "Mat Cost";
                    ws.Cells["T5"].Value = "Total Cost";
                    ws.Cells["U5"].Value = "Qty";
                    ws.Cells["V5"].Value = "Description";
                    ws.Cells["W5"].Value = "Approver Description";
                    ws.Cells["A5:W5"].Style.Font.Bold = true;
                    ws.Cells["A5:W5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["A5:W5"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    int rw = 5, rw_begin = 0, rw_end = 0;
                    double num = 0.00;

                    #endregion
                    for (int k = 0; k < dtv.Rows.Count; k++)
                    {
                        DataTable dt = GetMNRCostDtls(dtv.Rows[k]["ID"].ToString());

                        if (dt != null && dt.Rows.Count > 0)
                        {
                            Server.ScriptTimeout = 3600;
                            rw++;
                            rw_begin = rw;
                            num = 0.00;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                ws.Cells["A" + rw].Value = dtv.Rows[k]["ReqRefNo"].ToString();
                                ws.Cells["B" + rw].Value = dtv.Rows[k]["CntrNo"].ToString();

                                ws.Cells["C" + rw].Value = "";
                                ws.Cells["D" + rw].Value = dt.Rows[i]["ComponentCode"];
                                ws.Cells["E" + rw].Value = dt.Rows[i]["LocCode"].ToString();
                                ws.Cells["F" + rw].Value = dt.Rows[i]["DamageCode"].ToString();
                                ws.Cells["G" + rw].Value = dt.Rows[i]["RepairCode"].ToString();
                                ws.Cells["H" + rw].Value ="";
                                ws.Cells["I" + rw].Value = dt.Rows[i]["Measurement"].ToString();
                                ws.Cells["J" + rw].Value = dt.Rows[i]["munit"].ToString();

                                double.TryParse(dt.Rows[i]["LabourHrs"].ToString(), out num);
                                ws.Cells["K" + rw].Value = num;
                                ws.Cells["K" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["EstLabcost"].ToString(), out num);
                                ws.Cells["L" + rw].Value = num;
                                ws.Cells["L" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["EstTotalLabCost"].ToString(), out num);
                                ws.Cells["M" + rw].Value = num;
                                ws.Cells["M" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["EstMatcost"].ToString(), out num);
                                ws.Cells["N" + rw].Value = num;
                                ws.Cells["N" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["EstTotalCost"].ToString(), out num);
                                ws.Cells["O" + rw].Value = num;
                                ws.Cells["O" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["AppvdLabHr"].ToString(), out num);
                                ws.Cells["P" + rw].Value = num;
                                ws.Cells["P" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["AppLabcost"].ToString(), out num);
                                ws.Cells["Q" + rw].Value = num;
                                ws.Cells["Q" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["AppTotalLabCost"].ToString(), out num);
                                ws.Cells["R" + rw].Value = num;
                                ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["AppMatcost"].ToString(), out num);
                                ws.Cells["S" + rw].Value = num;
                                ws.Cells["S" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["AppTotalCost"].ToString(), out num);
                                ws.Cells["T" + rw].Value = num;
                                ws.Cells["T" + rw].Style.Numberformat.Format = "#,##0.00";

                                double.TryParse(dt.Rows[i]["qty"].ToString(), out num);
                                ws.Cells["U" + rw].Value = num;
                                ws.Cells["U" + rw].Style.Numberformat.Format = "#";

                                ws.Cells["V" + rw].Value = dt.Rows[i]["description"].ToString();
                                ws.Cells["W" + rw].Value = dt.Rows[i]["Appdescription"].ToString();
                                rw++;

                            }
                            rw_end = rw;

                            #region RefNo wise sub total
                            ExcelRange rng = ws.Cells["A" + rw_end + ":J" + rw_end];
                            rng.Value = "SubTotal of Ref No. : " + dtv.Rows[k]["ReqRefNo"].ToString();
                            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            rng.Merge = true;

                            ws.Cells["A" + rw_end + ":W" + rw_end].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells["A" + rw_end + ":W" + rw_end].Style.Fill.BackgroundColor.SetColor(Color.White);
                            ws.Cells["A" + rw_end + ":W" + rw_end].Style.Font.Bold = true;
                            ws.Cells["A" + rw_end + ":W" + rw_end].Style.Font.Color.SetColor(Color.MediumVioletRed);

                            //Estimated-Lab Hr
                            ws.Cells["K" + rw_end].Formula = string.Format("=SUM(K{0}:K{1})", rw_begin, (rw_end - 1));
                            ws.Cells["K" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Estimated-Lab Cost
                            ws.Cells["L" + rw_end].Formula = string.Format("=SUM(L{0}:L{1})", rw_begin, (rw_end - 1));
                            ws.Cells["L" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Estimated-Total Lab Cost
                            ws.Cells["M" + rw_end].Formula = string.Format("=SUM(M{0}:M{1})", rw_begin, (rw_end - 1));
                            ws.Cells["M" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Estimated-Mat Cost 
                            ws.Cells["N" + rw_end].Formula = string.Format("=SUM(N{0}:N{1})", rw_begin, (rw_end - 1));
                            ws.Cells["N" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Estimated-Total Cost
                            ws.Cells["O" + rw_end].Formula = string.Format("=SUM(O{0}:O{1})", rw_begin, (rw_end - 1));
                            ws.Cells["O" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Approved-Lab Hr
                            ws.Cells["P" + rw_end].Formula = string.Format("=SUM(P{0}:P{1})", rw_begin, (rw_end - 1));
                            ws.Cells["P" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Approved-Lab Cost
                            ws.Cells["Q" + rw_end].Formula = string.Format("=SUM(Q{0}:Q{1})", rw_begin, (rw_end - 1));
                            ws.Cells["Q" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Approved-Total Lab cost
                            ws.Cells["R" + rw_end].Formula = string.Format("=SUM(R{0}:R{1})", rw_begin, (rw_end - 1));
                            ws.Cells["R" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Approved-Mat cost
                            ws.Cells["S" + rw_end].Formula = string.Format("=SUM(S{0}:S{1})", rw_begin, (rw_end - 1));
                            ws.Cells["S" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            //Approved-Total Cost
                            ws.Cells["T" + rw_end].Formula = string.Format("=SUM(T{0}:T{1})", rw_begin, (rw_end - 1));
                            ws.Cells["T" + rw_end].Style.Numberformat.Format = "#,##0.00";

                            #endregion

                        }
                    }
                    ws.Cells["A4:W" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A4:W" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A4:W" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A4:W" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    //Estimated 
                    ws.Cells["K4:O" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["K4:O" + rw].Style.Fill.BackgroundColor.SetColor(Color.DarkGoldenrod);

                    //Approved 
                    ws.Cells["P4:T" + rw].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["P4:T" + rw].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                    ws.Cells[1, 1, rw, 42].AutoFitColumns();

                    pck.SaveAs(Response.OutputStream);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;  filename=MNRReport.xlsx");
                    Response.End();
                }
               
            }
            catch (Exception ex)
            {

            }
        }



        public DataTable GetMNRDetailSearchValues(string DepotID, string ReqRefNo, string CntrID, string DtFrom, string DtTo, string Status)
        {
            string strWhere = "";
            string _Query = "Select R.ID,R.ReqRefNo,replace(convert(NVARCHAR, R.DtRequest, 103), ' ', '-') as DtReq,replace(convert(NVARCHAR,R.DtApproved, 103), ' ', '-') as DtApprv,C.CntrNo,T.TYPE +'-' + T.SIZE as TypeSize,case when  R.IsVendor=0 then (select top 1 AgencyName from NVO_AgencyMaster where ID =R.RequestedByID) ELSE (select top 1 CustomerName from NVO_CustomerMaster where ID = R.RequestedByID AND CustomerType = 44) END AS RequestedBy,GM.GeneralName As Status, rt.LabCost,rt.MatCost,rt.EstTotalCost,rt.ApprTotalCost," +
                "  (SELECT TOP 1 CurrencyCode FROM NVO_CurrencyMaster WHERE ID=R.CurrID) AS CURRENCY from NVO_MNRCntrRepairReq R Inner join NVO_Containers C ON C.ID = R.CntrID " +
               " Inner join NVO_tblCntrTypes T ON T.ID = C.TypeID " +
               " Inner join NVO_GeneralMaster GM ON GM.ID = R.Status AND GM.SeqNo = 36 " +
             " OUTER APPLY(SELECT cast(round(Sum(MaterialCostx100 / 100.00), 2, 0) as decimal(18, 2)) AS MatCost, cast(round(Sum(TotalLabourCostx100 / 100.00), 2, 0) as decimal(18, 2)) AS LabCost, cast(round(Sum(MaterialCostx100 / 100.00) + Sum(TotalLabourCostx100 / 100.00), 2, 0) as decimal(18, 2)) as EstTotalCost,   cast(round(Sum(AppvdMaterialCostx100 / 100.00) + Sum(TotalLabourCostx100 / 100.00), 2, 0) as decimal(18, 2)) as ApprTotalCost from NVO_MNRCntrRepairReqDtls where RepairReqID = r.ID) as rt ";

            if (ReqRefNo != "" && ReqRefNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where R.ReqRefNo like '%" + ReqRefNo + "%'";
                else
                    strWhere += " and R.ReqRefNo like '%" + ReqRefNo + "%'";

            if (DepotID != "" && DepotID != "null" && DepotID != "?" && DepotID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where R.DepotID=" + DepotID;
                else
                    strWhere += " and R.DepotID =" + DepotID;


            if (CntrID != "" && CntrID != "null" && CntrID != "?" && CntrID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where R.CntrID=" + CntrID;
                else
                    strWhere += " and R.CntrID =" + CntrID;

            if (Status != "" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where R.Status=" + Status;
                else
                    strWhere += " and R.Status =" + Status;

            switch (Status)
            {
                case "120":
                    if (DtFrom != "")
                        strWhere += _Query + " where R.DtRepaired >='" + DtFrom + "'";
                    if (DtTo != "")
                        strWhere += " and R.DtRepaired <='" + DtTo + "'";
                    break;
                case "117":
                case "119":
                    if (DtFrom != "")
                        strWhere += _Query + "where R.DtApproved >='" + DtFrom + "'";
                    if (DtTo != "")
                        strWhere += " and R.DtApproved <='" + DtTo + "'";
                    break;
                default:
                    if (DtFrom != "")
                        strWhere += _Query + " where R.DtRequest >='" + DtFrom + "'";
                    if (DtTo != "")
                        strWhere += " and R.DtRequest <='" + DtTo + "'";
                    break;
            }

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
        public DataTable GetMNRCostDtls(string RepairReqID)
        {
            
            string _Query = "Select  MD.DID,(select top 1 ComponentCode from NVO_MNRComponentMaster  where ID= MD.ComponentID) AS ComponentCode,(select top 1 LocationCode from NVO_MNRLocationMaster where ID = MD.locCodeID) AS LocCode ," +
             " (select top 1 DamageCode from NVO_MNRDamageMaster  where ID = MD.DamageTypeID) AS DamageCode,(select top 1 RepairCode from NVO_MNRRepairMaster  where ID = MD.RepairTypeID) AS RepairCode, MD.Measurement,MD.MeasureUnit as Munit,CONVERT(money, MD.LabourHrs) LabourHrs, "+
               " CONVERT(money, MD.AppvdLabHr) AppvdLabHr, CONVERT(money, MD.LabourCostx100 / 100.00) as EstLabcost," +
              " CONVERT(money, MD.AppvdLabCostx100 / 100.00) as AppLabcost, CONVERT(money, MD.TotalLabourCostx100 / 100.00) as EstTotalLabCost, "+
               " CONVERT(money, MD.AppvdTotalLabCostx100 / 100.00) as AppTotalLabCost, CONVERT(money, MD.MaterialCostx100 / 100.00) AS EstMatcost, "+
               " CONVERT(money, MD.AppvdMaterialCostx100 / 100.00) AS AppMatcost, MD.CostToID,EstTotalCostX100 / 100 AS EstTotalCost, " +
               " MD.AppvdTotalCostx100 / 100.00 AS AppTotalCost, MD.Qty, MD.Description,MD.AppDescription "+
               "  from NVO_MNRCntrRepairReqDtls MD WHERE MD.RepairReqID = " + RepairReqID;




            return Manag.GetViewData(_Query, "");
        }


        public void MNRWeeklyReport( string DtFrom, string DtTo, string Status , string User)
        {

            try
            {
                string DtFromV = "", DtToV = "";
                DateTime outDate = new DateTime();
                if (DateTime.TryParse(DtFrom, out outDate))
                    DtFromV = outDate.ToString("MM/dd/yyyy");
                if (DateTime.TryParse(DtTo, out outDate))
                    DtToV = outDate.ToString("MM/dd/yyyy");

                DataTable dtv = GetMNRWeeklyReport( DtFrom, DtTo, Status);
                if (dtv.Rows.Count > 0)
                {

                    ExcelPackage pck = new ExcelPackage();

                    var ws = pck.Workbook.Worksheets.Add("MNRRepair");

                    ws.Cells["A2"].Value = "MNR Report List";
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
                    ws.Cells["B7"].Value = "REQUEST#";
                    ws.Cells["C7"].Value = "REQUEST DATE";
                    ws.Cells["E7"].Value = "CONTAINER#";
                    ws.Cells["D7"].Value = "SIZE & TYPE";
                    ws.Cells["F7"].Value = "POL";
                    ws.Cells["G7"].Value = "POD";
                    ws.Cells["H7"].Value = "GATE IN DATE";
                    ws.Cells["I7"].Value = "EOR DATE";
                    ws.Cells["J7"].Value = "APPROVED DATE";
                    ws.Cells["K7"].Value = "AV DATE";
                    ws.Cells["L7"].Value = "LINE";
                    ws.Cells["M7"].Value = "EOR COST";
                    ws.Cells["N7"].Value = "APPROVED COST";
                    ws.Cells["O7"].Value = "ACCOUNTABILITY - CUSTOMER";
                    ws.Cells["P7"].Value = "ACCOUNTABILITY - LINE";
                    ws.Cells["Q7"].Value = "REMARKS";
                    r = ws.Cells["A7:Q7"];
                    r.Style.Font.Bold = true;
                    r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                    int sl = 1;

                    int rw = 8;
                    double amt = 0.00;
                    double amt1 = 0.00;
                    for (int i = 0; i < dtv.Rows.Count; i++)
                    {
                        ws.Cells["A" + rw].Value = sl;
                        ws.Cells["B" + rw].Value = dtv.Rows[i]["ReqRefNo"].ToString();
                        ws.Cells["C" + rw].Value = dtv.Rows[i]["DtReq"].ToString();
                        ws.Cells["D" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                        ws.Cells["E" + rw].Value = dtv.Rows[i]["TypeSize"].ToString();
                        ws.Cells["F" + rw].Value = dtv.Rows[i]["POL"].ToString();
                        ws.Cells["G" + rw].Value = dtv.Rows[i]["POD"].ToString();
                        ws.Cells["H" + rw].Value = dtv.Rows[i]["GateINDate"].ToString();
                        ws.Cells["I" + rw].Value = dtv.Rows[i]["DtReq"].ToString();
                        ws.Cells["J" + rw].Value = dtv.Rows[i]["DtApprv"].ToString();
                        ws.Cells["K" + rw].Value = dtv.Rows[i]["AVDate"].ToString();
                        ws.Cells["L" + rw].Value = "";
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["EstTotalCost"].ToString();
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["ApprTotalCost"].ToString();

                        double.TryParse(dtv.Rows[i]["EstTotalCost"].ToString(), out amt);
                        ws.Cells["M" + rw].Value = amt;
                        ws.Cells["M" + rw].Style.Numberformat.Format = "#,##0.00";

                        double.TryParse(dtv.Rows[i]["ApprTotalCost"].ToString(), out amt1);
                        ws.Cells["N" + rw].Value = amt1;
                        ws.Cells["N" + rw].Style.Numberformat.Format = "#,##0.00";


                        ws.Cells["O" + rw].Value = dtv.Rows[i]["CusCost"].ToString();
                        ws.Cells["P" + rw].Value = dtv.Rows[i]["LineCost"].ToString();

                        double.TryParse(dtv.Rows[i]["CusCost"].ToString(), out amt);
                        ws.Cells["O" + rw].Value = amt;
                        ws.Cells["O" + rw].Style.Numberformat.Format = "#,##0.00";

                        double.TryParse(dtv.Rows[i]["LineCost"].ToString(), out amt1);
                        ws.Cells["P" + rw].Value = amt1;
                        ws.Cells["P" + rw].Style.Numberformat.Format = "#,##0.00";

                        ws.Cells["Q" + rw].Value ="";
                        //// Approved Cost *50 PERCENT / 100

                        //ws.Cells["L" + rw].Formula = " = " + ws.Cells["K" + rw].Address + "  + (" + ws.Cells["K" + rw].Address + " * 50 / 100  )";
                        //ws.Cells["K" + rw].Style.Numberformat.Format = "#,##0.00";

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
                    Response.AddHeader("content-disposition", "attachment;  filename=MNRReport.xlsx");
                    Response.End();

                }

            }
            catch (Exception ex)
            {

            }
        }

        public DataTable GetMNRWeeklyReport( string DtFrom, string DtTo, string Status)
        {
            string strWhere = "";
            string _Query = "Select R.ID,R.ReqRefNo,replace(convert(NVARCHAR, R.DtRequest, 103), ' ', '-') as DtReq,replace(convert(NVARCHAR, R.DtApproved, 103), ' ', '-') as DtApprv,C.CntrNo,T.TYPE + '-' + T.SIZE as TypeSize, "+
         " isnull(replace(convert(NVARCHAR, (select top 1 Dtmovement from NVO_ContainerTxns Where ContainerID = CntrID and StatusCode = 'MA'),103), '', '-'),'') As GateInDate, isnull(replace(convert(NVARCHAR, (select top 1 Dtmovement from NVO_ContainerTxns Where ContainerID = CntrID and StatusCode = 'AV'), 103), '', '-' ),'') As AVDate, isnull((Select top 1 POL  from NVO_Booking WHERE ID = R.BLID ),'') as POL, "+
       " isnull((Select top 1 POD  from NVO_Booking WHERE ID = R.BLID ),'') as POD,  case when R.IsVendor = 0 then(select top 1 AgencyName from NVO_AgencyMaster where ID = R.RequestedByID) ELSE(select top 1 CustomerName from NVO_CustomerMaster where ID = R.RequestedByID "+
        "AND CustomerType = 44) END AS RequestedBy,  GM.GeneralName As Status, rt.LabCost,rt.MatCost,rt.EstTotalCost,TotalCostAppr.ApprTotalCost,CosToCus.CusCost, CosToLine.LineCost, (SELECT TOP 1 CurrencyCode FROM NVO_CurrencyMaster WHERE ID = R.CurrID) AS CURRENCY from NVO_MNRCntrRepairReq R " + 
        " Inner join NVO_Containers C ON C.ID = R.CntrID Inner join NVO_tblCntrTypes T ON T.ID = C.TypeID Inner join NVO_GeneralMaster GM ON GM.ID = R.Status AND GM.SeqNo = 36  OUTER APPLY(SELECT cast(round(Sum(MaterialCostx100 / 100.00), 2, 0)  "+
      " as decimal(18, 2)) AS MatCost, cast(round(Sum(TotalLabourCostx100), 2, 0) as decimal(18, 2)) AS LabCost, cast(round(Sum(MaterialCostx100) + Sum(TotalLabourCostx100), 2, 0) as decimal(18, 2)) as EstTotalCost from NVO_MNRCntrRepairReqDtls where RepairReqID = r.ID) as rt "+
      " outer apply(SELECT cast(round(Sum(AppvdMaterialCostx100) + Sum(TotalLabourCostx100), 2, 0) as decimal(18, 2)) as LineCost from NVO_MNRCntrRepairReqDtls where RepairReqID = r.ID and CostToID = 129) CosToLine outer apply(SELECT cast(round(Sum(AppvdMaterialCostx100) + Sum(TotalLabourCostx100), 2, 0) as decimal(18, 2)) as CusCost from NVO_MNRCntrRepairReqDtls where RepairReqID = r.ID and CostToID = 126) CosToCus outer apply(SELECT cast(round(Sum(AppvdMaterialCostx100) + Sum(TotalLabourCostx100), 2, 0) as decimal(18, 2)) as ApprTotalCost from NVO_MNRCntrRepairReqDtls where RepairReqID = r.ID and CostToID  in (126,124,129) ) TotalCostAppr ";

          

            if (Status != "" && Status != "null" && Status != "?" && Status != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where R.Status=" + Status;
                else
                    strWhere += " and R.Status =" + Status;

            switch (Status)
            {
                case "120":
                    if (DtFrom != "")
                        strWhere += _Query + " where R.DtRepaired >='" + DtFrom + "'";
                    if (DtTo != "")
                        strWhere += " and R.DtRepaired <='" + DtTo + "'";
                    break;
                case "117":
                case "119":
                    if (DtFrom != "")
                        strWhere += _Query + "where R.DtApproved >='" + DtFrom + "'";
                    if (DtTo != "")
                        strWhere += " and R.DtApproved <='" + DtTo + "'";
                    break;
                default:
                    if (DtFrom != "")
                        strWhere += _Query + " where R.DtRequest >='" + DtFrom + "'";
                    if (DtTo != "")
                        strWhere += " and R.DtRequest <='" + DtTo + "'";
                    break;
            }

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
    }
}