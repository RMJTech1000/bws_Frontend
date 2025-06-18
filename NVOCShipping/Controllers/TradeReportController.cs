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
    public class TradeReportController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: TradeReport
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult SlotContractReport()
        {
            return View();
        }

        public void SlotContractReportValues(string FeederID, string FromPort, string ToPort, string EffectDate, string CtryID, string User)
        {


            ExcelPackage pck = new ExcelPackage();
            DataTable dt = GetCountryValues(FeederID, FromPort, ToPort, EffectDate, CtryID);


            for (int J = 0; J < dt.Rows.Count; J++)
            {
                var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["CountryName"].ToString());

                ws.Cells["A2"].Value = "Slot Contract Report ";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:S2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                ws.Cells["A7"].Value = "User :";
                ws.Cells["A7"].Style.Font.Bold = true;
                ws.Cells["B7"].Value = User;
                ws.Cells["B7"].Style.Font.Bold = true;
                ws.Cells["C7"].Value = "Effective Date :";
                ws.Cells["C7"].Style.Font.Bold = true;
                ws.Cells["D7"].Value = System.DateTime.Today.Date.ToShortDateString();
                ws.Cells["D7"].Style.Font.Bold = true;
                //Record Headers

                ws.Cells["A11:A13"].Value = "S. No.";
                ws.Cells["A11:A13"].Merge = true;
                ws.Cells["A11:A13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["B11:B13"].Value = "CARRIER";
                ws.Cells["B11:B13"].Merge = true;
                ws.Cells["B11:B13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["C11:C13"].Value = "NEGOTIATION PARTY";
                ws.Cells["C11:C13"].Merge = true;
                ws.Cells["C11:C13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["C11:C13"].AutoFitColumns();

                ws.Cells["D11:D13"].Value = "POL";
                ws.Cells["D11:D13"].Merge = true;
                ws.Cells["D11:D13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["E11:E13"].Value = "POD";
                ws.Cells["E11:E13"].Merge = true;
                ws.Cells["E11:E13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["F11:F13"].Value = "TERM";
                ws.Cells["F11:F13"].Merge = true;
                ws.Cells["F11:F13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["G11:G13"].Value = "SLOT REFERENCE";
                ws.Cells["G11:G13"].Merge = true;
                ws.Cells["G11:G13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["G11:G13"].AutoFitColumns();

                ws.Cells["H11:H13"].Value = "ROUTE";
                ws.Cells["H11:H13"].Merge = true;
                ws.Cells["H11:H13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                ws.Cells["I11:P11"].Value = "SLOT RATES";
                ws.Cells["I11:P11"].Merge = true;
                ws.Cells["I11:P11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                ws.Cells["I12:J12"].Value = "LADEN-GENERAL";
                ws.Cells["I12:J12"].Merge = true;
                ws.Cells["I12:J12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["K12:L12"].Value = "LADEN-HAZ";
                ws.Cells["K12:L12"].Merge = true;
                ws.Cells["K12:L12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["M12:N12"].Value = "EMPTY";
                ws.Cells["M12:N12"].Merge = true;
                ws.Cells["M12:N12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["O12:P12"].Value = "REEFER";
                ws.Cells["O12:P12"].Merge = true;
                ws.Cells["O12:P12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells["I13"].Value = "20GP";
                ws.Cells["J13"].Value = "40GP/HC";
                ws.Cells["K13"].Value = "20GP";
                ws.Cells["L13"].Value = "40GP/HC";
                ws.Cells["M13"].Value = "20GP";
                ws.Cells["N13"].Value = "40GP/HC";
                ws.Cells["O13"].Value = "20RF";
                ws.Cells["P13"].Value = "40RF";

                ws.Cells["Q11:Q13"].Value = "EFFECTIVE DATE";
                ws.Cells["Q11:Q13"].Merge = true;
                ws.Cells["Q11:Q13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["Q11:Q13"].AutoFitColumns();

                ws.Cells["R11:R13"].Value = "EXPIRY DATE";
                ws.Cells["R11:R13"].Merge = true;
                ws.Cells["R11:R13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["R11:R13"].AutoFitColumns();

                ws.Cells["S11:S13"].Value = "REMARKS";
                ws.Cells["S11:S13"].Merge = true;
                ws.Cells["S11:S13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["S11:S13"].AutoFitColumns();

                r = ws.Cells["A11:S13"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                DataTable dtv = GetSlotReportValues(FeederID, FromPort, ToPort, EffectDate, dt.Rows[J]["CtryID"].ToString());
                if (dtv.Rows.Count > 0)
                {
                    int sl = 1;
                    int rw = 14;
                    for (int i = 0; i < dtv.Rows.Count; i++)
                    {
                        ws.Cells["A" + rw].Value = sl;
                        ws.Cells["B" + rw].Value = dtv.Rows[i]["Carrier"].ToString();

                        ws.Cells["C" + rw].Value = dtv.Rows[i]["NegotitionParty"].ToString();
                        ws.Cells["D" + rw].Value = dtv.Rows[i]["POL"].ToString();
                        ws.Cells["E" + rw].Value = dtv.Rows[i]["POD"].ToString();
                        ws.Cells["F" + rw].Value = dtv.Rows[i]["SlotTerm"].ToString();
                        ws.Cells["G" + rw].Value = dtv.Rows[i]["SlotRef"].ToString();
                        ws.Cells["H" + rw].Value = dtv.Rows[i]["RouteType"].ToString();
                        ws.Cells["I" + rw].Value = dtv.Rows[i]["LADGen20GP"];
                        ws.Cells["J" + rw].Value = dtv.Rows[i]["LADGen40GP40HC"];
                        ws.Cells["K" + rw].Value = dtv.Rows[i]["LADHAZ20GP"];
                        ws.Cells["L" + rw].Value = dtv.Rows[i]["LADHAZ40GP40HC"];
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["MTY20GP"];
                        ws.Cells["N" + rw].Value = dtv.Rows[i]["MTY40GP40HC"];
                        ws.Cells["O" + rw].Value = dtv.Rows[i]["MTYHAZ20GP"];
                        ws.Cells["P" + rw].Value = dtv.Rows[i]["MTYHAZ40GP40HC"];
                        ws.Cells["Q" + rw].Value = dtv.Rows[i]["ValidFrom"].ToString();
                        ws.Cells["R" + rw].Value = dtv.Rows[i]["ExpiryDate"].ToString();
                        ws.Cells["S" + rw].Value = dtv.Rows[i]["Remarks"].ToString();
                        sl++;
                        rw += 1;
                    }

                    rw -= 1;


                    ws.Cells["A11:S13" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A11:S13" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A11:S13" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A11:S13" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    ws.Cells[1, 1, rw, 50].AutoFitColumns();
                }
            }
                pck.SaveAs(Response.OutputStream);
            
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=SlotContractReport.xlsx");
            Response.End();

        }

        public DataTable GetSlotReportValues(string FeederID, string FromPort, string ToPort, string EffectDate, string CtryID)
        {
            string strWhere = "";
            string _Query = " select * from NVO_VIEWSlotContractReport  where Status='Active Rates' ";


            if (EffectDate != "" && EffectDate != "0" && EffectDate != null && EffectDate != "?" && EffectDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and ValidFromDt <='" + EffectDate + "'";
                else
                    strWhere += " and ValidFromDt <='" + EffectDate + "'";




            if (FeederID != "" && FeederID != "0" && FeederID != "null" && FeederID != "?")

                if (strWhere == "")
                    strWhere += _Query + " and SlotOperatorID=" + FeederID;
                else
                    strWhere += " and SlotOperatorID = " + FeederID;


            if (CtryID != "" && CtryID != "0" && CtryID != "null" && CtryID != "?")

                if (strWhere == "")
                    strWhere += _Query + " AND CtryID=" + CtryID;
                else
                    strWhere += " and CtryID = " + CtryID;

            if (FromPort != "" && FromPort != "0" && FromPort != "null" && FromPort != "?")

                if (strWhere == "")
                    strWhere += _Query + " and POLID=" + FromPort;
                else
                    strWhere += " and POLID = " + FromPort;

            if (ToPort != "" && ToPort != "0" && ToPort != "null" && ToPort != "?")

                if (strWhere == "")
                    strWhere += _Query + " and PODID=" + ToPort;
                else
                    strWhere += " and PODID = " + ToPort;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere + " order By SlotOperatorID,ID DESC ", "");
        }

        public DataTable GetCountryValues(string FeederID, string FromPort, string ToPort, string EffectDate, string CtryID)
        {
            string strWhere = "";

            string _Query = "  select Distinct NVO_CountryMaster.ID AS CtryID,NVO_CountryMaster.CountryName from" +
                " NVO_SLOTMaster SM  inner Join NVO_PortMainMaster on NVO_PortMainMaster.ID = SM.POL inner Join " +
               " NVO_CountryMaster on NVO_CountryMaster.ID = NVO_PortMainMaster.CountryID ";

            if (EffectDate != "" && EffectDate != "0" && EffectDate != null && EffectDate != "?" && EffectDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where SM.ValidFrom <='" + EffectDate + "'";
                else
                    strWhere += " and SM.ValidFrom <='" + EffectDate + "'";

            if (FeederID != "" && FeederID != "0" && FeederID != "null" && FeederID != "?")

                if (strWhere == "")
                    strWhere += _Query + " where SM.SlotOperator=" + FeederID;
                else
                    strWhere += " and SM.SlotOperator = " + FeederID;


            if (CtryID != "" && CtryID != "0" && CtryID != "null" && CtryID != "?")

                if (strWhere == "")
                    strWhere += _Query + " where NVO_CountryMaster.ID=" + CtryID;
                else
                    strWhere += " and NVO_CountryMaster.ID = " + CtryID;

            if (FromPort != "" && FromPort != "0" && FromPort != "null" && FromPort != "?")

                if (strWhere == "")
                    strWhere += _Query + " where SM.POL=" + FromPort;
                else
                    strWhere += " and SM.POL = " + FromPort;

            if (ToPort != "" && ToPort != "0" && ToPort != "null" && ToPort != "?")

                if (strWhere == "")
                    strWhere += _Query + " where SM.POD=" + ToPort;
                else
                    strWhere += " and SM.POD = " + ToPort;

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere + " ", "");
        }


    }
}