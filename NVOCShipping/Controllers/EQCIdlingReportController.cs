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
    public class EQCIdlingReportController : Controller
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
        public void EQCIdlingReportNew(string Agency, string Type, string User)
        {
            DataTable dtv = GetEQCIdlingReportNew(Agency, Type);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                #region 1st page
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

               

                var ws = pck.Workbook.Worksheets.Add("EQCIdlingSummary");

                Color colFromHex;

                ws.Cells["A1"].Value = "EQC IDLING REPORT";
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

                //ws.Cells["A6"].Value = "Status :";
                //ws.Cells["B6"].Value = cntrStatus;
                //ws.Cells["A6"].Style.Font.Bold = true;
                //ws.Cells["B6"].Style.Font.Bold = true;
                // e_msg += "<br>" + cntrStatus;

                if (Session["UserName"] == null)
                    Session["UserName1"] = "SYS GEN";
                else
                    Session["UserName1"] = Session["UserName"];

                rw = 8;
                string cntrStatus = "";
                if (Type == "1")
                    cntrStatus = "'MS','FL','FB','FI','FC'";
                if (Type == "2")
                    cntrStatus = "'FV','FU','FVICD','FI','DV','MA'";
                if (Type == "3")
                    cntrStatus = "'TZ','TZFB'";
                if (Type == "4")
                    cntrStatus = "'DL','UR'";
                if (Type == "5")
                    cntrStatus = "'AV'";

                if (Type == "?" )
                    cntrStatus = "'MS','FL','FB','FI','FC','FV','FU','FVICD','FI','DV','MA','TZ','TZFB','DL','UR','AV'";

                var Cntrs = cntrStatus.Split(',');
                DataTable _dtValue = null;
                string country1 = "", GeoLocation = "";
                string mainsector = ""; string mainsectorfr = ""; string mainsectorOt = ""; string mainsectortnk = ""; string mainsectorrf = "";
                int country_start = 0, country_end = 0, GeoLoc_start = 0, GeoLoc_end = 0;
                int mainsector_start = 0, mainsector_end = 0;

                for (int y = 0; y < Cntrs.Length; y++)
                {
                    _dtValue = _dtAgewiseRepValue(cntrStatus, Agency);
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
                ws.Cells["D" + rw].Value = "Agency";
                ws.Cells["E" + rw].Value = "Shipment Type";
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
                    ws.Cells["D" + rw].Value = dtNew.Rows[i]["Agency"].ToString();
                    ws.Cells["E" + rw].Value = dtNew.Rows[i]["ShipmentType"].ToString();
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

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#2874A6");
                rng = ws.Cells["G" + (rw_bgn - 1) + ":J" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#9B59B6");
                rng = ws.Cells["K" + (rw_bgn - 2) + ":N" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#5D6D7E");
                rng = ws.Cells["O" + (rw_bgn - 2) + ":R" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#F0B27A");
                rng = ws.Cells["S" + (rw_bgn - 2) + ":V" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#ABEBC6");
                rng = ws.Cells["W" + (rw_bgn - 2) + ":Z" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#EDBB99");
                rng = ws.Cells["AA" + (rw_bgn - 2) + ":AD" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                //ws.Cells["A"+rw_bgn + ":Z" + rw].Style.Numberformat.Format = "#,##0";
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
                ws.Cells["D" + rw].Value = "Agency";
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
                    ws.Cells["D" + rw].Value = dtNew.Rows[i]["Agency"].ToString();
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
                            ws.Cells["V" + rw_end].Formula = string.Format("=SUM(V{0}:V{1})", rw_bgn_sub3, (rw_end - 1));
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

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#2874A6");
                rng = ws.Cells["E" + (rw_bgn - 1) + ":G" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#9B59B6");
                rng = ws.Cells["H" + (rw_bgn - 2) + ":J" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#5D6D7E");
                rng = ws.Cells["K" + (rw_bgn - 2) + ":M" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#F0B27A");
                rng = ws.Cells["N" + (rw_bgn - 2) + ":P" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#ABEBC6");
                rng = ws.Cells["Q" + (rw_bgn - 2) + ":S" + rw];
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(colFromHex);

                colFromHex = System.Drawing.ColorTranslator.FromHtml("#EDBB99");
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

                //#region 2ND SHEET
                //ws = pck.Workbook.Worksheets.Add("IdlingBreakUp");
                //ws.Cells["A2"].Value = "Idling Detail Report List";
                //ws.Cells["A2"].Style.Font.Bold = true;
                //ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //r = ws.Cells["A2:X2"];
                //r.Merge = true;
                //r.Style.Font.Size = 12;
                //r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                //ws.Cells["A4"].Value = "User :";
                //ws.Cells["A4"].Style.Font.Bold = true;
                //ws.Cells["B4"].Value = User;
                //ws.Cells["B4"].Style.Font.Bold = true;
                //ws.Cells["C4"].Value = "Date :";
                //ws.Cells["C4"].Style.Font.Bold = true;
                //ws.Cells["D4"].Value = System.DateTime.Today.Date.ToShortDateString();
                //ws.Cells["D4"].Style.Font.Bold = true;

                ////Record Headers


                //ws.Cells["A7"].Value = "S. No.";
                //ws.Cells["B7"].Value = "Type";
                //ws.Cells["C7"].Value = "Country";
                //ws.Cells["D7"].Value = "Geo Location";
                //ws.Cells["E7"].Value = "Agency";
                //ws.Cells["F7"].Value = "Container No";
                //ws.Cells["G7"].Value = "Container Type";
                //ws.Cells["H7"].Value = "StatusCode";
                //ws.Cells["I7"].Value = "Date Movement";
                //ws.Cells["J7"].Value = "Ageing";
                //ws.Cells["K7"].Value = "Location";
                //ws.Cells["L7"].Value = "Vessel/Voyage";
                //ws.Cells["M7"].Value = "Transit";
                //ws.Cells["N7"].Value = "Depot";
                //ws.Cells["O7"].Value = "BLNumber";
                //ws.Cells["P7"].Value = "Orgin";
                //ws.Cells["Q7"].Value = "POL";
                //ws.Cells["R7"].Value = "POD";
                //ws.Cells["S7"].Value = "FPOD";
                //ws.Cells["T7"].Value = "Customer";
                //ws.Cells["U7"].Value = "Vendor";
                //ws.Cells["V7"].Value = "Status";
                //ws.Cells["W7"].Value = "Created By";
                //ws.Cells["X7"].Value = "Created On";
                //ws.Cells["Y7"].Value = "Container Owner";
                //ws.Cells["Z7"].Value = "Leasing Partner";
                //ws.Cells["AA7"].Value = "Leasing Term";
                //r = ws.Cells["A7:AA7"];
                //r.Style.Font.Bold = true;
                //r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                //int slno = 1;
                //int row = 8;

                //for (int i = 0; i < dtv.Rows.Count; i++)
                //{
                //    ws.Cells["A" + row].Value = slno;
                //    ws.Cells["B" + row].Value = "";
                //    ws.Cells["C" + row].Value = "";
                //    ws.Cells["C" + row].Value = "";
                //    ws.Cells["D" + row].Value = "";
                //    ws.Cells["E" + row].Value = "";
                //    ws.Cells["F" + row].Value = dtv.Rows[i]["CntrNo"].ToString();
                //    ws.Cells["G" + row].Value = dtv.Rows[i]["CntrType"].ToString();
                //    ws.Cells["H" + row].Value = dtv.Rows[i]["StatusCode"].ToString();
                //    ws.Cells["I" + row].Value = dtv.Rows[i]["DtMovement"].ToString();
                //    ws.Cells["J" + row].Value = "";
                //    ws.Cells["K" + row].Value = "";
                //    ws.Cells["L" + row].Value = "";
                //    ws.Cells["M" + row].Value = "";
                //    ws.Cells["N" + row].Value = "";
                //    ws.Cells["O" + row].Value = "";
                //    ws.Cells["P" + row].Value = "";
                //    ws.Cells["Q" + row].Value = "";
                //    ws.Cells["R" + row].Value = "";
                //    ws.Cells["S" + row].Value = "";
                //    ws.Cells["T" + row].Value = "";
                //    ws.Cells["U" + row].Value = "";
                //    ws.Cells["V" + row].Value = "";
                //    ws.Cells["W" + row].Value = "";
                //    ws.Cells["X" + row].Value = "";
                //    ws.Cells["Y" + row].Value = "";
                //    ws.Cells["Z" + row].Value = "";
                //    ws.Cells["AA" + row].Value = "";
                //    slno++;
                //    row += 1;
                //}

                //row -= 1;

                //ws.Cells["A7:AA" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                //ws.Cells["A7:AA" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                //ws.Cells["A7:AA" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                //ws.Cells["A7:AA" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                //ws.Cells[1, 1, row, 24].AutoFitColumns();
                //#endregion

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=EQCIdlingReport.xlsx");
                Response.End();

            }

        }

        public DataTable _dtAgewiseRepValue(string CntrStCode, string Agency)
        {
            string strWhere = "";

            string _Query = "SELECT CountryName,GeoLocName,StatusCode,Agency, " +
                " case when StatusCode  in ('MS','FL','FB','FI','FC') then 'EXPORT'  when StatusCode  in ('FV', 'FU', 'FVICD', 'FI', 'DV', 'MA') then 'IMPORT' "+
                 " when StatusCode  in ('TZ', 'TZFB') then 'TRANSHIPMENT' when StatusCode  in ('DL', 'UR') then 'DAMAGE' " +
                 " when StatusCode  in ('AV') then 'AVAILABLE' END AS ShipmentType ";



                _Query += ",SUM(ISNULL(Dry20L7,0)) Dry20L7, SUM(ISNULL(Dry40L7,0)) Dry40L7, SUM(ISNULL(DryHQL7,0)) DryHQL7,  " +
                                          " SUM(ISNULL(Dry208To14,0)) Dry208To14, SUM(ISNULL(Dry408To14,0)) Dry408To14, SUM(ISNULL(DryHQ8To14,0)) DryHQ8To14,  " +
                                          " SUM(ISNULL(Dry2015To29,0)) Dry2015To29, SUM(ISNULL(Dry4015To29,0)) Dry4015To29, " +
                                          " SUM(ISNULL(DryHQ15To29,0)) DryHQ15To29, SUM(ISNULL(Dry2030To59,0)) Dry2030To59, SUM(ISNULL(Dry4030To59,0)) Dry4030To59, " +
                                          " SUM(ISNULL(DryHQ30To59,0)) DryHQ30To59, SUM(ISNULL(Dry20G59,0)) Dry20G59,  " +
                                          " SUM(ISNULL(Dry40G59,0)) Dry40G59,SUM(ISNULL(DryHQG59,0)) DryHQG59, " +
                                          " SUM(ISNULL(Dry20HQL7,0)) Dry20HQL7,  " +
                                          " SUM(ISNULL(Dry20HQ8To14,0)) Dry20HQ8To14,SUM(ISNULL(Dry20HQ15To29,0)) Dry20HQ15To29, " +
                                          " SUM(ISNULL(Dry20HQ30To59,0)) Dry20HQ30To59, SUM(ISNULL(Dry20HQG59,0)) Dry20HQG59 ";

                _Query += ",SUM(ISNULL(RF20L7,0)) RF20L7,SUM(ISNULL(RF40L7,0)) RF40L7,SUM(ISNULL(RFHQL7,0)) RFHQL7,SUM(ISNULL(RF208To14,0)) RF208To14,SUM(ISNULL(RF408To14,0)) RF408To14," +
                                      " SUM(ISNULL(RFHQ8To14,0)) RFHQ8To14,SUM(ISNULL(RF2015To29,0)) RF2015To29,SUM(ISNULL(RF4015To29,0)) RF4015To29,SUM(ISNULL(RFHQ15To29,0)) RFHQ15To29, SUM(ISNULL(RF2030To59,0)) RF2030To59, " +
                                      " SUM(ISNULL(RF4030To59,0)) RF4030To59, SUM(ISNULL(RFHQ30To59,0)) RFHQ30To59, SUM(ISNULL(RF20G59,0)) RF20G59,SUM(ISNULL(RF40G59,0)) RF40G59, SUM(ISNULL(RFHQG59 ,0)) RFHQG59 ";


            if (CntrStCode != "")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinalIdling  where StatusCode  in (" + CntrStCode + ")";
                else
                    strWhere += " and StatusCode in (" + CntrStCode + ")";

            if (Agency != "" && Agency != "0" && Agency != "null" && Agency != "?")
                if (strWhere == "")
                    strWhere += _Query + " from VInvSnShotFinalIdling WHERE AgencyID=" + Agency;
                else
                    strWhere += " and AgencyID =" + Agency;

            if (strWhere == "")
                strWhere += _Query + " from VInvSnShotFinalIdling GROUP BY  CountryName, GeoLocName,Agency, StatusCode ORDER BY  CountryName, GeoLocName,StatusCode,ShipmentType,Agency ";

            else
                strWhere += "  GROUP BY  CountryName, GeoLocName,Agency, StatusCode ORDER BY CountryName, GeoLocName ,StatusCode,ShipmentType,Agency ";


            return Manag.GetViewData(strWhere, "");
        }
        public DataTable GetEQCIdlingReportNew(string Agency, string Type)
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
    }
    
}

