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
    public class EQCDetentionReportController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: EQCDetentionReport
        public ActionResult Index()
        {
            return View();
        }

        public void EQCImpDetentionReportValues(string DtFrom, string DtTo, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("ImportDetentionReport");

            ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:G2"];
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

            ws.Cells["A7:A10"].Value = "S. No.";
            ws.Cells["A7:A10"].Merge = true;
            ws.Cells["A7:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["B7:B10"].Value = "GEO-LOCATION";
            ws.Cells["B7:B10"].Merge = true;
            ws.Cells["B7:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["C7:C10"].Value = "AGENCY NAME";
            ws.Cells["C7:C10"].Merge = true;
            ws.Cells["C7:C0"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["D7:D10"].Value = "VESSEL/VOYAGE";
            ws.Cells["D7:D10"].Merge = true;
            ws.Cells["D7:D0"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["E7:E10"].Value = "POL";
            ws.Cells["E7:E10"].Merge = true;
            ws.Cells["E7:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["F7:F10"].Value = "POD";
            ws.Cells["F7:F10"].Merge = true;
            ws.Cells["F7:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["G7:G10"].Value = "CONSIGNEE";
            ws.Cells["G7:G10"].Merge = true;
            ws.Cells["G7:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["H7:H10"].Value = "BL NUMBER";
            ws.Cells["H7:H10"].Merge = true;
            ws.Cells["H7:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["I7:I10"].Value = "CONTAINER NO";
            ws.Cells["I7:I10"].Merge = true;
            ws.Cells["I7:I10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["J7:J10"].Value = "TYPE SIZE";
            ws.Cells["J7:J10"].Merge = true;
            ws.Cells["J7:J10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["K7:AA7"].Value = "With Free Time";
            ws.Cells["K7:AA7"].Merge = true;
            ws.Cells["K7:AA7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["AB7:AP7"].Value = "Without Free Time";
            ws.Cells["AB7:AP7"].Merge = true;
            ws.Cells["AB7:AP7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["K8"].Value = "Landing";
            ws.Cells["K9"].Value = " Date";
            ws.Cells["K10"].Value = "(FV)";

            ws.Cells["L8"].Value = "EMPTY IN";
            ws.Cells["L9"].Value = " Date";
            ws.Cells["L10"].Value = "(MA)";

            ws.Cells["M8"].Value = "FREE TIME";
            ws.Cells["M9"].Value = "(DAYS)";
            ws.Cells["M10"].Value = "";

            ws.Cells["N8"].Value = "DTN FROM";
            ws.Cells["N9"].Value = "A";
            ws.Cells["N10"].Value = "(FV  + F/Time)";


            ws.Cells["O8"].Value = "PLOT IN";
            ws.Cells["O9"].Value = "B";
            ws.Cells["O10"].Value = "MA";

            ws.Cells["P8"].Value = "DTN";
            ws.Cells["P9"].Value = "DAYS";
            ws.Cells["P10"].Value = "B-A";

            ws.Cells["Q8:W8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["Q8:W8"].Merge = true;
            ws.Cells["Q8:W8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["Q9"].Value = "SLAB1";
            ws.Cells["R9"].Value = "SLAB2";
            ws.Cells["S9"].Value = "SLAB3";
            ws.Cells["T9"].Value = "SLAB4";
            ws.Cells["U9"].Value = "SLAB5";
            ws.Cells["V9"].Value = "SLAB6";
            ws.Cells["W9"].Value = "TOTAL";

            ws.Cells["X8"].Value = "TARIFF";
            ws.Cells["X9"].Value = "CURRENCY";
            ws.Cells["X10"].Value = "";

            ws.Cells["Y8"].Value = "EX.RATE";
            ws.Cells["Y9"].Value = "";
            ws.Cells["Y10"].Value = "";

            ws.Cells["Z8:AA8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["Z8:AA8"].Merge = true;
            ws.Cells["Z8:AA8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["Z9"].Value = "USD";
            ws.Cells["AA9"].Value = "LCY";


            ws.Cells["AB8"].Value = "FREE TIME";
            ws.Cells["AB9"].Value = "(DAYS)";
            ws.Cells["AB10"].Value = "";

            ws.Cells["AC8"].Value = "DTN FROM";
            ws.Cells["AC9"].Value = "A";
            ws.Cells["AC10"].Value = "(FV  + F/Time)";


            ws.Cells["AD8"].Value = "PLOT IN";
            ws.Cells["AD9"].Value = "B";
            ws.Cells["AD10"].Value = "MA";

            ws.Cells["AE8"].Value = "DTN";
            ws.Cells["AE9"].Value = "DAYS";
            ws.Cells["AE10"].Value = "B-A";

            ws.Cells["AF8:AL8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["AF8:AL8"].Merge = true;
            ws.Cells["AF8:AL8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["AF9"].Value = "SLAB1";
            ws.Cells["AG9"].Value = "SLAB2";
            ws.Cells["AH9"].Value = "SLAB3";
            ws.Cells["AI9"].Value = "SLAB4";
            ws.Cells["AJ9"].Value = "SLAB5";
            ws.Cells["AK9"].Value = "SLAB6";
            ws.Cells["AL9"].Value = "TOTAL";

            ws.Cells["AM8"].Value = "TARIFF";
            ws.Cells["AM9"].Value = "CURRENCY";
            ws.Cells["AM10"].Value = "";

            ws.Cells["AN8"].Value = "EX.RATE";
            ws.Cells["AN9"].Value = "";
            ws.Cells["AN10"].Value = "";

            ws.Cells["AO8:AP8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["AO8:AP8"].Merge = true;
            ws.Cells["AO8:AP8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["AO9"].Value = "USD";
            ws.Cells["AP9"].Value = "LCY";


            ws.Cells["AQ7:AQ10"].Value = "SHIPPER";
            ws.Cells["AQ7:AQ10"].Merge = true;
            ws.Cells["AQ7:AQ10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["AR7:AR10"].Value = "INVOICE NO";
            ws.Cells["AR7:AR10"].Merge = true;
            ws.Cells["AR7:AR10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["AS7:AS10"].Value = "WORKSHEET USD";
            ws.Cells["AS7:AS10"].Merge = true;
            ws.Cells["AS7:AS10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["AT7:AT10"].Value = "WORKSHEET INR";
            ws.Cells["AT7:AT10"].Merge = true;
            ws.Cells["AT7:AT10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            //r = ws.Cells["A7:AA7"];
            //r.Style.Font.Bold = true;
            //r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            ExcelRange rng1 = ws.Cells["K7:AA10"];
            rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng1.Style.Fill.BackgroundColor.SetColor(Color.LightSalmon);

            ExcelRange rng2 = ws.Cells["AB7:AP10"];
            rng2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng2.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

            ws.Cells["A7:AT10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            int sl = 1;
            int rw = 11;
            string Currency = "USD";
            DataTable dtv = GetImpDetentionReport(DtFrom, DtTo);
            for( int i=0; i<dtv.Rows.Count; i++)
            {
                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = "";
                ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["POL"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["POD"].ToString();
                ws.Cells["G" + rw].Value = "";
                ws.Cells["H" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["FVDate"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["MADate"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["ImpFreeDays"].ToString();

                ws.Cells["N" + rw].Value = dtv.Rows[i]["FromDate"].ToString();

                ws.Cells["O" + rw].Value = dtv.Rows[i]["MADate"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["Daysv"].ToString();

                string FDatev = "";
                string TDatev = "";
                int RowsCount = 0;
                string ID = "";
                string LLimit = "";
                string ULimit = "";
                Currency = "";
                string Amount = "";
                double Days = 0;
                int _FreeDays = 0;
                decimal ExRate = 0;
                FDatev = "";
                TDatev = "";
                DateTime IncreamentalFromDt, FromDt;
                DateTime IncrementalToDt, ToDt;
                TimeSpan TMS; bool IsexceedFreedays = false; string WaiverQty = "0";
                DateTime.TryParse(dtv.Rows[i]["FVDatev"].ToString(), out IncreamentalFromDt);
                DateTime.TryParse(dtv.Rows[i]["FVDatev"].ToString(), out FromDt);
                DateTime.TryParse(dtv.Rows[i]["MADatev"].ToString().ToString(), out ToDt);
                WaiverQty = dtv.Rows[i]["ImpFreeDays"].ToString();
                int.TryParse(WaiverQty, out _FreeDays);
                DataTable _dtSlap = GetContractSlap(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                if(_dtSlap.Rows.Count >0)
                {
                    for (int y = 0; y < _dtSlap.Rows.Count; y++)
                    {
                        RowsCount = y;
                        
                        ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                        LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                        ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                        Currency = _dtSlap.Rows[y]["Currency"].ToString();
                        Amount = _dtSlap.Rows[y]["Amount"].ToString();
                        FDatev = IncreamentalFromDt.ToShortDateString();
                        IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                        TDatev = IncrementalToDt.ToShortDateString();

                        if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                        {
                            TDatev = ToDt.ToShortDateString();
                            TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                            Days = (int)((TMS.TotalDays + 1));

                            if (!IsexceedFreedays)
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                if (Days <= 0)
                                    Days = 1;
                            }
                        }
                        else
                        {
                            if (!IsexceedFreedays)
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));

                            }
                        }
                        IncreamentalFromDt = IncrementalToDt.AddDays(1);
                       
                        string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                        if (y == 0)
                        {
                            ws.Cells["q" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["q" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 1)
                        {
                            ws.Cells["R" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 2)
                        {
                            ws.Cells["S" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["s" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 3)
                        {
                            ws.Cells["T" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["T" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 4)
                        {
                            ws.Cells["U" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 5)
                        {
                            ws.Cells["V" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["V" + rw].Style.Numberformat.Format = "#,##0.00";
                        }

                        if (DateTime.Parse(TDatev) >= DateTime.Parse((ToDt.ToString("dd/MM/yyyy"))))
                        {
                            break;

                        }

                    }
                    if (RowsCount == 1)
                    {
                        ws.Cells["R" + rw].Value = 0.00;
                        ws.Cells["s" + rw].Value = 0.00;
                        ws.Cells["t" + rw].Value = 0.00;
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 2)
                    {
                        ws.Cells["s" + rw].Value = 0.00;
                        ws.Cells["t" + rw].Value = 0.00;
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 3)
                    {
                        ws.Cells["t" + rw].Value = 0.00;
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 4)
                    {
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 5)
                    {
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                }
               else
                {
                    ws.Cells["q" + rw].Value = 0.00;
                    ws.Cells["R" + rw].Value = 0.00;
                    ws.Cells["s" + rw].Value = 0.00;
                    ws.Cells["t" + rw].Value = 0.00;
                    ws.Cells["u" + rw].Value = 0.00;
                    ws.Cells["v" + rw].Value = 0.00;
                }

                string froMulaAddQV = string.Format("=SUM(Q" + rw + ":V" + rw + ")");
                ws.Cells["w" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["w" + rw].Style.Numberformat.Format = "#,##0.00";
                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["x" + rw].Value = _dtSlap.Rows[0]["Currency"].ToString();
                else
                    ws.Cells["x" + rw].Value = "NA";

                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["y" + rw].Value = _dtSlap.Rows[0]["ExRate"].ToString();
                else
                    ws.Cells["y" + rw].Value = "0.00";

                ws.Cells["z" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["z" + rw].Style.Numberformat.Format = "#,##0.00";

                string froMulaAddCal = string.Format("=SUM(w" + rw + "*y" + rw + ")");
                ws.Cells["aa" + rw].Formula = froMulaAddCal.ToString();
                ws.Cells["aa" + rw].Style.Numberformat.Format = "#,##0.00";


                DataTable _dtFree = GetSlapDays(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                if (_dtFree.Rows.Count > 0)
                    ws.Cells["ab" + rw].Value = _dtSlap.Rows[0]["SlabTo"];
                else
                    ws.Cells["ab" + rw].Value = 0;



                ws.Cells["AC" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                ws.Cells["AD" + rw].Value = dtv.Rows[i]["MADate"].ToString();
                ws.Cells["AE" + rw].Value = dtv.Rows[i]["Daysv"].ToString();


                #region TwoSlap

                WaiverQty = "5";
                int.TryParse(WaiverQty, out _FreeDays);
                if (_dtSlap.Rows.Count > 0)
                {
                    for (int y = 0; y < _dtSlap.Rows.Count; y++)
                    {
                        RowsCount = y;

                        ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                        LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                        ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                        Currency = _dtSlap.Rows[y]["Currency"].ToString();
                        Amount = _dtSlap.Rows[y]["Amount"].ToString();
                        FDatev = IncreamentalFromDt.ToShortDateString();
                        IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                        TDatev = IncrementalToDt.ToShortDateString();

                        if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                        {
                            TDatev = ToDt.ToShortDateString();
                            TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                            Days = (int)((TMS.TotalDays + 1));

                            if (!IsexceedFreedays)
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                if (Days <= 0)
                                    Days = 1;
                            }
                        }
                        else
                        {
                            if (!IsexceedFreedays)
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));

                            }
                        }
                        IncreamentalFromDt = IncrementalToDt.AddDays(1);
                       
                        string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                        if (y == 0)
                        {
                            ws.Cells["AF" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AF" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 1)
                        {
                            ws.Cells["AG" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AG" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 2)
                        {
                            ws.Cells["AH" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AH" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 3)
                        {
                            ws.Cells["AI" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AI" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 4)
                        {
                            ws.Cells["AJ" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AJ" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 5)
                        {
                            ws.Cells["AK" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AK" + rw].Style.Numberformat.Format = "#,##0.00";
                        }

                        if (DateTime.Parse(TDatev) >= DateTime.Parse((ToDt.ToString("dd/MM/yyyy"))))
                        {
                            break;

                        }

                    }
                    if (RowsCount == 1)
                    {
                        ws.Cells["AG" + rw].Value = 0.00;
                        ws.Cells["AH" + rw].Value = 0.00;
                        ws.Cells["AI" + rw].Value = 0.00;
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 2)
                    {
                        ws.Cells["AH" + rw].Value = 0.00;
                        ws.Cells["AI" + rw].Value = 0.00;
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 3)
                    {
                        ws.Cells["AI" + rw].Value = 0.00;
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 4)
                    {
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 5)
                    {
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                }
                else
                {
                    ws.Cells["AF" + rw].Value = 0.00;
                    ws.Cells["AG" + rw].Value = 0.00;
                    ws.Cells["AH" + rw].Value = 0.00;
                    ws.Cells["AI" + rw].Value = 0.00;
                    ws.Cells["AJ" + rw].Value = 0.00;
                    ws.Cells["AK" + rw].Value = 0.00;
                }

                #endregion



                string froMulaAL = string.Format("=SUM(AF" + rw + ":AK" + rw + ")");
                ws.Cells["AL" + rw].Formula = froMulaAddQV;
                ws.Cells["AL" + rw].Style.Numberformat.Format = "#,##0.00";
                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["AM" + rw].Value = _dtSlap.Rows[0]["Currency"].ToString();
                else
                    ws.Cells["AM" + rw].Value = "NA";

                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["AN" + rw].Value = _dtSlap.Rows[0]["ExRate"].ToString();
                else
                    ws.Cells["AN" + rw].Value = "0.00";

                ws.Cells["AO" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["AO" + rw].Style.Numberformat.Format = "#,##0.00";

                string froMulaAP = string.Format("=SUM(AL" + rw + "*AN" + rw + ")");
                ws.Cells["AP" + rw].Formula = froMulaAP.ToString();

                ws.Cells["AQ" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                ws.Cells["AR" + rw].Value = "";
                ws.Cells["AS" + rw].Value = "";
                ws.Cells["AT" + rw].Value = "";




                rw++;
            }


         
            ws.Cells["A7:AT10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 50].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=ImportDetentionReport.xlsx");
            Response.End();

        }
        public DataTable GetImpDetentionReport(string DtFrom, string DtTo)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
                            " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,  " +
                            " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FVDate, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MADate, " +

                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FVDatev, " +

                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as MADatev, " +

                            " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +
                            " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ImpFreeDays,  " +
                            " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ExpFreeDays, " +
                            " DATEADD(DAY,(select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

                            " case when(DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                            " then DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +

                            " from NVO_ContainerTxns " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                            " where NVO_ContainerTxns.StatusCode in ('FV')";

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND NVO_ContainerTxns.DtMovement between '" + DtFrom + "' and '" + DtTo + "'";
               


            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetContractSlap(string AgencyID, string CntrTypes)
        {
            string strWhere = "";
            string _Query = " select isnull((select top(1)  Rate from NVO_ExRate where AgencyID= NVO_ContRentContract.AgencyID and FromCurrency= NVO_ContRentTariffDtls.CurrencyID order by Id desc),0) as ExRate," +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where ID =NVO_ContRentTariffDtls.CurrencyID) as Currency,* from NVO_ContRentContract inner join NVO_ContRentTariffDtls on NVO_ContRentTariffDtls.RentID = NVO_ContRentContract.ID " +
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 2 and ChargesID=28";
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetDemurageContractSlap(string AgencyID, string CntrTypes)
        {
            string strWhere = "";
            string _Query = " select isnull((select top(1)  Rate from NVO_ExRate where AgencyID= NVO_ContRentContract.AgencyID and FromCurrency= NVO_ContRentTariffDtls.CurrencyID order by Id desc),0) as ExRate," +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where ID =NVO_ContRentTariffDtls.CurrencyID) as Currency,* from NVO_ContRentContract inner join NVO_ContRentTariffDtls on NVO_ContRentTariffDtls.RentID = NVO_ContRentContract.ID " +
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 2 and ChargesID=29";
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }


        public DataTable GetContractSlapExp(string AgencyID, string CntrTypes)
        {
            string strWhere = "";
            string _Query = " select isnull((select top(1)  Rate from NVO_ExRate where AgencyID= NVO_ContRentContract.AgencyID and FromCurrency= NVO_ContRentTariffDtls.CurrencyID order by Id desc),0) as ExRate," +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where ID =NVO_ContRentTariffDtls.CurrencyID) as Currency,* from NVO_ContRentContract inner join NVO_ContRentTariffDtls on NVO_ContRentTariffDtls.RentID = NVO_ContRentContract.ID " +
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 1 and ChargesID=28";
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetSlapDays(string AgencyID, string CntrTypes)
        {
            string strWhere = "";
            string _Query = " select top(1) SlabTo from NVO_ContRentContract inner join NVO_ContRentTariffDtls on NVO_ContRentTariffDtls.RentID = NVO_ContRentContract.ID "+
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 2";
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetSlapDaysExp(string AgencyID, string CntrTypes)
        {
            string strWhere = "";
            string _Query = " select top(1) SlabTo from NVO_ContRentContract inner join NVO_ContRentTariffDtls on NVO_ContRentTariffDtls.RentID = NVO_ContRentContract.ID " +
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 1";
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }




        public void EQCExpDetentionReportValues(string DtFrom, string DtTo, string User)
        {

            ExcelPackage pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("ExportDetentionReport");

            ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:G2"];
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

            ws.Cells["A7:A10"].Value = "S. No.";
            ws.Cells["A7:A10"].Merge = true;
            ws.Cells["A7:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["B7:B10"].Value = "GEO-LOCATION";
            ws.Cells["B7:B10"].Merge = true;
            ws.Cells["B7:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["C7:C10"].Value = "AGENCY NAME";
            ws.Cells["C7:C10"].Merge = true;
            ws.Cells["C7:C0"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["D7:D10"].Value = "VESSEL/VOYAGE";
            ws.Cells["D7:D10"].Merge = true;
            ws.Cells["D7:D0"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["E7:E10"].Value = "POL";
            ws.Cells["E7:E10"].Merge = true;
            ws.Cells["E7:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["F7:F10"].Value = "POD";
            ws.Cells["F7:F10"].Merge = true;
            ws.Cells["F7:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["G7:G10"].Value = "CONSIGNEE";
            ws.Cells["G7:G10"].Merge = true;
            ws.Cells["G7:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["H7:H10"].Value = "BL NUMBER";
            ws.Cells["H7:H10"].Merge = true;
            ws.Cells["H7:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["I7:I10"].Value = "CONTAINER NO";
            ws.Cells["I7:I10"].Merge = true;
            ws.Cells["I7:I10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["J7:J10"].Value = "TYPE SIZE";
            ws.Cells["J7:J10"].Merge = true;
            ws.Cells["J7:J10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["K7:AA7"].Value = "With Free Time";
            ws.Cells["K7:AA7"].Merge = true;
            ws.Cells["K7:AA7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["AB7:AP7"].Value = "Without Free Time";
            ws.Cells["AB7:AP7"].Merge = true;
            ws.Cells["AB7:AP7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["K8"].Value = "Landing";
            ws.Cells["K9"].Value = " Date";
            ws.Cells["K10"].Value = "(MS)";

            ws.Cells["L8"].Value = "EMPTY IN";
            ws.Cells["L9"].Value = " Date";
            ws.Cells["L10"].Value = "(FL)";

            ws.Cells["M8"].Value = "FREE TIME";
            ws.Cells["M9"].Value = "(DAYS)";
            ws.Cells["M10"].Value = "";

            ws.Cells["N8"].Value = "DTN FROM";
            ws.Cells["N9"].Value = "A";
            ws.Cells["N10"].Value = "(MS  + F/Time)";


            ws.Cells["O8"].Value = "PLOT IN";
            ws.Cells["O9"].Value = "B";
            ws.Cells["O10"].Value = "FL";

            ws.Cells["P8"].Value = "DTN";
            ws.Cells["P9"].Value = "DAYS";
            ws.Cells["P10"].Value = "B-A";

            ws.Cells["Q8:W8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["Q8:W8"].Merge = true;
            ws.Cells["Q8:W8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["Q9"].Value = "SLAB1";
            ws.Cells["R9"].Value = "SLAB2";
            ws.Cells["S9"].Value = "SLAB3";
            ws.Cells["T9"].Value = "SLAB4";
            ws.Cells["U9"].Value = "SLAB5";
            ws.Cells["V9"].Value = "SLAB6";
            ws.Cells["W9"].Value = "TOTAL";

            ws.Cells["X8"].Value = "TARIFF";
            ws.Cells["X9"].Value = "CURRENCY";
            ws.Cells["X10"].Value = "";

            ws.Cells["Y8"].Value = "EX.RATE";
            ws.Cells["Y9"].Value = "";
            ws.Cells["Y10"].Value = "";

            ws.Cells["Z8:AA8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["Z8:AA8"].Merge = true;
            ws.Cells["Z8:AA8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["Z9"].Value = "USD";
            ws.Cells["AA9"].Value = "LCY";


            ws.Cells["AB8"].Value = "FREE TIME";
            ws.Cells["AB9"].Value = "(DAYS)";
            ws.Cells["AB10"].Value = "";

            ws.Cells["AC8"].Value = "DTN FROM";
            ws.Cells["AC9"].Value = "A";
            ws.Cells["AC10"].Value = "(MS  + F/Time)";


            ws.Cells["AD8"].Value = "PLOT IN";
            ws.Cells["AD9"].Value = "B";
            ws.Cells["AD10"].Value = "FL";

            ws.Cells["AE8"].Value = "DTN";
            ws.Cells["AE9"].Value = "DAYS";
            ws.Cells["AE10"].Value = "B-A";

            ws.Cells["AF8:AL8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["AF8:AL8"].Merge = true;
            ws.Cells["AF8:AL8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["AF9"].Value = "SLAB1";
            ws.Cells["AG9"].Value = "SLAB2";
            ws.Cells["AH9"].Value = "SLAB3";
            ws.Cells["AI9"].Value = "SLAB4";
            ws.Cells["AJ9"].Value = "SLAB5";
            ws.Cells["AK9"].Value = "SLAB6";
            ws.Cells["AL9"].Value = "TOTAL";

            ws.Cells["AM8"].Value = "TARIFF";
            ws.Cells["AM9"].Value = "CURRENCY";
            ws.Cells["AM10"].Value = "";

            ws.Cells["AN8"].Value = "EX.RATE";
            ws.Cells["AN9"].Value = "";
            ws.Cells["AN10"].Value = "";

            ws.Cells["AO8:AP8"].Value = "DETENTION (IN TARIFF CURRENCY)";
            ws.Cells["AO8:AP8"].Merge = true;
            ws.Cells["AO8:AP8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells["AO9"].Value = "USD";
            ws.Cells["AP9"].Value = "LCY";


            ws.Cells["AQ7:AQ10"].Value = "SHIPPER";
            ws.Cells["AQ7:AQ10"].Merge = true;
            ws.Cells["AQ7:AQ10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["AR7:AR10"].Value = "INVOICE NO";
            ws.Cells["AR7:AR10"].Merge = true;
            ws.Cells["AR7:AR10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["AS7:AS10"].Value = "WORKSHEET USD";
            ws.Cells["AS7:AS10"].Merge = true;
            ws.Cells["AS7:AS10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["AT7:AT10"].Value = "WORKSHEET INR";
            ws.Cells["AT7:AT10"].Merge = true;
            ws.Cells["AT7:AT10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            //r = ws.Cells["A7:AA7"];
            //r.Style.Font.Bold = true;
            //r.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            ExcelRange rng1 = ws.Cells["K7:AA10"];
            rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng1.Style.Fill.BackgroundColor.SetColor(Color.LightSalmon);

            ExcelRange rng2 = ws.Cells["AB7:AP10"];
            rng2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng2.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

            ws.Cells["A7:AT10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            int sl = 1;
            int rw = 11;
            DataTable dtv = GetExpDetentionReport(DtFrom, DtTo);
            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = "";
                ws.Cells["C" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["POL"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["POD"].ToString();
                ws.Cells["G" + rw].Value = "";
                ws.Cells["H" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                ws.Cells["K" + rw].Value = dtv.Rows[i]["MSDate"].ToString();
                ws.Cells["L" + rw].Value = dtv.Rows[i]["FLDate"].ToString();
                ws.Cells["M" + rw].Value = dtv.Rows[i]["ExpFreeDays"].ToString();

                ws.Cells["N" + rw].Value = dtv.Rows[i]["FromDate"].ToString();

                ws.Cells["O" + rw].Value = dtv.Rows[i]["FLDate"].ToString();
                ws.Cells["P" + rw].Value = dtv.Rows[i]["Daysv"].ToString();

                string FDatev = "";
                string TDatev = "";
                int RowsCount = 0;
                string ID = "";
                string LLimit = "";
                string ULimit = "";
                string Currency = "";
                string Amount = "";
                double Days = 0;
                int _FreeDays = 0;
                decimal ExRate = 0;
                FDatev = "";
                TDatev = "";
                DateTime IncreamentalFromDt, FromDt;
                DateTime IncrementalToDt, ToDt;
                TimeSpan TMS; bool IsexceedFreedays = false; string WaiverQty = "0";
                DateTime.TryParse(dtv.Rows[i]["MSDatev"].ToString(), out IncreamentalFromDt);
                DateTime.TryParse(dtv.Rows[i]["MSDatev"].ToString(), out FromDt);
                DateTime.TryParse(dtv.Rows[i]["FLDatev"].ToString().ToString(), out ToDt);
                WaiverQty = dtv.Rows[i]["ImpFreeDays"].ToString();
                int.TryParse(WaiverQty, out _FreeDays);
                DataTable _dtSlap = GetContractSlapExp(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                if (_dtSlap.Rows.Count > 0)
                {
                    for (int y = 0; y < _dtSlap.Rows.Count; y++)
                    {
                        RowsCount = y;

                        ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                        LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                        ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                        Currency = _dtSlap.Rows[y]["Currency"].ToString();
                        Amount = _dtSlap.Rows[y]["Amount"].ToString();
                        FDatev = IncreamentalFromDt.ToShortDateString();
                        IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                        TDatev = IncrementalToDt.ToShortDateString();

                        if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                        {
                            TDatev = ToDt.ToShortDateString();
                            TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                            Days = (int)((TMS.TotalDays + 1));

                            if (!IsexceedFreedays)
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                if (Days <= 0)
                                    Days = 1;
                            }
                        }
                        else
                        {
                            if (!IsexceedFreedays)
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));

                            }
                        }
                        IncreamentalFromDt = IncrementalToDt.AddDays(1);

                        string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                        if (y == 0)
                        {
                            ws.Cells["q" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["q" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 1)
                        {
                            ws.Cells["R" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 2)
                        {
                            ws.Cells["S" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["s" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 3)
                        {
                            ws.Cells["T" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["T" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 4)
                        {
                            ws.Cells["U" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 5)
                        {
                            ws.Cells["V" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["V" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (DateTime.Parse(TDatev) >= DateTime.Parse((ToDt.ToString("dd/MM/yyyy"))))
                        {
                            break;

                        }

                    }
                    if (RowsCount == 1)
                    {
                        ws.Cells["R" + rw].Value = 0.00;
                        ws.Cells["s" + rw].Value = 0.00;
                        ws.Cells["t" + rw].Value = 0.00;
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 2)
                    {
                        ws.Cells["s" + rw].Value = 0.00;
                        ws.Cells["t" + rw].Value = 0.00;
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 3)
                    {
                        ws.Cells["t" + rw].Value = 0.00;
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 4)
                    {
                        ws.Cells["u" + rw].Value = 0.00;
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                    if (RowsCount == 5)
                    {
                        ws.Cells["v" + rw].Value = 0.00;
                    }
                }
                else
                {
                    ws.Cells["q" + rw].Value = 0.00;
                    ws.Cells["R" + rw].Value = 0.00;
                    ws.Cells["s" + rw].Value = 0.00;
                    ws.Cells["t" + rw].Value = 0.00;
                    ws.Cells["u" + rw].Value = 0.00;
                    ws.Cells["v" + rw].Value = 0.00;
                }

                string froMulaAddQV = string.Format("=SUM(Q" + rw + ":V" + rw + ")");
                ws.Cells["w" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["w" + rw].Style.Numberformat.Format = "#,##0.00";
                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["x" + rw].Value = _dtSlap.Rows[0]["Currency"].ToString();
                else
                    ws.Cells["x" + rw].Value = "NA";

                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["y" + rw].Value = _dtSlap.Rows[0]["ExRate"].ToString();
                else
                    ws.Cells["y" + rw].Value = "0.00";

                ws.Cells["z" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["z" + rw].Style.Numberformat.Format = "#,##0.00";

                string froMulaAddCal = string.Format("=SUM(w" + rw + "*y" + rw + ")");
                ws.Cells["aa" + rw].Formula = froMulaAddCal.ToString();
                ws.Cells["aa" + rw].Style.Numberformat.Format = "#,##0.00";


                DataTable _dtFree = GetSlapDays(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                if (_dtFree.Rows.Count > 0)
                    ws.Cells["ab" + rw].Value = _dtSlap.Rows[0]["SlabTo"];
                else
                    ws.Cells["ab" + rw].Value = 0;



                ws.Cells["AC" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                ws.Cells["AD" + rw].Value = dtv.Rows[i]["FLDate"].ToString();
                ws.Cells["AE" + rw].Value = dtv.Rows[i]["Daysv"].ToString();


                #region TwoSlap

                WaiverQty = "5";
                int.TryParse(WaiverQty, out _FreeDays);
                if (_dtSlap.Rows.Count > 0)
                {
                    for (int y = 0; y < _dtSlap.Rows.Count; y++)
                    {
                        RowsCount = y;

                        ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                        LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                        ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                        Currency = _dtSlap.Rows[y]["Currency"].ToString();
                        Amount = _dtSlap.Rows[y]["Amount"].ToString();
                        FDatev = IncreamentalFromDt.ToShortDateString();
                        IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                        TDatev = IncrementalToDt.ToShortDateString();

                        if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                        {
                            TDatev = ToDt.ToShortDateString();
                            TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                            Days = (int)((TMS.TotalDays + 1));

                            if (!IsexceedFreedays)
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                if (Days <= 0)
                                    Days = 1;
                            }
                        }
                        else
                        {
                            if (!IsexceedFreedays)
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));

                            }
                        }
                        IncreamentalFromDt = IncrementalToDt.AddDays(1);
                      
                        string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                        if (y == 0)
                        {
                            ws.Cells["AF" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AF" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 1)
                        {
                            ws.Cells["AG" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AG" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 2)
                        {
                            ws.Cells["AH" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AH" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 3)
                        {
                            ws.Cells["AI" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AI" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 4)
                        {
                            ws.Cells["AJ" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AJ" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 5)
                        {
                            ws.Cells["AK" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["AK" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (DateTime.Parse(TDatev) >= DateTime.Parse((ToDt.ToString("dd/MM/yyyy"))))
                        {
                            break;

                        }

                    }
                    if (RowsCount == 1)
                    {
                        ws.Cells["AG" + rw].Value = 0.00;
                        ws.Cells["AH" + rw].Value = 0.00;
                        ws.Cells["AI" + rw].Value = 0.00;
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 2)
                    {
                        ws.Cells["AH" + rw].Value = 0.00;
                        ws.Cells["AI" + rw].Value = 0.00;
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 3)
                    {
                        ws.Cells["AI" + rw].Value = 0.00;
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 4)
                    {
                        ws.Cells["AJ" + rw].Value = 0.00;
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                    if (RowsCount == 5)
                    {
                        ws.Cells["AK" + rw].Value = 0.00;
                    }
                }
                else
                {
                    ws.Cells["AF" + rw].Value = 0.00;
                    ws.Cells["AG" + rw].Value = 0.00;
                    ws.Cells["AH" + rw].Value = 0.00;
                    ws.Cells["AI" + rw].Value = 0.00;
                    ws.Cells["AJ" + rw].Value = 0.00;
                    ws.Cells["AK" + rw].Value = 0.00;
                }

                #endregion



                string froMulaAL = string.Format("=SUM(AF" + rw + ":AK" + rw + ")");
                ws.Cells["AL" + rw].Formula = froMulaAddQV;
                ws.Cells["AL" + rw].Style.Numberformat.Format = "#,##0.00";
                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["AM" + rw].Value = _dtSlap.Rows[0]["Currency"].ToString();
                else
                    ws.Cells["AM" + rw].Value = "NA";

                if (_dtSlap.Rows.Count > 0)
                    ws.Cells["AN" + rw].Value = _dtSlap.Rows[0]["ExRate"].ToString();
                else
                    ws.Cells["AN" + rw].Value = "0.00";

                ws.Cells["AO" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["AO" + rw].Style.Numberformat.Format = "#,##0.00";

                string froMulaAP = string.Format("=SUM(AL" + rw + "*AN" + rw + ")");
                ws.Cells["AP" + rw].Formula = froMulaAP.ToString();

                ws.Cells["AQ" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                ws.Cells["AR" + rw].Value = "";
                ws.Cells["AS" + rw].Value = "";
                ws.Cells["AT" + rw].Value = "";




                rw++;
            }



            ws.Cells["A7:AT10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A7:AT10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 50].AutoFitColumns();

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=ExpDetentionReport.xlsx");
            Response.End();

        }

        public DataTable GetExpDetentionReport(string DtFrom, string DtTo)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
                            " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,  " +
                            " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MSDate, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FLDate, " +

                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MSDatev, " +

                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as FLDatev, " +

                            " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +
                            " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ImpFreeDays,  " +
                            " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ExpFreeDays, " +
                            " DATEADD(DAY,(select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

                            " case when(DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                            " then DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +

                            " from NVO_ContainerTxns " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                            " where NVO_ContainerTxns.StatusCode in ('MS')";

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND NVO_ContainerTxns.DtMovement between '" + DtFrom + "' and '" + DtTo + "'";



            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public void EQCExpDetentionSummaryReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {

            ExcelPackage pck = new ExcelPackage();

            DataTable dt = GetAgencyLocation(DtFrom, DtTo, GeoLocID);
            for (int J = 0; J < dt.Rows.Count; J++)
            {
                var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
                ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:G2"];
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

                ws.Cells["A8:A10"].Value = "S. No.";
                ws.Cells["A8:A10"].Merge = true;
                ws.Cells["A8:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["B8:B10"].Value = "VESSEL/VOYAGE";
                ws.Cells["B8:B10"].Merge = true;
                ws.Cells["B8:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["C8:C10"].Value = "POL";
                ws.Cells["C8:C10"].Merge = true;
                ws.Cells["C8:C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["D8:D10"].Value = "POD";
                ws.Cells["D8:D10"].Merge = true;
                ws.Cells["D8:D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["E8:E10"].Value = "BL NUMBER";
                ws.Cells["E8:E10"].Merge = true;
                ws.Cells["E8:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                ws.Cells["F8:F10"].Value = "AGENCY NAME";
                ws.Cells["F8:F10"].Merge = true;
                ws.Cells["F8:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["G8:G10"].Value = "CONTAINER NO";
                ws.Cells["G8:G10"].Merge = true;
                ws.Cells["G8:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["H8:H10"].Value = "TYPE-SIZE";
                ws.Cells["H8:H10"].Merge = true;
                ws.Cells["H8:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                ws.Cells["I8"].Value = "PICK UP ";
                ws.Cells["I9"].Value = " Date";
                ws.Cells["I10"].Value = "A (MS)";

                ws.Cells["J8"].Value = "Free";
                ws.Cells["J9"].Value = " Time";
                ws.Cells["J10"].Value = "B";

                ws.Cells["K8"].Value = "DTN From";
                ws.Cells["K9"].Value = "(MS + F/Time)";
                ws.Cells["K10"].Value = "A+B = C";

                ws.Cells["L8"].Value = "DTN Upto";
                ws.Cells["L9"].Value = "PORT IN";
                ws.Cells["L10"].Value = "D (FL)";


                ws.Cells["M8"].Value = "DTN";
                ws.Cells["M9"].Value = "Days";
                ws.Cells["M10"].Value = "D-C";


                ws.Cells["N9"].Value = "SLAB1";
                ws.Cells["O9"].Value = "SLAB2";
                ws.Cells["P9"].Value = "SLAB3";
                ws.Cells["Q9"].Value = "SLAB4";
                ws.Cells["R9"].Value = "SLAB5";
                ws.Cells["S9"].Value = "SLAB6";

                ws.Cells["T9"].Value = "CURRENCY";

                ws.Cells["U8"].Value = "TOTAL";
                ws.Cells["U9"].Value = "USD";
                ws.Cells["U10"].Value = "COLLECTION";


                ExcelRange rng1 = ws.Cells["K8:U10"];
                rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng1.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);



                ws.Cells["A8:U10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                int sl = 1;
                int rw = 11;
                string Currency = "USD";
                DataTable dtv = GetExpDetentionSummaryReport(DtFrom, DtTo, dt.Rows[J]["GeoLocationID"].ToString());
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    int CountDays = 0;
                    ws.Cells["A" + rw].Value = sl++;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["MSDate"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["ExpFreeDays"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["FLDate"].ToString();


                    int RowDays = 0;
                    if ((Int32.Parse(dtv.Rows[i]["Days"].ToString())) > 1)
                    {
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["Days"].ToString();
                        RowDays = Int32.Parse(dtv.Rows[i]["Days"].ToString());
                    }
                    else
                        ws.Cells["M" + rw].Value = 0;


                    if (RowDays.ToString() != "0")
                    {
                        string FDatev = "";
                        string TDatev = "";
                        int RowsCount = 0;
                        string ID = "";
                        string LLimit = "";
                        string ULimit = "";
                        Currency = "";
                        string Amount = "";
                        double Days = 0;
                        int _FreeDays = 0;
                        decimal ExRate = 0;
                        FDatev = "";
                        TDatev = "";
                        DateTime IncreamentalFromDt, FromDt;
                        DateTime IncrementalToDt, ToDt;
                        TimeSpan TMS; bool IsexceedFreedays = false; string WaiverQty = "0";
                        DateTime.TryParse(dtv.Rows[i]["MSDatev"].ToString(), out IncreamentalFromDt);
                        DateTime.TryParse(dtv.Rows[i]["MSDatev"].ToString(), out FromDt);
                        DateTime.TryParse(dtv.Rows[i]["FLDatev"].ToString(), out ToDt);
                        WaiverQty = dtv.Rows[i]["ExpFreeDays"].ToString();
                        int.TryParse(WaiverQty, out _FreeDays);
                        DataTable _dtSlap = GetContractSlapExp(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                        if (_dtSlap.Rows.Count > 0)
                        {
                            for (int y = 0; y < _dtSlap.Rows.Count; y++)
                            {
                                RowsCount = y;
                                ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                                LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                                ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                                Currency = _dtSlap.Rows[y]["Currency"].ToString();
                                Amount = _dtSlap.Rows[y]["Amount"].ToString();
                                FDatev = IncreamentalFromDt.ToShortDateString();
                                IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                                TDatev = IncrementalToDt.ToShortDateString();

                                if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                                {
                                    TDatev = ToDt.ToShortDateString();
                                    TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                    Days = (int)((TMS.TotalDays+1));

                                    if (!IsexceedFreedays)
                                    {
                                        TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays+1));
                                        Days = Days - _FreeDays;
                                        if (Days < 0)
                                            Days = 0;
                                        else
                                            IsexceedFreedays = true;
                                    }
                                    else
                                    {
                                        //TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        //Days = (int)((TMS.TotalDays + 1));
                                        //if (Days <= 0)
                                        //    Days = 1;

                                        //Muthu Chage +1 values
                                        //TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        //Days = (int)((TMS.TotalDays + 1));
                                        //if (Days <= 0)
                                        //    Days = 0;

                                        TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays+1));
                                        if (Days <= 0)
                                            Days = 0;
                                    }
                                }
                                else
                                {
                                    if (!IsexceedFreedays)
                                    {
                                        TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays));
                                        Days = Days - _FreeDays;
                                        if (Days < 0)
                                            Days = 0;
                                        else
                                            IsexceedFreedays = true;
                                    }
                                    else
                                    {
                                        TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays));

                                    }
                                }
                                IncreamentalFromDt = IncrementalToDt.AddDays(1);
                              
                                string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();
                               

                                if (y == 0 )
                                {
                                    ws.Cells["N" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["N" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 1)
                                {
                                    ws.Cells["O" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["O" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 2)
                                {
                                    ws.Cells["P" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["P" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 3)
                                {
                                    ws.Cells["Q" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["Q" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 4)
                                {
                                    ws.Cells["R" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 5)
                                {
                                    ws.Cells["S" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["S" + rw].Style.Numberformat.Format = "#,##0.00";
                                }

                                
                            }
                            if (RowsCount == 1)
                            {
                                ws.Cells["P" + rw].Value = 0.00;
                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 2)
                            {

                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 3)
                            {

                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 4)
                            {

                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 5)
                            {
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                        }
                        else
                        {
                            ws.Cells["N" + rw].Value = 0.00;
                            ws.Cells["O" + rw].Value = 0.00;
                            ws.Cells["P" + rw].Value = 0.00;
                            ws.Cells["Q" + rw].Value = 0.00;
                            ws.Cells["R" + rw].Value = 0.00;
                            ws.Cells["S" + rw].Value = 0.00;
                        }
                    }
                    else
                    {
                        ws.Cells["N" + rw].Value = 0.00;
                        ws.Cells["O" + rw].Value = 0.00;
                        ws.Cells["P" + rw].Value = 0.00;
                        ws.Cells["Q" + rw].Value = 0.00;
                        ws.Cells["R" + rw].Value = 0.00;
                        ws.Cells["S" + rw].Value = 0.00;
                    }
                    ws.Cells["T" + rw].Value = Currency;
                    string froMulaAddQV = string.Format("=SUM(M" + rw + ":R" + rw + ")");
                    ws.Cells["U" + rw].Formula = froMulaAddQV.ToString();
                    ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                    rw++;
                }



                ws.Cells["A8:U10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 50].AutoFitColumns();
            }
            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=ExpDetentionSummaryReport.xlsx");
            Response.End();

        }

        public DataTable GetAgencyLocation(string DtFrom, string DtTo, string GeoLocID)
        {
            string strWhere = "";
            string _Query = " select distinct GeoLocationID, (select top(1) GeoLocation from NVO_GeoLocations where ID =NVO_AgencyMaster.GeoLocationID) as GeoLocation " +
                            " From NVO_AgencyMaster " +
                            " inner join NVO_ContainerTxns on NVO_ContainerTxns.AgencyID=NVO_AgencyMaster.ID  ";

            if (GeoLocID != "" && GeoLocID != "0" && GeoLocID != "null" && GeoLocID != "?")
               
                    strWhere += _Query + " where GeoLocationID=" + GeoLocID;
                else
                    strWhere += _Query + " where GeoLocationID !=0 ";

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND NVO_ContainerTxns.DtMovement between '" + DtFrom + "' and '" + DtTo + "'";

        

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetExpDetentionSummaryReport(string DtFrom, string DtTo, string GeoLocID)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
                            " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,  " +
                            " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MSDate, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FLDate, " +

                            " (select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MSDatev, " +

                            " isnull((select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as FLDatev, " +

                            " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS DaysT, " +
                            " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)) as ImpFreeDays,  " +
                            " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)) as ExpFreeDays, " +
                            " convert(varchar,DATEADD(DAY,(select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)), " +
                            " (select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)),103) as FromDate, " +

                            " case when(DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID=2), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                            " then DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,3)), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv, " +

                            " datediff(day,  DATEADD(DAY,(select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID  " +
                            " and ModeID in(2, 1)),  (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)), isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) as Days1, " +
                            " isnull((DATEDIFF(DAY,   " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)+1)) " +
                            " -(select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID  and ModeID in(2, 1)),0) AS Days " +

                            " from NVO_ContainerTxns " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                            " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID=NVO_ContainerTxns.AgencyID " +
                            " where NVO_ContainerTxns.StatusCode in ('FL') and  GeoLocationID= " + GeoLocID;
                            

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND convert(varchar, NVO_ContainerTxns.DtMovement, 23) between '" + DtFrom + "' and '" + DtTo + "'";



            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }


        public void EQCImpDetentionSummaryReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {

            ExcelPackage pck = new ExcelPackage();

            DataTable dt = GetAgencyLocation(DtFrom,DtTo,GeoLocID);
            if (dt.Rows.Count >0)
            {
                for (int J = 0; J < dt.Rows.Count; J++)
                {
                    var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
                    ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
                    ws.Cells["A2"].Style.Font.Bold = true;
                    ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ExcelRange r = ws.Cells["A2:G2"];
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

                    ws.Cells["A8:A10"].Value = "S. No.";
                    ws.Cells["A8:A10"].Merge = true;
                    ws.Cells["A8:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["B8:B10"].Value = "VESSEL/VOYAGE";
                    ws.Cells["B8:B10"].Merge = true;
                    ws.Cells["B8:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["C8:C10"].Value = "POL";
                    ws.Cells["C8:C10"].Merge = true;
                    ws.Cells["C8:C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["D8:D10"].Value = "POD";
                    ws.Cells["D8:D10"].Merge = true;
                    ws.Cells["D8:D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["E8:E10"].Value = "BL NUMBER";
                    ws.Cells["E8:E10"].Merge = true;
                    ws.Cells["E8:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["F8:F10"].Value = "AGENCY NAME";
                    ws.Cells["F8:F10"].Merge = true;
                    ws.Cells["F8:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["G8:G10"].Value = "CONTAINER NO";
                    ws.Cells["G8:G10"].Merge = true;
                    ws.Cells["G8:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    ws.Cells["H8:H10"].Value = "TYPE-SIZE";
                    ws.Cells["H8:H10"].Merge = true;
                    ws.Cells["H8:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    ws.Cells["I8"].Value = "Arrival";
                    ws.Cells["I9"].Value = " Date";
                    ws.Cells["I10"].Value = "A (FV)";

                    ws.Cells["J8"].Value = "Free";
                    ws.Cells["J9"].Value = " Time";
                    ws.Cells["J10"].Value = "B";

                    ws.Cells["K8"].Value = "DTN From";
                    ws.Cells["K9"].Value = "(FV + F/Time)";
                    ws.Cells["K10"].Value = "A+B = C";

                    ws.Cells["L8"].Value = "DTN Upto";
                    ws.Cells["L9"].Value = "PORT IN";
                    ws.Cells["L10"].Value = "D (MA)";


                    ws.Cells["M8"].Value = "DTN";
                    ws.Cells["M9"].Value = "Days";
                    ws.Cells["M10"].Value = "D-C";


                    ws.Cells["N9"].Value = "SLAB1";
                    ws.Cells["O9"].Value = "SLAB2";
                    ws.Cells["P9"].Value = "SLAB3";
                    ws.Cells["Q9"].Value = "SLAB4";
                    ws.Cells["R9"].Value = "SLAB5";
                    ws.Cells["S9"].Value = "SLAB6";
                    ws.Cells["T9"].Value = "CURRENCY";

                    ws.Cells["U8"].Value = "TOTAL";
                    ws.Cells["U9"].Value = "USD";
                    ws.Cells["U10"].Value = "COLLECTION";


                    ExcelRange rng1 = ws.Cells["K8:T10"];
                    rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng1.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);



                    ws.Cells["A8:U10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A8:U10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A8:U10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A8:U10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    int sl = 1;
                    int rw = 11;
                    string Currency = "USD";
                    DataTable dtv = GetImpDetentionSummaryReport(DtFrom, DtTo, dt.Rows[J]["GeoLocationID"].ToString());
                    for (int i = 0; i < dtv.Rows.Count; i++)
                    {
                        ws.Cells["A" + rw].Value = sl++;
                        ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                        ws.Cells["C" + rw].Value = dtv.Rows[i]["POL"].ToString();
                        ws.Cells["D" + rw].Value = dtv.Rows[i]["POD"].ToString();
                        ws.Cells["E" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                        ws.Cells["F" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                        ws.Cells["G" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                        ws.Cells["H" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                        ws.Cells["I" + rw].Value = dtv.Rows[i]["FVDate"].ToString();
                        ws.Cells["J" + rw].Value = dtv.Rows[i]["ImpFreeDays"].ToString();
                        ws.Cells["K" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                        ws.Cells["L" + rw].Value = dtv.Rows[i]["MADate"].ToString();
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["Daysv"].ToString();
                        if (dtv.Rows[i]["Daysv"].ToString() != "0")
                        {

                            string FDatev = "";
                            string TDatev = "";
                            int RowsCount = 0;
                            string ID = "";
                            string LLimit = "";
                            string ULimit = "";
                            Currency = "";
                            string Amount = "";
                            double Days = 0;
                            int _FreeDays = 0;
                            decimal ExRate = 0;
                            FDatev = "";
                            TDatev = "";
                            DateTime IncreamentalFromDt, FromDt;
                            DateTime IncrementalToDt, ToDt;
                            TimeSpan TMS; bool IsexceedFreedays = false; string WaiverQty = "0";
                            DateTime.TryParse(dtv.Rows[i]["FVDatev"].ToString(), out IncreamentalFromDt);
                            DateTime.TryParse(dtv.Rows[i]["FVDatev"].ToString(), out FromDt);
                            DateTime.TryParse(dtv.Rows[i]["MADatev"].ToString().ToString(), out ToDt);
                            WaiverQty = dtv.Rows[i]["ImpFreeDays"].ToString();
                            int.TryParse(WaiverQty, out _FreeDays);
                            DataTable _dtSlap = GetContractSlap(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                            if (_dtSlap.Rows.Count > 0)
                            {
                                for (int y = 0; y < _dtSlap.Rows.Count; y++)
                                {
                                    RowsCount = y;

                                    ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                                    LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                                    ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                                    Currency = _dtSlap.Rows[y]["Currency"].ToString();
                                    Amount = _dtSlap.Rows[y]["Amount"].ToString();
                                    FDatev = IncreamentalFromDt.ToShortDateString();
                                    IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                                    TDatev = IncrementalToDt.ToShortDateString();

                                    if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                                    {
                                        TDatev = ToDt.ToShortDateString();
                                        TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays + 1));

                                        if (!IsexceedFreedays)
                                        {
                                            TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                            Days = (int)((TMS.TotalDays + 1));
                                            Days = Days - _FreeDays;
                                            if (Days < 0)
                                                Days = 0;
                                            else
                                                IsexceedFreedays = true;
                                        }
                                        else
                                        {
                                            TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                            Days = (int)((TMS.TotalDays + 1));
                                            if (Days <= 0)
                                                Days = 1;
                                        }
                                    }
                                    else
                                    {
                                        if (!IsexceedFreedays)
                                        {
                                            TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                            Days = (int)((TMS.TotalDays + 1));
                                            Days = Days - _FreeDays;
                                            if (Days < 0)
                                                Days = 0;
                                            else
                                                IsexceedFreedays = true;
                                        }
                                        else
                                        {
                                            TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                            Days = (int)((TMS.TotalDays + 1));

                                        }
                                    }
                                    IncreamentalFromDt = IncrementalToDt.AddDays(1);


                                    string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                                    if (y == 0)
                                    {
                                        ws.Cells["N" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                        ws.Cells["N" + rw].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    if (y == 1)
                                    {
                                        ws.Cells["O" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                        ws.Cells["O" + rw].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    if (y == 2)
                                    {
                                        ws.Cells["P" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                        ws.Cells["P" + rw].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    if (y == 3)
                                    {
                                        ws.Cells["Q" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                        ws.Cells["Q" + rw].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    if (y == 4)
                                    {
                                        ws.Cells["R" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                        ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    if (y == 5)
                                    {
                                        ws.Cells["R" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                        ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    if (DateTime.Parse(TDatev) >= DateTime.Parse(ToDt.ToString()))
                                    {
                                        break;

                                    }
                                    //if (DateTime.Parse(TDatev) >= DateTime.Parse((ToDt.ToString("dd/MM/yyyy"))))
                                    //{
                                    //    break;

                                    //}

                                }
                                if (RowsCount == 1)
                                {

                                    ws.Cells["P" + rw].Value = 0.00;
                                    ws.Cells["Q" + rw].Value = 0.00;
                                    ws.Cells["R" + rw].Value = 0.00;
                                    ws.Cells["S" + rw].Value = 0.00;
                                }
                                if (RowsCount == 2)
                                {

                                    ws.Cells["Q" + rw].Value = 0.00;
                                    ws.Cells["R" + rw].Value = 0.00;
                                    ws.Cells["S" + rw].Value = 0.00;
                                }
                                if (RowsCount == 3)
                                {

                                    ws.Cells["R" + rw].Value = 0.00;
                                    ws.Cells["S" + rw].Value = 0.00;
                                }
                                if (RowsCount == 4)
                                {

                                    ws.Cells["S" + rw].Value = 0.00;
                                }
                                if (RowsCount == 5)
                                {
                                    ws.Cells["S" + rw].Value = 0.00;
                                }
                            }
                            else
                            {
                                ws.Cells["N" + rw].Value = 0.00;
                                ws.Cells["O" + rw].Value = 0.00;
                                ws.Cells["P" + rw].Value = 0.00;
                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            ws.Cells["T" + rw].Value = Currency;
                        }
                        else
                        {
                            ws.Cells["N" + rw].Value = 0.00;
                            ws.Cells["O" + rw].Value = 0.00;
                            ws.Cells["P" + rw].Value = 0.00;
                            ws.Cells["Q" + rw].Value = 0.00;
                            ws.Cells["R" + rw].Value = 0.00;
                            ws.Cells["S" + rw].Value = 0.00;
                        }


                        string froMulaAddQV = string.Format("=SUM(M" + rw + ":R" + rw + ")");
                        ws.Cells["U" + rw].Formula = froMulaAddQV.ToString();
                        ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                        rw++;
                    }
                    ws.Cells["A8:U10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A8:U10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A8:U10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["A8:U10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    ws.Cells[1, 1, rw, 50].AutoFitColumns();
                }
                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ImpDetentionSummaryReport.xlsx");
                Response.End();
            }
           else
            {
                var ws = pck.Workbook.Worksheets.Add(GeoLocID);
                ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:G2"];
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

                ws.Cells["A8:A10"].Value = "S. No.";
                ws.Cells["A8:A10"].Merge = true;
                ws.Cells["A8:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["B8:B10"].Value = "VESSEL/VOYAGE";
                ws.Cells["B8:B10"].Merge = true;
                ws.Cells["B8:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["C8:C10"].Value = "POL";
                ws.Cells["C8:C10"].Merge = true;
                ws.Cells["C8:C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["D8:D10"].Value = "POD";
                ws.Cells["D8:D10"].Merge = true;
                ws.Cells["D8:D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["E8:E10"].Value = "BL NUMBER";
                ws.Cells["E8:E10"].Merge = true;
                ws.Cells["E8:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["F8:F10"].Value = "AGENCY NAME";
                ws.Cells["F8:F10"].Merge = true;
                ws.Cells["F8:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["G8:G10"].Value = "CONTAINER NO";
                ws.Cells["G8:G10"].Merge = true;
                ws.Cells["G8:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["H8:H10"].Value = "TYPE-SIZE";
                ws.Cells["H8:H10"].Merge = true;
                ws.Cells["H8:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                ws.Cells["I8"].Value = "Arrival";
                ws.Cells["I9"].Value = " Date";
                ws.Cells["I10"].Value = "A (FV)";

                ws.Cells["J8"].Value = "Free";
                ws.Cells["J9"].Value = " Time";
                ws.Cells["J10"].Value = "B";

                ws.Cells["K8"].Value = "DTN From";
                ws.Cells["K9"].Value = "(FV + F/Time)";
                ws.Cells["K10"].Value = "A+B = C";

                ws.Cells["L8"].Value = "DTN Upto";
                ws.Cells["L9"].Value = "PORT IN";
                ws.Cells["L10"].Value = "D (MA)";


                ws.Cells["M8"].Value = "DTN";
                ws.Cells["M9"].Value = "Days";
                ws.Cells["M10"].Value = "D-C";


                ws.Cells["N9"].Value = "SLAB1";
                ws.Cells["O9"].Value = "SLAB2";
                ws.Cells["P9"].Value = "SLAB3";
                ws.Cells["Q9"].Value = "SLAB4";
                ws.Cells["R9"].Value = "SLAB5";
                ws.Cells["S9"].Value = "SLAB6";
                ws.Cells["T9"].Value = "CURRENCY";

                ws.Cells["U8"].Value = "TOTAL";
                ws.Cells["U9"].Value = "USD";
                ws.Cells["U10"].Value = "COLLECTION";


                ExcelRange rng1 = ws.Cells["K8:T10"];
                rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng1.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);



                ws.Cells["A8:U10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ImpDetentionSummaryReport.xlsx");
                Response.End();
            }


        }
        public DataTable GetImpDetentionSummaryReport(string DtFrom, string DtTo, string GeoLocID)
        {
            string strWhere = "";
            //string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
            //                " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,  " +
            //                " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +


            //                " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDate, " +

            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MADate, " +

            //                 " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                 " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                 " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //                 " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                 " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDatev,  " +

            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as MADatev, " +

            //                " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +

            //                " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ImpFreeDays,  " +
            //                " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ExpFreeDays, " +
            //                " DATEADD(DAY,(select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

            //                " case when(DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
            //                " then DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +
            //                " from NVO_ContainerTxns " +
            //                " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
            //                " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
            //                " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
            //                " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID=NVO_ContainerTxns.AgencyID " +
            //                " where NVO_ContainerTxns.StatusCode in ('MA') and  BLTypes in (40,42) and GeoLocationID= " + GeoLocID;


            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
                           " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,  " +
                           " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +


                           " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                           " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDate, " +

                           " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                           " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MADate, " +

                            " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDatev,  " +

                           " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as MADatev, " +

 //" DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
 //" where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
 //" isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
 //" where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +


                            " DATEDIFF(DAY, (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID "+
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID "+
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else "+
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID "+
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end),  "+
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx "+
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, "+


                           " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ImpFreeDays,  " +
                           " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as ExpFreeDays, " +
                           " convert(varchar,DATEADD(DAY,(select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                           " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)),103) as FromDate, " +


                           //" case when(DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                           //" (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           //" where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                           //" isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           //" where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                           //" then DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), " +
                           //" (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           //" where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                           //" isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           //" where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +


                           
                          " case when(DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID), "+
                          " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                           " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))), " +
                           " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                           " then DATEDIFF(DAY, (DATEADD(DAY, (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID),  " +
                           " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                           " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                           " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))),  " +
                           " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                           " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +


                           " from NVO_ContainerTxns " +
                           " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
                           " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
                           " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                           " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID=NVO_ContainerTxns.AgencyID " +
                           " where NVO_ContainerTxns.StatusCode in ('MA') and  BLTypes in (40,42) and GeoLocationID= " + GeoLocID;

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND convert(varchar, NVO_ContainerTxns.DtMovement, 23) between '" + DtFrom + "' and '" + DtTo + "'";



            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }



        public void DetentionDemurageSummaryReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {

            ExcelPackage pck = new ExcelPackage();

            DataTable dt = GetAgencyLocation(DtFrom, DtTo, GeoLocID);
           
            for (int J = 0; J < dt.Rows.Count; J++)
            {
                var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
                ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:G2"];
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

                ws.Cells["A8:A10"].Value = "S. No.";
                ws.Cells["A8:A10"].Merge = true;
                ws.Cells["A8:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["B8:B10"].Value = "VESSEL/VOYAGE";
                ws.Cells["B8:B10"].Merge = true;
                ws.Cells["B8:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["C8:C10"].Value = "POL";
                ws.Cells["C8:C10"].Merge = true;
                ws.Cells["C8:C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["D8:D10"].Value = "POD";
                ws.Cells["D8:D10"].Merge = true;
                ws.Cells["D8:D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["E8:E10"].Value = "BL NUMBER";
                ws.Cells["E8:E10"].Merge = true;
                ws.Cells["E8:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                ws.Cells["F8:F10"].Value = "AGENCY NAME";
                ws.Cells["F8:F10"].Merge = true;
                ws.Cells["F8:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["G8:G10"].Value = "CONTAINER NO";
                ws.Cells["G8:G10"].Merge = true;
                ws.Cells["G8:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["H8:H10"].Value = "TYPE-SIZE";
                ws.Cells["H8:H10"].Merge = true;
                ws.Cells["H8:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                ws.Cells["I8"].Value = "PICK UP ";
                ws.Cells["I9"].Value = " Date";
                ws.Cells["I10"].Value = "A (MS)";

                ws.Cells["J8"].Value = "Free";
                ws.Cells["J9"].Value = " Time";
                ws.Cells["J10"].Value = "B";

                ws.Cells["K8"].Value = "DTN From";
                ws.Cells["K9"].Value = "(MS + F/Time)";
                ws.Cells["K10"].Value = "A+B = C";

                ws.Cells["L8"].Value = "DTN Upto";
                ws.Cells["L9"].Value = "PORT IN";
                ws.Cells["L10"].Value = "D (FL)";


                ws.Cells["M8"].Value = "DTN";
                ws.Cells["M9"].Value = "Days";
                ws.Cells["M10"].Value = "D-C";


                ws.Cells["N9"].Value = "SLAB1";
                ws.Cells["O9"].Value = "SLAB2";
                ws.Cells["P9"].Value = "SLAB3";
                ws.Cells["Q9"].Value = "SLAB4";
                ws.Cells["R9"].Value = "SLAB5";
                ws.Cells["S9"].Value = "SLAB6";

                ws.Cells["T9"].Value = "CURRENCY";

                ws.Cells["U8"].Value = "TOTAL";
                ws.Cells["U9"].Value = "USD";
                ws.Cells["U10"].Value = "COLLECTION";


                ExcelRange rng1 = ws.Cells["K8:U10"];
                rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng1.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);



                ws.Cells["A8:U10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                int sl = 1;
                int rw = 11;
                string Currency = "USD";
                DataTable dtv = GetExpDetentionSummaryReport(DtFrom, DtTo, dt.Rows[J]["GeoLocationID"].ToString());
                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    int CountDays = 0;
                    ws.Cells["A" + rw].Value = sl++;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["MSDate"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["ExpFreeDays"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["FLDate"].ToString();
                    int RowDays = 0;
                    if ((Int32.Parse(dtv.Rows[i]["Days1"].ToString())) > 1)
                    {
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["Days"].ToString();
                        RowDays = Int32.Parse(dtv.Rows[i]["Days"].ToString());
                    }
                    else
                        ws.Cells["M" + rw].Value = 0;


                    if (RowDays.ToString() != "0")
                    {
                        string FDatev = "";
                        string TDatev = "";
                        int RowsCount = 0;
                        string ID = "";
                        string LLimit = "";
                        string ULimit = "";
                        Currency = "";
                        string Amount = "";
                        double Days = 0;
                        int _FreeDays = 0;
                        decimal ExRate = 0;
                        FDatev = "";
                        TDatev = "";
                        DateTime IncreamentalFromDt, FromDt;
                        DateTime IncrementalToDt, ToDt;
                        TimeSpan TMS; bool IsexceedFreedays = false; string WaiverQty = "0";
                        DateTime.TryParse(dtv.Rows[i]["MSDatev"].ToString(), out IncreamentalFromDt);
                        DateTime.TryParse(dtv.Rows[i]["MSDatev"].ToString(), out FromDt);
                        DateTime.TryParse(dtv.Rows[i]["FLDatev"].ToString(), out ToDt);
                        WaiverQty = dtv.Rows[i]["ExpFreeDays"].ToString();
                        int.TryParse(WaiverQty, out _FreeDays);
                        DataTable _dtSlap = GetContractSlapExp(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                        if (_dtSlap.Rows.Count > 0)
                        {
                            for (int y = 0; y < _dtSlap.Rows.Count; y++)
                            {
                                RowsCount = y;
                                ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                                LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                                ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                                Currency = _dtSlap.Rows[y]["Currency"].ToString();
                                Amount = _dtSlap.Rows[y]["Amount"].ToString();
                                FDatev = IncreamentalFromDt.ToShortDateString();
                                IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                                TDatev = IncrementalToDt.ToShortDateString();

                                if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                                {
                                    TDatev = ToDt.ToShortDateString();
                                    TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                    Days = (int)((TMS.TotalDays + 1));

                                    if (!IsexceedFreedays)
                                    {
                                        TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays + 1));
                                        Days = Days - _FreeDays;
                                        if (Days < 0)
                                            Days = 0;
                                        else
                                            IsexceedFreedays = true;
                                    }
                                    else
                                    {
                                        //TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        //Days = (int)((TMS.TotalDays + 1));
                                        //if (Days <= 0)
                                        //    Days = 1;
                                        TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays + 1));
                                        if (Days <= 0)
                                            Days = 0;
                                    }
                                }
                                else
                                {
                                    if (!IsexceedFreedays)
                                    {
                                        TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays + 1));
                                        Days = Days - _FreeDays;
                                        if (Days < 0)
                                            Days = 0;
                                        else
                                            IsexceedFreedays = true;
                                    }
                                    else
                                    {
                                        TMS = IncrementalToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays + 1));

                                    }
                                }
                                IncreamentalFromDt = IncrementalToDt.AddDays(1);

                                string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                                if (y == 0)
                                {
                                    ws.Cells["N" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["N" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 1)
                                {
                                    ws.Cells["O" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["O" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 2)
                                {
                                    ws.Cells["P" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["P" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 3)
                                {
                                    ws.Cells["Q" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["Q" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 4)
                                {
                                    ws.Cells["R" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["R" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 5)
                                {
                                    ws.Cells["S" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["S" + rw].Style.Numberformat.Format = "#,##0.00";
                                }


                            }
                            if (RowsCount == 1)
                            {
                                ws.Cells["P" + rw].Value = 0.00;
                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 2)
                            {

                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 3)
                            {

                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 4)
                            {

                                ws.Cells["S" + rw].Value = 0.00;
                            }
                            if (RowsCount == 5)
                            {
                                ws.Cells["S" + rw].Value = 0.00;
                            }
                        }
                        else
                        {
                            ws.Cells["N" + rw].Value = 0.00;
                            ws.Cells["O" + rw].Value = 0.00;
                            ws.Cells["P" + rw].Value = 0.00;
                            ws.Cells["Q" + rw].Value = 0.00;
                            ws.Cells["R" + rw].Value = 0.00;
                            ws.Cells["S" + rw].Value = 0.00;
                        }
                    }
                    else
                    {
                        ws.Cells["N" + rw].Value = 0.00;
                        ws.Cells["O" + rw].Value = 0.00;
                        ws.Cells["P" + rw].Value = 0.00;
                        ws.Cells["Q" + rw].Value = 0.00;
                        ws.Cells["R" + rw].Value = 0.00;
                        ws.Cells["S" + rw].Value = 0.00;
                    }
                    ws.Cells["T" + rw].Value = Currency;
                    string froMulaAddQV = string.Format("=SUM(M" + rw + ":R" + rw + ")");
                    ws.Cells["U" + rw].Formula = froMulaAddQV.ToString();
                    ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                    rw++;
                }



                ws.Cells["A8:U10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:U10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 50].AutoFitColumns();
            }
            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=ExpDetentionSummaryReport.xlsx");
            Response.End();

        }


        public void EQC_Dumarage_Detention_ReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {


            ExcelPackage pck = new ExcelPackage();

            //DataTable dt = GetAgencyLocation(DtFrom, DtTo, GeoLocID);
            //for (int J = 0; J < dt.Rows.Count; J++)
            //{
            //var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
            var ws = pck.Workbook.Worksheets.Add("AAA");
            ws.Cells["A2"].Value = "DEMURRAGE / DETENTION COLLECTIONS";
            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ExcelRange r = ws.Cells["A2:G2"];
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

            ws.Cells["A8:A10"].Value = "S. No.";
            ws.Cells["A8:A10"].Merge = true;
            ws.Cells["A8:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["B8:B10"].Value = "VESSEL/VOYAGE";
            ws.Cells["B8:B10"].Merge = true;
            ws.Cells["B8:B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["C8:C10"].Value = "POL";
            ws.Cells["C8:C10"].Merge = true;
            ws.Cells["C8:C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["D8:D10"].Value = "POD";
            ws.Cells["D8:D10"].Merge = true;
            ws.Cells["D8:D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["E8:E10"].Value = "BL NUMBER";
            ws.Cells["E8:E10"].Merge = true;
            ws.Cells["E8:E10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["F8:F10"].Value = "AGENCY NAME";
            ws.Cells["F8:F10"].Merge = true;
            ws.Cells["F8:F10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["G8:G10"].Value = "CONTAINER NO";
            ws.Cells["G8:G10"].Merge = true;
            ws.Cells["G8:G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["H8:H10"].Value = "TYPE-SIZE";
            ws.Cells["H8:H10"].Merge = true;
            ws.Cells["H8:H10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


            ws.Cells["I8"].Value = "Arrival";
            ws.Cells["I9"].Value = " Date";
            ws.Cells["I10"].Value = "A (FV)";

            ws.Cells["J8"].Value = "DEM Free";
            ws.Cells["J9"].Value = " Time";
            ws.Cells["J10"].Value = "B";


            ExcelRange rng1 = ws.Cells["A8:J10"];
            rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng1.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            // FIREST HEAD



            ws.Cells["K8"].Value = "DEM From";
            ws.Cells["K9"].Value = " (FV + F/Time)";
            ws.Cells["K10"].Value = "A+B = C";

            ws.Cells["L8"].Value = "DEM Upto";
            ws.Cells["L9"].Value = "";
            ws.Cells["L10"].Value = "D (FU)";

            ws.Cells["M8"].Value = "DEM";
            ws.Cells["M9"].Value = "Days";
            ws.Cells["M10"].Value = "D-C";


            ws.Cells["N8:R8"].Value = "DEMURRAGE";
            ws.Cells["N8:R8"].Merge = true;
            ws.Cells["N8:R8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["N9"].Value = "SLAB1";
            ws.Cells["O9"].Value = "SLAB2";
            ws.Cells["P9"].Value = "SLAB3";
            ws.Cells["Q9"].Value = "SLAB4";
            ws.Cells["R9"].Value = "Total Demurrage";


            ExcelRange rng2 = ws.Cells["K8:R10"];
            rng2.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng2.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);


            ws.Cells["S8"].Value = "DET Free";
            ws.Cells["S9"].Value = "Time";
            ws.Cells["S10"].Value = "E";

            ExcelRange rng3 = ws.Cells["S8:S10"];
            rng3.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng3.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

            ws.Cells["T8"].Value = "DET From";
            ws.Cells["T9"].Value = "(FU + F/Time)";
            ws.Cells["T10"].Value = "D+E = F";

            ws.Cells["U8"].Value = "DTN Upto";
            ws.Cells["U9"].Value = "PORT IN";
            ws.Cells["U10"].Value = "G (MA)";

            ws.Cells["V8"].Value = "DTN";
            ws.Cells["V9"].Value = "Days";
            ws.Cells["V10"].Value = "G-F";

            ws.Cells["W8:AB8"].Value = "DETENTION";
            ws.Cells["W8:AB8"].Merge = true;
            ws.Cells["W8:AB8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["W9"].Value = "SLAB1";
            ws.Cells["X9"].Value = "SLAB2";
            ws.Cells["Y9"].Value = "SLAB3";
            ws.Cells["Z9"].Value = "SLAB4";
            ws.Cells["AA9"].Value = "Total Detention";
            ws.Cells["AA10"].Value = "";
            ws.Cells["AB8"].Value = "";
            ws.Cells["AB9"].Value = "CURRENCY";
            ws.Cells["AB10"].Value = "";

            ws.Cells["AC8"].Value = "TOTAL";
            ws.Cells["AC9"].Value = "COLLECTION IN USD";
            ws.Cells["AC10"].Value = "DEM+DET";

            ExcelRange rng4 = ws.Cells["T8:AC10"];
            rng4.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rng4.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            ws.Cells["A8:AC10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A8:AC10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A8:AC10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A8:AC10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            int sl = 1;
            int rw = 11;
            DataTable dtv = GetEqc_Detention_Dumarage_Report(DtFrom, DtTo, GeoLocID);
            for (int i = 0; i < dtv.Rows.Count; i++)
            {
                ws.Cells["A" + rw].Value = sl++;
                ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                ws.Cells["C" + rw].Value = dtv.Rows[i]["POL"].ToString();
                ws.Cells["D" + rw].Value = dtv.Rows[i]["POD"].ToString();
                ws.Cells["E" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                ws.Cells["F" + rw].Value = dtv.Rows[i]["AgencyName"].ToString();
                ws.Cells["G" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                ws.Cells["H" + rw].Value = dtv.Rows[i]["CntrTypes"].ToString();
                ws.Cells["I" + rw].Value = dtv.Rows[i]["FVDate"].ToString();
                ws.Cells["J" + rw].Value = dtv.Rows[i]["DEMFreeDays"];
                ws.Cells["K" + rw].Value = dtv.Rows[i]["FVAB"].ToString();
                ws.Cells["l" + rw].Value = dtv.Rows[i]["FUDate"].ToString();

                if (dtv.Rows[i]["DEMDays"].ToString() != "")
                {
                    if (Int32.Parse(dtv.Rows[i]["DEMDays"].ToString()) > 0)
                        ws.Cells["m" + rw].Value = dtv.Rows[i]["DEMDays"];
                    else
                        ws.Cells["m" + rw].Value = 0;
                }
                else
                {
                    ws.Cells["m" + rw].Value = 0;
                }

                ws.Cells["s" + rw].Value = dtv.Rows[i]["DETFreeDays"].ToString();
                ws.Cells["t" + rw].Value = dtv.Rows[i]["DETFrom"].ToString();
                ws.Cells["u" + rw].Value = dtv.Rows[i]["MADate"].ToString();
                ws.Cells["v" + rw].Value = dtv.Rows[i]["DETDays"].ToString();

                if (dtv.Rows[i]["DETDays"].ToString() != "")
                {
                    if (Int32.Parse(dtv.Rows[i]["DETDays"].ToString()) > 0)
                        ws.Cells["v" + rw].Value = dtv.Rows[i]["DETDays"];
                    else
                        ws.Cells["v" + rw].Value = 0;
                }
                else
                {
                    ws.Cells["v" + rw].Value = 0;
                }

                string FDatev = "";
                string TDatev = "";
                int RowsCount = 0;
                string ID = "";
                string LLimit = "";
                string ULimit = "";
                string Currency = "";
                string Amount = "";
                double Days = 0;
                double DemDays = 0;
                int _FreeDays = 0;
                int DemFreeday = 0;
                decimal ExRate = 0;
                FDatev = "";
                TDatev = "";
                DateTime IncreamentalFromDt, FromDt;
                DateTime IncrementalToDt, ToDt;

                DateTime DEMIncreamentalFromDt, DEMFromDt;
                DateTime DEMIncrementalToDt, DEMToDt;

                TimeSpan TMS; bool IsexceedFreedays = false; string WaiverQty = "0";
                TimeSpan DemTMS; bool DemIsexceedFreedays = false;
                string DemWaiverQty = "0";
                DateTime.TryParse(dtv.Rows[i]["FUDate"].ToString(), out IncreamentalFromDt);
                DateTime.TryParse(dtv.Rows[i]["FUDate"].ToString(), out FromDt);
                DateTime.TryParse(dtv.Rows[i]["MADate"].ToString().ToString(), out ToDt);

                DateTime.TryParse(dtv.Rows[i]["FVDate"].ToString(), out DEMIncreamentalFromDt);
                DateTime.TryParse(dtv.Rows[i]["FVDate"].ToString(), out DEMFromDt);
                DateTime.TryParse(dtv.Rows[i]["FUDate"].ToString().ToString(), out DEMToDt);


                WaiverQty = dtv.Rows[i]["DETFreeDays"].ToString();
                DemWaiverQty = dtv.Rows[i]["DEMFreeDays"].ToString();
                int.TryParse(WaiverQty, out _FreeDays);
                int.TryParse(DemWaiverQty, out DemFreeday);
                DataTable _dtSlap = GetContractSlap(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                if (_dtSlap.Rows.Count > 0)
                {
                    for (int y = 0; y < _dtSlap.Rows.Count; y++)
                    {
                        RowsCount = y;

                        ExRate = decimal.Parse(_dtSlap.Rows[y]["ExRate"].ToString());
                        LLimit = _dtSlap.Rows[y]["SlabFrom"].ToString();
                        ULimit = _dtSlap.Rows[y]["SlabTo"].ToString();
                        Currency = _dtSlap.Rows[y]["Currency"].ToString();
                        Amount = _dtSlap.Rows[y]["Amount"].ToString();
                        FDatev = IncreamentalFromDt.ToShortDateString();
                        IncrementalToDt = FromDt.AddDays(float.Parse(_dtSlap.Rows[y]["SlabTo"].ToString()) - 1);
                        TDatev = IncrementalToDt.ToShortDateString();

                        if (IncrementalToDt >= ToDt || IncrementalToDt < FromDt)
                        {
                            TDatev = ToDt.ToShortDateString();
                            TMS = ToDt.Subtract(DateTime.Parse(IncreamentalFromDt.ToShortDateString()));
                            Days = (int)((TMS.TotalDays + 1));

                            if (!IsexceedFreedays)
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = ToDt.Subtract(DateTime.Parse(FDatev));
                                Days = (int)((TMS.TotalDays + 1));
                                if (Days <= 0)
                                    Days = 0;
                            }
                        }
                        else
                        {
                            if (!IsexceedFreedays)
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                Days = (int)((TMS.TotalDays + 1));
                                Days = Days - _FreeDays;
                                if (Days < 0)
                                    Days = 0;
                                else
                                    IsexceedFreedays = true;
                            }
                            else
                            {
                                TMS = IncrementalToDt.Subtract(DateTime.Parse(FDatev));
                                Days = (int)((TMS.TotalDays + 1));

                            }
                        }
                        IncreamentalFromDt = IncrementalToDt.AddDays(1);
                        string Total = (decimal.Parse(Amount) * decimal.Parse(Days.ToString())).ToString();

                        if (y == 0)
                        {
                            ws.Cells["w" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["w" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 1)
                        {
                            ws.Cells["x" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["x" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 2)
                        {
                            ws.Cells["y" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["y" + rw].Style.Numberformat.Format = "#,##0.00";
                        }
                        if (y == 3)
                        {
                            ws.Cells["z" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                            ws.Cells["z" + rw].Style.Numberformat.Format = "#,##0.00";
                        }

                    }
                    if (RowsCount == 1)
                    {
                        ws.Cells["w" + rw].Value = 0.00;
                        ws.Cells["x" + rw].Value = 0.00;
                        ws.Cells["y" + rw].Value = 0.00;
                        ws.Cells["z" + rw].Value = 0.00;
                        // ws.Cells["R" + rw].Value = 0.00;
                    }
                    if (RowsCount == 2)
                    {
                        ws.Cells["x" + rw].Value = 0.00;
                        ws.Cells["y" + rw].Value = 0.00;
                        ws.Cells["z" + rw].Value = 0.00;
                        //ws.Cells["R" + rw].Value = 0.00;
                    }
                    if (RowsCount == 3)
                    {
                        ws.Cells["y" + rw].Value = 0.00;
                        //ws.Cells["z" + rw].Value = 0.00;
                        //ws.Cells["R" + rw].Value = 0.00;
                    }
                    //if (RowsCount == 4)
                    //{
                    //    ws.Cells["Q" + rw].Value = 0.00;
                    //    ws.Cells["R" + rw].Value = 0.00;
                    //}
                    //if (RowsCount == 5)
                    //{
                    //    ws.Cells["R" + rw].Value = 0.00;
                    //}
                }
                else
                {
                    ws.Cells["w" + rw].Value = 0.00;
                    ws.Cells["x" + rw].Value = 0.00;
                    ws.Cells["y" + rw].Value = 0.00;
                    ws.Cells["z" + rw].Value = 0.00;

                }

                //Dem
                if (dtv.Rows[i]["DEMDays"].ToString() != "")
                {
                    string fff = dtv.Rows[i]["DEMDays"].ToString();
                    if (Int32.Parse(dtv.Rows[i]["DEMDays"].ToString()) > 0)
                    {
                        DataTable _dtDemSlap = GetDemurageContractSlap(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString());
                        if (_dtDemSlap.Rows.Count > 0)
                        {
                            for (int y = 0; y < _dtDemSlap.Rows.Count; y++)
                            {
                                RowsCount = y;

                                ExRate = decimal.Parse(_dtDemSlap.Rows[y]["ExRate"].ToString());
                                LLimit = _dtDemSlap.Rows[y]["SlabFrom"].ToString();
                                ULimit = _dtDemSlap.Rows[y]["SlabTo"].ToString();
                                Currency = _dtDemSlap.Rows[y]["Currency"].ToString();
                                Amount = _dtDemSlap.Rows[y]["Amount"].ToString();
                                FDatev = DEMIncreamentalFromDt.ToShortDateString();
                                var slap = _dtDemSlap.Rows[y]["SlabTo"].ToString();
                                DEMIncrementalToDt = DEMFromDt.AddDays(float.Parse(_dtDemSlap.Rows[y]["SlabTo"].ToString()) - 1);
                                TDatev = DEMIncrementalToDt.ToShortDateString();

                                if (DEMIncrementalToDt >= DEMToDt || DEMIncrementalToDt < DEMFromDt)
                                {
                                    TDatev = DEMToDt.ToShortDateString();
                                    DemTMS = DEMToDt.Subtract(DateTime.Parse(DEMIncreamentalFromDt.ToShortDateString()));
                                    DemDays = (int)((DemTMS.TotalDays + 1));

                                    if (!IsexceedFreedays)
                                    {
                                        DemTMS = DEMToDt.Subtract(DateTime.Parse(DEMFromDt.ToShortDateString()));
                                        DemDays = (int)((DemTMS.TotalDays + 1));
                                        DemDays = DemDays - DemFreeday;
                                        if (DemDays < 0)
                                            DemDays = 0;
                                        else
                                            IsexceedFreedays = true;
                                    }
                                    else
                                    {
                                        DemTMS = DEMToDt.Subtract(DateTime.Parse(DEMIncreamentalFromDt.ToShortDateString()));
                                        DemDays = (int)((DemTMS.TotalDays + 1));
                                        if (DemDays <= 0)
                                            DemDays = 0;
                                    }
                                }
                                else
                                {
                                    if (!IsexceedFreedays)
                                    {
                                        DemTMS = DEMIncrementalToDt.Subtract(DateTime.Parse(DEMFromDt.ToShortDateString()));
                                        DemDays = (int)((DemTMS.TotalDays + 1));
                                        DemDays = DemDays - DemFreeday;
                                        if (DemDays < 0)
                                            DemDays = 0;
                                        else
                                            IsexceedFreedays = true;
                                    }
                                    else
                                    {
                                        DemTMS = DEMIncrementalToDt.Subtract(DateTime.Parse(FDatev));
                                        DemDays = (int)((DemTMS.TotalDays + 1));

                                    }
                                }
                                DEMIncreamentalFromDt = DEMIncrementalToDt.AddDays(1);
                                string Total = (decimal.Parse(Amount) * decimal.Parse(DemDays.ToString())).ToString();

                                //string Total = (decimal.Parse(Amount) * decimal.Parse(dtv.Rows[i]["DEMDays"].ToString())).ToString();


                                if (y == 0)
                                {
                                    ws.Cells["n" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(DemDays.ToString()));
                                    ws.Cells["n" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 1)
                                {
                                    ws.Cells["o" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(DemDays.ToString()));
                                    ws.Cells["o" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 2)
                                {
                                    ws.Cells["p" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(DemDays.ToString()));
                                    ws.Cells["p" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 3)
                                {
                                    ws.Cells["q" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(DemDays.ToString()));
                                    ws.Cells["q" + rw].Style.Numberformat.Format = "#,##0.00";
                                }


                            }
                            if (RowsCount == 1)
                            {
                                // ws.Cells["n" + rw].Value = 0.00;
                                ws.Cells["o" + rw].Value = 0.00;
                                ws.Cells["p" + rw].Value = 0.00;
                                ws.Cells["q" + rw].Value = 0.00;
                                // ws.Cells["R" + rw].Value = 0.00;
                            }
                            if (RowsCount == 2)
                            {
                                //ws.Cells["o" + rw].Value = 0.00;
                                ws.Cells["p" + rw].Value = 0.00;
                                ws.Cells["q" + rw].Value = 0.00;
                                //ws.Cells["R" + rw].Value = 0.00;
                            }
                            if (RowsCount == 3)
                            {
                                //ws.Cells["q" + rw].Value = 0.00;
                                //ws.Cells["q" + rw].Value = 0.00;
                                //ws.Cells["R" + rw].Value = 0.00;
                            }


                        }
                        else
                        {
                            ws.Cells["n" + rw].Value = 0.00;
                            ws.Cells["o" + rw].Value = 0.00;
                            ws.Cells["p" + rw].Value = 0.00;
                            ws.Cells["q" + rw].Value = 0.00;

                        }
                    }
                    else
                    {
                        ws.Cells["n" + rw].Value = 0.00;
                        ws.Cells["o" + rw].Value = 0.00;
                        ws.Cells["p" + rw].Value = 0.00;
                        ws.Cells["q" + rw].Value = 0.00;

                    }
                }
                else
                {
                    ws.Cells["n" + rw].Value = 0.00;
                    ws.Cells["o" + rw].Value = 0.00;
                    ws.Cells["p" + rw].Value = 0.00;
                    ws.Cells["q" + rw].Value = 0.00;
                }
                string froMulaAddQV1 = string.Format("=SUM(n" + rw + ":q" + rw + ")");
                ws.Cells["r" + rw].Formula = froMulaAddQV1.ToString();
                ws.Cells["r" + rw].Style.Numberformat.Format = "#,##0.00";
                //end
                ws.Cells["ab" + rw].Value = Currency;
                string froMulaAddQV = string.Format("=SUM(w" + rw + ":z" + rw + ")");
                ws.Cells["aa" + rw].Formula = froMulaAddQV.ToString();
                ws.Cells["aa" + rw].Style.Numberformat.Format = "#,##0.00";

                string froMulaAddQV2 = string.Format("=SUM(r" + rw + "+aa" + rw + ")");
                ws.Cells["ac" + rw].Formula = froMulaAddQV2.ToString();
                ws.Cells["ac" + rw].Style.Numberformat.Format = "#,##0.00";
                rw++;
            }



            ws.Cells["A8:AC10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A8:AC10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A8:AC10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A8:AC10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            ws.Cells[1, 1, rw, 50].AutoFitColumns();
            // }
            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=DEMURRAGEDETENTIONCOLLECTIONSReport.xlsx");
            Response.End();

        }




        public DataTable GetEqc_Detention_Dumarage_Report(string DtFrom, string DtTo, string GeoLocID)
        {
            string strWhere = "";




            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID,   " +
                            " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,   " +
                            "(select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +
                            " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDate,3 as DEMFreeDays, " +
                            " convert(varchar, (DATEADD(day, (3 - 1), " +
                            " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else  " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))), 103) AS FVAB,  " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID   " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FUDate,   " +
                            " DATEDIFF(DAY, DATEADD(day, (3 - 1),  " +
                            " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else  " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end)), (select top(1) DtMovement from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) AS DEMDays,  " +
                            " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) as DETFreeDays, " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) AS DETFrom,  " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID   " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MADate,  " +

                            " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx   " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID   " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), (DATEADD(day, ((select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID)), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)))) AS DTNDays,   " +
                            " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else  " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDatev,  " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MADatev,  " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID  " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) +1 as MADatev1,  " +
                            " isnull(DATEDIFF(DAY, DATEADD(day, ((select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID) - 1),  (select top(1) DtMovement from NVO_ContainerTxns CnTx  " +
                            " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)), (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)),0) AS DETDays " +
                            " from NVO_ContainerTxns  " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID  " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber  " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID  " +
                            " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID = NVO_ContainerTxns.AgencyID  " +
                            " where NVO_ContainerTxns.StatusCode in ('MA', 'FU') and  " +
                            " BLTypes in (40, 42) " +
                            //" and NVO_BOL.BLNumber in ('OCLJEA23061443PKG') and CntrNo='FSCU6641572'";
                            " and GeoLocationID= " + GeoLocID;

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND convert(varchar, NVO_ContainerTxns.DtMovement, 23) between '" + DtFrom + "' and '" + DtTo + "'";


            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

    }
}