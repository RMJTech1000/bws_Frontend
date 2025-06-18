using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text;
using DataManager;

namespace NVOCShipping.Controllers
{
    public class StorageRentalController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: StorageRental
        public ActionResult StorageRental()
        {
            return View();
        }

        public ActionResult StorageRentalView()
        {
            return View();
        }

        public ActionResult ImportPortStorageSummaryReport()
        {
            return View();
        }

        public ActionResult ExportPortStorageSummaryReport()
        {
            return View();
        }

        public ActionResult FreedaysReports()
        {

            return View();
        }
        public void EQCImpPortStorageSummaryReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {

            ExcelPackage pck = new ExcelPackage();

            DataTable dt = GetAgencyLocation(DtFrom, DtTo, GeoLocID);
            for (int J = 0; J < dt.Rows.Count; J++)
            {
                var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
                ws.Cells["A2"].Value = "PORT STORAGE COLLECTIONS";
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




                ws.Cells["I8"].Value = "PORT STORAGE ";
                ws.Cells["I9"].Value = " FROM";
                ws.Cells["I10"].Value = "A (FV)";
                //OLD MS
                ws.Cells["J8"].Value = "PORT STORAGE";
                ws.Cells["J9"].Value = " TO";
                ws.Cells["J10"].Value = "B(FU)";

                ws.Cells["K8"].Value = "PORT STORAGE";
                ws.Cells["K9"].Value = "TOTAL DAYS";
                ws.Cells["K10"].Value = "A+B = C";

                ws.Cells["L8"].Value = "PORT FREE";
                ws.Cells["L9"].Value = "TIME";
                ws.Cells["L10"].Value = "D";


                ws.Cells["M8"].Value = "PORT FREE";
                ws.Cells["M9"].Value = "CHARGEABLE DAYS";
                ws.Cells["M10"].Value = "D-C";


                ws.Cells["N9"].Value = "SLAB1";
                ws.Cells["O9"].Value = "SLAB2";
                ws.Cells["P9"].Value = "SLAB3";
                ws.Cells["Q9"].Value = "SLAB4";
                ws.Cells["R9"].Value = "SLAB5";
                ws.Cells["S9"].Value = "SLAB6";
                ws.Cells["T9"].Value = "SLAB7";
                ws.Cells["U9"].Value = "SLAB8";
                ws.Cells["V9"].Value = "SLAB9";

                ws.Cells["W9"].Value = "CURRENCY";

                ws.Cells["X8"].Value = "TOTAL";
                ws.Cells["X9"].Value = "USD";
                ws.Cells["X10"].Value = "COLLECTION";


                ExcelRange rng1 = ws.Cells["K8:X10"];
                rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng1.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);



                ws.Cells["A8:X10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                int sl = 1;
                int rw = 11;
                string Currency = "USD";
                DataTable dtv = GetImpPortStorageSummaryReport(DtFrom, DtTo, dt.Rows[J]["GeoLocationID"].ToString());
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
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["PortFreeDays"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["FUDate"].ToString();
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
                        DateTime.TryParse(dtv.Rows[i]["FUDatev"].ToString().ToString(), out ToDt);
                        WaiverQty = dtv.Rows[i]["PortFreeDays"].ToString();
                        int.TryParse(WaiverQty, out _FreeDays);
                        DataTable _dtSlap = GetContractSlap(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString(), dtv.Rows[i]["POLID"].ToString());
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
                                    ws.Cells["S" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["S" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 6)
                                {
                                    ws.Cells["T" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["T" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 7)
                                {
                                    ws.Cells["U" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                                }

                                if (y == 8)
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

                                ws.Cells["P" + rw].Value = 0.00;
                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                                ws.Cells["T" + rw].Value = 0.00;
                                ws.Cells["U" + rw].Value = 0.00;
                                ws.Cells["V" + rw].Value = 0.00;
                            }
                            if (RowsCount == 2)
                            {

                                ws.Cells["Q" + rw].Value = 0.00;
                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                                ws.Cells["T" + rw].Value = 0.00;
                                ws.Cells["U" + rw].Value = 0.00;
                                ws.Cells["V" + rw].Value = 0.00;
                            }
                            if (RowsCount == 3)
                            {

                                ws.Cells["R" + rw].Value = 0.00;
                                ws.Cells["S" + rw].Value = 0.00;
                                ws.Cells["T" + rw].Value = 0.00;
                                ws.Cells["U" + rw].Value = 0.00;
                                ws.Cells["V" + rw].Value = 0.00;
                            }
                            if (RowsCount == 4)
                            {

                                ws.Cells["S" + rw].Value = 0.00;
                                ws.Cells["T" + rw].Value = 0.00;
                                ws.Cells["U" + rw].Value = 0.00;
                                ws.Cells["V" + rw].Value = 0.00;
                            }
                            if (RowsCount == 5)
                            {
                                ws.Cells["T" + rw].Value = 0.00;
                           
                                ws.Cells["U" + rw].Value = 0.00;
                                ws.Cells["V" + rw].Value = 0.00;
                            }
                            if (RowsCount == 6)
                            {
                                ws.Cells["U" + rw].Value = 0.00;
                                ws.Cells["V" + rw].Value = 0.00;
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
                            ws.Cells["T" + rw].Value = 0.00;
                            ws.Cells["U" + rw].Value = 0.00;
                            ws.Cells["V" + rw].Value = 0.00;
                        }
                        ws.Cells["W" + rw].Value = Currency;
                    }
                    else
                    {
                        ws.Cells["N" + rw].Value = 0.00;
                        ws.Cells["O" + rw].Value = 0.00;
                        ws.Cells["P" + rw].Value = 0.00;
                        ws.Cells["Q" + rw].Value = 0.00;
                        ws.Cells["R" + rw].Value = 0.00;
                        ws.Cells["S" + rw].Value = 0.00;
                        ws.Cells["T" + rw].Value = 0.00;
                        ws.Cells["U" + rw].Value = 0.00;
                        ws.Cells["V" + rw].Value = 0.00;
                    }


                    string froMulaAddQV = string.Format("=SUM(M" + rw + ":R" + rw + ")");
                    ws.Cells["X" + rw].Formula = froMulaAddQV.ToString();
                    ws.Cells["X" + rw].Style.Numberformat.Format = "#,##0.00";
                    rw++;
                }



                ws.Cells["A8:X10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 50].AutoFitColumns();
            }
            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Import_PortStorage_SummaryReport.xlsx");
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

        public DataTable GetImpPortStorageSummaryReport(string DtFrom, string DtTo, string GeoLocID)
        {
            string strWhere = "";


            //string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
            //               " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,NVO_BOL.POLID,NVO_BOL.PODID,  " +
            //               " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +


            //               " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //               " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDate, " +

            //               " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
            //               " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MADate, " +

            //                " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDatev,  " +

            //               " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //               " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as MADatev, " +


            //                " DATEDIFF(DAY, (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end),  " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +


            //                " 3 as PortFreeDays,"+
            //               " DATEADD(DAY,3, " +
            //               " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //               " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

            //              " case when(DATEDIFF(DAY, (DATEADD(DAY, 3, " +
            //              " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //              " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //               " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))), " +
            //               " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //               " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
            //               " then DATEDIFF(DAY, (DATEADD(DAY, 3,  " +
            //               " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
            //               " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //               " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))),  " +
            //               " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //               " where CnTx.StatusCode = 'MA' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +


            //               " from NVO_ContainerTxns " +
            //               " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
            //               " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
            //               " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
            //               " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID=NVO_ContainerTxns.AgencyID " +
            //               " where NVO_ContainerTxns.StatusCode in ('MA') and  BLTypes in (40,42) and GeoLocationID= " + GeoLocID;


            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
                         " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,NVO_BOL.POLID,NVO_BOL.PODID,  " +
                         " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +


                         " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                         " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDate, " +

                         " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                         " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FUDate, " +

                          " case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                          " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end FVDatev,  " +

                         " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                         " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as FUDatev, " +


                          " DATEDIFF(DAY, (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                          " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                          " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end),  " +
                          " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                          " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +


                          " 3 as PortFreeDays," +
                         " DATEADD(DAY,3, " +
                         " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                         " where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

                        " case when(DATEDIFF(DAY, (DATEADD(DAY, 3, " +
                        " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                        " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), 0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                         " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))), " +
                         " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                         " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                         " then DATEDIFF(DAY, (DATEADD(DAY, 3,  " +
                         " (case when isnull((select top(1) ID from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber),0) = 0 then(select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FV' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) else " +
                         " (select top(1) DtMovement from NVO_ContainerTxns CnTx  where CnTx.StatusCode = 'FVICD' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                         " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) end))),  " +
                         " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                         " where CnTx.StatusCode = 'FU' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv " +


                         " from NVO_ContainerTxns " +
                         " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
                         " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
                         " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                         " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID=NVO_ContainerTxns.AgencyID " +
                         " where NVO_ContainerTxns.StatusCode in ('FU') and  BLTypes in (40,42) and GeoLocationID= " + GeoLocID;

            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND convert(varchar, NVO_ContainerTxns.DtMovement, 23) between '" + DtFrom + "' and '" + DtTo + "'";



            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetContractSlap(string AgencyID, string CntrTypes, string PortID)
        {
            string strWhere = "";
            string _Query = " select isnull((select top(1)  Rate from NVO_ExRate where AgencyID= NVO_StorageContract.AgencyID and FromCurrency= NVO_StorageContractTariffDtls.CurrencyID order by Id desc),0) as ExRate," +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where ID =NVO_StorageContractTariffDtls.CurrencyID) as Currency,* from NVO_StorageContract inner join NVO_StorageContractTariffDtls on NVO_StorageContractTariffDtls.RentID = NVO_StorageContract.ID " +
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 2 and ChargesID=42 and PortID =" + PortID;
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }


        public void ExpPortStorageSummaryReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {

            ExcelPackage pck = new ExcelPackage();

            DataTable dt = GetAgencyLocation(DtFrom, DtTo, GeoLocID);
            for (int J = 0; J < dt.Rows.Count; J++)
            {
                var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
                ws.Cells["A2"].Value = "EXPORT PORT STORAGE COLLECTIONS";
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


                ws.Cells["I8"].Value = "PORT STORAGE ";
                ws.Cells["I9"].Value = " FROM";
                ws.Cells["I10"].Value = "A (FL)";
                //OLD MS
                ws.Cells["J8"].Value = "PORT STORAGE";
                ws.Cells["J9"].Value = " TO";
                ws.Cells["J10"].Value = "B(FB)";

                ws.Cells["K8"].Value = "PORT STORAGE";
                ws.Cells["K9"].Value = "TOTAL DAYS";
                ws.Cells["K10"].Value = "A+B = C";

                ws.Cells["L8"].Value = "PORT FREE";
                ws.Cells["L9"].Value = "TIME";
                ws.Cells["L10"].Value = "D";


                ws.Cells["M8"].Value = "PORT FREE";
                ws.Cells["M9"].Value = "CHARGEABLE DAYS";
                ws.Cells["M10"].Value = "D-C";


                ws.Cells["N9"].Value = "SLAB1";
                ws.Cells["O9"].Value = "SLAB2";
                ws.Cells["P9"].Value = "SLAB3";
                ws.Cells["Q9"].Value = "SLAB4";
                ws.Cells["R9"].Value = "SLAB5";
                ws.Cells["S9"].Value = "SLAB6";
                ws.Cells["T9"].Value = "SLAB7";
                ws.Cells["U9"].Value = "SLAB8";
                ws.Cells["V9"].Value = "SLAB9";

                ws.Cells["W9"].Value = "CURRENCY";

                ws.Cells["X8"].Value = "TOTAL";
                ws.Cells["X9"].Value = "USD";
                ws.Cells["X10"].Value = "COLLECTION";


                ExcelRange rng1 = ws.Cells["K8:X10"];
                rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng1.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);



                ws.Cells["A8:X10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                int sl = 1;
                int rw = 11;
                string Currency = "USD";
                DataTable dtv = GetExpPortStorageSummaryReport(DtFrom, DtTo, dt.Rows[J]["GeoLocationID"].ToString());
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
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["FBDate"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["PortFreeDays"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["FromDate"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["FLDate"].ToString();


                    int RowDays = 0;
                    if ((Int32.Parse(dtv.Rows[i]["PortFreeDays"].ToString())) > 1)
                    {
                        ws.Cells["M" + rw].Value = dtv.Rows[i]["PortFreeDays"].ToString();
                        RowDays = Int32.Parse(dtv.Rows[i]["PortFreeDays"].ToString());
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
                        DateTime.TryParse(dtv.Rows[i]["FBDatev"].ToString(), out IncreamentalFromDt);
                        DateTime.TryParse(dtv.Rows[i]["FBDatev"].ToString(), out FromDt);
                        DateTime.TryParse(dtv.Rows[i]["FLDatev"].ToString(), out ToDt);
                        WaiverQty = dtv.Rows[i]["PortFreeDays"].ToString();
                        int.TryParse(WaiverQty, out _FreeDays);
                        DataTable _dtSlap = GetContractSlapExp(dtv.Rows[i]["AgencyID"].ToString(), dtv.Rows[i]["TypeID"].ToString(), dtv.Rows[i]["PODID"].ToString());
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
                                    Days = (int)((TMS.TotalDays));

                                    if (!IsexceedFreedays)
                                    {
                                        TMS = ToDt.Subtract(DateTime.Parse(FromDt.ToShortDateString()));
                                        Days = (int)((TMS.TotalDays));
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
                                        Days = (int)((TMS.TotalDays));
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
                                if (y == 6)
                                {
                                    ws.Cells["T" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["T" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 7)
                                {
                                    ws.Cells["U" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["U" + rw].Style.Numberformat.Format = "#,##0.00";
                                }
                                if (y == 8)
                                {
                                    ws.Cells["V" + rw].Value = (decimal.Parse(Amount) * decimal.Parse(Days.ToString()));
                                    ws.Cells["V" + rw].Style.Numberformat.Format = "#,##0.00";
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
                            ws.Cells["T" + rw].Value = 0.00;
                            ws.Cells["U" + rw].Value = 0.00;
                            ws.Cells["V" + rw].Value = 0.00;
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
                        ws.Cells["T" + rw].Value = 0.00;
                        ws.Cells["U" + rw].Value = 0.00;
                        ws.Cells["V" + rw].Value = 0.00;
                    }
                    ws.Cells["W" + rw].Value = Currency;
                    string froMulaAddQV = string.Format("=SUM(M" + rw + ":R" + rw + ")");
                    ws.Cells["X" + rw].Formula = froMulaAddQV.ToString();
                    ws.Cells["X" + rw].Style.Numberformat.Format = "#,##0.00";
                    rw++;
                }



                ws.Cells["A8:X10" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A8:X10" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 50].AutoFitColumns();
            }
            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Export_PortStorage_SummaryReport.xlsx");
            Response.End();

        }


        public DataTable GetExpPortStorageSummaryReport(string DtFrom, string DtTo, string GeoLocID)
        {

   


            //string strWhere = "";
            //string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
            //                " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,NVO_BOL.POLID,NVO_BOL.PODID,  " +
            //                " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +
            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MSDate, " +
            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FLDate, " +

            //                " (select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as MSDatev, " +

            //                " isnull((select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as FLDatev, " +

            //                " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +
            //                " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)) as ImpFreeDays,  " +
            //                " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)) as ExpFreeDays, " +
            //                " 3 as PortFreeDays,"+
            //                " DATEADD(DAY,(3), " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

            //                " case when(DATEDIFF(DAY, (DATEADD(DAY,(3), " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
            //                " then DATEDIFF(DAY, (DATEADD(DAY, (3), " +
            //                " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
            //                " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv, " +

            //                " datediff(day,  DATEADD(DAY,(3),  (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'MS' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)), isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
            //                " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
            //                " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) as Days1 " +

            //                " from NVO_ContainerTxns " +
            //                " inner join NVO_Containers on NVO_Containers.ID = NVO_ContainerTxns.ContainerID " +
            //                " inner join NVO_Booking on NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
            //                " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
            //                " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID=NVO_ContainerTxns.AgencyID " +
            //                " where NVO_ContainerTxns.StatusCode in ('FL') and  GeoLocationID= " + GeoLocID;


            string strWhere = "";
            string _Query = " select distinct NVO_BOL.BLNumber,Shipper, (select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_ContainerTxns.AgencyID) as AgencyName,NVO_ContainerTxns.AgencyID,TypeID, " +
                            " VesVoy,NVO_ContainerTxns.ContainerID,CntrNo,POL,POD,NVO_BOL.POLID,NVO_BOL.PODID,  " +
                            " (select top(1) Type + '-' + Size from NVO_tblCntrTypes  where NVO_tblCntrTypes.ID = NVO_Containers.TypeID) as CntrTypes, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FBDate, " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FLDate, " +

                            " (select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber) as FBDatev, " +

                            " isnull((select top(1) CAST(DtMovement AS DATE) from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate()) as FLDatev, " +

                            " DATEDIFF(DAY, (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 AS Days, " +
                            " (select top(1)ImpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)) as ImpFreeDays,  " +
                            " (select top(1)ExpFreeDays from NVO_RatesheetMode where NVO_RatesheetMode.RRID = NVO_Booking.RRID and ModeID in(2,1)) as ExpFreeDays, " +
                            " 3 as PortFreeDays," +
                            " DATEADD(DAY,(3), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)) as FromDate, " +

                            " case when(DATEDIFF(DAY, (DATEADD(DAY,(3), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1) >= 1 " +
                            " then DATEDIFF(DAY, (DATEADD(DAY, (3), " +
                            " (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber))), " +
                            " isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) +1 else 0 end AS Daysv, " +

                            " datediff(day,  DATEADD(DAY,(3),  (select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FB' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber)), isnull((select top(1) DtMovement from NVO_ContainerTxns CnTx " +
                            " where CnTx.StatusCode = 'FL' and CnTx.ContainerID = NVO_ContainerTxns.ContainerID " +
                            " and CnTx.BLNumber = NVO_ContainerTxns.BLNumber), getdate())) as Days1 " +

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

        public DataTable GetContractSlapExp(string AgencyID, string CntrTypes, string PortID)
        {
            string strWhere = "";
            string _Query = " select isnull((select top(1)  Rate from NVO_ExRate where AgencyID= NVO_StorageContract.AgencyID and FromCurrency= NVO_StorageContractTariffDtls.CurrencyID order by Id desc),0) as ExRate," +
                            " (select top(1) CurrencyCode from NVO_CurrencyMaster where ID =NVO_StorageContractTariffDtls.CurrencyID) as Currency,* from NVO_StorageContract inner join NVO_StorageContractTariffDtls on NVO_StorageContractTariffDtls.RentID = NVO_StorageContract.ID " +
                            " where AgencyID = " + AgencyID + " and ContainerType = " + CntrTypes + " and ShipmentTypeID = 1 and ChargesID=42 and PortID =" + PortID;
            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }


        public void FreeDaySummaryReportValues(string DtFrom, string DtTo, string User, string GeoLocID)
        {

            ExcelPackage pck = new ExcelPackage();

            DataTable dt = GetAgencyLocation(DtFrom, DtTo, GeoLocID);
            for (int J = 0; J < dt.Rows.Count; J++)
            {
                var ws = pck.Workbook.Worksheets.Add(dt.Rows[J]["GeoLocation"].ToString());
                ws.Cells["a2"].Value = "FREEDAYS REPORT";
                ws.Cells["a2"].Style.Font.Bold = true;
                ws.Cells["a2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:X2"];
                r.Merge = true;
                r.Style.Font.Size = 12;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int rw = 2;


                ws.Cells["S3:U3"].Value = "FREEDAYS REPORT.";
                ws.Cells["S3:U3"].Merge = true;
                ws.Cells["S3:U3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                Color colFromHexExp = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");
                ExcelRange r1 = ws.Cells["S3:U3"];
                r1.Merge = true;
                r1.Style.Font.Size = 12;
                r1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r1.Style.Fill.BackgroundColor.SetColor(colFromHexExp);


                ws.Cells["V3:X3"].Value = "IMPORT FREEDAYS.";
                ws.Cells["V3:X3"].Merge = true;
                ws.Cells["V3:X3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                Color colFromHexImp = System.Drawing.ColorTranslator.FromHtml("#FFDAE1");
                ExcelRange r2 = ws.Cells["V3:X3"];
                r2.Merge = true;
                r2.Style.Font.Size = 12;
                r2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r2.Style.Fill.BackgroundColor.SetColor(colFromHexImp);


                ws.Cells["A4"].Value = "SR.No";
                ws.Cells["B4"].Value = "Location.";
                ws.Cells["C4"].Value = "Agent Name.";
                ws.Cells["D4"].Value = "Rate Request No.";
                ws.Cells["E4"].Value = "Booking No";
                ws.Cells["F4"].Value = "Booking Status";
                ws.Cells["G4"].Value = "POL";
                ws.Cells["H4"].Value = "POD";
                ws.Cells["I4"].Value = "T/S";
                ws.Cells["J4"].Value = "FPOD";

                ws.Cells["K4"].Value = "Vessel /Voy ";
                ws.Cells["L4"].Value = "Terminal Name";
                ws.Cells["M4"].Value = "Slot Operator";
                ws.Cells["N4"].Value = "ETD'";
                ws.Cells["O4"].Value = "Commodity Type";
                ws.Cells["P4"].Value = "No.of.20'";
                ws.Cells["Q4"].Value = "No.Of.40'";
                ws.Cells["R4"].Value = "FRIGHT TERMS";
                ws.Cells["S4"].Value = "Combined";
                ws.Cells["T4"].Value = "Detention";
                ws.Cells["U4"].Value = "Demmurage";
                ws.Cells["V4"].Value = "Combined";
                ws.Cells["W4"].Value = "Detention";
                ws.Cells["X4"].Value = "Demmurage";
                rw = 5;
                int Rows = 1;
                string sft = dt.Rows[J]["GeoLocation"].ToString();
                DataTable dtx = GetFreeDaysSummaryReport(DtFrom, DtTo, dt.Rows[J]["GeoLocationID"].ToString());
                for (int z = 0; z < dtx.Rows.Count; z++)
                {

                    ws.Cells["A" + rw].Value = Rows++;
                    ws.Cells["B" + rw].Value = dt.Rows[J]["GeoLocation"].ToString();
                    ws.Cells["C" + rw].Value = dtx.Rows[z]["AgencyName"].ToString();
                    ws.Cells["D" + rw].Value = dtx.Rows[z]["RatesheetNo"].ToString();
                    ws.Cells["E" + rw].Value = dtx.Rows[z]["BookingNo"].ToString();
                    ws.Cells["F" + rw].Value = dtx.Rows[z]["BkgStatus"].ToString();
                    ws.Cells["G" + rw].Value = dtx.Rows[z]["POL"].ToString();
                    ws.Cells["H" + rw].Value = dtx.Rows[z]["POD"].ToString();
                    ws.Cells["I" + rw].Value = dtx.Rows[z]["TSPORT"].ToString();
                    ws.Cells["J" + rw].Value = dtx.Rows[z]["FPOD"].ToString();
                    ws.Cells["K" + rw].Value = dtx.Rows[z]["VesVoy"].ToString();

                    ws.Cells["L" + rw].Value = dtx.Rows[z]["TerminalName"].ToString();
                    ws.Cells["M" + rw].Value = dtx.Rows[z]["SlotOperators"].ToString();
                    ws.Cells["N" + rw].Value = dtx.Rows[z]["ETD"].ToString();
                    ws.Cells["O" + rw].Value = dtx.Rows[z]["CommodityType"].ToString();
                    ws.Cells["P" + rw].Value = dtx.Rows[z]["GP20"];
                    ws.Cells["Q" + rw].Value = dtx.Rows[z]["GP40"];
                    ws.Cells["R" + rw].Value = dtx.Rows[z]["FRIGHTTerms"].ToString();

                    ws.Cells["S" + rw].Value = dtx.Rows[z]["ExpCombined"];
                    ws.Cells["T" + rw].Value = dtx.Rows[z]["ExpDetention"];
                    ws.Cells["U" + rw].Value = dtx.Rows[z]["ExpDemmurage"];
                    ws.Cells["V" + rw].Value = dtx.Rows[z]["ImpCombined"];
                    ws.Cells["W" + rw].Value = dtx.Rows[z]["ImpDetention"];
                    ws.Cells["X" + rw].Value = dtx.Rows[z]["ImpDemmurage"];
                    rw++;
                }



            
                ws.Cells["A2:X" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A2:X" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A2:X" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A2:X" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 50].AutoFitColumns();
            }
            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Freedays_SummaryReport.xlsx");
            Response.End();

        }



        public DataTable GetFreeDaysSummaryReport(string DtFrom, string DtTo, string GeoLocID)
        {

            string strWhere = "";
            string _Query = " select distinct NVO_Ratesheet.ID,RatesheetNo,(select top(1) AgencyName from NVO_AgencyMaster where ID = NVO_Ratesheet.AgentId) as AgencyName, " +
                            " BookingNo,case when BkgStatus = 1  then 'DRAFT' ELSE CASE WHEN BkgStatus = 2  then 'FINAL' ELSE CASE WHEN BkgStatus = 3  then 'REJECTED' END end end AS BkgStatus, " +
                            " POL,POD,FPOD,(select top(1) PortName from NVO_PortMaster where Id = TSPORTID) as TSPORT,VesVoy,(select top(1) CustomerName from NVO_view_CustomerDetails where CID = NVO_Booking.SlotOperatorID) as SlotOperators,CommodityType, " +
                            " (select sum(Qty) from NVO_BookingCntrTypes where CntrTypes in (1, 4, 6, 8, 10, 11) and BkgId = NVO_Booking.Id) as GP20, " +
                            " (select sum(Qty) from NVO_BookingCntrTypes where CntrTypes in (2, 3, 5, 7, 9, 12)  and BkgId = NVO_Booking.Id) as GP40, " +
                            " (select top(1)  case when PaymentModeID = 18 then 'PREPAID' else 'COLLECT' end from NVO_RatesheetCharges where ChargeCodeID = 1 and RRID = NVO_Ratesheet.ID) as FRIGHTTerms, " +
                            " (select top(1) ExpFreeDays from NVO_RatesheetMode where ModeID = 1 and RRID = NVO_Ratesheet.ID) as ExpCombined, " +
                            " (select top(1) ExpFreeDays from NVO_RatesheetMode where ModeID = 2 and RRID = NVO_Ratesheet.ID) as ExpDetention, " +
                            " (select top(1) ExpFreeDays from NVO_RatesheetMode where ModeID = 3 and RRID = NVO_Ratesheet.ID) as ExpDemmurage, " +
                            " (select top(1) ImpFreeDays from NVO_RatesheetMode where ModeID = 1 and RRID = NVO_Ratesheet.ID) as ImpCombined, " +
                            " (select top(1) ImpFreeDays from NVO_RatesheetMode where ModeID = 2 and RRID = NVO_Ratesheet.ID) as ImpDetention, " +
                            " (select top(1) ImpFreeDays from NVO_RatesheetMode where ModeID = 3 and RRID = NVO_Ratesheet.ID) as ImpDemmurage, " +
                            " (select top(1) (select top(1) TerminalName from NVO_TerminalMaster where ID =TerminalID) from NVO_VoyageRoute where VoyageID=NVO_Booking.VesVoyID) as TerminalName , "+
                            " (select top(1) convert(varchar, ETD, 103) from NVO_VoyageRoute where VoyageID = NVO_Booking.VesVoyID) as ETD " +
                            " from NVO_Ratesheet " +
                            " inner join NVO_Booking on NVO_Booking.RRID = NVO_Ratesheet.ID" +
                            " inner join NVO_AgencyMaster on NVO_AgencyMaster.ID = NVO_Ratesheet.AgentId INNER JOIN NVO_ContainerTxns ON NVO_ContainerTxns.BLNumber=NVO_Booking.ID "+
                            " where  NVO_ContainerTxns.StatusCode in ('FB')  and NVO_AgencyMaster.GeoLocationID=" + GeoLocID;
            if (DtFrom != "" && DtFrom != "undefined" && DtFrom != null || DtTo != "" && DtTo != "undefined" && DtTo != null)
                if (strWhere == "")
                    strWhere += _Query + " AND convert(varchar, NVO_ContainerTxns.DtMovement, 23) between '" + DtFrom + "' and '" + DtTo + "'";



            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

    }


}