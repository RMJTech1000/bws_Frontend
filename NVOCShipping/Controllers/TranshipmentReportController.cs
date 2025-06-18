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
    public class TranshipmentReportController : Controller
    {
        // GET: TransipmentReport
        ExportReportManager RegManage = new ExportReportManager();
        public ActionResult TranshipmentSummaryReport()
        {
            return View();
        }

        public ActionResult TranshipmentSummaryReportView(string DtFrom, string DtTo)
        {
            BindTranshipmentSummaryReportGetValue(DtFrom, DtTo);
            return View();
        }

        public void BindTranshipmentSummaryReportGetValue(string DtFrom, string DtTo)
        {
            using (ExcelPackage _xlPV = new ExcelPackage())
            {
                _xlPV.Workbook.Properties.Author = "TRANSHIPMENTSUMMARYREPORT";
                _xlPV.Workbook.Properties.Title = "TRANSHIPMENTSUMMARYREPORT";
                _xlPV.Workbook.Worksheets.Add("TRANSHIPMENTSUMMARYREPORT");
                ExcelWorksheet ws = _xlPV.Workbook.Worksheets[1];
                Color colCost;
                ws.Name = "REPORT"; //Setting Sheet's name
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
                int MaxCol = 25;
                int RowIndex = 2;
                ws.Cells[RowIndex, 1].Value = "TRANSHIPMENT SUMMARY REPORT";
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
                //ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 14.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                RowIndex++;

                ws.Cells[RowIndex, 1, RowIndex, 6].Value = "";
                ws.Cells[RowIndex, 1, RowIndex, 6].Merge = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 12.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;


                colCost = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
                ws.Cells[RowIndex, 7, RowIndex, 12].Value = "1ST LEG VESSEL DETAILS";
                ws.Cells[RowIndex, 7, RowIndex, 12].Merge = true;
                ws.Cells[RowIndex, 7, RowIndex, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 7, RowIndex, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 7, RowIndex, 12].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#FFDAE1");
                ws.Cells[RowIndex, 13, RowIndex, 18].Value = "2ND LEG VESSEL DETAILS";
                ws.Cells[RowIndex, 13, RowIndex, 18].Merge = true;
                ws.Cells[RowIndex, 13, RowIndex, 18].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 13, RowIndex, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 13, RowIndex, 18].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#A7C7E7");
                ws.Cells[RowIndex, 19, RowIndex, 24].Value = "3RD LEG VESSEL DETAILS";
                ws.Cells[RowIndex, 19, RowIndex, 24].Merge = true;
                ws.Cells[RowIndex, 19, RowIndex, 24].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 19, RowIndex, 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 19, RowIndex, 24].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#FFD966");
                ws.Cells[RowIndex, 25].Value = "";
                ws.Cells[RowIndex, 25, RowIndex, 25].Merge = true;
                ws.Cells[RowIndex, 25, RowIndex, 25].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 25, RowIndex, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 25, RowIndex, 25].Style.Fill.BackgroundColor.SetColor(colCost);

                RowIndex++;
                ws.Cells[RowIndex, 1].Value = "SR NO.";
                ws.Cells[RowIndex, 2].Value = "CNTR NO";
                ws.Cells[RowIndex, 3].Value = "TYPE";
                ws.Cells[RowIndex, 4].Value = "BL NUMBER";
                ws.Cells[RowIndex, 5].Value = "POL";
                ws.Cells[RowIndex, 6].Value = "POD";

                ws.Cells[RowIndex, 7].Value = "LOCATION";
                ws.Cells[RowIndex, 8].Value = "1ST LEG-VSL/VOY";
                ws.Cells[RowIndex, 9].Value = "ETA-POL";
                ws.Cells[RowIndex, 10].Value = "ETD-POL(FB)";
                ws.Cells[RowIndex, 11].Value = "SLOT OPERATOR";
                ws.Cells[RowIndex, 12].Value = "SLOT AMOUNT";

                ws.Cells[RowIndex, 13].Value = "LOCATION";
                ws.Cells[RowIndex, 14].Value = "2ND LEG-VSL/VOY";
                ws.Cells[RowIndex, 15].Value = "ETA(TZ)";
                ws.Cells[RowIndex, 16].Value = "ETD(TZFB)";
                ws.Cells[RowIndex, 17].Value = "SLOT OPERATOR";
                ws.Cells[RowIndex, 18].Value = "SLOT AMOUNT";

                ws.Cells[RowIndex, 19].Value = "LOCATION";
                ws.Cells[RowIndex, 20].Value = "3RD LEG-VSL/VOY";
                ws.Cells[RowIndex, 21].Value = "ETA(TZ)";
                ws.Cells[RowIndex, 22].Value = "ETD(TZFB)";
                ws.Cells[RowIndex, 23].Value = "SLOT OPERATOR";
                ws.Cells[RowIndex, 24].Value = "SLOT AMOUNT";

                ws.Cells[RowIndex, 25].Value = "STATUS";

                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.WrapText = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 10.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                RowIndex++; int InvCount = 0;
                int rw_bgn = 0;
                rw_bgn = RowIndex;


                DataTable dtx = GetALLTranshipmentValues(DtFrom, DtTo);
                for (int i = 0; i < dtx.Rows.Count; i++)
                {
                    InvCount++;
                    int v =RowIndex;
                    ws.Cells[RowIndex, 1].Value = InvCount;
                    ws.Cells[RowIndex, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 2].Value = dtx.Rows[i]["CntrNo"].ToString();
                    ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 3].Value = dtx.Rows[i]["Size"].ToString();
                    ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 4].Value = dtx.Rows[i]["BLNumber"].ToString();
                    ws.Cells[RowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 5].Value = dtx.Rows[i]["POL"].ToString();
                    ws.Cells[RowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 6].Value = dtx.Rows[i]["POD"].ToString();
                    ws.Cells[RowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    int RowColumInx = 7;
                    DataTable _dtx = GetALLTranshipmentDtlsValues(dtx.Rows[i]["BLID"].ToString(), dtx.Rows[i]["CntrID"].ToString());
                    for( int j =0; j<_dtx.Rows.Count; j++)
                    {
                        ws.Cells[RowIndex, RowColumInx].Value = _dtx.Rows[j]["GeoLocation"].ToString();
                        ws.Cells[RowIndex, RowColumInx].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        RowColumInx++;
                        ws.Cells[RowIndex, RowColumInx].Value = _dtx.Rows[j]["VelVoy"].ToString();
                        ws.Cells[RowIndex, RowColumInx].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        RowColumInx++;
                        ws.Cells[RowIndex, RowColumInx].Value = _dtx.Rows[j]["ETA"].ToString();
                        ws.Cells[RowIndex, RowColumInx].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        RowColumInx++;
                        ws.Cells[RowIndex, RowColumInx].Value = _dtx.Rows[j]["ETDMovement"].ToString();
                        ws.Cells[RowIndex, RowColumInx].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        RowColumInx++;
                        ws.Cells[RowIndex, RowColumInx].Value = _dtx.Rows[j]["Carrier"].ToString();
                        ws.Cells[RowIndex, RowColumInx].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        RowColumInx++;
                        ws.Cells[RowIndex, RowColumInx].Value = _dtx.Rows[j]["SlotAmt"].ToString();
                        ws.Cells[RowIndex, RowColumInx].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        RowColumInx++;
                    }
                    RowIndex++;
                }



                string style1stleg = "#DAFFED";
                string style2ndleg = "#FFDAE1";
                string style3rdleg = "#A7C7E7";
                string style4rdleg = "#FFD966";
                rw_bgn = 7;
                colCost = System.Drawing.ColorTranslator.FromHtml(style1stleg);
                ws.Cells["G" + (rw_bgn - 2) + ":L" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["G" + (rw_bgn - 2) + ":L" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(style2ndleg);
                ws.Cells["M" + (rw_bgn - 2) + ":R" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["M" + (rw_bgn - 2) + ":R" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(style3rdleg);
                ws.Cells["S" + (rw_bgn - 2) + ":X" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["S" + (rw_bgn - 2) + ":X" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);


                colCost = System.Drawing.ColorTranslator.FromHtml(style4rdleg);
                ws.Cells["y" + (rw_bgn - 2) + ":y" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["y" + (rw_bgn - 2) + ":y" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                ws.Cells["A3:Y" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:Y" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:Y" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:Y" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, 25, 25].AutoFitColumns();

                string _filePath = Server.MapPath("~/PreXL/");
                string _fileName = "TranshipmentReportSummary " + Session.SessionID + ".xlsx";
                Byte[] bin = _xlPV.GetAsByteArray();
                System.IO.File.WriteAllBytes(_filePath + "\\" + _fileName, bin);
                Response.Clear();
                Response.AppendHeader("content-disposition", "attachment; filename=" + _fileName);
                Response.ContentType = "application/octet-stream";
                Response.WriteFile(_filePath + "\\" + _fileName);
                Response.Flush();
                Response.End();


            }

        }


        public DataTable GetALLTranshipmentValues(string FromDate, string ToDate)
        {
            string strWhere = "";
            string _Query = " select distinct NVO_Booking.Id as BkgID,NVO_BOL.ID as BLID,BookingNo,NVO_BOL.BLNumber,CntrID,NVO_Containers.CntrNo,POL,POD,NVO_Booking.AgentID, " +
                            " (select top(1) VesVoy from NVO_BLRelease where NVO_BLRelease.BLID = NVO_BOL.ID) as VesVoy,NVO_tblCntrTypes.Size " +
                            " FROM NVO_Containers " +
                            " inner join NVO_ContainerTxns on NVO_ContainerTxns.ContainerID = NVO_Containers.ID " +
                            " INNER JOIN NVO_Booking ON NVO_Booking.ID = NVO_ContainerTxns.BLNumber " +
                            " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
                            " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOL.ID and NVO_BOLCntrDetails.BkgId = NVO_Booking.ID AND NVO_BOLCntrDetails.CntrID = NVO_ContainerTxns.ContainerID " +
                            " inner join NVO_tblCntrTypes on NVO_tblCntrTypes.ID = NVO_Containers.TypeID " +
                            " where TSPORTID != 0 and NVO_BOL.BLTypes = 40 AND NVO_ContainerTxns.StatusCode = 'TZFB'";

            if (FromDate != null && FromDate != "undefined" || ToDate != null && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " and  convert(varchar,NVO_ContainerTxns.DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,NVO_ContainerTxns.DtMovement,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;
            return RegManage.GetViewData(strWhere, "");
        }


        public DataTable GetALLTranshipmentDtlsValues(string BLID, string CntrID)
        {
            string strWhere = "";
            //string _Query = " select distinct NVO_BOLVoyageDetails.BKgId, NVO_BOLVoyageDetails.BLID,LegInformation,NVO_Containers.TypeID,Teus, " +
            //                " (select top(1) PortName from NVO_PortMaster where ID = NVO_BOLVoyageDetails.LoadPort) as GeoLocation, " +
            //                " (select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID = NVO_BOLVoyageDetails.VesVoyID) as VelVoy, " +
            //                " convert(varchar, ETA, 103) as ETA, " +
            //                " case when LegInformation = 74 then " +
            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns where BLNumber = NVO_BOLVoyageDetails.BKgId " +
            //                " and StatusCode = 'FB' and ContainerID = NVO_BOLCntrDetails.CntrID) else " +
            //                " case when LegInformation = 75 then " +
            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns where BLNumber = NVO_BOLVoyageDetails.BKgId " +
            //                " and StatusCode = 'TZFB' and ContainerID = NVO_BOLCntrDetails.CntrID) else " +
            //                " case when LegInformation = 76 then " +
            //                " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns where BLNumber = NVO_BOLVoyageDetails.BKgId " +
            //                " and StatusCode = 'TZFB' and ContainerID = NVO_BOLCntrDetails.CntrID order by NVO_ContainerTxns.ID desc) " +
            //                " end end end as ETDMovement, " +
            //                " (select top(1) Carrier from NVO_V_VoyageAllocationSlotAmount where NVO_V_VoyageAllocationSlotAmount.BLId = NVO_BOLVoyageDetails.BLID " +
            //                " and NVO_V_VoyageAllocationSlotAmount.LegInformation = NVO_BOLVoyageDetails.LegInformation) as Carrier, " +
            //                " case when Teus = 1 then(select top(1) SlotAmt20 from NVO_V_VoyageAllocationSlotAmount where NVO_V_VoyageAllocationSlotAmount.BLId = NVO_BOLVoyageDetails.BLID " +
            //                " and NVO_V_VoyageAllocationSlotAmount.LegInformation = NVO_BOLVoyageDetails.LegInformation) " +
            //                " else " +
            //                " (select top(1) SlotAmt40 from NVO_V_VoyageAllocationSlotAmount where NVO_V_VoyageAllocationSlotAmount.BLId = NVO_BOLVoyageDetails.BLID " +
            //                " and NVO_V_VoyageAllocationSlotAmount.LegInformation = NVO_BOLVoyageDetails.LegInformation) end as SlotAmt " +
            //                " from NVO_BOLVoyageDetails " +
            //                " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOLVoyageDetails.BLID and NVO_BOLCntrDetails.BkgId = NVO_BOLVoyageDetails.BKgId " +
            //                " inner join NVO_Containers on NVO_Containers.ID = NVO_BOLCntrDetails.CntrID " +
            //                " inner join NVO_tblCntrTypes on NVO_tblCntrTypes.ID = NVO_Containers.TypeID " +
            //                " where NVO_BOLVoyageDetails.BLID = " + BLID + " and NVO_Containers.ID=" + CntrID;


            string _Query = " select distinct NVO_BOLVoyageDetails.BKgId, NVO_BOLVoyageDetails.BLID,LegInformation,NVO_Containers.TypeID,Teus, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = NVO_BOLVoyageDetails.PortID) as GeoLocation, " +
                            " (select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID = NVO_BOLVoyageDetails.VesVoyID) as VelVoy, " +
                            " convert(varchar, ETA, 103) as ETA, " +
                            " case when LegInformation = 74 then " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns where BLNumber = NVO_BOLVoyageDetails.BKgId " +
                            " and StatusCode = 'FB' and ContainerID = NVO_BOLCntrDetails.CntrID) else " +
                            " case when LegInformation = 75 then " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns where BLNumber = NVO_BOLVoyageDetails.BKgId " +
                            " and StatusCode = 'TZFB' and ContainerID = NVO_BOLCntrDetails.CntrID) else " +
                            " case when LegInformation = 76 then " +
                            " (select top(1) convert(varchar, DtMovement, 103) from NVO_ContainerTxns where BLNumber = NVO_BOLVoyageDetails.BKgId " +
                            " and StatusCode = 'TZFB' and ContainerID = NVO_BOLCntrDetails.CntrID order by NVO_ContainerTxns.ID desc) " +
                            " end end end as ETDMovement, " +
                            " (select top(1) Carrier from NVO_V_VoyageAllocationSlotAmount where NVO_V_VoyageAllocationSlotAmount.BLId = NVO_BOLVoyageDetails.BLID " +
                            " and NVO_V_VoyageAllocationSlotAmount.LegInformation = NVO_BOLVoyageDetails.LegInformation) as Carrier, " +
                            " case when Teus = 1 then(select top(1) SlotAmt20 from NVO_V_VoyageAllocationSlotAmount where NVO_V_VoyageAllocationSlotAmount.BLId = NVO_BOLVoyageDetails.BLID " +
                            " and NVO_V_VoyageAllocationSlotAmount.LegInformation = NVO_BOLVoyageDetails.LegInformation) " +
                            " else " +
                            " (select top(1) SlotAmt40 from NVO_V_VoyageAllocationSlotAmount where NVO_V_VoyageAllocationSlotAmount.BLId = NVO_BOLVoyageDetails.BLID " +
                            " and NVO_V_VoyageAllocationSlotAmount.LegInformation = NVO_BOLVoyageDetails.LegInformation) end as SlotAmt " +
                            " from Nvo_View_BLVoyageLegDetails NVO_BOLVoyageDetails " +
                            " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOLVoyageDetails.BLID and NVO_BOLCntrDetails.BkgId = NVO_BOLVoyageDetails.BKgId " +
                            " inner join NVO_Containers on NVO_Containers.ID = NVO_BOLCntrDetails.CntrID " +
                            " inner join NVO_tblCntrTypes on NVO_tblCntrTypes.ID = NVO_Containers.TypeID " +
                            " where NVO_BOLVoyageDetails.BLID = " + BLID + " and NVO_Containers.ID=" + CntrID;

            if (strWhere == "")
                strWhere = _Query;
            return RegManage.GetViewData(strWhere, "");
        }
    }
}