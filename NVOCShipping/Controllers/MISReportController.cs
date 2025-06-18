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

    public class MISReportController : Controller
    {
        // GET: MISReport

        ExportReportManager RegManage = new ExportReportManager();
        public ActionResult MISReportView()
        {
            return View();
        }
        public ActionResult MISReport()
        {
            return View();
        }

        public ActionResult MISReportExcel(string AgencyID)
        {
            BindMISReportGetValue(AgencyID);
            return View();
        }

        public void BindMISReportGetValue(string AgencyID)
        {
            using (ExcelPackage _xlPV = new ExcelPackage())
            {
                _xlPV.Workbook.Properties.Author = "MIS REPORT";
                _xlPV.Workbook.Properties.Title = "MIS REPORT";
                _xlPV.Workbook.Worksheets.Add("MISREPORT");
                ExcelWorksheet ws = _xlPV.Workbook.Worksheets[1];
                Color colCost;
                ws.Name = "REPORT"; //Setting Sheet's name
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
                int MaxCol = 37;
                int RowIndex = 2;
                ws.Cells[RowIndex, 1].Value = "MIS REPORT";
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
                //ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 14.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                RowIndex++;

                ws.Cells[RowIndex, 1, RowIndex,14].Value = "";
                ws.Cells[RowIndex, 1, RowIndex, 14].Merge = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 12.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;

                
                colCost = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
                ws.Cells[RowIndex, 15, RowIndex, 22].Value = "REVENUE";
                ws.Cells[RowIndex, 15, RowIndex, 22].Merge = true;
                ws.Cells[RowIndex, 15, RowIndex, 22].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 15, RowIndex, 22].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 15, RowIndex, 22].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#FFDAE1");
                ws.Cells[RowIndex, 23, RowIndex, 33].Value = "COST";
                ws.Cells[RowIndex, 23, RowIndex, 33].Merge = true;
                ws.Cells[RowIndex, 23, RowIndex, 33].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 23, RowIndex, 33].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 23, RowIndex, 33].Style.Fill.BackgroundColor.SetColor(colCost);

                ws.Cells[RowIndex, 34].Value = "";

               

                RowIndex++;
                ws.Cells[RowIndex, 1].Value = "SR No.";
                ws.Cells[RowIndex, 2].Value = "Rate Request No.";
                ws.Cells[RowIndex, 3].Value = "BLNo";
                ws.Cells[RowIndex, 4].Value = "Container No";
                ws.Cells[RowIndex, 5].Value = "POL";
                ws.Cells[RowIndex, 6].Value = "POD";

                ws.Cells[RowIndex, 7].Value = "T/S";

                ws.Cells[RowIndex, 8].Value = "FPOD";
                ws.Cells[RowIndex, 9].Value = "Vessel /Voy ";
                ws.Cells[RowIndex, 10].Value = "Terminal Name";
                ws.Cells[RowIndex, 11].Value = "Slot Operator";
                ws.Cells[RowIndex, 12].Value = "ETD'";
                ws.Cells[RowIndex, 13].Value = "Commodity Type ";
                ws.Cells[RowIndex, 14].Value = "20'";
                ws.Cells[RowIndex, 15].Value = "40'";
                ws.Cells[RowIndex, 16].Value = "Container Types";
                ws.Cells[RowIndex, 17].Value = "Lease Type";
                ws.Cells[RowIndex, 18].Value = "Ocen Freight ($)";
                ws.Cells[RowIndex, 19].Value = "BAF ($)";
                ws.Cells[RowIndex, 20].Value = "Surcharges ($)";
                ws.Cells[RowIndex, 21].Value = "Load Port THC ($)";
                ws.Cells[RowIndex, 22].Value = "Discharge Port THC ($)";
                ws.Cells[RowIndex, 23].Value = "HDL/TPT ($)";
                ws.Cells[RowIndex, 24].Value = "MISC ($)";
                ws.Cells[RowIndex, 25].Value = "Total Revenue ($)";
                ws.Cells[RowIndex, 26].Value = "SLOT ($)";
                ws.Cells[RowIndex, 27].Value = "T/S THC COST ($)";
                ws.Cells[RowIndex, 28].Value = "T/S COMMISSION ($)";
                ws.Cells[RowIndex, 29].Value = "T/S SLOT COST ($)";
                ws.Cells[RowIndex, 30].Value = "LOAD PORT PHC ($)";
                ws.Cells[RowIndex, 31].Value = "DISCHARGE PORT PHC ($)";
                ws.Cells[RowIndex, 32].Value = "LOAD COMMISSION ($)";
                ws.Cells[RowIndex, 33].Value = "DISCHARGE COMMISSION ($)";
                ws.Cells[RowIndex, 34].Value = "MISC COST ($)";
                ws.Cells[RowIndex, 35].Value = "LOG. FEE ($)";
                ws.Cells[RowIndex, 36].Value = "TOTAL COST ($)";
                ws.Cells[RowIndex, 37].Value = "PROFIT / LOSS ($)";


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
                DataTable dtx = GetCntrDtlsTerminalDepPdfValues(AgencyID);
                for (int i = 0; i < dtx.Rows.Count; i++)
                {
                    decimal RevenusAmt = decimal.Parse(dtx.Rows[i]["OFTt"].ToString()) + decimal.Parse(dtx.Rows[i]["BAFt"].ToString()) + decimal.Parse(dtx.Rows[i]["SURCHARGESt"].ToString()) + decimal.Parse(dtx.Rows[i]["LOADTHCt"].ToString()) + decimal.Parse(dtx.Rows[i]["HLDTPTt"].ToString()) + decimal.Parse(dtx.Rows[i]["MISCt"].ToString()) + decimal.Parse(dtx.Rows[i]["DESTTHCt"].ToString());
                    decimal CostAmt = 5 + decimal.Parse(dtx.Rows[i]["SlotAmtt"].ToString()) + decimal.Parse(dtx.Rows[i]["TSCOSTt"].ToString()) + decimal.Parse(dtx.Rows[i]["EXCOMMt"].ToString()) + decimal.Parse(dtx.Rows[i]["ICOMMt"].ToString()) + decimal.Parse(dtx.Rows[i]["TSCOMMt"].ToString()) + decimal.Parse(dtx.Rows[i]["LOADPHCt"].ToString()) + decimal.Parse(dtx.Rows[i]["DESTPHCt"].ToString()) + decimal.Parse(dtx.Rows[i]["MISCCOSTt"].ToString());


                    InvCount++;
                    ws.Cells[RowIndex, 1].Value = InvCount;
                    ws.Cells[RowIndex, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 2].Value = dtx.Rows[i]["RRNO"].ToString();
                    ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 3].Value = dtx.Rows[i]["BookingNo"].ToString();
                    ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 4].Value = dtx.Rows[i]["CntrNo"].ToString();
                    ws.Cells[RowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 5].Value = dtx.Rows[i]["POL"].ToString();
                    ws.Cells[RowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 6].Value = dtx.Rows[i]["POD"].ToString();
                    ws.Cells[RowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 7].Value = dtx.Rows[i]["TSPORT"].ToString();
                    ws.Cells[RowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 8].Value = dtx.Rows[i]["FPOD"].ToString();
                    ws.Cells[RowIndex, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 9].Value = dtx.Rows[i]["VesVoy"].ToString();
                    ws.Cells[RowIndex, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 10].Value = dtx.Rows[i]["Terminal"].ToString();
                    ws.Cells[RowIndex, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 11].Value = dtx.Rows[i]["SlotOperator"].ToString();
                    ws.Cells[RowIndex, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 12].Value = dtx.Rows[i]["ETDDate"].ToString();
                    ws.Cells[RowIndex, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 13].Value = dtx.Rows[i]["Commodity"].ToString();
                    ws.Cells[RowIndex, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 14].Value = dtx.Rows[i]["Size20"].ToString();
                    ws.Cells[RowIndex, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 15].Value = dtx.Rows[i]["Size40"].ToString();
                    ws.Cells[RowIndex, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 16].Value = dtx.Rows[i]["TypeSize"].ToString();
                    ws.Cells[RowIndex, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 17].Value = dtx.Rows[i]["LeasTerms"].ToString();
                    ws.Cells[RowIndex, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 18].Value = dtx.Rows[i]["OFTt"];
                    ws.Cells[RowIndex, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 19].Value = dtx.Rows[i]["BAFt"];
                    ws.Cells[RowIndex, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 20].Value = dtx.Rows[i]["SURCHARGESt"];
                    ws.Cells[RowIndex, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 21].Value = dtx.Rows[i]["LOADTHCt"];
                    ws.Cells[RowIndex, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 22].Value = dtx.Rows[i]["DESTTHCt"];
                    ws.Cells[RowIndex, 22].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 23].Value = dtx.Rows[i]["HLDTPTt"];
                    ws.Cells[RowIndex, 23].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 24].Value = dtx.Rows[i]["MISCt"];
                    ws.Cells[RowIndex, 24].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);



                    ws.Cells[RowIndex, 25].Value = RevenusAmt;
                    ws.Cells[RowIndex, 25].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 26].Value = dtx.Rows[i]["SlotAmtt"];
                    ws.Cells[RowIndex, 26].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["TSCOSTt"];
                    ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 28].Value = dtx.Rows[i]["TSCOMMt"];
                    ws.Cells[RowIndex, 28].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 29].Value = dtx.Rows[i]["LOADPHCt"];
                    ws.Cells[RowIndex, 29].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 30].Value = "";
                    ws.Cells[RowIndex, 30].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 31].Value = dtx.Rows[i]["DESTPHCt"];
                    ws.Cells[RowIndex, 31].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 32].Value = dtx.Rows[i]["EXCOMMt"];
                    ws.Cells[RowIndex, 32].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 33].Value = dtx.Rows[i]["ICOMMt"];
                    ws.Cells[RowIndex, 33].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 34].Value = dtx.Rows[i]["MISCCOSTt"];
                    ws.Cells[RowIndex, 34].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 35].Value = 5.00;
                    ws.Cells[RowIndex, 35].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 36].Value = CostAmt;
                    ws.Cells[RowIndex, 36].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 37].Value = (RevenusAmt - CostAmt);
                    ws.Cells[RowIndex, 37].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    RowIndex++;

                }


                string styleRev = "#FFDAE1";
                string styleCost = "#DAFFED";
                string styleprofit = "#E1E199";
                rw_bgn = 7;
                colCost = System.Drawing.ColorTranslator.FromHtml(styleCost);
                ws.Cells["O" + (rw_bgn - 2) + ":W" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["O" + (rw_bgn - 2) + ":W" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(styleRev);
                ws.Cells["W" + (rw_bgn - 2) + ":AI" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["W" + (rw_bgn - 2) + ":AI" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(styleprofit);
                ws.Cells["AH" + (rw_bgn - 2) + ":AH" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["AH" + (rw_bgn - 2) + ":AH" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                ws.Cells["A3:AJ" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:AJ" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:AJ" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:AJ" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, 34, 36].AutoFitColumns();

                string _filePath = Server.MapPath("~/PreXL/");
                string _fileName = "MISReport " + Session.SessionID + ".xlsx";
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



        public DataTable GetCntrDtlsTerminalDepPdfValues(string AgencyID)
        {
            string strWhere = "";
            //string _Query = " select BookingNo,POL,POD,FPOD,TSPORT,(select top(1) CntrNo from NVO_Containers where Id =CntrID) as CntrNo,VesVoy, " +
            //                " (select top(1) CntrNo from NVO_Containers where Id = CntrID) as CntrNo, " +
            //                " (select(select sum(SizeID) from NVO_tblCntrTypes where SizeID = 1 and EQTypeID = 1 and Id = NVO_Containers.TypeID) from NVO_Containers where Id = CntrID) as Size20, " +
            //                " isnull((select(select sum(SizeID) from NVO_tblCntrTypes where SizeID in (2) and EQTypeID = 1 and Id = NVO_Containers.TypeID) from NVO_Containers where Id = CntrID),0) as Size40, " +
            //                " isnull((select(select sum(SizeID) from NVO_tblCntrTypes where SizeID in (2) and EQTypeID = 2 and Id = NVO_Containers.TypeID) from NVO_Containers where Id = CntrID),0) as SizeRF40, " +
            //                " isnull((select(select sum(SizeID) from NVO_tblCntrTypes where SizeID in (2) and EQTypeID = 3 and Id = NVO_Containers.TypeID) from NVO_Containers where Id = CntrID),0) as SizeOT40, " +
            //                " (select(select top(1) GeneralName from NVO_GeneralMaster where Id = NVO_Containers.LeaseTermID) from NVO_Containers where Id = CntrID) as LeasTerms, " +
            //                " (select top(1) ManifRate from NVO_BLCharges where ChargeCodeID = 1  and BkgID = NVO_Booking.Id) as OFT, " +
            //                " (select top(1) (select top(1) CurrencyCode from NVO_CurrencyMaster where ID =CurrencyId) from NVO_BLCharges where ChargeCodeID = 1  and BkgID = NVO_Booking.Id) as OFTCURR, " +
            //                " (select top(1) ManifRate from NVO_BLCharges where ChargeCodeID = 22  and BkgID = NVO_Booking.Id) as BAF, " +
            //                " (select top(1) ManifRate from NVO_BLCharges where ChargeCodeID = 4  and BkgID = NVO_Booking.Id) as LOADTHC, " +
            //                " (select top(1) ManifRate from NVO_BLCharges where ChargeCodeID = 9  and BkgID = NVO_Booking.Id) as DESTTHC, '0' as HLDTPT, " +
            //                " (select sum(ManifRate) from NVO_BLCharges where ChargeCodeID in(17, 25, 13, 11)  and BkgID = NVO_Booking.Id) as MISC, " +
            //                " (select top(1)Amount from NVO_SLOTDDtls where NVO_SLOTDDtls.SLID = NVO_Booking.SlotContractID) as SlotAmt," +
            //                " (select top(1)(select top(1) CurrencyCode from NVO_CurrencyMaster where ID = CurrencyId) from NVO_SLOTDDtls where NVO_SLOTDDtls.SLID = NVO_Booking.SlotContractID) as SlotCURR, " +
            //                " (select top(1) ManifRate from NVO_BLVenCharges where TariffTypeID = 136 and PaymentModeID = 19 and ChargeCodeID =9 and BkgID=NVO_Booking.ID) as DESPORTTHC, " +
            //                " (select top(1) ManifRate from NVO_BLVenCharges where TariffTypeID = 136 and ChargeCodeID =4 and BkgID=NVO_Booking.ID) as LOADPORTTHC, " +
            //                "  (select sum(ManifRate) from NVO_BLVenCharges where  ChargeCodeID in(17,13,11) and BkgID=NVO_Booking.ID) as MISCCOST " +
            //                " from NVO_Booking " +
            //                " inner join NVO_BOL on NVO_BOL.BkgID = NVO_Booking.ID " +
            //                " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOL.ID";

            //string _Query = " select AgentID,Id,RRID,RRNO,BookingNo,POL,POD,FPOD,TSPORT,CntrNo,VesVoy,Terminal,SlotOperator,TypeSize,ETDDate,	" +
            //                " Size20,Size40,SizeRF40,SizeOT40,LeasTerms,OFT,OFTCURR,BAF,LOADTHC,LOADTHCCURR,DESTTHC,DESTTHCCURR,HLDTPT,MISC,MISCCCURR,SlotAmt,SlotCURR,	" +
            //                " DESPORTTHC,LOADPORTTHC,MISCCOST,OFTCURR,BAFCURR,SURCHARGES,SURCHARGESCURR,(OFT + BAF + SURCHARGES + LOADTHC + DESTTHC + HLDTPT + MISC ) as RevenuAmt, " +
            //                " TSCOST,TSCOSTCURR,EXCOMM,EXCOMMCURR,ICOMM,ICOMMCURR,TSCOMM,TSCOMMCURR ,LOADPHC,LOADPHCCURR,DESTPHC,DESTPHCCURR,MISCCOST,MISCCOSTCURR, " +
            //                " (TSCOST+EXCOMM+ICOMM+TSCOMM+LOADPHC+DESTPHC+MISCCOST) as COSTAmt, " +
            //                " ((OFT + BAF + SURCHARGES + LOADTHC + DESTTHC + HLDTPT + MISC ) - (TSCOST+EXCOMM+ICOMM+TSCOMM+LOADPHC+DESTPHC+MISCCOST)) as ProfitAmt,Commodity from NVO_V_MiSViewALLDataReport";
            string _Query = " select distinct AgentID,Id,RRID,RRNO,BookingNo,POL,POD,FPOD,TSPORT,CntrNo,VesVoy,Terminal,SlotOperator,TypeSize,ETDDate,	 " +
                          " Size20,Size40,SizeRF40,SizeOT40,LeasTerms,Commodity," +
                          " cast(isnull((OFT / (case when OFTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.OFTCurrId)  end)),0) as decimal(10,2)) as OFTt," +
                          " cast(isnull((BAF / (case when BAFCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.BAFCurrId)  end)),0) as decimal(10,2)) as BAFt, " +
                          " cast(isnull((SURCHARGES / (case when SURCHARGESCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SURCHARGESCurrId)  end)),0) as decimal(10,2))  as SURCHARGESt, " +
                          " cast(isnull((LOADTHC/ (case when LOADTHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOADTHCCurrId)  end)),0) as decimal(10,2))  as LOADTHCt, " +
                          " cast(isnull((DESTTHC / (case when DESTTHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DESTTHCCurrId)  end)),0) as decimal(10,2))  as DESTTHCt, " +
                          " cast(isnull((HLDTPT / (case when HLDTPTCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.HLDTPTCCurrId)  end)),0) as decimal(10,2))  as HLDTPTt, " +
                          " cast(isnull((MISC / (case when MISCCcurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCcurrId)  end)),0) as decimal(10,2)) as MISCt1, " +
                          " cast(isnull((SUV_MISC / (case when SUV_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SUV_MISCCCurrId)  end)),0)  + " +
                          " isnull((LOLO_MISC / (case when LOLO_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOLO_MISCCCurrId)  end)),0)  + " +
                          " isnull((WASH_MISC / (case when WASH_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.WASH_MISCCCurrId)  end)),0)  + " +
                          " isnull((CMC_MISC / (case when CMC_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.CMC_MISCCCurrId)  end)),0) as decimal(10,2)) as MISCt, " +
                          " cast(isnull((SlotAmt / (case when SlotCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SlotCurrId)  end)),0) as decimal(10,2)) as SlotAmtt," +
                          " cast(isnull((TSCOST / (case when TSCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOSTCurrId)  end)),0)  as decimal(10,2)) as TSCOSTt, " +
                          " cast(isnull((EXCOMM / (case when EXCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.EXCOMMCurrId)  end)),0) as decimal(10,2)) as EXCOMMt, " +
                          " cast(isnull((ICOMM / (case when ICOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.ICOMMCurrId)  end)),0) as decimal(10,2)) as ICOMMt, " +
                          " cast(isnull((TSCOMM / (case when TSCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOMMCurrId)  end)),0) as decimal(10,2)) as TSCOMMt, " +
                          " cast(isnull((LOADPHC / (case when LOADPHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOADPHCCurrId)  end)),0)  as decimal(10,2)) as LOADPHCt, " +
                          " cast(isnull((DESTPHC / (case when DESTPHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DESTPHCCurrId)  end)),0)  as decimal(10,2)) as DESTPHCt, " +
                          " cast(isnull((MISCCOST / (case when MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCOSTCurrId)  end)),0)  as decimal(10,2)) as MISCCOSTt1, " +
                          " cast(isnull((SUV_MISCCOST / (case when SUV_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SUV_MISCCOSTCurrId)  end)),0)  + " +
                          " isnull((LOLO_MISCCOST / (case when LOLO_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOLO_MISCCOSTCurrId)  end)),0)  + " +
                          " isnull((WASH_MISCCOST / (case when WASH_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.WASH_MISCCOSTCurrId)  end)),0) as decimal(10, 2)) as MISCCOSTt " +
                          " from NVO_V_MiSViewALLDataReport_New_All NVO_V_MiSViewALLDataReport";

            if (AgencyID != "" && AgencyID != "null" && AgencyID != "?" && AgencyID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where AgentID=" + AgencyID;
                else
                    strWhere += " and AgentID =" + AgencyID;
            if (strWhere == "")
                strWhere = _Query;
            return RegManage.GetViewData(strWhere, "");
        }


        public ActionResult MISReportAllExcel(string AgentID, string FromDate, string ToDate)
        {
            BindMISALLReportGetValue(AgentID, FromDate, ToDate);
            return View();
        }
        public void BindMISALLReportGetValue(string AgentID, string FromDate, string ToDate)
        {
            using (ExcelPackage _xlPV = new ExcelPackage())
            {
                _xlPV.Workbook.Properties.Author = "MIS REPORT";
                _xlPV.Workbook.Properties.Title = "MIS REPORT";
                _xlPV.Workbook.Worksheets.Add("MIS REPORT");
                ExcelWorksheet ws = _xlPV.Workbook.Worksheets[1];
                Color colCost;
                ws.Name = "REPORT"; //Setting Sheet's name
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
                int MaxCol = 43;
                int RowIndex = 2;
                ws.Cells[RowIndex, 1].Value = "MIS REPORT";
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Merge = true;
                //ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 14.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Color.SetColor(System.Drawing.Color.Navy);
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                RowIndex++;

                ws.Cells[RowIndex, 1, RowIndex, 14].Value = "";
                ws.Cells[RowIndex, 1, RowIndex, 14].Merge = true;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Size = 12.0F;
                ws.Cells[RowIndex, 1, RowIndex, MaxCol].Style.Font.Bold = true;


                colCost = System.Drawing.ColorTranslator.FromHtml("#DAFFED");
                ws.Cells[RowIndex, 20, RowIndex, 28].Value = "REVENUE";
                ws.Cells[RowIndex, 20, RowIndex, 28].Merge = true;
                ws.Cells[RowIndex, 20, RowIndex, 28].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 20, RowIndex, 28].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 20, RowIndex, 28].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#FFDAE1");
                ws.Cells[RowIndex, 29, RowIndex, 39].Value = "COST";
                ws.Cells[RowIndex, 29, RowIndex, 39].Merge = true;
                ws.Cells[RowIndex, 29, RowIndex, 39].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 29, RowIndex, 39].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 29, RowIndex, 39].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml("#E1E199");
                ws.Cells[RowIndex, 40, RowIndex, 40].Value = "PROFIT";
                ws.Cells[RowIndex, 40, RowIndex, 40].Merge = true;
                ws.Cells[RowIndex, 40, RowIndex, 40].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[RowIndex, 40, RowIndex, 40].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[RowIndex, 40, RowIndex, 40].Style.Fill.BackgroundColor.SetColor(colCost);



                //ws.Cells[RowIndex, 34].Value = "";



                RowIndex++;
                ws.Cells[RowIndex, 1].Value = "SR No.";
                ws.Cells[RowIndex, 2].Value = "Location.";
                ws.Cells[RowIndex, 3].Value = "Agent Name.";
                ws.Cells[RowIndex, 4].Value = "Rate Request No.";
                ws.Cells[RowIndex, 5].Value = "BLNo";
                ws.Cells[RowIndex, 6].Value = "Container No";
                ws.Cells[RowIndex, 7].Value = "POL";
                ws.Cells[RowIndex, 8].Value = "POD";

                ws.Cells[RowIndex, 9].Value = "T/S";

                ws.Cells[RowIndex, 10].Value = "FPOD";
                ws.Cells[RowIndex, 11].Value = "Vessel /Voy ";
                ws.Cells[RowIndex, 12].Value = "Terminal Name";
                ws.Cells[RowIndex, 13].Value = "Slot Operator";
                ws.Cells[RowIndex, 14].Value = "ETD'";
                ws.Cells[RowIndex, 15].Value = "Commodity Type ";
                ws.Cells[RowIndex, 16].Value = "20'";
                ws.Cells[RowIndex, 17].Value = "40'";
                ws.Cells[RowIndex, 18].Value = "Container Types";
                ws.Cells[RowIndex, 19].Value = "Lease Type";
                ws.Cells[RowIndex, 20].Value = "Ocen Freight ($)";
                ws.Cells[RowIndex, 21].Value = "BAF ($)";
                ws.Cells[RowIndex, 22].Value = "DGS ($)";
                ws.Cells[RowIndex, 23].Value = "Surcharges ($)";
                ws.Cells[RowIndex, 24].Value = "Load Port THC ($)";
                ws.Cells[RowIndex, 25].Value = "Discharge Port THC ($)";
                ws.Cells[RowIndex, 26].Value = "HDL/TPT ($)";
                ws.Cells[RowIndex, 27].Value = "MISC ($)";
                ws.Cells[RowIndex, 28].Value = "Total Revenue ($)";
                ws.Cells[RowIndex, 29].Value = "SLOT ($)";
                ws.Cells[RowIndex, 30].Value = "T/S THC COST ($)";
                ws.Cells[RowIndex, 31].Value = "T/S COMMISSION ($)";
                ws.Cells[RowIndex, 32].Value = "T/S SLOT COST ($)";
                ws.Cells[RowIndex, 33].Value = "LOAD PORT PHC ($)";
                ws.Cells[RowIndex, 34].Value = "DISCHARGE PORT PHC ($)";
                ws.Cells[RowIndex, 35].Value = "LOAD COMMISSION ($)";
                ws.Cells[RowIndex, 36].Value = "DISCHARGE COMMISSION ($)";
                ws.Cells[RowIndex, 37].Value = "MISC COST ($)";
                ws.Cells[RowIndex, 38].Value = "LOG. FEE ($)";
                ws.Cells[RowIndex, 39].Value = "TOTAL COST ($)";
                ws.Cells[RowIndex, 40].Value = "PROFIT / LOSS ($)";
                ws.Cells[RowIndex, 41].Value = "FRIGHT TERMS";
                ws.Cells[RowIndex, 42].Value = "SLOT TERMS";
                ws.Cells[RowIndex, 43].Value = "DESTINATION AGENT";


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
                DataTable dtx = GetALLCntrDtlsTerminalDepPdfValues(AgentID, FromDate,ToDate);
                for (int i = 0; i < dtx.Rows.Count; i++)
                {
                    decimal RevenusAmt = decimal.Parse(dtx.Rows[i]["OFTt"].ToString()) + decimal.Parse(dtx.Rows[i]["DGSt"].ToString()) + decimal.Parse(dtx.Rows[i]["BAFt"].ToString()) + decimal.Parse(dtx.Rows[i]["SURCHARGESt"].ToString()) + decimal.Parse(dtx.Rows[i]["LOADTHCt"].ToString()) + decimal.Parse(dtx.Rows[i]["HLDTPTt"].ToString()) + decimal.Parse(dtx.Rows[i]["MISRevAmt"].ToString()) + decimal.Parse(dtx.Rows[i]["DESTTHCt"].ToString());
                    decimal CostAmt = 5 + decimal.Parse(dtx.Rows[i]["SlotAmtt"].ToString()) + decimal.Parse(dtx.Rows[i]["TSCOSTt"].ToString()) + decimal.Parse(dtx.Rows[i]["EXCOMMt"].ToString()) + decimal.Parse(dtx.Rows[i]["ICOMMt"].ToString()) + decimal.Parse(dtx.Rows[i]["TSCOMMt"].ToString()) + decimal.Parse(dtx.Rows[i]["LOADPHCt"].ToString()) + decimal.Parse(dtx.Rows[i]["DESTPHCt"].ToString()) + decimal.Parse(dtx.Rows[i]["MISCCOSTt"].ToString());


                    InvCount++;
                    ws.Cells[RowIndex, 1].Value = InvCount;
                    ws.Cells[RowIndex, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 2].Value = dtx.Rows[i]["Locations"].ToString();
                    ws.Cells[RowIndex, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 3].Value = dtx.Rows[i]["AgentName"].ToString();
                    ws.Cells[RowIndex, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 4].Value = dtx.Rows[i]["RRNO"].ToString();
                    ws.Cells[RowIndex, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 5].Value = dtx.Rows[i]["BookingNo"].ToString();
                    ws.Cells[RowIndex, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 6].Value = dtx.Rows[i]["CntrNo"].ToString();
                    ws.Cells[RowIndex, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 7].Value = dtx.Rows[i]["POL"].ToString();
                    ws.Cells[RowIndex, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 8].Value = dtx.Rows[i]["POD"].ToString();
                    ws.Cells[RowIndex, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 9].Value = dtx.Rows[i]["TSPORT"].ToString();
                    ws.Cells[RowIndex, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 10].Value = dtx.Rows[i]["FPOD"].ToString();
                    ws.Cells[RowIndex, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 11].Value = dtx.Rows[i]["VesVoy"].ToString();
                    ws.Cells[RowIndex, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 12].Value = dtx.Rows[i]["Terminal"].ToString();
                    ws.Cells[RowIndex, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 13].Value = dtx.Rows[i]["SlotOperator"].ToString();
                    ws.Cells[RowIndex, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 14].Value = dtx.Rows[i]["ETDDate"].ToString();
                    ws.Cells[RowIndex, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 15].Value = dtx.Rows[i]["Commodity"].ToString();
                    ws.Cells[RowIndex, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 16].Value = dtx.Rows[i]["Size20"].ToString();
                    ws.Cells[RowIndex, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 17].Value = dtx.Rows[i]["Size40"].ToString();
                    ws.Cells[RowIndex, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 18].Value = dtx.Rows[i]["TypeSize"].ToString();
                    ws.Cells[RowIndex, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 19].Value = dtx.Rows[i]["LeasTerms"].ToString();
                    ws.Cells[RowIndex, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 20].Value = dtx.Rows[i]["OFTt"];
                    ws.Cells[RowIndex, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 21].Value = dtx.Rows[i]["BAFt"];
                    ws.Cells[RowIndex, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 22].Value = dtx.Rows[i]["DGSt"];
                    ws.Cells[RowIndex, 22].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 23].Value = dtx.Rows[i]["SURCHARGESt"];
                    ws.Cells[RowIndex, 23].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 24].Value = dtx.Rows[i]["LOADTHCt"];
                    ws.Cells[RowIndex, 24].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 25].Value = dtx.Rows[i]["DESTTHCt"];
                    ws.Cells[RowIndex, 25].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 26].Value = dtx.Rows[i]["HLDTPTt"];
                    ws.Cells[RowIndex, 26].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["MISCt"];
                    //ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 27].Value = dtx.Rows[i]["MISRevAmt"];
                    ws.Cells[RowIndex, 27].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    


                    ws.Cells[RowIndex, 28].Value = RevenusAmt;
                    ws.Cells[RowIndex, 28].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    ws.Cells[RowIndex, 29].Value = dtx.Rows[i]["SlotAmtt"];
                    ws.Cells[RowIndex, 29].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 30].Value = dtx.Rows[i]["TSCOSTt"];
                    ws.Cells[RowIndex, 30].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    //ws.Cells[RowIndex, 30].Value = // dtx.Rows[i]["TSCOMMt"];
                    if (dtx.Rows[i]["TSPORT"].ToString() == "")
                        ws.Cells[RowIndex, 31].Value = 0.00;
                    else
                        ws.Cells[RowIndex, 31].Value = 10.00;
                    ws.Cells[RowIndex, 31].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 32].Value = dtx.Rows[i]["TSSlotCost"]; 
                    ws.Cells[RowIndex, 32].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 33].Value = dtx.Rows[i]["LOADPHCt"];
                    ws.Cells[RowIndex, 33].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 34].Value = dtx.Rows[i]["DESTPHCt"];
                    ws.Cells[RowIndex, 34].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    decimal ExCommt = decimal.Parse(dtx.Rows[i]["EXCOMMt"].ToString());
                    if(ExCommt <= 10)
                        ws.Cells[RowIndex, 35].Value = 10;
                    else
                        ws.Cells[RowIndex, 35].Value = dtx.Rows[i]["EXCOMMt"];


                    ws.Cells[RowIndex, 35].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    //ws.Cells[RowIndex, 35].Value = dtx.Rows[i]["EXCOMMt"];
                    //ws.Cells[RowIndex, 35].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    decimal ExICommt = decimal.Parse(dtx.Rows[i]["ICOMMt"].ToString());

                    if (ExICommt <= 10)
                        ws.Cells[RowIndex, 36].Value = 10;
                    else
                        ws.Cells[RowIndex, 36].Value = dtx.Rows[i]["ICOMMt"];


                    ws.Cells[RowIndex, 36].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    //ws.Cells[RowIndex, 36].Value = dtx.Rows[i]["ICOMMt"];
                    //ws.Cells[RowIndex, 36].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 37].Value = dtx.Rows[i]["MISCCOSTt"];
                    ws.Cells[RowIndex, 37].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 38].Value = 5.00;
                    ws.Cells[RowIndex, 38].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 39].Value = CostAmt;
                    ws.Cells[RowIndex, 39].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 40].Value = (RevenusAmt - CostAmt);
                    ws.Cells[RowIndex, 40].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 41].Value = dtx.Rows[i]["FreightTerms"].ToString();
                    ws.Cells[RowIndex, 41].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 42].Value = dtx.Rows[i]["SlotTerms"].ToString();
                    ws.Cells[RowIndex, 42].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    ws.Cells[RowIndex, 43].Value = dtx.Rows[i]["DesAgencyName"].ToString();
                    ws.Cells[RowIndex, 43].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


                    RowIndex++;

                }


                string styleRev = "#FFDAE1";
                string styleCost = "#DAFFED";
                string styleprofit = "#E1E199";
                rw_bgn = 7;
                colCost = System.Drawing.ColorTranslator.FromHtml(styleCost);
                ws.Cells["T" + (rw_bgn - 2) + ":AB" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["T" + (rw_bgn - 2) + ":AB" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(styleRev);
                ws.Cells["AC" + (rw_bgn - 2) + ":AM" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["AC" + (rw_bgn - 2) + ":AM" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                colCost = System.Drawing.ColorTranslator.FromHtml(styleprofit);
                ws.Cells["AN" + (rw_bgn - 2) + ":AN" + RowIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["AN" + (rw_bgn - 2) + ":AN" + RowIndex].Style.Fill.BackgroundColor.SetColor(colCost);

                ws.Cells["A3:AP" + RowIndex].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:AP" + RowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:AP" + RowIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A3:AP" + RowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, 34, 43].AutoFitColumns();

                string _filePath = Server.MapPath("~/PreXL/");
                string _fileName = "MISReport " + Session.SessionID + ".xlsx";
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


        public DataTable GetALLCntrDtlsTerminalDepPdfValues(string AgentID, string FromDate, string ToDate)
        {
            string strWhere = "";

            string _Query = " select AgentID,Id,RRID,RRNO,BookingNo,POL,POD,FPOD,TSPORT,CntrNo,VesVoy,Terminal,SlotOperator,TypeSize,ETDDate,ETDDatev,FreightTerms,	 " +
                            " (select top(1)(select top(1) GeoLocation from NVO_GeoLocations  where NVO_GeoLocations.ID=NVO_AgencyMaster.GeoLocationID) from NVO_AgencyMaster where NVO_AgencyMaster.Id = NVO_V_MiSViewALLDataReport.AgentID) as Locations, " +
                            " (select top(1) AgencyName from NVO_AgencyMaster where NVO_AgencyMaster.ID = NVO_V_MiSViewALLDataReport.AgentID) as AgentName, " +
                          " Size20,Size40,SizeRF40,SizeOT40,LeasTerms,Commodity," +
                          " cast(isnull((OFT / (case when OFTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.OFTCurrId)  end)),0) as decimal(10,2)) as OFTt," +
                          " cast(isnull((BAF / (case when BAFCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.BAFCurrId)  end)),0) as decimal(10,2)) as BAFt, " +
                          " cast(isnull((SURCHARGES / (case when SURCHARGESCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SURCHARGESCurrId)  end)),0) as decimal(10,2))  as SURCHARGESt, " +
                          " cast(isnull((DGS / (case when DGS_CCURRId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DGS_CCURRId)  end)),0) as decimal(10,2))  as DGSt, " +
                          " cast(isnull((LOADTHC/ (case when LOADTHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOADTHCCurrId)  end)),0) as decimal(10,2))  as LOADTHCt, " +
                          " cast(isnull((DESTTHC / (case when DESTTHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DESTTHCCurrId)  end)),0) as decimal(10,2))  as DESTTHCt, " +
                          " cast(isnull((HLDTPT / (case when HLDTPTCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.HLDTPTCCurrId)  end)),0) as decimal(10,2))  as HLDTPTt, " +
                          " cast(isnull((MISC / (case when MISCCcurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCcurrId)  end)),0) as decimal(10,2)) as MISCt1, " +

                          " cast(isnull((SUV_MISC / (case when SUV_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SUV_MISCCCurrId)  end)),0)  + " +
                          " isnull((LOLO_MISC / (case when LOLO_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOLO_MISCCCurrId)  end)),0)  + " +
                          " isnull((WASH_MISC / (case when WASH_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.WASH_MISCCCurrId)  end)),0)  + " +
                          " isnull((CMC_MISC / (case when CMC_MISCCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.CMC_MISCCCurrId)  end)),0) as decimal(10,2)) as MISCt, " +
                          " SlotAmt  as SlotAmtt," +
                          //" cast(isnull((SlotAmt / (case when SlotCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SlotCurrId)  end)),0) as decimal(10,2)) as SlotAmtt," +
                          " cast(isnull((TSCOST / (case when TSCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOSTCurrId)  end)),0)  as decimal(10,2)) as TSCOSTt, " +
                          " TSSlotCost, " +
                          // " cast(isnull((TSSlotCost / (case when TSCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOSTCurrId)  end)),0)  as decimal(10,2)) as TSSlotCost,"+

                          " cast(isnull((EXCOMM / (case when EXCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.EXCOMMCurrId)  end)),0) as decimal(10,2)) as EXCOMMt, " +
                          " cast(isnull((ICOMM / (case when ICOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.ICOMMCurrId)  end)),0) as decimal(10,2)) as ICOMMt, " +
                        
                          " cast(isnull((TSCOMM / (case when TSCOMMCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.TSCOMMCurrId)  end)),0) as decimal(10,2)) as TSCOMMt, " +
                          " cast(isnull((LOADPHC / (case when LOADPHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOADPHCCurrId)  end)),0)  as decimal(10,2)) as LOADPHCt, " +
                          " cast(isnull((DESTPHC / (case when DESTPHCCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.DESTPHCCurrId)  end)),0)  as decimal(10,2)) as DESTPHCt, " +
                          " cast(isnull((MISCCOST / (case when MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCOSTCurrId)  end)),0)  as decimal(10,2)) as MISCCOSTt1, " +

                          " cast(isnull((SUV_MISCCOST / (case when SUV_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.SUV_MISCCOSTCurrId)  end)),0)  + " +
                          " isnull((LOLO_MISCCOST / (case when LOLO_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.LOLO_MISCCOSTCurrId)  end)),0)  + " +
                          " isnull((WASH_MISCCOST / (case when WASH_MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.WASH_MISCCOSTCurrId)  end)),0) as decimal(10, 2)) as MISCCOSTt, " +
                          " MISRevAmt, "+
                          " cast(isnull((MISCostAmt / (case when MISCCOSTCurrId = 146  then 1 else (select top(1) ExRate from NVO_MISExRateCountrywise where NVO_MISExRateCountrywise.Id = NVO_V_MiSViewALLDataReport.MISCCOSTCurrId)  end)),0)  as decimal(10,2)) as MISCostAmtv,SlotTerms, " +
                          " DesAgencyName " +

                          " from NVO_V_MiSViewALLDataReport_New_All NVO_V_MiSViewALLDataReport";

            if (AgentID != "" && AgentID != "null" && AgentID != "?" && AgentID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where AgentID=" + AgentID;
                else
                    strWhere += " and AgentID =" + AgentID;

            if (FromDate != "" && FromDate != "undefined" || ToDate != "" && ToDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where  convert(varchar,ETDDatev,23) between '" + FromDate + "' and '" + ToDate + "'";
                else
                    strWhere += "  and convert(varchar,ETDDatev,23) between '" + FromDate + "' and '" + ToDate + "'";

            if (strWhere == "")
                strWhere = _Query;
            return RegManage.GetViewData(strWhere, "");
        }
    }
}
