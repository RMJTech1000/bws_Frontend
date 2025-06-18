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
    public class ImportExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: ImportExcel
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ImportLoadingReport()
        {
            return View();
        }
        public ActionResult ImportSummaryReport()
        {
            return View();
        }
        public ActionResult DODetailsReport()
        {
            return View();
        }
        public void ImportSummaryReportView(string DtFrom, string DtTo, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetImportLoadingView(DtFrom, DtTo, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ImportSummaryReport");

                ws.Cells["A2"].Value = "Import Summary Report";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:S2"];
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
                ws.Cells["B7"].Value = "Import Vsl/Voy";
                ws.Cells["C7"].Value = "IGM NO";
                ws.Cells["D7"].Value = "IGM DATE";
                ws.Cells["E7"].Value = "Discharge Terminal";
                ws.Cells["F7"].Value = "Berth on(FV)";
                ws.Cells["G7"].Value = "Container No";
                ws.Cells["H7"].Value = "Size";

                ws.Cells["I7"].Value = "CFS DETAILS";
                ws.Cells["J7"].Value = "BL Number";
                ws.Cells["K7"].Value = "STATUS / gen / haz";
                ws.Cells["L7"].Value = "Principal";
                ws.Cells["M7"].Value = "POL";
                ws.Cells["N7"].Value = "T/P";
                ws.Cells["O7"].Value = "POD";
                ws.Cells["P7"].Value = "FPOD";
                ws.Cells["Q7"].Value = "SHIPPER";
                ws.Cells["R7"].Value = "CONSIGNEE";
                ws.Cells["S7"].Value = "NOM/FREEHAND";
                ws.Cells["T7"].Value = "CFS NOM";
                ws.Cells["U7"].Value = "Free Days";
                ws.Cells["V7"].Value = " DETENTION ";
                ws.Cells["W7"].Value = "OCEAN FREIGHT";
                ws.Cells["X7"].Value = "BAF";
                ws.Cells["Y7"].Value = "THC";
                ws.Cells["Z7"].Value = "IHC";
                ws.Cells["AA7"].Value = "WASHING";
                ws.Cells["AB7"].Value = "CMC";
                ws.Cells["AC7"].Value = "EIS";
                ws.Cells["AD7"].Value = "SURVEY";
                ws.Cells["AE7"].Value = "ISPS";
                ws.Cells["AF7"].Value = "SEAL";
                ws.Cells["AG7"].Value = "TOLL";
                ws.Cells["AH7"].Value = "MUC";
                ws.Cells["AI7"].Value = "ADMIN";
                ws.Cells["AJ7"].Value = "PCS";
                ws.Cells["AK7"].Value = "DELIVERY ORDER";
                ws.Cells["AL7"].Value = "HBL";
                ws.Cells["AM7"].Value = "DOC / BL";
                ws.Cells["AN7"].Value = "HIGH SEA SALES ";
                ws.Cells["AO7"].Value = " DO REVALIDATION FEE ";
                ws.Cells["AP7"].Value = " FREIGHT TERM ";

                r = ws.Cells["A7:AP7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 0;

                int rw = 8;
                int frowid = 0;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    frowid = rw;


                    //ExcelRange rng = ws.Cells["A" + frowid + ":V" + frowid];
                    //rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    sl++;

                     ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["BLVesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["IGMNo"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["IGMDate"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Terminal"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["FVDate"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["CntrType"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["CFS"].ToString();                                     
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();                                 
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["Principal"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["N" + rw].Value = "";  //Transhipment
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["FPOD"].ToString();
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["Shipper"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["Consignee"].ToString();
                    ws.Cells["S" + rw].Value = dtv.Rows[i]["SalesPerson"].ToString();
                    ws.Cells["T" + rw].Value = dtv.Rows[i]["CFSNominated"].ToString();
                    ws.Cells["U" + rw].Value = dtv.Rows[i]["FreeDays"].ToString();
                    ws.Cells["V" + rw].Value = "";                       //detention
                    ws.Cells["W" + rw].Value = dtv.Rows[i]["FRT"].ToString();
                    ws.Cells["X" + rw].Value = dtv.Rows[i]["BAF"].ToString();
                    ws.Cells["Y" + rw].Value = dtv.Rows[i]["THC"].ToString();
                    ws.Cells["Z" + rw].Value = dtv.Rows[i]["IHC"].ToString();
                    ws.Cells["AA" + rw].Value = dtv.Rows[i]["WASHING"].ToString();
                    ws.Cells["AB" + rw].Value = dtv.Rows[i]["CMC"].ToString();
                    ws.Cells["AC" + rw].Value = dtv.Rows[i]["EIS"].ToString();
                    ws.Cells["AD" + rw].Value = dtv.Rows[i]["SUV"].ToString();
                    ws.Cells["AE" + rw].Value = dtv.Rows[i]["ISPS"].ToString();
                    ws.Cells["AF" + rw].Value = dtv.Rows[i]["SEAL"].ToString();
                    ws.Cells["AG" + rw].Value = dtv.Rows[i]["TOLL"].ToString();
                    ws.Cells["AH" + rw].Value = dtv.Rows[i]["MUC"].ToString();
                    ws.Cells["AI" + rw].Value = dtv.Rows[i]["ADM"].ToString();
                    ws.Cells["AJ" + rw].Value = dtv.Rows[i]["PCS"].ToString();
                    ws.Cells["AK" + rw].Value = dtv.Rows[i]["DOF"].ToString();
                    ws.Cells["AL" + rw].Value = dtv.Rows[i]["HBL"].ToString();
                    ws.Cells["AM" + rw].Value = dtv.Rows[i]["DOC_BL"].ToString();
                    ws.Cells["AN" + rw].Value = dtv.Rows[i]["HSS"].ToString();
                    ws.Cells["AO" + rw].Value = dtv.Rows[i]["DOR"].ToString();
                    ws.Cells["AP" + rw].Value = dtv.Rows[i]["FREIGHTTERM"].ToString();
                    

                    rw++;
                }

                rw -= 1;

                ws.Cells["A7:AP7" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AP7" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AP7" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:AP7" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 40].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ImportSummaryReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetImportLoadingView(string DtFrom, string DtTo, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = " select(select top 1 GeoLocation from NVO_GeoLocations Where id = NVO_AgencyMaster.GeoLocationID) AS Bkgloc, BLNumber, BLVesVoy, " +
                            " (select top(1) IGMNo from NVO_ImpCAN where NVO_ImpCAN.BkgId = IMPS.BkgID) as IGMNo, " +
                            " (select top(1) convert(varchar, IGMDate, 103) from NVO_ImpCAN where NVO_ImpCAN.BkgId = IMPS.BkgID) as IGMDate, " +
                            " (select top(1)(select top(1) TerminalName from NVO_TerminalMaster where ID = TerminalID) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = IMPS.BLVesVoyID) as Terminal, " +
                            " (select top(1) ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = IMPS.BLVesVoyID) ETA,POL,POD, " +
                            " (select top(1) PortName from NVO_PortMaster where ID = IMPS.FPODID) as FPOD, " +
                            " (select top(1)((select top(1) CustomerName from NVO_view_CustomerDetails where CID = PartID)) from NVO_BOLCustomerDetails where BLID = IMPS.ID and PartyTypeID = 1) as Shipper,  " +
                            " (select top(1)((select top(1) CustomerName from NVO_view_CustomerDetails where CID = PartID)) from NVO_BOLCustomerDetails where BLID = IMPS.ID and PartyTypeID = 2) as Consignee, " +
                            " Size as CntrType, " +
                            " (select top(1) CntrNo from NVO_Containers where NVO_Containers.Id = NVO_BOLCntrDetails.CntrID and NVO_BOLCntrDetails.BkgId = IMPS.BKgID) as CntrNo, " +
                            " isnull((select(select count(SizeID) from NVO_tblCntrTypes where SizeID = 1 and EQTypeID = 1 and Id = NVO_Containers.TypeID) from NVO_Containers where NVO_Containers.ID = NVO_BOLCntrDetails.CntrID),0) as Size20,  " +
                            " isnull((select(select count(SizeID) from NVO_tblCntrTypes where SizeID in (2, 3) and EQTypeID = 1 and Id = NVO_Containers.TypeID) from NVO_Containers where NVO_Containers.Id = NVO_BOLCntrDetails.CntrID),0) as Size40, " +
                            " isnull ((Select TOP 1 DtMovement from NVO_ContainerTxns where ContainerID=NVO_BOLCntrDetails.CntrID AND StatusCode='FV' AND BLNumber =IMPS.BkgID),'') FVDate, " +
                            " CommodityType,(select top(1) RRID from NVO_Booking where NVO_Booking.ID= IMPS.BkgID) as RRID,(select top(1) SalesPerson from NVO_Booking where NVO_Booking.ID= IMPS.BkgID) as SalesPerson, " +
                            " ISNULL((select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID =  IMPS.BkgID AND ChargeCodeID = 1),0) FRT, " +
                            " ISNULL((select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 22),0) BAF, " +
                              " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 4),0) THC, " +
                             " ISNULL((select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 24),0) IHC, " +
                                " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 17),0) WASHING, " +
                               " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 25),0) CMC, " +
                                  " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 37),0) EIS, " +
                               " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 11),0) SUV, " +
                                " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 47),0) ISPS, " +
                                  " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 10),0) SEAL, " +
                                 " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 12),0) TOLL, " +
                                   " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 2),0) MUC, " +
                                     " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 18),0) ADM, " +
                                   " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 38),0) PCS, " +
                                " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 19),0) DOF, " +
                              " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 44),0) HBL, " +
                                   " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 6),0) DOC_BL, " +
                                  " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 39),0) DOR, " +
                                     " ISNULL( (select TOP 1 CustomerRate from NVO_BLCharges WHERE BkgID = IMPS.BkgID AND ChargeCodeID = 87),0) HSS, " +
                               " ISNULL((select TOP 1 GeneralName from NVO_GeneralMaster GM  inner join NVO_BLCharges on  NVO_BLCharges.PaymentModeID = GM.ID  WHERE NVO_BLCharges.BkgID =  IMPS.BkgID AND ChargeCodeID = 1),0) FREIGHTTERM,IMPS.FreeDays," +
                                " ( Select top 1 CityName from NVO_CityMaster Where ID =CFS) AS CFS, " +
                                " ( Select top 1 GeneralName from NVO_GeneralMaster Where ID =NominationCFS) AS CFSNominated, " +
                               " ( Select top 1 CustomerName from NVO_CustomerMaster Where ID =PrincipalID) AS Principal " +

                          " from NVO_v_ImportSummaryReportView IMPS inner join NVO_Age" +
                          "ncyMaster on NVO_AgencyMaster.ID = IMPS.AgencyID " +
                            " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = IMPS.ID  ";

            


            if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "2" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

                if (strWhere == "")
                    strWhere += _Query + " where IMPS.AgencyID = " + AgencyID.ToString();
                else
                    strWhere += " and IMPS.AgencyID = " + AgencyID.ToString();

            //if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
            //    if (strWhere == "")
            //        strWhere += _Query + " and NVO_Booking.BkgDate between '" + DtFrom + "' and '" + DtTo + "'";
            //    else
            //        strWhere += "  and NVO_Booking.BkgDate between '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public void ImportDOReportView(string DtFrom, string DtTo, string Status, string User, string AgencyID)
        {
            DataTable dtv = GetDODetailsView(DtFrom, DtTo, Status, AgencyID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ImportDOReport");

                ws.Cells["A2"].Value = "Import DO Report";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:I2"];
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
                ws.Cells["B7"].Value = "D/O Number";
                ws.Cells["C7"].Value = "D/O Date";
                ws.Cells["D7"].Value = "Vessel/Voyage";
                ws.Cells["E7"].Value = "BL Number";
                ws.Cells["F7"].Value = "Container No";
                ws.Cells["G7"].Value = "Size";
                ws.Cells["H7"].Value = "Customer Name";
                ws.Cells["I7"].Value = "Principal";
              
                r = ws.Cells["A7:I7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int sl = 0;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    sl++;

                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["DONo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["IssueDate"].ToString();
                    ws.Cells["D" + rw].Value = "";
                    ws.Cells["E" + rw].Value = "";
                    ws.Cells["F" + rw].Value = "";
                    ws.Cells["G" + rw].Value = "";
                    ws.Cells["H" + rw].Value = "";
                    ws.Cells["I" + rw].Value = "";
                   

                    rw++;
                }

                rw -= 1;

                ws.Cells["A7:I7" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I7" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I7" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I7" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 10].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=DODetailsReport.xlsx");
                Response.End();

            }

        }

        public DataTable GetDODetailsView(string DtFrom, string DtTo, string Status, string AgencyID)
        {
            string strWhere = "";
            string _Query = "select * from NVO_ImpDeliveryOrder DO ";




            //if (AgencyID.ToString() != "" && AgencyID.ToString() != "0" && AgencyID.ToString() != "2" && AgencyID.ToString() != "undefined" && AgencyID.ToString() != null)

            //    if (strWhere == "")
            //        strWhere += _Query + " where IMPS.AgencyID = " + AgencyID.ToString();
            //    else
            //        strWhere += " and IMPS.AgencyID = " + AgencyID.ToString();

            if (DtFrom != "" && DtFrom != "undefined" || DtTo != "" && DtTo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where DO.IssueDate between '" + DtFrom + "' and '" + DtTo + "'";
                else
                    strWhere += "  and DO.BkgDate IssueDate '" + DtFrom + "' and '" + DtTo + "'";

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

    }
}