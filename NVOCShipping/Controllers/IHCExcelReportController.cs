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
    public class IHCExcelReportController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: IHCExcelReport
        public ActionResult Index()
        {
            return View();
        }

        public void CreateIHCExcelReport(string ICDLocID, string PortID, string Status, string ContainerTypeID, string CommodityTypeID, string ChargeTypeID, string User)
        {
            DataTable dtv = GetIHCExcelReportValues(ICDLocID, PortID, Status, ContainerTypeID, CommodityTypeID, ChargeTypeID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("IHCTariffList");

             
                //Record Headers

                ws.Cells["A1"].Value = "ID";
                ws.Cells["B1"].Value = "ICD_NAME";
                ws.Cells["C1"].Value = "PORT";
                ws.Cells["D1"].Value = "CHARGE_TYPE";
                ws.Cells["E1"].Value = "CHARGES";
                ws.Cells["F1"].Value = "COMMODITY_TYPE";
                ws.Cells["G1"].Value = "CONTAINER_TYPE";
                ws.Cells["H1"].Value = "THC_INCLUDED";
                ws.Cells["I1"].Value = "STATUS";
                ws.Cells["J1"].Value = "VALID_FROM";
                ws.Cells["K1"].Value = "VALID_TILL";
                ws.Cells["L1"].Value = "SLAB_FROM";
                ws.Cells["M1"].Value = "SLAB_TO";
                ws.Cells["N1"].Value = "CURRENCY";
                ws.Cells["O1"].Value = "TID";
                ws.Cells["P1"].Value = "BREAKUP_CHARGE";
                ws.Cells["Q1"].Value = "PAYMENT_TO";
                ws.Cells["R1"].Value = "AMOUNT";
                ExcelRange r = ws.Cells["A1:R1"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                int sl = 1;

                int rw = 2;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"];
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ICDName"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["PORT"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["CHARGE_TYPE"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["CHARGES"].ToString();
                    
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["COMMODITY_TYPE"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["CONTAINER_TYPE"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["THC_INCLUDED"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["STATUSV"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["VALID_FROM"].ToString(); 
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["VALID_TILL"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["SLAB_FROM"];
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["SLAB_TO"];
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["CURRENCY"].ToString(); 
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["TID"]; 
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["BREAKUP_CHARGE"].ToString(); 
                    ws.Cells["Q" + rw].Value = dtv.Rows[i]["PAYMENT_TO"].ToString();
                    ws.Cells["R" + rw].Value = dtv.Rows[i]["AMOUNT"];
                   
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A1:S" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:S" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:S" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:S" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 20].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=IHCTariffList.xlsx");
                Response.End();

            }

        }

        public DataTable GetIHCExcelReportValues(string ICDLocID, string PortID, string Status, string ContainerTypeID, string CommodityTypeID, string ChargeTypeID)
        {
            string strWhere = "";
            string _Query = "select IH.ID,IH.ICDName,(select top 1 PortName from NVO_PortMaster WHERE ID = IH.PortID ) AS PORT, "+
            " (Select TOP 1 GeneralName FROM NVO_GeneralMaster WHERE ID = IH.ChargeTypeID) AS CHARGE_TYPE, "+
            " (Select top 1 ChgCode from NVO_ChargeTB WHERE ID=IH.ChargesID) AS CHARGES, " +
           " (Select TOP 1 GeneralName FROM NVO_GeneralMaster WHERE ID = IH.CommodityTypeID ) AS COMMODITY_TYPE, "+
          " (Select TOP 1 Size FROM NVO_tblCntrTypes WHERE ID = IH.ContainerTypeID ) AS CONTAINER_TYPE, "+
          " case when IH.THCIncluded = 1 then 'YES' Else 'NO' End as THC_INCLUDED, case when IH.STATUS = 1 then 'ACTIVE' Else 'INACTIVE' End as STATUSV," +
         " convert(varchar, IH.ValidFrom, 105) AS VALID_FROM, convert(varchar, IH.ValidTill, 105) AS VALID_TILL, "+
        " IDT.SlabFrom as SLAB_FROM,IDT.SlabTo as SLAB_TO, (Select TOP 1 CurrencyCode FROM NVO_CurrencyMaster WHERE ID = IDT.CurrencyID ) AS CURRENCY,IDT.TID,isnull(IC.TariffChID, 0) TariffChID, "+
        " isnull((Select TOP 1 ChgCode FROM NVO_ChargeTB WHERE ID = IC.ChargeCodeID) ,'') AS BREAKUP_CHARGE, "+
         "isnull ((Select TOP 1 GeneralName FROM NVO_GeneralMaster WHERE ID = IC.PaymentTo ),'') AS PAYMENT_TO, Case when IH.ChargeTypeID=102 then   isnull (IDT.Amount,0) when IH.ChargeTypeID=103 then isnull(IC.Amount, 0) end as AMOUNT "+
       " FROM NVO_IHCHaulageTariff IH Inner Join NVO_IHCHaulageTariffDtls IDT on  IDT.IHCTariffID = IH.ID LEFT OUTER Join NVO_PortTariffIHCChargedtls IC on  IC.TariffChID = IDT.TID  ";


            if (ICDLocID != "" && ICDLocID != "0" && ICDLocID != "null" && ICDLocID != "?" && ICDLocID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where IH.ICDLocID=" + ICDLocID;
                else
                    strWhere += " and IH.ICDLocID =" + ICDLocID;

            if (PortID != "" && PortID != "0" && PortID != "null" && PortID != "?" && PortID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where IH.PortID=" + PortID;
                else
                    strWhere += " and IH.PortID =" + PortID;

            if (ChargeTypeID != "" && ChargeTypeID != "0" && ChargeTypeID != "null" && ChargeTypeID != "?" && ChargeTypeID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where IH.ChargeTypeID=" + ChargeTypeID;
                else
                    strWhere += " and IH.ChargeTypeID =" + ChargeTypeID;

            if (ContainerTypeID != "" && ContainerTypeID != "0" && ContainerTypeID != "null" && ContainerTypeID != "?" && ContainerTypeID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where IH.ContainerTypeID=" + ContainerTypeID;
                else
                    strWhere += " and IH.ContainerTypeID =" + ContainerTypeID;

            if (CommodityTypeID != "" && CommodityTypeID != "0" && CommodityTypeID != "null" && CommodityTypeID != "?" && CommodityTypeID != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " Where IH.CommodityTypeID=" + CommodityTypeID;
                else
                    strWhere += " and IH.CommodityTypeID =" + CommodityTypeID;

            if (Status != "" && Status != "0" && Status != "null" && Status != "?" && Status != "undefined")

                if (strWhere == "")
                    strWhere += _Query + " where  IH.Status = " + Status;
                else
                    strWhere += " and IH.Status = " + Status;

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }
    }
}