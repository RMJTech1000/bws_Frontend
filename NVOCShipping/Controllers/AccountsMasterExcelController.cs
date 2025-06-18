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
    public class AccountsMasterExcelController : Controller
    {
        MasterManager Manag = new MasterManager();
        // GET: RRExcel
        public ActionResult Index()
        {
            return View();
        }
        public void CreateGLMasterExcel(string GLCode, string GLDesc, string MainAccType, string Category, string User)
        {
            DataTable dtv = GetGLMasterSearchValues(GLCode, GLDesc, MainAccType, Category);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("GLMaster");
                //Record Headers

                ws.Cells["A1"].Value = "ID";
                ws.Cells["B1"].Value = "CODE";
                ws.Cells["C1"].Value = "GLDESCRIPTION";
                ws.Cells["D1"].Value = "COMPANY";
                ws.Cells["E1"].Value = "MAINACCOUNT";
                ws.Cells["F1"].Value = "NATURE";
                ws.Cells["G1"].Value = "CATEGORY";
                ws.Cells["H1"].Value = "GLMATCHING";
                ws.Cells["I1"].Value = "GLSTATUS";
                ExcelRange r ;
                r = ws.Cells["A1:I1"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                int sl = 1;

                int rw = 2;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"];
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["GLCode"];
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["GLDesc"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["MainAccType"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["GLNature"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Category"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["GLMatching"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Status"].ToString();

                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A1:I" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:I" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:I" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:I" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=GLMasterList.xlsx");
                Response.End();

            }

        }

        public DataTable GetGLMasterSearchValues(string GLCode, string GLDesc, string MainAccType, string Category)
        {
            string strWhere = "";
            string _Query = " select ID,GLCode,GLDesc,(Select top 1 AgencyName from NVO_AgencyMaster where ID = companyID) Agency," +
                " (Select top 1 GeneralName from NVO_GeneralMaster where ID = MainAccType) MainAccType,(Select top 1 GeneralName from NVO_GeneralMaster where ID = GLNature) GLNature," +
                " (Select top 1 GeneralName from NVO_GeneralMaster where ID = Category) Category,(Select top 1 GeneralName from NVO_GeneralMaster where ID = GLMatching) GLMatching," +
                " Case when NVO_GLMaster.StatusID = 1 Then 'ACTIVE' WHEN NVO_GLMaster.StatusID = 2 Then 'INACTIVE' End as Status from NVO_GLMaster ";


            if (GLCode != "" && GLCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where GLCode like '%" + GLCode + "%'";
                else
                    strWhere += " and GLCode like '%" + GLCode + "%'";


            if (GLDesc != "" && GLDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where GLDesc like '%" + GLDesc + "%'";
                else
                    strWhere += " and GLDesc like '%" + GLDesc + "%'";

            if (MainAccType != "0" && MainAccType != "null" && MainAccType != "?" && MainAccType != "")
                if (strWhere == "")
                    strWhere += _Query + " Where MainAccType=" + MainAccType;
                else
                    strWhere += " and MainAccType =" + MainAccType;

            if (Category != "0" && Category != "" && Category != "null" && Category != "?")
                if (strWhere == "")
                    strWhere += _Query + " where Category =" + Category;
                else
                    strWhere += " and Category =" + Category;

            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }


        public void CreateChgCodeExcel(string ChgCode, string ChgDesc,string StatusID,string OwnershipID, string User)
        {
            DataTable dtv = GetChgCodeMasterValues(ChgCode, ChgDesc, StatusID, OwnershipID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ChargeMaster");

                ws.Cells["A2"].Value = "Charge Master List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:K2"];
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
                ws.Cells["B7"].Value = "Charge Code";
                ws.Cells["C7"].Value = "Charge Description";
                ws.Cells["D7"].Value = "SACCODE";
                ws.Cells["E7"].Value = "Charge Group";
                ws.Cells["F7"].Value = "Basis";
                ws.Cells["G7"].Value = "Ownership";
                ws.Cells["H7"].Value = "DtValidFrom";
                ws.Cells["I7"].Value = "DtValidTill";
                ws.Cells["J7"].Value = "ROLE";
                ws.Cells["K7"].Value = "Status";
                r = ws.Cells["A7:K7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ChgCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["ChgDesc"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["SacCode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ChargeGroup"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Basis"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Ownership"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["DtValidFrom"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["DtValidTill"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Role"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:K" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ChargeMasterList.xlsx");
                Response.End();

            }

        }

        public DataTable GetChgCodeMasterValues(string ChgCode, string ChgDesc, string StatusID, string OwnershipID)
        {
            string strWhere = "";
            string _Query = " select Chg.ID,Chg.ChgCode,Chg.ChgDesc,SacCode,(select top 1 GeneralName from NVO_generalMaster where ID = ChargeGroupID) as ChargeGroup,( select top 1 GeneralName from NVO_generalMaster where ID = BasisID) as Basis,(select top 1 GeneralName from NVO_generalMaster where ID = ownershipID) as Ownership,convert(varchar, Chg.DtValidFrom, 106) As DtValidFrom, convert(varchar, Chg.DtValidTill, 106) As DtValidTill," +
              " case when IsNVOCC=1 THEN 'NVOCC' WHEN IsForwarding =1 then 'Forwarding' WHEN IsDeposit=1 THEN 'Deposit' ELSE '' End as Role, case when Chg.status = 1 then 'Active' when Chg.status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_ChargeTB Chg  ";
         

            if (ChgCode != "" && ChgCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where Chg.ChgCode like '%" + ChgCode + "%'";
                else
                    strWhere += " and Chg.ChgCode like '%" + ChgCode + "%'";

            if (ChgDesc != "" && ChgDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where Chg.ChgDesc like '%" + ChgDesc + "%'";
                else
                    strWhere += " and Chg.ChgDesc like '%" + ChgDesc + "%'";

            if (StatusID != "0" && StatusID != "null" && StatusID != "?" && StatusID != "")
                if (strWhere == "")
                    strWhere += _Query + " Where Chg.Status=" + StatusID;
                else
                    strWhere += " and Chg.Status =" + StatusID;

            if (OwnershipID != "0" && OwnershipID != "" && OwnershipID != "null" && OwnershipID != "?")
                if (strWhere == "")
                    strWhere += _Query + " where Chg.OwnershipID =" + OwnershipID;
                else
                    strWhere += " and Chg.OwnershipID =" + OwnershipID;

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");
        }

        public void CreateGLMappingExcel(string ChargeCodeV, string ProductTypeID, string ShipmentTypeID, string User)
        {
            DataTable dtv = GetGLMappingValues(ChargeCodeV, ProductTypeID, ShipmentTypeID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("GLMapping");

                ws.Cells["A2"].Value = "GLMapping List";
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
                ws.Cells["B7"].Value = "Charge Code";
                ws.Cells["C7"].Value = "Product Type";
                ws.Cells["D7"].Value = "Shipment Type";
                ws.Cells["E7"].Value = "Principal Revenue GL";
                ws.Cells["F7"].Value = "Principal Expenses GL";
                ws.Cells["G7"].Value = "Agency Revenue GL";
                ws.Cells["H7"].Value = "Agency Expenses GL";
                ws.Cells["I7"].Value = "Valid From";

                r = ws.Cells["A7:I7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ChgCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["ProductTypeV"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ShipmentTypeV"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["PrincipalRevGL"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["PrincipalEXPGL"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["AgencyRevGL"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["AgencyExpGL"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["ValidFromV"].ToString();
                  
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:I" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=GLMappingList.xlsx");
                Response.End();

            }

        }

        public DataTable GetGLMappingValues(string ChargeCodeV, string ProductTypeID, string ShipmentTypeID)
        {
            string strWhere = "";

            string _Query = " select NVO_GLMapping.ID,ProductTypeID,ShipmentTypeID,Case when ProductTypeID=14 then 'SOC' else 'COC' end as ProductTypeV, (select top 1 ChgCode from NVO_ChargeTB where ID = ChargeCodeID ) ChgCode,Case When ShipmentTypeID = 1 then 'EXPORT' else 'IMPORT' end as ShipmentTypeV, ( Select top 1 GLCODE from NVO_GLMaster WHERE ID =PrincipalRevGL ) AS  PrincipalRevGL,(Select top 1 GLCODE from NVO_GLMaster WHERE ID = PrincipalEXPGL ) AS PrincipalEXPGL, ( Select top 1 GLCODE from NVO_GLMaster WHERE ID = AgencyExpGL ) AS AgencyExpGL,(Select top 1 GLCODE from NVO_GLMaster WHERE ID =AgencyRevGL ) AS  AgencyRevGL, replace(convert(NVARCHAR, ValidFrom, 106), ' ', '-') as ValidFromV from NVO_GLMapping ";



            if (ChargeCodeV != "" && ChargeCodeV != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where ChgCode like '%" + ChargeCodeV + "%'";
                else
                    strWhere += " and ChgCode like '%" + ChargeCodeV + "%'";


            if (ProductTypeID != "0" && ProductTypeID != "null" && ProductTypeID != "?" && ProductTypeID != "")
                if (strWhere == "")
                    strWhere += _Query + " Where ProductTypeID=" + ProductTypeID;
                else
                    strWhere += " and ProductTypeID =" + ProductTypeID;

            if (ShipmentTypeID != "0" && ShipmentTypeID != "" && ShipmentTypeID != null && ShipmentTypeID != "?")
                if (strWhere == "")
                    strWhere += _Query + " where ShipmentTypeID =" + ShipmentTypeID;
                else
                    strWhere += " and ShipmentTypeID =" + ShipmentTypeID;


            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }

        public void CreateTaxEngineExcel(string ChargeTBID, string CountryID ,string User)
        {
            DataTable dtv = GetTaxEngineValues(ChargeTBID, CountryID);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TaxEngine");

                ws.Cells["A2"].Value = "TaxEngine List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:E2"];
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
                ws.Cells["B7"].Value = "Shipment Type";
                ws.Cells["C7"].Value = "SACCODE";
                ws.Cells["D7"].Value = "Charge Code";
                ws.Cells["E7"].Value = "Tax Name";
               

                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {
                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["SACCODE"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ChgCode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["TaxName"].ToString();
      
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 24].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TaxEngine List.xlsx");
                Response.End();

            }

        }

        public DataTable GetTaxEngineValues(string ChargeTBID, string CountryID)
        {
            string strWhere = "";

            string _Query = " select TE.ID,( select top 1 GeneralName from NVO_generalMaster where ID=ShipmentTypeID)  ShipmentType,TE.SACCODE,  (select top 1 ChgCode from NVO_ChargeTB where ID = ChargeTBID ) ChgCode,(select top 1 TaxName from NVO_ChgTaxDeclaration where ID = TaxPercentageID ) TaxName From  NVO_ChargeTaxEngineDtls TE ";


            if (ChargeTBID != "0" && ChargeTBID != "null" && ChargeTBID != "?" && ChargeTBID != "")
                if (strWhere == "")
                    strWhere += _Query + " Where TE.ChargeTBID=" + ChargeTBID;
                else
                    strWhere += " and TE.ChargeTBID =" + ChargeTBID;

            if (CountryID != "0" && CountryID != "null" && CountryID != "?" && CountryID != "")
                if (strWhere == "")
                    strWhere += _Query + " Where (select top 1 countryID from NVO_ChgTaxDeclaration where ID = TaxPercentageID )=" + CountryID;
                else
                    strWhere += " and (select top 1 countryID from NVO_ChgTaxDeclaration where ID = TaxPercentageID ) =" + CountryID;

            if (strWhere == "")
                strWhere = _Query;


            if (strWhere == "")
                strWhere = _Query;
            return Manag.GetViewData(strWhere, "");
        }
        
    }
}