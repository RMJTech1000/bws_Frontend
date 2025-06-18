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
    public class PartyExcelController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: PartyExcel
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult AgencyExcel()
        {
            //CreateExcel(CountryID,);
            return View();

        }
        public void CreateAgencyExcel(string CountryID, string CityID, string AgencyCode, string AgencyName, string OwnOffice, string OrganizationType, string User)
        {
            DataTable dtv = GetAgencyValues(CountryID, CityID, AgencyCode, AgencyName, OwnOffice, OrganizationType);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("AgencySearch");

                ws.Cells["A2"].Value = "Agency Search List";
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
                ws.Cells["B7"].Value = "Agency Name";
                ws.Cells["C7"].Value = "Agency Code";
                ws.Cells["D7"].Value = "Country";
                ws.Cells["E7"].Value = "City";
                ws.Cells["F7"].Value = "Own Office";
                ws.Cells["G7"].Value = "Organization Type";
                ws.Cells["H7"].Value = "Address";
                ws.Cells["I7"].Value = "GeoLocation";
                ws.Cells["J7"].Value = "Role";
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
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["agencyname"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["agencycode"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["CountryName"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["cityname"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["OwnOffices"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["OrganizationTypes"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["Address"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["GeoLocation"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Role"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["Status"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:K" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:K" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 12].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=AgencyListSearch.xlsx");
                Response.End();

            }

        }

        public DataTable GetAgencyValues(string CountryID, string CityID, string AgencyCode, string AgencyName, string OwnOffice, string OrganizationType)
        {
            string strWhere = "";

            string _Query = "select am.id,agencycode,agencyname, " +
                "  (select top 1 CountryName from NVO_CountryMaster where ID =CountryID)as CountryName,Address,(select top 1 cityname from NVO_CityMaster where ID =CityID) as cityname ," +
                "  (select top 1 GeoLocation from NVO_GeoLocations where ID =GeoLocationID) as GeoLocation, " +
                " (select top 1 GeneralName from NVO_GeneralMaster where ID =OrganizationType) as OrganizationTypes,Case when ownoffice=1 then 'YES' Else 'No' END AS OwnOffices,Case when Status=1 then 'ACTIVE' Else 'INACTIVE' END AS Status,Case when IsOceanus=1 then 'Oceanus' when Is3rdParty=1 then '3rdParty' when IsBoth=1 then 'Both' Else '' END AS Role " +
                   " from NVO_AgencyMaster am ";

            if (CountryID != "" && CountryID != "null" && CountryID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where am.CountryID=" + CountryID;
                else
                    strWhere += " and am.CountryID =" + CountryID;

            if (CityID != "" && CityID != "null" && CityID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where am.CityID=" + CityID;
                else
                    strWhere += " and am.CityID =" + CityID;

            if (AgencyCode != "" && AgencyCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where agencycode like '%" + AgencyCode + "%'";
                else
                    strWhere += " and agencycode like '%" + AgencyCode + "%'";

            if (AgencyName != "" && AgencyName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where agencyname like '%" + AgencyName + "%'";
                else
                    strWhere += " and agencyname like '%" + AgencyName + "%'";

            if (OwnOffice != "" && OwnOffice != "null" && OwnOffice != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where OwnOffice =" + OwnOffice;
                else
                    strWhere += " and OwnOffice =" + OwnOffice;

            if (OrganizationType != "" && OrganizationType != "null" && OrganizationType != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where OrganizationType =" + OrganizationType;
                else
                    strWhere += " and OrganizationType =" + OrganizationType;


            if (strWhere == "")
                strWhere = _Query + " order by am.ID desc "; ;


            return Manag.GetViewData(strWhere, "");
        }


        public void CreatePartyExcel(string CountryID, string CustomerName, string CustomerType, string User)
        {
            DataTable dtv = GetPartyValues(CountryID, CustomerName, CustomerType);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("PartyList");

                ws.Cells["A2"].Value = "Party Search List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:J2"];
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
                ws.Cells["B7"].Value = "Customer Name";
                ws.Cells["C7"].Value = "Country";
                ws.Cells["D7"].Value = "Party Type";
                ws.Cells["E7"].Value = "City";
                ws.Cells["F7"].Value = "State";
                ws.Cells["G7"].Value = "Address";
                ws.Cells["H7"].Value = "EmailID";
                ws.Cells["I7"].Value = "Tel No";
                ws.Cells["J7"].Value = "Status";

                r = ws.Cells["A7:K7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    ws.Cells["A" + rw].Value = sl;
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Customer"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CountryName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["PartyType"].ToString();

                    //DataTable dtB = GetBRANCHValues(dtv.Rows[i]["ID"].ToString());
                    //for (int k = 0; k < dtB.Rows.Count; k++)
                    //{
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["City"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["State"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Address"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["EmailID"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["TelNo"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Status"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:J" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:J" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=PartyListSearch.xlsx");
                Response.End();

            }

        }

        public DataTable GetPartyValues(string CountryID, string CustomerName, string CustomerType)
        {
            string strWhere = "";

            string _Query = "select  ID,(CustomerName + '-' + B.City) as Customer,(Select top(1) CountryName from NVO_CountryMaster WHERE ID = CountryID)  CountryName,(Select top(1) CityName from NVO_CityMaster WHERE ID = B.CityID)  CityName,(Select top(1) StateName from NVO_StateMaster WHERE ID = B.StateID)  StateName, (Select top(1) GeneralName from NVO_GeneralMaster WHERE ID = CustomerType) PartyType,CM.CustomerType,B.Address,B.State,B.City,B.EmailID,B.TelNo,B.Status from NVO_CustomerMaster CM inner join NVO_CusBranchLocation B on B.CustomerID = CM.ID  ";

            if (CustomerName != "" && CustomerName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CustomerName like '%" + CustomerName + "%'";
                else
                    strWhere += " and CustomerName like '%" + CustomerName + "%'";

            if (CountryID != "" && CountryID != "null" && CountryID != "0" && CountryID != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where CountryID=" + CountryID;
                else
                    strWhere += " and CountryID =" + CountryID;

            if (CustomerType.ToString() != "" && CustomerType.ToString() != "null" && CustomerType.ToString() != "0" && CustomerType.ToString() != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where CustomerType=" + CustomerType.ToString();
                else
                    strWhere += " and CustomerType =" + CustomerType.ToString();


            if (strWhere == "")
                strWhere = _Query + " order by CM.ID desc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetBRANCHValues(string CustomerID)
        {


            string _Query = "select * From NVO_CusBranchLocation where CustomerID=" + CustomerID;

            return Manag.GetViewData(_Query, "");
        }
    }
}