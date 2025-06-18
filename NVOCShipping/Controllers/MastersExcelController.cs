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
    public class MastersExcelController : Controller
    {
        DocumentManager Manag = new DocumentManager();
        // GET: MastersExcel
        public ActionResult Index()
        {
            return View();
        }

        public void CountryExcel(string CountryCode, string CountryName, string User)
        {
            DataTable dtv = GetCountryValues(CountryCode, CountryName);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CountryList");

                ws.Cells["A2"].Value = "Country List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Country Code";
                ws.Cells["C7"].Value = "Country Name";
                ws.Cells["D7"].Value = "Status";

                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CountryCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CountryName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusV"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CountryList.xlsx");
                Response.End();

            }
        }

        public void CurrencyExcel(string CurrencyCode, string CurrencyName, string User)
        {
            DataTable dtv = GetCurrencyValues(CurrencyCode, CurrencyName);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CurrencyList");

                ws.Cells["A2"].Value = "Currency List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S. No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Currency Code";
                ws.Cells["C7"].Value = "Currency Name";
                ws.Cells["D7"].Value = "Country";

                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CurrencyCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CurrencyName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Country"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CurrencyList.xlsx");
                Response.End();

            }
        }

        public void CityExcel(string CityCode, string CityName, string User)
        {
            DataTable dtv = GetCityValues(CityCode, CityName);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CityList");

                ws.Cells["A2"].Value = "City List";
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

                //ws.Cells["A7"].Value = "S. No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "City Code";
                ws.Cells["C7"].Value = "City Name";
                ws.Cells["D7"].Value = "Country Code";
                ws.Cells["E7"].Value = "Status";

                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CityCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CityName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["countryCode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["status"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CityList.xlsx");
                Response.End();

            }
        }

        public void CommodityExcel(string CmdName, string HsCode, string CmdType, string User)
        {
            DataTable dtv = GetCommodityValues(CmdName, HsCode, CmdType);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CommodityList");

                ws.Cells["A2"].Value = "Commodity List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:F2"];
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

                //ws.Cells["A7"].Value = "S. No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Commodity Code";
                ws.Cells["C7"].Value = "Commodity Name";
                ws.Cells["D7"].Value = "HS Code";
                ws.Cells["E7"].Value = "DG Flag";
                ws.Cells["F7"].Value = "Commodity Type";

                r = ws.Cells["A7:F7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CommodityUnCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CommodityName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["HScode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DangerousFlag"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:F" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CommodityList.xlsx");
                Response.End();

            }
        }

        public void CargoPackageExcel(string PkgCode, string PkgDesc, string User)
        {
            DataTable dtv = GetCargoPackageValues(PkgCode, PkgDesc);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CargoPackageList");

                ws.Cells["A2"].Value = "Cargo Package List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S. No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Package Code";
                ws.Cells["C7"].Value = "Package Description";
                ws.Cells["D7"].Value = "Status";

                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["PkgCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["PkgDescription"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();
                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CargoPackageList.xlsx");
                Response.End();

            }
        }

        public void DepotExcel(string DepotName, string Country, string City, string Status, string User)
        {
            DataTable dtv = GetDepotValues(DepotName, Country, City, Status);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("DepotList");

                ws.Cells["A2"].Value = "Depot List";
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

                //ws.Cells["A7"].Value = "S. No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Depot Name";
                ws.Cells["C7"].Value = "Country";
                ws.Cells["D7"].Value = "City";
                ws.Cells["E7"].Value = "Status";

                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["DepName"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CountryV"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["CityV"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["StatusV"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=DepotList.xlsx");
                Response.End();

            }
        }

        public void PortExcel(string PortCode, string PortName, string Status, string User)
        {
            DataTable dtv = GetPortValues(PortCode, PortName, Status);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();
                var ws = pck.Workbook.Worksheets.Add("PortList");

                ws.Cells["A2"].Value = "Port List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:H2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Port Code";
                ws.Cells["C7"].Value = "Port Name";
                ws.Cells["D7"].Value = "Country Code";
                ws.Cells["E7"].Value = "MainPort";
                ws.Cells["F7"].Value = "Status";
                ws.Cells["G7"].Value = "Sea Port";
                ws.Cells["H7"].Value = "ICD Port";

                r = ws.Cells["A7:H7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["PortCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["PortName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["countryCode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["MainPort"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["SeaPort"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["AirPort"].ToString();
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:H" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=PortList.xlsx");
                Response.End();

            }
        }

        public void TerminalExcel(string TerminalCode, string TerminalName, string User)
        {
            DataTable dtv = GetTerminalValues(TerminalCode, TerminalName);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TerminalMasterList");

                ws.Cells["A2"].Value = "Terminal Master List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Terminal Code";
                ws.Cells["C7"].Value = "Terminal Name";
                ws.Cells["D7"].Value = "Port";
                ws.Cells["E7"].Value = "Status";

                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["TerminalCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["TerminalName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["PortV"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["StatusV"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TerminalMasterList.xlsx");
                Response.End();

            }
        }

        public void GeoLocationExcel(string GeoLoc, string Country, string User)
        {
            DataTable dtv = GetGeoLocationValues(GeoLoc, Country);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("GeoLocationMasterList");

                ws.Cells["A2"].Value = "GeoLocation Master List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Geo Location";
                ws.Cells["C7"].Value = "Country";
                ws.Cells["D7"].Value = "Status";

                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["GeoLocation"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CountryName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=GeoLocationMasterList.xlsx");
                Response.End();

            }
        }

        public void MainPortExcel(string PortCode, string PortName, string Status, string User)
        {
            DataTable dtv = GetMainPortValues(PortCode, PortName, Status);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MainPortList");

                ws.Cells["A2"].Value = "MainPort Master List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Port Code";
                ws.Cells["C7"].Value = "Port Name";
                ws.Cells["D7"].Value = "Country Code";
                ws.Cells["E7"].Value = "Status";

                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["PortCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["PortName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["countryCode"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MainPortMasterList.xlsx");
                Response.End();

            }
        }

        public void MISExchangeRateExcel(string EffDate, string User)
        {
            DataTable dtv = GetMISExchangeRateValues(EffDate);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MISExchangeRateList");

                ws.Cells["A2"].Value = "MIS Exchange Rate Master List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Effective Date";


                r = ws.Cells["A7:B7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["EffectiveDate"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:B" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:B" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:B" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:B" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MISExchangeRateMasterList.xlsx");
                Response.End();

            }
        }

        public void ExchangeRateExcel(string Date, string User)
        {
            DataTable dtv = GetExchangeRateValues(Date);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ExchangeRateList");

                ws.Cells["A2"].Value = "Exchange Rate Master List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Date";
                ws.Cells["C7"].Value = "Base Currency";
                ws.Cells["D7"].Value = "Conversion Currency";
                ws.Cells["E7"].Value = "Rate";


                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Date"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["FromCurrency"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ToCurrency"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Rate"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ExchangeRateMasterList.xlsx");
                Response.End();

            }
        }

        public void MRGExcel(string POL, string POD, string Commodity, string CntrTypes, string User)
        {
            DataTable dtv = GetMRGValues(POL, POD, Commodity, CntrTypes);
            if (dtv.Rows.Count > 0)
            {

                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MRGList");

                ws.Cells["A2"].Value = "MRG List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:H2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "MRG No";
                ws.Cells["C7"].Value = "POL";
                ws.Cells["D7"].Value = "POD";
                ws.Cells["E7"].Value = "Container Type";
                ws.Cells["F7"].Value = "Commodity Type";
                ws.Cells["G7"].Value = "Freight Rate";
                ws.Cells["H7"].Value = "Service Type";


                r = ws.Cells["A7:H7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["MRGRate"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["POO"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Destination"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Cntrsize"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["Commodity"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Amount"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ServiceTypes"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:H" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MRGList.xlsx");
                Response.End();

            }
        }

        public void SlotExcel(string SlotOperator, string SlotRef, string SlotTerms, string POL, string POD, string User)
        {
            DataTable dtv = GetSlotValues(SlotOperator, SlotRef, SlotTerms, POL, POD);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("SlotList");


                //Record Headers

                ws.Cells["A1"].Value = "ID";
                ws.Cells["B1"].Value = "SLOTCONTRACTREF";
                ws.Cells["C1"].Value = "EFFECTIVEDATE";
                //ws.Cells["D1"].Value = "VALID TO";
                ws.Cells["D1"].Value = "SERVICE NAME";
                ws.Cells["E1"].Value = "SLOT OPERATOR";
                ws.Cells["F1"].Value = "SLOT TERM";
                ws.Cells["G1"].Value = "RESPONSIBLE AGENCY";
                ws.Cells["H1"].Value = "POL";
                ws.Cells["I1"].Value = "POD";
                ws.Cells["J1"].Value = "CHARGE";
                ws.Cells["K1"].Value = "BASIS";
                ws.Cells["L1"].Value = "SIZE TYPE";
                ws.Cells["M1"].Value = "COMMODITY";
                ws.Cells["N1"].Value = "CURRENCY";
                ws.Cells["O1"].Value = "AMOUNT";
                ws.Cells["P1"].Value = "ROUTETYPE";


                ExcelRange r = ws.Cells["A1:P1"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                int sl = 1;

                int rw = 2;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["SlotRef"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["FDate"].ToString();
                    //ws.Cells["D" + rw].Value = dtv.Rows[i]["TDate"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ServiceName"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["SlotOperatorv"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["SlotTerm"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["Charge"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["Basics"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["Commodity"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["Amount"];
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["RoutType"];
                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A1:P" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:P" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:P" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:P" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 12].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=SlotList.xlsx");
                Response.End();

            }
        }

        public void TariffExcel(string Loc, string Shipment, string TariffTypes, string Status, string ServiceType, string DGClass, string CollMode, string User)
        {
            DataTable dtv = GetTariffValues(Loc, Shipment, TariffTypes, Status, ServiceType, DGClass, CollMode);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TariffList");

                //Record Headers

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A1"].Value = "ID";
                ws.Cells["B1"].Value = "Location";
                ws.Cells["C1"].Value = "Tariff Type";
                ws.Cells["D1"].Value = "ShipmentType";
                ws.Cells["E1"].Value = "MODULE";
                ws.Cells["F1"].Value = "CommodityType";
                ws.Cells["G1"].Value = "TariffMode";
                ws.Cells["H1"].Value = "ServiceType";
                ws.Cells["I1"].Value = "CollectionMode";
                ws.Cells["J1"].Value = "ChargeType";
                ws.Cells["K1"].Value = "ChargeCode";
                ws.Cells["L1"].Value = "Basis";
                ws.Cells["M1"].Value = "Cntr";
                ws.Cells["N1"].Value = "Currency";
                ws.Cells["O1"].Value = "Amount";
                ws.Cells["P1"].Value = "Dest.Port";
                //ws.Cells["Q1"].Value = "HandlingCharge";
                //ws.Cells["R1"].Value = "ChargeCodeBreakup";
                //ws.Cells["S1"].Value = "ShipmentBreakup";
                //ws.Cells["T1"].Value = "CommodityBreakup";
                //ws.Cells["U1"].Value = "CntrTypeBreakup";
                //ws.Cells["V1"].Value = "BreakupAmount";

                ExcelRange r = ws.Cells["A1:P1"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                int sl = 1;

                int rw = 2;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["Location"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["TariffType"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["MODULE"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["TariffMode"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ServiceType"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["CollectionMode"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["ChargeType"].ToString();
                    ws.Cells["K" + rw].Value = dtv.Rows[i]["ChargeCode"].ToString();
                    ws.Cells["L" + rw].Value = dtv.Rows[i]["Basis"].ToString();
                    ws.Cells["M" + rw].Value = dtv.Rows[i]["Cntr"].ToString();
                    ws.Cells["N" + rw].Value = dtv.Rows[i]["Currency"].ToString();
                    ws.Cells["O" + rw].Value = dtv.Rows[i]["TarifChgAmount"];
                    ws.Cells["P" + rw].Value = dtv.Rows[i]["DestPort"].ToString();
                    //ws.Cells["Q" + rw].Value = dtv.Rows[i]["HandlingCharge"].ToString();
                    //ws.Cells["R" + rw].Value = dtv.Rows[i]["ChargeCodeBreakup"].ToString();
                    //ws.Cells["S" + rw].Value = dtv.Rows[i]["ShipmentBreakup"].ToString();
                    //ws.Cells["T" + rw].Value = dtv.Rows[i]["CommodityBreakup"].ToString();
                    //ws.Cells["U" + rw].Value = dtv.Rows[i]["CntrTypeBreakup"].ToString();
                    //ws.Cells["V" + rw].Value = dtv.Rows[i]["BreakupAmount"].ToString();
                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A1:P" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:P" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:P" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:P" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 25].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TariffList.xlsx");
                Response.End();

            }
        }

        public void WaiverRequestExcel(string ReqType, string BookingNo, string FreeDays, string WaiverAmt, string User)
        {
            DataTable dtv = GetWaiverValues(ReqType, BookingNo, FreeDays, WaiverAmt);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("WaiverRequestList");

                ws.Cells["A2"].Value = "Waiver Request List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Request Type";
                ws.Cells["C7"].Value = "Charge Type";
                ws.Cells["D7"].Value = "Transaction Type";
                ws.Cells["E7"].Value = "Booking No";
                ws.Cells["F7"].Value = "Container No";
                ws.Cells["G7"].Value = "Free Days";
                ws.Cells["H7"].Value = "Waiver Amount";
                ws.Cells["I7"].Value = "Status";


                r = ws.Cells["A7:I7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["RequestTypes"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["ChargeType"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["TransType"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["BookingNo"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["CntrNo"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Freedays"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["WaiverAmt"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["Status"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:I" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:I" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=WaiverRequestList.xlsx");
                Response.End();

            }
        }

        public void ContainerRentExcel(string ShipmentType, string TariffType, string ContainerType, string User)
        {
            DataTable dtv = GetContainerContractValues(ShipmentType, TariffType, ContainerType);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("WaiverRequestList");

                ws.Cells["A2"].Value = "Container Rental Contract List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Charge Type";
                ws.Cells["C7"].Value = "Charges";
                ws.Cells["D7"].Value = "Charge Owner";
                ws.Cells["E7"].Value = "Shipment Type";
                ws.Cells["F7"].Value = "Tariff Type";
                ws.Cells["G7"].Value = "Size";
                ws.Cells["H7"].Value = "Valid From";
                ws.Cells["I7"].Value = "Valid Till";
                ws.Cells["J7"].Value = "Status";


                r = ws.Cells["A7:J7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ChargeType"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Charges"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["ChargeOwnerIDV"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ShipmentTypeV"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["TariffTypeV"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["Size"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["ValidFrom"].ToString();
                    ws.Cells["I" + rw].Value = dtv.Rows[i]["ValidTill"].ToString();
                    ws.Cells["J" + rw].Value = dtv.Rows[i]["StatusV"].ToString();

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
                Response.AddHeader("content-disposition", "attachment;  filename=ContainerRentalContractList.xlsx");
                Response.End();

            }
        }

        public void CommissionContractExcel(string ShipmentType, string Agency, string CommissionCharge, string Status, string User)
        {
            DataTable dtv = GetCommissionContractValues(ShipmentType, Agency, CommissionCharge, Status);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CommissionContractList");

                ws.Cells["A2"].Value = "Commission Contract List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Shipment Type";
                ws.Cells["C7"].Value = "Agency";
                ws.Cells["D7"].Value = "Commission Charge";
                ws.Cells["E7"].Value = "Valid From";
                ws.Cells["F7"].Value = "Valid Till";
                ws.Cells["G7"].Value = "Status";


                r = ws.Cells["A7:G7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ShipmentType"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Agency"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["CommissionCharge"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["ValidFrom"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["ValidTill"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["STATUSResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:G" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:G" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:G" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:G" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CommissionContractList.xlsx");
                Response.End();

            }
        }

        public void VoyageLockingExcel(string VesVoy, string User)
        {
            DataTable dtv = GetVoyageLockingValues(VesVoy);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("VoyageLockingList");

                ws.Cells["A2"].Value = "Voyage Locking List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Vessel Voyage";
                ws.Cells["C7"].Value = "Current Port";
                ws.Cells["D7"].Value = "Next Port";
                ws.Cells["E7"].Value = "Status";


                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["VoyageID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["CurrentPort"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["NextPort"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Status"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=VoyageLockingList.xlsx");
                Response.End();

            }
        }

        public void VoyageOpeningExcel(string Agency, string User)
        {
            DataTable dtv = GetVoyageOpeningValues(Agency);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("VoyageOpeningList");

                ws.Cells["A2"].Value = "Voyage Opening List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:F2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Vessel Voyage";
                ws.Cells["C7"].Value = "Discharge Voyage No";
                ws.Cells["D7"].Value = "ETA";
                ws.Cells["E7"].Value = "Discharge Terminal";
                ws.Cells["F7"].Value = "Status";


                r = ws.Cells["A7:F7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["VesVoy"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["ETA"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["DisVoyageNo"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DiscTerminal"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["OpenStatus"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:F" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=VoyageOpeningList.xlsx");
                Response.End();

            }
        }

        public void ImportExcel(string BookingNo, string POL, string POD, string VesVoy, string Agency, string User)
        {
            DataTable dtv = GetImportValues(BookingNo, POL, POD, VesVoy, Agency);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ImportList");

                ws.Cells["A2"].Value = "Import List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:H2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Booking No";
                ws.Cells["C7"].Value = "BL Type";
                ws.Cells["D7"].Value = "Mode";
                ws.Cells["E7"].Value = "POL";
                ws.Cells["F7"].Value = "POD";
                ws.Cells["G7"].Value = "VesVoy";
                ws.Cells["H7"].Value = "Commodity Type";


                r = ws.Cells["A7:H7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["BLNumber"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["BLType"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["BlDirect"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["POL"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["POD"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["BLVesVoy"].ToString();
                    ws.Cells["H" + rw].Value = dtv.Rows[i]["CommodityType"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:H" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:H" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ImportList.xlsx");
                Response.End();

            }
        }

        public void BankMasterExcel(string BankName, string AccNo, string User)
        {
            DataTable dtv = GetBankMasterValues(BankName, AccNo);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("BankMasterList");

                ws.Cells["A2"].Value = "Bank Master List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Bank Code";
                ws.Cells["C7"].Value = "Bank Name";
                ws.Cells["D7"].Value = "Account No";
                ws.Cells["E7"].Value = "Status";


                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["BankCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["BankName"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["AccountNo"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Status"].ToString();


                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=BankMasterList.xlsx");
                Response.End();

            }
        }

        public void CreditControlExcel(string Party, string PartyType, string ExemTax, string User)
        {
            DataTable dtv = GetCreditControlValues(Party, PartyType, ExemTax);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("CreditControlList");

                ws.Cells["A2"].Value = "Credit Control List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Party";
                ws.Cells["C7"].Value = "Party Type";
                ws.Cells["D7"].Value = "Exemp From Tax";


                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["CustomerName"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["PartyType"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["TaxExempt"].ToString();


                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=CreditControlList.xlsx");
                Response.End();

            }
        }

        public void TaxDeclrationExcel(string TaxName, string Status, string User)
        {
            DataTable dtv = GetTaxDeclrationValues(TaxName, Status);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TaxDeclrationList");

                ws.Cells["A2"].Value = "Tax Declration List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Tax Name";
                ws.Cells["C7"].Value = "Tax Percentage";
                ws.Cells["D7"].Value = "Country";
                ws.Cells["E7"].Value = "Status";


                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["TaxName"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["TaxPercentage"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["CountryName"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["StatusV"].ToString();


                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TaxDeclrationList.xlsx");
                Response.End();

            }
        }

        public void TaxNamesExcel(string TaxCode, string User)
        {
            DataTable dtv = GetTaxNamesValues(TaxCode);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("TaxNamesList");

                ws.Cells["A2"].Value = "Tax Names List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Tax Code";
                ws.Cells["C7"].Value = "Tax Description";
                ws.Cells["D7"].Value = "Country";


                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["TaxCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["TaxDescription"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Country"].ToString();


                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=TaxNamesList.xlsx");
                Response.End();

            }
        }

        public void ReceiptsExcel(string RecpNo, string RecParty, string Date, string RecType, string User)
        {
            DataTable dtv = GetReceiptsValues(RecpNo, RecParty, Date, RecType);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ReceiptsList");

                ws.Cells["A2"].Value = "Receipts List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:F2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Receipt No";
                ws.Cells["C7"].Value = "Receipt Date";
                ws.Cells["D7"].Value = "Party Name";
                ws.Cells["E7"].Value = "Receipt Amount";
                ws.Cells["F7"].Value = "Payment Type";


                r = ws.Cells["A7:F7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ReceiptNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["RecDate"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["LocalAmount"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["PaymentType"].ToString();


                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:F" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ReceiptsList.xlsx");
                Response.End();

            }
        }

        public void ReceiptCancelExcel(string RecpNo, string RecParty, string Date, string RecType, string Agency, string User)
        {
            DataTable dtv = GetReceiptCancelValues(RecpNo, RecParty, Date, RecType, Agency);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("ReceiptsCancelList");

                ws.Cells["A2"].Value = "Receipts Cancel List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:H2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Receipt Voucher";
                ws.Cells["C7"].Value = "Receipt Voucher Date";
                ws.Cells["D7"].Value = "Receipt Type";
                ws.Cells["E7"].Value = "Party Name";
                ws.Cells["F7"].Value = "Amount";
                ws.Cells["G7"].Value = "Cancelled On";
                ws.Cells["H7"].Value = "Cancelled By";


                r = ws.Cells["A7:H7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ReceiptNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["RecDate"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["PaymentType"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["PartyName"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["Amount"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["CancelledOn"].ToString();
                    ws.Cells["G" + rw].Value = dtv.Rows[i]["CancelledBy"].ToString();


                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:F" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ReceiptCancelList.xlsx");
                Response.End();

            }
        }

        public void MNRDamageMasterExcel(string DmgCode, string DmgDesc, string User)
        {
            DataTable dtv = GetMNRDamageMasterValues(DmgCode, DmgDesc);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MNRDamageMasterList");

                ws.Cells["A2"].Value = "MNR Damage Master List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Damage Code";
                ws.Cells["C7"].Value = "Damage Description";
                ws.Cells["D7"].Value = "Status";


                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["DamageCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["DamageDescription"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MNRDamageMasterList.xlsx");
                Response.End();

            }
        }

        public void MNRRepairExcel(string RprCode, string RprDesc, string User)
        {
            DataTable dtv = GetMNRRepairValues(RprCode, RprDesc);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MNRRepairList");

                ws.Cells["A2"].Value = "MNR Repair List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Repair Code";
                ws.Cells["C7"].Value = "Repair Description";
                ws.Cells["D7"].Value = "Status";


                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["RepairCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["RepairDescription"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MNRDamageMasterList.xlsx");
                Response.End();

            }
        }

        public void MNRComponentExcel(string CompCode, string CompDesc, string Assembly, string User)
        {
            DataTable dtv = GetMNRComponentValues(CompCode, CompDesc, Assembly);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MNRComponentList");

                ws.Cells["A2"].Value = "MNR Component List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Component Code";
                ws.Cells["C7"].Value = "Component Description";
                ws.Cells["D7"].Value = "Assembly";
                ws.Cells["E7"].Value = "Status";


                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ComponentCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["ComponentDescription"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["Assembly"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MNRComponentList.xlsx");
                Response.End();

            }
        }

        public void MNRLocationExcel(string LocCode, string LocDesc, string User)
        {
            DataTable dtv = GetMNRLocationValues(LocCode, LocDesc);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("MNRLocationList");

                ws.Cells["A2"].Value = "MNR Location List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:D2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Location Code";
                ws.Cells["C7"].Value = "Location Description";
                ws.Cells["D7"].Value = "Status";


                r = ws.Cells["A7:D7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["LocationCode"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["Description"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["StatusResult"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:D" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:D" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=MNRLocationList.xlsx");
                Response.End();

            }
        }

        public void LeaseContractExcel(string LeaseRef, string LeasePartner, string Status, string User)
        {
            DataTable dtv = GetLeaseContractValues(LeaseRef, LeasePartner, Status);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("LeaseContractList");

                ws.Cells["A2"].Value = "Lease Contract List";
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelRange r = ws.Cells["A2:F2"];
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "Leasing Reference";
                ws.Cells["C7"].Value = "Leasing Partner";
                ws.Cells["D7"].Value = "Pick Up Criteria";
                ws.Cells["E7"].Value = "Valid From";
                ws.Cells["F7"].Value = "Valid Till";


                r = ws.Cells["A7:F7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["ID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["ContractRefNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["LeasingPartner"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["PickupCriteria"].ToString();
                    ws.Cells["E" + rw].Value = dtv.Rows[i]["DtContractFrom"].ToString();
                    ws.Cells["F" + rw].Value = dtv.Rows[i]["DtContractTill"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:F" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:F" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=LeaseContractList.xlsx");
                Response.End();

            }
        }

        public void OnHireRequestExcel(string PickupRef, string User)
        {
            DataTable dtv = GetOnHireValues(PickupRef);
            if (dtv.Rows.Count > 0)
            {
                ExcelPackage pck = new ExcelPackage();

                var ws = pck.Workbook.Worksheets.Add("OnHireRequestList");

                ws.Cells["A2"].Value = "On Hire Request List";
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

                //ws.Cells["A7"].Value = "S.No.";
                ws.Cells["A7"].Value = "ID";
                ws.Cells["B7"].Value = "OnHire Request No";
                ws.Cells["C7"].Value = "Lease Reference No";
                ws.Cells["D7"].Value = "Leasing Partner";
                ws.Cells["E7"].Value = "Leasing Term";


                r = ws.Cells["A7:E7"];
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                int sl = 1;

                int rw = 8;

                for (int i = 0; i < dtv.Rows.Count; i++)
                {

                    //ws.Cells["A" + rw].Value = sl;
                    ws.Cells["A" + rw].Value = dtv.Rows[i]["LeaseContractID"].ToString();
                    ws.Cells["B" + rw].Value = dtv.Rows[i]["RequestNo"].ToString();
                    ws.Cells["C" + rw].Value = dtv.Rows[i]["LeasingPartner"].ToString();
                    ws.Cells["D" + rw].Value = dtv.Rows[i]["LeaseTerm"].ToString();

                    // }
                    sl++;
                    rw += 1;
                }

                rw -= 1;

                ws.Cells["A7:E" + rw].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A7:E" + rw].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[1, 1, rw, 11].AutoFitColumns();

                pck.SaveAs(Response.OutputStream);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=OnHireRequestList.xlsx");
                Response.End();

            }
        }


        //Functions
        public DataTable GetCountryValues(string CountryCode, string CountryName)
        {
            string strWhere = "";
            string _Query = " select case when Status=1 then 'Active' else 'Inactive' end as StatusV, * from NVO_CountryMaster";

            if (CountryCode != "" && CountryCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CountryCode like '%" + CountryCode + "%'";
                else
                    strWhere += " and CountryCode like '%" + CountryCode + "%'";

            if (CountryName != "" && CountryName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CountryName like '%" + CountryName + "%'";
                else
                    strWhere += " and CountryName like '%" + CountryName + "%'";

            if (strWhere == "")
                strWhere = _Query + " order by ID ASC ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetCurrencyValues(string CurrencyCode, string CurrencyName)
        {
            string strWhere = "";
            string _Query = " select isnull((select top 1 CountryName from NVO_CountryMaster Where ID=CountryID),'') AS Country,* from NVO_CurrencyMaster";

            if (CurrencyCode != "" && CurrencyCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CurrencyCode like '%" + CurrencyCode + "%'";
                else
                    strWhere += " and CurrencyCode like '%" + CurrencyCode + "%'";

            if (CurrencyName != "" && CurrencyName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CurrencyName like '%" + CurrencyName + "%'";
                else
                    strWhere += " and CurrencyName like '%" + CurrencyName + "%'";

            if (strWhere == "")
                strWhere = _Query + " order by ID ASC ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetCityValues(string CityCode, string CityName)
        {
            string strWhere = "";
            string _Query = " select CTM.id,CTM.CityCode,CTM.CityName,cm.countryCode,CTM.status,CTM.Statename,cm.ID as countryID, " +
                " case when CTM.status=1 then 'Active' when CTM.status=0 then 'Inactive' ELSE '' END as StatusResult " +
                " from NVO_CityMaster CTM " +
                " inner join NVO_CountryMaster cm on cm.ID= CTM.countryid";

            if (CityCode != "" && CityCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CityCode like '%" + CityCode + "%'";
                else
                    strWhere += " and CityCode like '%" + CityCode + "%'";

            if (CityName != "" && CityName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CityName like '%" + CityName + "%'";
                else
                    strWhere += " and CityName like '%" + CityName + "%'";

            if (strWhere == "")
                strWhere = _Query + " order by ID ASC ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetCommodityValues(string CmdName, string HsCode, string CmdType)
        {
            string strWhere = "";
            string _Query = " select CM.ID,CM.CommodityUnCode,GM.GeneralName,CM.CommodityName,CM.HScode,CM.Remarks,GM.GeneralName as CommodityType," +
                     " case when CM.Dangerousflag =1 then 'Yes' when  CM.Dangerousflag =0 then 'No' else '' end DangerousFlag  " +
                     //" case when CommodityType =1 then 'General' when  CommodityType =2 then 'Hazardous' else '' end CommodityType " +
                     " from NVO_CommodityMaster CM" +
                     " LEFT OUTER JOIN NVO_generalMaster GM  ON GM.id=CM.CommodityType AND GM.SeqNo=2 ";

            if (CmdName != "" && CmdName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CM.CommodityName like '%" + CmdName + "%'";
                else
                    strWhere += " and CM.CommodityName like '%" + CmdName + "%'";

            if (HsCode != "" && HsCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CM.HScode like '%" + HsCode + "%'";
                else
                    strWhere += " and CM.HScode like '%" + HsCode + "%'";

            if (CmdType != "" && CmdType != "null" && CmdType != "0" && CmdType != "?")
                if (strWhere == "")
                    strWhere += _Query + " Where GM.GeneralName=" + CmdType;
                else
                    strWhere += " and GM.GeneralName =" + CmdType;

            if (strWhere == "")
                strWhere = _Query + " order by ID ASC ";


            return Manag.GetViewData(strWhere, "");
        }
        public DataTable GetCargoPackageValues(string PkgCode, string PkgDesc)
        {
            string strWhere = "";
            string _Query = " select Id,PkgCode,PkgDescription,status, " +
                " case when status = 1 then 'Active' when status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_CargoPkgMaster ";

            if (PkgCode != "" && PkgCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PkgCode like '%" + PkgCode + "%'";
                else
                    strWhere += " and PkgCode like '%" + PkgCode + "%'";

            if (PkgDesc != "" && PkgDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PkgDescription like '%" + PkgDesc + "%'";
                else
                    strWhere += " and PkgDescription like '%" + PkgDesc + "%'";


            if (strWhere == "")
                strWhere = _Query + " order by ID ASC ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetDepotValues(string DepotName, string Country, string City, string Status)
        {
            string strWhere = "";
            string _Query = " Select ID,DepName,(select CountryName from NVO_CountryMaster where ID= DepCountry) as CountryV, " +
                            " (select CityName from NVO_CityMaster where ID = DepCity) as CityV, " +
                            " Case when Status =1 then 'Active' else 'InActive' end as StatusV " +
                            " from NVO_DepotMaster";

            if (DepotName != "" && DepotName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where DepName like '%" + DepotName + "%'";
                else
                    strWhere += " and DepName like '%" + DepotName + "%'";

            if (Country != "" && Country != "null" && Country != "0" && Country != "?")
                if (strWhere == "")
                    strWhere += _Query + " where DepCountry=" + Country;
                else
                    strWhere += " and DepCountry=" + Country;

            if (City != "" && City != "null" && City != "0" && City != "?")
                if (strWhere == "")
                    strWhere += _Query + " where DepCity=" + City;
                else
                    strWhere += " and DepCity=" + City;

            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " where Status=" + Status;
                else
                    strWhere += " and Status=" + Status;


            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetPortValues(string PortCode, string PortName, string Status)
        {
            string strWhere = "";
            string _Query = "select PTM.id,PTM.PortCode,PTM.PortName,cm.countryCode,PTM.status,cm.id as countryid,case when PTM.IsICDPort = 1 then 'ICDPort' when PTM.IsICDPort = 0 then '' ELSE '' END as AirPort, case when PTM.IsSeaPort = 1 then 'SeaPort' when PTM.IsSeaPort = 0 then '' ELSE '' END as SeaPort, case when PTM.status = 1 then 'Active' when PTM.status = 0 then 'Inactive' ELSE '' END as StatusResult,(SELECT Top 1 PortName from NVO_PortMainMaster Where ID=MainPortID) As MainPort  from NVO_PortMaster PTM inner join NVO_CountryMaster cm on cm.ID = PTM.countryid";

            if (PortCode != "" && PortCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PTM.PortCode like '%" + PortCode + "%'";
                else
                    strWhere += " and PTM.PortCode like '%" + PortCode + "%'";

            if (PortName != "" && PortName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PTM.PortName like '%" + PortName + "%'";
                else
                    strWhere += " and PTM.PortName like '%" + PortName + "%'";

            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PTM.status=" + Status;
                else
                    strWhere += " and PTM.status=" + Status;


            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetTerminalValues(string TerminalCode, string TerminalName)
        {
            string strWhere = "";
            string _Query = " select ID, TerminalCode, TerminalName,(select PortName from NVO_PortMaster where ID = PortID)as PortV, " +
                            " Case when Status = 0 then 'Inactive' else 'Active' end as StatusV from NVO_TerminalMaster";

            if (TerminalCode != "" && TerminalCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where TerminalCode like '%" + TerminalCode + "%'";
                else
                    strWhere += " and TerminalCode like '%" + TerminalCode + "%'";

            if (TerminalName != "" && TerminalName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where TerminalName like '%" + TerminalName + "%'";
                else
                    strWhere += " and TerminalName like '%" + TerminalName + "%'";

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetGeoLocationValues(string GeoLoc, string Country)
        {
            string strWhere = "";
            string _Query = " select G.ID,G.GeoLocation,CM.CountryName,case when G.Status = 1 then 'Active' when G.Status = 0 then 'Inactive' ELSE '' END as StatusResult from NVO_GeoLocations G Inner Join NVO_CountryMaster CM ON CM.ID =G.CountryID ";

            if (GeoLoc != "" && GeoLoc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where G.GeoLocation like '%" + GeoLoc + "%'";
                else
                    strWhere += " and G.GeoLocation like '%" + GeoLoc + "%'";

            if (Country != "" && Country != "null" && Country != "0" && Country != "?")
                if (strWhere == "")
                    strWhere += _Query + " where G.CountryID=" + Country;
                else
                    strWhere += " and G.CountryID=" + Country;

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMainPortValues(string PortCode, string PortName, string Status)
        {
            string strWhere = "";
            string _Query = "select PTM.id,PTM.PortCode,PTM.PortName,cm.countryCode,PTM.status,cm.id as countryid,case when PTM.IsICDPort = 1 then 'ICDPort' when PTM.IsICDPort = 0 then '' ELSE '' END as AirPort, case when PTM.IsSeaPort = 1 then 'SeaPort' when PTM.IsSeaPort = 0 then '' ELSE '' END as SeaPort, case when PTM.status = 1 then 'Active' when PTM.status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_PortMainMaster PTM inner join NVO_CountryMaster cm on cm.ID = PTM.countryid ";

            if (PortCode != "" && PortCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PTM.PortCode like '%" + PortCode + "%'";
                else
                    strWhere += " and PTM.PortCode like '%" + PortCode + "%'";

            if (PortName != "" && PortName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where PTM.PortName like '%" + PortName + "%'";
                else
                    strWhere += " and PTM.PortName like '%" + PortName + "%'";

            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PTM.status=" + Status;
                else
                    strWhere += " and PTM.status=" + Status;

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMISExchangeRateValues(string EffDate)
        {
            string strWhere = "";
            string _Query = " select ID, convert(varchar, EffectiveDate, 103) as EffectiveDate from NVO_MISExchangeRate";

            if (EffDate != "" && EffDate != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where EffectiveDate <= ' " + EffDate + "'";
                else
                    strWhere += " and EffectiveDate <= ' " + EffDate + "'";


            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetExchangeRateValues(string Date)
        {
            string strWhere = "";
            string _Query = "select Ex.ID,convert(varchar, Ex.Date, 106) As Date ,cm1.CurrencyCode AS FromCurrency,Cm2.CurrencyCode AS ToCurrency,Ex.Rate from NVO_ExRate Ex " +
                " inner join NVO_CurrencyMaster cm1 on cm1.id =Ex.FromCurrency " +
                " inner join NVO_CurrencyMaster cm2 on cm2.id =Ex.ToCurrency";

            if (Date != "" && Date != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where Ex.Date <= ' " + Date + "'";
                else
                    strWhere += " and Ex.Date <= ' " + Date + "'";


            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMRGValues(string POL, string POD, string Commodity, string CntrTypes)
        {
            string strWhere = "";
            string _Query = " SELECT ID,MRGRate, convert(varchar,Fdate, 103) as FDate, convert(varchar,Tdate, 103) as TDate, (select top(1) PortName from NVO_PortMaster where Id = PofLoading) POO, " +
                            " (select  top(1) GeneralName from NVO_GeneralMaster where Id = Commodity) as Commodity,(select top(1) CurrencyCode from NVO_CurrencyMaster where Id = Currency) as Currency, " +
                            " (select top(1) size from NVO_tblCntrTypes where ID = CntrTypes) as Cntrsize,(select top(1) PortName from NVO_PortMaster where Id = PofDischarge) Destination, " +
                            " (select top(1) Description from NVO_tblDLValues where Id = ServiceTypes) as ServiceTypes,Amount FROM NVO_MRGRate";

            if (POL != "" && POL != "null" && POL != "0" && POL != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PofLoading=" + POL;
                else
                    strWhere += " and PofLoading=" + POL;

            if (POD != "" && POD != "null" && POD != "0" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PofDischarge=" + POD;
                else
                    strWhere += " and PofDischarge=" + POD;

            if (Commodity != "" && Commodity != "null" && Commodity != "0" && Commodity != "?")
                if (strWhere == "")
                    strWhere += _Query + " where Commodity=" + Commodity;
                else
                    strWhere += " and Commodity=" + Commodity;

            if (CntrTypes != "" && CntrTypes != "null" && CntrTypes != "0" && CntrTypes != "?")
                if (strWhere == "")
                    strWhere += _Query + " where CntrTypes=" + CntrTypes;
                else
                    strWhere += " and CntrTypes=" + CntrTypes;


            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetSlotValues(string SlotOperator, string SlotRef, string SlotTerms, string POL, string POD)
        {
            string strWhere = "";
            string _Query = "select * from V_SlatMasterView_Report";

            if (POL != "" && POL != "null" && POL != "0" && POL != "?")
                if (strWhere == "")
                    strWhere += _Query + " where intPOL=" + POL;
                else
                    strWhere += " and intPOL=" + POL;

            if (POD != "" && POD != "null" && POD != "0" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " where intPOD=" + POD;
                else
                    strWhere += " and intPOD=" + POD;

            if (SlotTerms != "" && SlotTerms != "null" && SlotTerms != "0" && SlotTerms != "?")
                if (strWhere == "")
                    strWhere += _Query + " where SlotTermID=" + SlotTerms;
                else
                    strWhere += " and SlotTermID=" + SlotTerms;

            if (SlotRef != "" && SlotRef != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where SlotRef like '%" + SlotRef + "%'";
                else
                    strWhere += " and SlotRef like '%" + SlotRef + "%'";

            if (SlotOperator != "" && SlotOperator != "null" && SlotOperator != "0" && SlotOperator != "?")
                if (strWhere == "")
                    strWhere += _Query + " where SlotTermID=" + SlotOperator;
                else
                    strWhere += " and SlotTermID=" + SlotOperator;

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetTariffValues(string Loc, string Shipment, string TariffTypes, string Status, string ServiceType, string DGClass, string CollMode)
        {
            string strWhere = "";
            string _Query = "select PM.ID ,(SELECT TOP 1  PORTNAME FROM  NVO_PortMaster WHERE ID = PortLocationID) as Location,(SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = ShipmentTypeID) as ShipmentType, (SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = ModuleID) as MODULE, (SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = CommodityTypeID) as CommodityType, " +
             " (SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = TraiffRegular) as TariffType, " +
           " (SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = TariffModeID) as TariffMode, (SELECT TOP 1  Description FROM  NVO_tblDLValues WHERE ID = ServiceTypeID) as ServiceType, (SELECT TOP 1  PortName FROM  NVO_PortMaster WHERE ID = DestCountryID) as DestPort," +
           " (SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = CollectionModeID) as CollectionMode ,(SELECT TOP 1  ChargeCode FROM  NVO_ChargeTB WHERE ID = PT.ChargeCodeID) as ChargeCodeBreakup,(SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = PT.ShipmetID) as ShipmentBreakup,(SELECT TOP 1  GeneralName FROM  NVO_GeneralMaster WHERE ID = PT.CommodityID) as CommodityBreakup,(SELECT TOP 1  Size FROM  NVO_tblCntrTypes WHERE ID = PT.CntrTypeID) as CntrTypeBreakup,PT.Amount as BreakupAmount,NVO_PortTariffdtls.Amount As TarifChgAmount  ,* from NVO_PortTariffMaster PM " +
           " INNER JOIN NVO_PortTariffdtls ON NVO_PortTariffdtls.PTID = PM.ID   left outer JOIN NVO_PortTariffChargedtls PT ON PT.TCID = NVO_PortTariffdtls.TID  ";

            if (Loc != "" && Loc != "null" && Loc != "0" && Loc != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.PortLocationID=" + Loc;
                else
                    strWhere += " and PM.PortLocationID=" + Loc;

            if (Shipment != "" && Shipment != "null" && Shipment != "0" && Shipment != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.ShipmentTypeID=" + Shipment;
                else
                    strWhere += " and PM.ShipmentTypeID=" + Shipment;

            if (TariffTypes != "" && TariffTypes != "null" && TariffTypes != "0" && TariffTypes != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.TraiffRegular=" + TariffTypes;
                else
                    strWhere += " and PM.TraiffRegular=" + TariffTypes;

            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.StatusID=" + Status;
                else
                    strWhere += " and PM.StatusID=" + Status;

            if (ServiceType != "" && ServiceType != "null" && ServiceType != "0" && ServiceType != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.ServiceTypeID=" + ServiceType;
                else
                    strWhere += " and PM.ServiceTypeID=" + ServiceType;

            if (DGClass != "" && DGClass != "null" && DGClass != "0" && DGClass != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.GroupID=" + DGClass;
                else
                    strWhere += " and PM.GroupID=" + DGClass;

            if (CollMode != "" && CollMode != "null" && CollMode != "0" && CollMode != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PM.CollectionModeID=" + CollMode;
                else
                    strWhere += " and PM.CollectionModeID=" + CollMode;

            if (strWhere == "")
                strWhere = _Query + " ";


            return Manag.GetViewData(strWhere + " order by PM.ID", "");
        }
        public DataTable GetWaiverValues(string ReqType, string BookingNo, string FreeDays, string WaiverAmt)
        {
            string strWhere = "";
            string _Query = " select distinct BookingID as ID,(select top(1) GeneralName from NVO_GeneralMaster where Id = RequestTypeID) as RequestTypes, " +
                            " (select top(1) BookingNo from NVO_Booking where Id = BookingID) as BookingNo, Freedays, WaiverAmt,'' Status, " +
                            " (select CntrNo from NVO_Containers where ID = NVO_WaiverDetails.ContainerID) as CntrNo, " +
                            " (select ChgDesc from NVO_ChargeTB where ID = NVO_WaiverDetails.ChargeType) as ChargeType, " +
                            " case when TransTypeID = 1 then 'Agency Export' else 'Agency Import' end as TransType from NVO_WaiverDetails ";

            if (ReqType != "" && ReqType != "null" && ReqType != "0" && ReqType != "?")
                if (strWhere == "")
                    strWhere += _Query + " where RequestTypeID=" + ReqType;
                else
                    strWhere += " and RequestTypeID=" + ReqType;

            if (BookingNo != "" && BookingNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where (select top(1) BookingNo from NVO_Booking where Id = BookingID) like '%" + BookingNo + "%'";
                else
                    strWhere += " and (select top(1) BookingNo from NVO_Booking where Id = BookingID) like '%" + BookingNo + "%'";

            if (FreeDays != "" && FreeDays != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where Freedays like '%" + FreeDays + "%'";
                else
                    strWhere += " and Freedays like '%" + FreeDays + "%'";

            if (WaiverAmt != "" && WaiverAmt != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where WaiverAmt like '%" + WaiverAmt + "%'";
                else
                    strWhere += " and WaiverAmt like '%" + WaiverAmt + "%'";

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetContainerContractValues(string ShipmentType, string TariffType, string ContainerType)
        {
            string strWhere = "";
            string _Query = "select NVO_ContRentContract.ID,(Select top 1 GeneralName  from NVO_GeneralMaster where ID = TariffTypeID ) AS TariffTypeV,(Select top 1 Size  from NVO_tblCntrTypes where ID = ContainerType ) AS Size,  convert(varchar, ValidFrom, 103) as ValidFrom,convert(varchar, ValidTill, 103) as ValidTill,(Select top 1 GeneralName from NVO_GeneralMaster where ID = ShipmentTypeID ) AS ShipmentTypeV,(Select top 1 GeneralName  from NVO_GeneralMaster where ID = ChargeOwnerID ) AS ChargeOwnerIDV,(Select top 1 GeneralName  from NVO_GeneralMaster where ID = ChargeTypeID ) AS ChargeType,(Select top 1 (ChgCode +' - '+ ChgDesc) from NVO_ChargeTB where ID = ChargesID ) AS Charges,case when status=1 then 'Active' else 'InActive' end StatusV  from NVO_ContRentContract ";

            if (ShipmentType != "" && ShipmentType != "null" && ShipmentType != "0" && ShipmentType != "?")
                if (strWhere == "")
                    strWhere += _Query + " where ShipmentTypeID=" + ShipmentType;
                else
                    strWhere += " and ShipmentTypeID=" + ShipmentType;

            if (TariffType != "" && TariffType != "null" && TariffType != "0" && TariffType != "?")
                if (strWhere == "")
                    strWhere += _Query + " where TariffTypeID=" + TariffType;
                else
                    strWhere += " and TariffTypeID=" + TariffType;

            if (ContainerType != "" && ContainerType != "null" && ContainerType != "0" && ContainerType != "?")
                if (strWhere == "")
                    strWhere += _Query + " where ContainerType=" + ContainerType;
                else
                    strWhere += " and ContainerType=" + ContainerType;

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetCommissionContractValues(string ShipmentType, string Agency, string CommissionCharge, string Status)
        {
            string strWhere = "";
            string _Query = "select ID,(Select top 1 GeneralName from NVO_GeneralMaster WHERE ID = ShipmentType) As ShipmentType,(Select top 1 AgencyName from NVO_AgencyMaster WHERE ID = AgencyID) As Agency," +
                "(Select top 1 ChgCode from NVO_ChargeTB WHERE ID = ComChargeID) As CommissionCharge,convert(varchar, ValidFrom, 103) as ValidFrom,convert(varchar, ValidTill, 103) as ValidTill, " +
                " CASE WHEN StatusID = 1 Then 'ACTIVE' WHEN StatusID = 2 THEN 'INACTIVE' END STATUSResult from NVO_Commissioncontract ";

            if (ShipmentType != "" && ShipmentType != "null" && ShipmentType != "0" && ShipmentType != "?")
                if (strWhere == "")
                    strWhere += _Query + " where ShipmentType=" + ShipmentType;
                else
                    strWhere += " and ShipmentType=" + ShipmentType;

            if (Agency != "" && Agency != "null" && Agency != "0" && Agency != "?")
                if (strWhere == "")
                    strWhere += _Query + " where AgencyID=" + Agency;
                else
                    strWhere += " and AgencyID=" + Agency;

            if (CommissionCharge != "" && CommissionCharge != "null" && CommissionCharge != "0" && CommissionCharge != "?")
                if (strWhere == "")
                    strWhere += _Query + " where ComChargeID=" + CommissionCharge;
                else
                    strWhere += " and ComChargeID=" + CommissionCharge;

            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " where StatusID=" + Status;
                else
                    strWhere += " and StatusID=" + Status;

            if (strWhere == "")
                strWhere = _Query + " order by ID asc ";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetVoyageLockingValues(string VesVoy)
        {
            string strWhere = "";
            string _Query = " Select  Distinct  V.ID As VoyageID,(select top(1) VesselName from NVO_VesselMaster where NVO_VesselMaster.ID = V.VesselID ) + '-' + " +
                            " (select top(1) VR.ExportVoyageCd from NVO_VoyageRoute VR where VR.VoyageID = V.ID order by VR.RID asc) as VesVoy,(Select Top 1 PORTNAME from NVO_PortMainMaster " +
                            " where NVO_PortMainMaster.ID = (select top(1) PortID from NVO_VoyageRoute VR where VR.VoyageID = V.ID order by VR.RID asc)) AS CurrentPort, " +
                            " (Select Top 1 PORTNAME from NVO_PortMainMaster where NVO_PortMainMaster.ID = (select top(1) PortID from NVO_VoyageRoute VR " +
                            " where VR.VoyageID = V.ID order by Vr.RID desc)) AS NextPort," +
                            " case when  (select Case when dd <= isBLLocked  then 'Locked' else 'Partially Locked'  end from NVO_V_VoyLockStatusMain where NVO_V_VoyLockStatusMain.VesVoyID = V.ID) != 'NULL' then   (select Case when dd <= isBLLocked  then 'Locked' else 'Partially Locked'  end from NVO_V_VoyLockStatusMain where NVO_V_VoyLockStatusMain.VesVoyID = V.ID) else 'Pending' end AS Status from NVO_Voyage V where IsExpImp=0";

            if (VesVoy != "" && VesVoy != "null" && VesVoy != "0" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " and V.ID=" + VesVoy;
                else
                    strWhere += " and V.ID=" + VesVoy;



            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetVoyageOpeningValues(string Agency)
        {
            string strWhere = "";
            string _Query = "  Select Distinct V.ID,isnull((select top(1) DisVoyageNo from NVO_VoyageOpen where PrevVoyageID = V.ID),'') as DisVoyageNo,(select top(1) VesselName from NVO_VesselMaster where ID = V.VesselID) + ' -' + (select top(1)ExportVoyageCd from NVO_VoyageRoute where VoyageID = V.ID) as VesVoy,(select top(1) PortName from NVO_PortMainMaster inner join NVO_VoyageRoute on NVO_VoyageRoute.PortID = NVO_PortMainMaster.ID where NVO_VoyageRoute.VoyageID = V.ID  ORDER BY RID ASC) as LoadPort," +
               "  isnull((select top(1) VID from NVO_VoyageOpen where PrevVoyageID = V.ID),0) as OpenID, CASE WHEN  isnull((select top(1) VID from NVO_VoyageOpen where PrevVoyageID = V.ID),0) = '0' Then 'Not Opened' Else 'Already Opened' End OpenStatus,isnull((select top(1) TerminalName from NVO_VoyageOpen inner join NVO_TerminalMaster TM ON TM.ID = NVO_VoyageOpen.TerminalID where PrevVoyageID = V.ID) ,0) as DiscTerminal,convert(varchar, (select Top(1) ETA from NVO_VoyageOpen where NVO_VoyageOpen.PrevVoyageID = V.ID), 103) as ETA " +
                 " from NVO_Voyage V inner join NVO_VoyageRoute on NVO_VoyageRoute.VoyageID = V.ID inner join NVO_V_VoyLockStatusMain on NVO_V_VoyLockStatusMain.VesVoyId = V.ID inner join NVO_V_VoyLockStatus on NVO_V_VoyLockStatus.VesVoyId = V.ID inner join NVO_Booking on NVO_Booking.ID = NVO_V_VoyLockStatus.ID where NVO_Booking.DestinationAgentID =" + Agency + " or  NVO_Booking.TranshipmetAgentID =" + Agency;


            if (strWhere == "")
                strWhere = _Query;


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetImportValues(string BookingNo, string POL, string POD, string VesVoy, string Agency)
        {
            string strWhere = "";
            string _Query = "select ID,BLNumber,BkgID,POO,POL,POD,FPOD,BLVesVoy,CommodityType,BlDirect,BLType from NVO_v_ImportBLView";

            strWhere = _Query += " where AgencyID=" + Agency;

            if (BookingNo != "" && BookingNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BLNumber like '%" + BookingNo + "%'";
                else
                    strWhere += " and BLNumber like '%" + BookingNo + "%'";

            if (POL != "" && POL != "null" && POL != "0" && POL != "?")
                if (strWhere == "")
                    strWhere += _Query + " and POLID=" + POL;
                else
                    strWhere += " and POLID=" + VesVoy;

            if (POD != "" && POD != "null" && POD != "0" && POD != "?")
                if (strWhere == "")
                    strWhere += _Query + " and PODID=" + POD;
                else
                    strWhere += " and PODID=" + VesVoy;

            if (VesVoy != "" && VesVoy != "null" && VesVoy != "0" && VesVoy != "?")
                if (strWhere == "")
                    strWhere += _Query + " and VesVoyID=" + VesVoy;
                else
                    strWhere += " and VesVoyID=" + VesVoy;



            if (strWhere == "")
                strWhere = _Query + "Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetBankMasterValues(string BankName, string AccNo)
        {
            string strWhere = "";
            string _Query = "SELECT ID,AccountNo,BankName,BankCode, case when StatusID = 1 then 'Active' when StatusID = 0 then 'Inactive' ELSE '' END as Status FROM NVO_FinBankMaster";

            if (BankName != "" && BankName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where BankName like '%" + BankName + "%'";
                else
                    strWhere += " and BankName like '%" + BankName + "%'";

            if (AccNo != "" && AccNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where AccountNo like '%" + AccNo + "%'";
                else
                    strWhere += " and AccountNo like '%" + AccNo + "%'";


            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetCreditControlValues(string Party, string PartyType, string ExemTax)
        {
            string strWhere = "";
            string _Query = "select FCC.ID," +
                 " (select top 1  upper(CustomerName + '-' + Branch) as CustomerName from NVO_CustomerMaster " +
                            " inner join NVO_CusBranchLocation on NVO_CusBranchLocation.CustomerID = NVO_CustomerMaster.Id where NVO_CusBranchLocation.CID =FCC.PartyID ) as  " +

                " CustomerName,GM.GeneralName As PartyType,case when FCC.ExemptFromTax =1 THEN 'YES' when FCC.ExemptFromTax =2 THEN 'NO' END TaxExempt from NVO_FinCustomerCreditControl FCC " +
                          //" inner join NVO_CusBranchLocation CB on CB.CID = FCC.PartyID "+
                          // " inner join  NVO_CustomerMaster CM ON CM.ID = FCC.PartyID " +
                          " inner join  NVO_GeneralMaster GM ON GM.ID = FCC.PartyType  AND GM.SeqNo = 18 ";

            if (Party != "" && Party != "null" && Party != "0" && Party != "?")
                if (strWhere == "")
                    strWhere += _Query + " and FCC.PartyID=" + Party;
                else
                    strWhere += " and FCC.PartyID=" + Party;

            if (PartyType != "" && PartyType != "null" && PartyType != "0" && PartyType != "?")
                if (strWhere == "")
                    strWhere += _Query + " and FCC.PartyType=" + PartyType;
                else
                    strWhere += " and FCC.PartyType=" + PartyType;

            if (ExemTax != "" && ExemTax != "null" && ExemTax != "0" && ExemTax != "?")
                if (strWhere == "")
                    strWhere += _Query + " and FCC.ExemptFromTax=" + ExemTax;
                else
                    strWhere += " and FCC.ExemptFromTax=" + ExemTax;

            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetTaxDeclrationValues(string TaxName, string Status)
        {
            string strWhere = "";
            string _Query = " Select CT.ID,CT.TaxName,CT.TaxPercentage,CM.CountryName,case when CT.StatusID=1 THEN 'ACTIVE' WHEN CT.StatusID=2 THEN 'INACTIVE' END StatusV from NVO_ChgTaxDeclaration CT Inner join NVO_CountryMaster CM ON CM.ID = CT.CountryID ";

            if (TaxName != "" && TaxName != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where CT.TaxName like '%" + TaxName + "%'";
                else
                    strWhere += " and CT.TaxName like '%" + TaxName + "%'";

            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " and CT.StatusID=" + Status;
                else
                    strWhere += " and CT.StatusID=" + Status;



            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetTaxNamesValues(string TaxCode)
        {
            string strWhere = "";
            string _Query = "select ID,TaxCode,TaxDescription,CountryID,Country  from NVO_FinanceTaxNames ";

            if (TaxCode != "" && TaxCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where TaxCode like '%" + TaxCode + "%'";
                else
                    strWhere += " and TaxCode like '%" + TaxCode + "%'";



            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetReceiptsValues(string RecpNo, string RecParty, string Date, string RecType)
        {
            string strWhere = "";
            string _Query = " Select ID,ReceiptNo, convert(varchar, DtReceipt, 103) as RecDate,PartyName, " +
                            " case when ReceiptCatgory = 1 then 'LOCAL' else 'OVERSEAS'  end as RecTypes ,case when ReceiptTypes = 193 then 'ON ACCOUNT'  when ReceiptTypes = 194 then 'BILL PAYMENT'  else 'UN DEPOSIT ACCOUNT' end as PaymentType,LocalAmount " +
                            " from NVO_Receipts";
            strWhere += _Query + " where ReceiptStatus in (1)";
            if (RecpNo != "" && RecpNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where ReceiptNo like '%" + RecpNo + "%'";
                else
                    strWhere += " and ReceiptNo like '%" + RecpNo + "%'";

            if (RecParty != "" && RecParty != "null" && RecParty != "0" && RecParty != "?")
                if (strWhere == "")
                    strWhere += _Query + " and PartyID=" + RecParty;
                else
                    strWhere += " and PartyID=" + RecParty;

            if (Date != "" && Date != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where DtReceipt <= ' " + Date + "'";
                else
                    strWhere += " and DtReceipt <= ' " + Date + "'";

            if (RecType != "" && RecType != "null" && RecType != "0" && RecType != "?")
                if (strWhere == "")
                    strWhere += _Query + " and ReceiptTypes=" + RecType;
                else
                    strWhere += " and ReceiptTypes=" + RecType;

            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetReceiptCancelValues(string RecpNo, string RecParty, string Date, string RecType, string Agency)
        {
            string strWhere = "";
            string _Query = "select R.ID,R.ReceiptNo,convert(varchar, R.DtReceipt, 103) as RecDate,R.PartyName,  case when R.ReceiptTypes = 193 then 'ON ACCOUNT'  when ReceiptTypes = 194 then 'BILL PAYMENT' else 'UN DEPOSIT ACCOUNT' end as PaymentType,Amount ,convert(varchar, RL.CancelledOn, 103) as CancelledOn, " +
                            " (select top 1 UserName from NVO_UserDetails Where ID = CancelledBy)  CancelledBy from NVO_ReceiptCancellationLog RL  INNER JOIN NVO_Receipts R on R.ID = RL.ReceiptID ";

            strWhere += _Query + " where R.ReceiptStatus in (2)";
            if (RecpNo != "" && RecpNo != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where R.ReceiptNo like '%" + RecpNo + "%'";
                else
                    strWhere += " and R.ReceiptNo like '%" + RecpNo + "%'";

            if (RecParty != "" && RecParty != "null" && RecParty != "0" && RecParty != "?")
                if (strWhere == "")
                    strWhere += _Query + " and R.PartyID=" + RecParty;
                else
                    strWhere += " and R.PartyID=" + RecParty;

            if (Date != "" && Date != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where R.DtReceipt <= ' " + Date + "'";
                else
                    strWhere += " and R.DtReceipt <= ' " + Date + "'";

            if (RecType != "" && RecType != "null" && RecType != "0" && RecType != "?")
                if (strWhere == "")
                    strWhere += _Query + " and R.ReceiptTypes=" + RecType;
                else
                    strWhere += " and R.ReceiptTypes=" + RecType;

            if (Agency != "" && Agency != "0" && Agency != "2" && Agency != "undefined" && Agency != null)

                if (strWhere == "")
                    strWhere += _Query + " where RL.AgencyID = " + Agency;
                else
                    strWhere += " and RL.AgencyID = " + Agency;

            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMNRDamageMasterValues(string DmgCode, string DmgDesc)
        {
            string strWhere = "";
            string _Query = " select Id,DamageCode,DamageDescription,status, " +
                " case when status = 1 then 'Active' when status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_MNRDamageMaster ";

            if (DmgCode != "" && DmgCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where DamageCode like '%" + DmgCode + "%'";
                else
                    strWhere += " and DamageCode like '%" + DmgCode + "%'";

            if (DmgDesc != "" && DmgDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where DamageDescription like '%" + DmgDesc + "%'";
                else
                    strWhere += " and DamageDescription like '%" + DmgDesc + "%'";



            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMNRRepairValues(string RprCode, string RprDesc)
        {
            string strWhere = "";
            string _Query = " select Id,RepairCode,RepairDescription,status, " +
                " case when status = 1 then 'Active' when status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_MNRRepairMaster ";

            if (RprCode != "" && RprCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where RepairCode like '%" + RprCode + "%'";
                else
                    strWhere += " and RepairCode like '%" + RprCode + "%'";

            if (RprDesc != "" && RprDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where RepairDescription like '%" + RprDesc + "%'";
                else
                    strWhere += " and RepairDescription like '%" + RprDesc + "%'";



            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMNRComponentValues(string CompCode, string CompDesc, string Assembly)
        {
            string strWhere = "";
            string _Query = " select NVO_MNRComponentMaster.Id,ComponentCode,ComponentDescription,NVO_MNRComponentMaster.status,NVO_GeneralMaster.Generalname as Assembly, " +
                " case when NVO_MNRComponentMaster.status = 1 then 'Active' when NVO_MNRComponentMaster.status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_MNRComponentMaster " +
                " Inner Join NVO_GeneralMaster on NVO_GeneralMaster.ID =NVO_MNRComponentMaster.AssemblyID and Seqno=35 ";

            if (CompCode != "" && CompCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where ComponentCode like '%" + CompCode + "%'";
                else
                    strWhere += " and ComponentCode like '%" + CompCode + "%'";

            if (CompDesc != "" && CompDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where ComponentDescription like '%" + CompDesc + "%'";
                else
                    strWhere += " and ComponentDescription like '%" + CompDesc + "%'";


            if (Assembly != "" && Assembly != "null" && Assembly != "0" && Assembly != "?")
                if (strWhere == "")
                    strWhere += _Query + " and AssemblyID=" + Assembly;
                else
                    strWhere += " and AssemblyID=" + Assembly;


            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetMNRLocationValues(string LocCode, string LocDesc)
        {
            string strWhere = "";
            string _Query = " select Id,LocationCode,Description,status, " +
                " case when status = 1 then 'Active' when status = 0 then 'Inactive' ELSE '' END as StatusResult  from NVO_MNRLocationMaster ";

            if (LocCode != "" && LocCode != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where LocationCode like '%" + LocCode + "%'";
                else
                    strWhere += " and LocationCode like '%" + LocCode + "%'";

            if (LocDesc != "" && LocDesc != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where Description like '%" + LocDesc + "%'";
                else
                    strWhere += " and Description like '%" + LocDesc + "%'";



            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetLeaseContractValues(string LeaseRef, string LeasePartner, string Status)
        {
            string strWhere = "";
            string _Query = "Select LC.ID,LC.ContractRefNo,CM.CustomerName as LeasingPartner,GM.GeneralName as PickupCriteria,convert(varchar,LC.DtContractFrom, 106) As DtContractFrom,convert(varchar,LC.DtContractTill, 106) As DtContractTill from NVO_LeaseContract LC " +
           " LEFT OUTER JOIN NVO_CustomerMaster CM ON CM.id = LC.LeasingPartnerID " +
           " LEFT OUTER JOIN NVO_GeneralMaster GM ON GM.id = LC.PickupCriteriaID AND GM.SeqNo = 15 ";

            if (LeaseRef != "" && LeaseRef != "undefined")
                if (strWhere == "")
                    strWhere += _Query + " where LC.ContractRefNo like '%" + LeaseRef + "%'";
                else
                    strWhere += " and LC.ContractRefNo like '%" + LeaseRef + "%'";


            if (LeasePartner != "" && LeasePartner != "null" && LeasePartner != "0" && LeasePartner != "?")
                if (strWhere == "")
                    strWhere += _Query + " and LC.LeasingPartnerID=" + LeasePartner;
                else
                    strWhere += " and LC.LeasingPartnerID=" + LeasePartner;


            if (Status != "" && Status != "null" && Status != "0" && Status != "?")
                if (strWhere == "")
                    strWhere += _Query + " and LC.Status=" + Status;
                else
                    strWhere += " and LC.Status=" + Status;



            if (strWhere == "")
                strWhere = _Query + " Order By ID Asc";


            return Manag.GetViewData(strWhere, "");
        }

        public DataTable GetOnHireValues(string PickupRef)
        {
            string strWhere = "";
            string _Query = "select LC.ID AS LeaseContractID,RequestNo,LC.ContractRefNo,cm.CustomerName as LeasingPartner,GM.GeneralName As LeaseTerm from NVO_ContainerOnHire OH " +
              " inner join NVO_LeaseContract LC on LC.ID = OH.LeasePickUpRefID " +
             " inner join NVO_CustomerMaster CM on CM.ID = OH.OnHireLeasingPartnerID " +
            " inner join NVO_GeneralMaster GM on GM.ID = OH.LeasingTermID  and GM.SeqNo = 13 ";




            if (PickupRef != "" && PickupRef != "null" && PickupRef != "0" && PickupRef != "?")
                if (strWhere == "")
                    strWhere += _Query + " and LC.ID=" + PickupRef;
                else
                    strWhere += " and LC.ID=" + PickupRef;



            if (strWhere == "")
                strWhere = _Query + " Order By LC.ID Asc";


            return Manag.GetViewData(strWhere, "");
        }
    }
}