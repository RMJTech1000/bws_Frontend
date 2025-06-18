using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;


using System.Threading.Tasks;
using System.Web.Mvc;
using DataTier;
using iTextSharp.text;
using QRCoder;
using DataManager;
using System.Net.Mail;
using System.Text;
using System.Data;
using System.IO;
//using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Web;
//using System.Drawing;

namespace NVOCShipping.api
{
    public class OnlineExcelPrintController : ApiController
    {
        DocumentManager Manag = new DocumentManager();
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/onlineexcel/BankMasterExcel")]
        public HttpResponseMessage BankMasterExcel(string BankName, string AccNo, string User)
        {
            MemoryStream memoryStream = new MemoryStream();
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
                r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

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
                r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

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
                pck.SaveAs(memoryStream);
               
            }

            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=example.xlsx");
            HttpContext.Current.Response.BinaryWrite(memoryStream.ToArray());
            HttpContext.Current.Response.End();

            HttpResponseMessage response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StreamContent(memoryStream)
            };

            //pck.SaveAs(response.Content);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment;  filename=BankMasterList.xlsx");
           // = new ContentDispositionHeaderValue("attachment;  filename=BankMasterList.xlsx");
            return response;
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
    }
}