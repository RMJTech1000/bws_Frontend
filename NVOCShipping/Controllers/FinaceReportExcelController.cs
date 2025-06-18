using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using DataManager;
using System.Text;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Data.SqlClient;


namespace NVOCShipping.Controllers
{
    public class FinaceReportExcelController : Controller
    {
        // GET: FinaceReportExcel
        StringBuilder sb = new StringBuilder();
        FinanceManager Manag = new FinanceManager();

        #region Statment  of Account
        public ActionResult StatementOfAccountExcelReport(string PartyID, string FromDate, string ToDate, string ReportType)
        {
            StatementAccountExce(PartyID, FromDate, ToDate, ReportType);
            return View();

        }

        public void StatementAccountExce(string PartyID, string FromDate, string ToDate, string ReportType)
        {

            string CompanyName = "";
            //DataTable _dtC = RegMang.GetViewData("select * from Nav_TblCompanyDetails", "");
            //if (_dtC.Rows.Count > 0)
            //    CompanyName = _dtC.Rows[0]["CompanyName"].ToString();

            string LeftAll = "<td align='center' style='background-color:#04208c; color:#ffffff; font-family:Arial; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; font-style: normal; font-weight: bold; font-size:13px;'>";
            string RColorRightLeft = "<td colspan='5' align='center' style='color:#04208c;  border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family: Arial; font-style: normal; font-weight: bold; font-size:16px;'>";
            string RightALm1Color = "<td align='Right' style='color: blue; background-color:#B6F2FC;  color:#FF0000; border-top: Black thin inset; border-bottom: Black thin inset; border-Right: Black thin inset; border-Left: Black thin inset; inset; font-family: Arial; font-style: normal;  font-size:16px;'>";
            string RowHeader = "<td align='center' style='color: black; font-family: Arial; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; font-style: normal; font-weight: bold; font-size:20px;'>";
            string Totalspace = "<td colspan='2' align='Center' style='color: black; background-color:#B6F2FC; color:#FF0000; border-top: Black thin inset; border-bottom: Black thin inset; border-Right: Black thin inset; border-Left: Black thin inset; inset; font-family: Arial; font-style: normal;  font-size:16px;'>";
            string Totalspace6 = "<td colspan='5' align='Center' style='color: black; background-color:#B6F2FC; color:#FF0000; border-top: Black thin inset; border-bottom: Black thin inset; border-Right: Black thin inset; border-Left: Black thin inset; inset; font-family: Arial; font-style: normal;  font-size:16px;'>";
            string Totalspace1 = "<td colspan='4' align='Center' style='color: black; border-top: Black thin inset; background-color:#B6F2FC; border-bottom: Black thin inset; border-Right: Black thin inset; border-Left: Black thin inset; inset; font-family: Arial; font-style: normal;  font-size:10px;'>";
            string DateRColorRightLeft = "<td colspan='10' align='center' style='color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family: Arial; font-style: normal; font-weight: bold; font-size:10px;'>";
            string RColorRightLeftDt = "<td colspan='8' align='center' style='color:#04208c;  border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family: Arial; font-style: normal; font-weight: bold; font-size:16px;'>";

            if (ReportType == "1")
            {
                #region Summary
                sb.Append("<table>");
                sb.Append("<tr></tr><tr></tr><tr>");
                sb.Append("<td colspan='5' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>SPRINT GLOBAL INC</TD></TR>");
                sb.Append("<tr></tr>");
                sb.Append("<tr><td>Location</td><td align='right'>" + "Locations" + "</td></tr>");
                sb.Append("<tr><td>From Date </td><td align='right'>" + FromDate + "</td></tr>");
                sb.Append("<tr><td>To Date </td><td align='right'>" + ToDate + "</td></tr>");
                sb.Append("<tr><td>Party Name</td><td align='right'>" + PartyID + "</td></tr>");
                sb.Append("<tr>");
                sb.Append(RColorRightLeft + "STATEMENT OF ACCOUNT</TD></TR>");
                sb.Append("<tr>");
                sb.Append(LeftAll + " SL NO</td>");
                sb.Append(LeftAll + " Party Name</td>");
                sb.Append(LeftAll + " Debit  </td>");
                sb.Append(LeftAll + " Credit</td>");
                sb.Append(LeftAll + " Balance</td>");
                sb.Append("</tr>");

                DataTable _dtEx = GetAccountSummary(FromDate, ToDate, PartyID);
                string ContentLeft = "";
                string ContentRight = "";
                int RowAl = 1;
                int RolNo = 0;
                decimal DebitAmt = 0;
                decimal CreditAmt = 0;

                decimal BlanceDebitAmt = 0;
                decimal BlanceCreditAmt = 0;
                decimal BalancT = 0;

                for (int i = 0; i < _dtEx.Rows.Count; i++)
                {
                    if (RowAl == 1)
                    {
                        RowAl = 2;
                        ContentLeft = "<td align='Left' style= ' background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style= ' background-color: #FFFFFF;  color: #000000;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";

                    }
                    else if (RowAl == 2)
                    {
                        RowAl = 1;
                        ContentLeft = "<td align='Left' style='background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style='background-color: #FFFFFF;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";


                    }

                    sb.Append("<tr>");
                    RolNo++;

                    sb.Append(ContentRight + RolNo + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["CustomerName"].ToString() + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["DebitAmt"].ToString()).ToString("#,#0.00")) + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["CreditAmt"].ToString()).ToString("#,#0.00")) + "</td>");

                    DebitAmt += decimal.Parse(_dtEx.Rows[i]["DebitAmt"].ToString());
                    CreditAmt += decimal.Parse(_dtEx.Rows[i]["CreditAmt"].ToString());
                    BlanceDebitAmt = decimal.Parse(_dtEx.Rows[i]["DebitAmt"].ToString());
                    BlanceCreditAmt = decimal.Parse(_dtEx.Rows[i]["CreditAmt"].ToString());

                    if (BlanceDebitAmt != 0)
                        BalancT += BlanceDebitAmt;
                    if (BlanceCreditAmt != 0)
                        BalancT -= BlanceCreditAmt;

                    sb.Append(ContentRight + (decimal.Parse((BalancT).ToString()).ToString("#,#0.00")) + "</td>");



                    sb.Append("</tr>");
                }
                sb.Append("</tr>");
                sb.Append(Totalspace + "Grand Total Amount" + "</td>");
                sb.Append(RightALm1Color + (decimal.Parse(DebitAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (decimal.Parse(CreditAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (DebitAmt - CreditAmt).ToString("#,#0.00") + " </td>");
                sb.Append("</tr>");
                sb.Append("</table>");
                Response.Write(sb.ToString());
                Response.Clear();
                Response.AddHeader("content-disposition", "attachment;filename=StatementofAccountSummary" + System.DateTime.Now + ".xls");
                Response.Charset = "";
                Response.ContentType = "application/vnd.xls";
                System.IO.StringWriter stringWrite = new System.IO.StringWriter();
                System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
                htmlWrite.Write(sb.ToString());
                Response.Write(stringWrite.ToString());
                Response.End();
                #endregion
            }
            else if (ReportType == "2")
            {
                int RowColumn = 8;
                int RowIndex = 0;
              

                #region Details
                sb.Append("<table>");
                sb.Append("<tr></tr><tr></tr><tr>");
                sb.Append("<td colspan='" + RowColumn + "' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>COMPNAY NAME</TD></TR>");
                sb.Append("<tr></tr>");
                sb.Append("<tr><td>Location</td><td align='right'>" + "Locations" + "</td></tr>");
                sb.Append("<tr><td>From Date </td><td align='right'>" + FromDate + "</td></tr>");
                sb.Append("<tr><td>To Date </td><td align='right'>" + ToDate + "</td></tr>");
                sb.Append("<tr><td>Party Name</td><td align='right'>" + PartyID + "</td></tr>");
                sb.Append("<tr>");
                sb.Append("<td colspan='" + RowColumn + "' align='center' style='color:#04208c;  border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family: Arial; font-style: normal; font-weight: bold; font-size:16px;'>STATEMENT OF ACCOUNT</TD></TR>");
                sb.Append("<tr>");
                sb.Append(LeftAll + " SL NO</td>");
                sb.Append(LeftAll + " Date</td>");
                sb.Append(LeftAll + " Voucher No</td>");
                sb.Append(LeftAll + " </td>");
                sb.Append(LeftAll + " AR/AP refernce no</td>");
                sb.Append(LeftAll + " Debit </td>");
                sb.Append(LeftAll + " Credit </td>");
                sb.Append(LeftAll + " Balance  </td>");
                sb.Append("</tr>");
                int RolNo = 0;
                int RowAl = 1;
                string ContentLeft = "";
                string ContentRight = "";

                decimal DebitAmt = 0;
                decimal CreditAmt = 0;

                decimal BlanceDebitAmt = 0;
                decimal BlanceCreditAmt = 0;
                decimal BalancT = 0;
                DataTable _dtOpn = GetAccountSummarydlsOpen(FromDate, ToDate, PartyID);
                if (_dtOpn.Rows.Count > 0)
                {
                    if (RowAl == 1)
                    {
                        RowAl = 2;
                        ContentLeft = "<td align='Left' style= ' background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:14px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style= 'background-color: #FFFFFF;  color: #000000;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:14px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";


                    }
                    else if (RowAl == 2)
                    {
                        RowAl = 1;
                        ContentLeft = "<td align='Left' style='background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:14px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style='background-color: #FFFFFF;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:14px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";


                    }

                    sb.Append("<tr>");
                    RolNo++;

                    sb.Append(ContentRight + RolNo + "</td>");
                    sb.Append(ContentLeft + "AS ON DATE: " + FromDate + "</td>");
                    sb.Append(ContentLeft + _dtOpn.Rows[0]["opening"].ToString() + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");
                    sb.Append(ContentRight + _dtOpn.Rows[0]["DebitAmt"].ToString() + "</td>");
                    sb.Append(ContentRight + _dtOpn.Rows[0]["CreditAmt"].ToString() + "</td>");
                    sb.Append(ContentRight + _dtOpn.Rows[0]["Balance"].ToString() + "</td>");

                    DebitAmt += decimal.Parse(_dtOpn.Rows[0]["DebitAmt"].ToString());
                    CreditAmt += decimal.Parse(_dtOpn.Rows[0]["CreditAmt"].ToString());
                    BlanceDebitAmt = decimal.Parse(_dtOpn.Rows[0]["DebitAmt"].ToString());
                    BlanceCreditAmt = decimal.Parse(_dtOpn.Rows[0]["CreditAmt"].ToString());

                    if (BlanceDebitAmt != 0)
                        BalancT += BlanceDebitAmt;
                    if (BlanceCreditAmt != 0)
                        BalancT -= BlanceCreditAmt;

                }

                DataTable _dtEx = GetAccountSummarydls(FromDate, ToDate, PartyID);
                for (int i = 0; i < _dtEx.Rows.Count; i++)
                {
                    if (RowAl == 1)
                    {
                        RowAl = 2;
                        ContentLeft = "<td align='Left' style= ' background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style= ' background-color: #FFFFFF;  color: #000000;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";


                    }
                    else if (RowAl == 2)
                    {
                        RowAl = 1;
                        ContentLeft = "<td align='Left' style='background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style='background-color: #FFFFFF;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";


                    }


                    sb.Append("<tr>");
                    RolNo++;

                    sb.Append(ContentRight + RolNo + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["Datev"].ToString() + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["VoucherNo"].ToString() + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["Refernce"].ToString() + "</td>");

                    sb.Append(ContentRight + string.Format("{0:0.00}", decimal.Parse(_dtEx.Rows[i]["DebitAmt"].ToString())) + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["CreditAmt"].ToString()).ToString("#,#0.00")) + "</td>");

                    DebitAmt += decimal.Parse(_dtEx.Rows[i]["DebitAmt"].ToString());
                    CreditAmt += decimal.Parse(_dtEx.Rows[i]["CreditAmt"].ToString());
                    BlanceDebitAmt = decimal.Parse(_dtEx.Rows[i]["DebitAmt"].ToString());
                    BlanceCreditAmt = decimal.Parse(_dtEx.Rows[i]["CreditAmt"].ToString());

                    if (BlanceDebitAmt != 0)
                        BalancT += BlanceDebitAmt;
                    if (BlanceCreditAmt != 0)
                        BalancT -= BlanceCreditAmt;

                    sb.Append(ContentRight + (decimal.Parse((BalancT).ToString()).ToString("#,#0.00")) + "</td>");


                    sb.Append("</tr>");
                }
                sb.Append("</tr>");
                sb.Append("<td colspan='" + (5 - RowIndex) + "' align='Center' style='color: black; background-color:#B6F2FC; color:#FF0000; border-top: Black thin inset; border-bottom: Black thin inset; border-Right: Black thin inset; border-Left: Black thin inset; inset; font-family: Arial; font-style: normal; font-size:16px;'>Grand Total Amount</td>");
                sb.Append(RightALm1Color + (decimal.Parse(DebitAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (decimal.Parse(CreditAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (DebitAmt - CreditAmt).ToString("#,#0.00") + " </td>");
                sb.Append("</tr>");
                sb.Append("</table>");
                Response.Write(sb.ToString());
                Response.Clear();
                Response.AddHeader("content-disposition", "attachment;filename=StatementofAccountSummary" + System.DateTime.Now + ".xls");
                Response.Charset = "";
                Response.ContentType = "application/vnd.xls";
                System.IO.StringWriter stringWrite = new System.IO.StringWriter();
                System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
                htmlWrite.Write(sb.ToString());
                Response.Write(stringWrite.ToString());
                Response.End();
                #endregion
            }
        }

        public DataTable GetAccountSummary(string FromDate, string ToDate, string Party)
        {
            string fromdate = ""; string ToDatev = "";
            if (FromDate != "")
                fromdate = FromDate;
            if (ToDate != "")
                ToDatev = ToDate;
            string strWhere = "";
            string _Query = " Select distinct CustomerName, " +
                            " isnull((select sum(DebitAmt) from Nav_Acc_Rep_AccountSummary RepDebit where RepDebit.PartyID = Nav_Acc_Rep_AccountSummary.PartyID),0.00) as DebitAmt, " +
                            " isnull((select sum(CreditAmt) from Nav_Acc_Rep_AccountSummary RepDebit where RepDebit.PartyID = Nav_Acc_Rep_AccountSummary.PartyID),0.00) as CreditAmt, " +
                            " ((isnull((select sum(DebitAmt) from Nav_Acc_Rep_AccountSummary RepDebit where RepDebit.PartyID = Nav_Acc_Rep_AccountSummary.PartyID), 0.00)) - " +
                            " (isnull((select sum(CreditAmt) from Nav_Acc_Rep_AccountSummary RepDebit where RepDebit.PartyID = Nav_Acc_Rep_AccountSummary.PartyID), 0.00))) as BalanceAmt " +
                            " from Nav_Acc_Rep_AccountSummary";
            if (Party != "")
                if (strWhere == "")
                    strWhere += _Query + " where PartyID=" + Party;
                else
                    strWhere += " AND PartyID=" + Party;

          
            if (FromDate != "" && ToDate != "")
                if (strWhere == "")
                    strWhere = _Query + " where VoucherDate BETWEEN '" + fromdate + "' AND '" + ToDate + "' ";
                else
                    strWhere += " AND VoucherDate BETWEEN '" + fromdate + "' AND '" + ToDate + "' ";




            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere, "");


        }


        public DataTable GetAccountSummarydls(string FromDate, string ToDate, string PartyID)
        {
            string fromdate = ""; string ToDatev = "";
            if (FromDate != "")
                fromdate = (DateTime.Parse(FromDate)).ToString("MM/dd/yyyy");
            if (ToDate!= "")
                ToDatev = (DateTime.Parse(ToDate)).ToString("MM/dd/yyyy");

            string strWhere = "";
            string _Query = "  select convert(varchar,VoucherDate, 103) as Datev,* from Nav_Acc_Rep_AccountSummary";
            if (PartyID != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PartyID=" + PartyID;
                else
                    strWhere += " AND PartyID=" + PartyID;

           

            if (FromDate != "" && FromDate != "")
                if (strWhere == "")
                    strWhere = _Query + " where VoucherDate BETWEEN '" + fromdate + "' AND '" + ToDatev + "' ";
                else
                    strWhere += " AND VoucherDate BETWEEN '" + fromdate + "' AND '" + ToDatev + "' ";

            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere + " order by VoucherDate asc", "");


        }



        public DataTable GetAccountSummarydlsOpen(string FromDate, string ToDate, string PartyID)
        {
            string fromdate = ""; string ToDatev = "";
            if (FromDate != "")
                fromdate = (DateTime.Parse(FromDate)).ToString("MM/dd/yyyy");
            if (ToDate != "")
                ToDatev = (DateTime.Parse(ToDate)).ToString("MM/dd/yyyy");

            string _Query = " select 'OPENING BALANCE' as opening, isnull(sum(DebitAmt),0) as DebitAmt, isnull(sum(CreditAmt),0) as  CreditAmt,  isnull(sum(DebitAmt),0) - isnull(sum(CreditAmt),0)  as Balance  from Nav_Acc_Rep_AccountSummary  where VoucherDate <= '" + fromdate + "'";
            return Manag.GetViewData(_Query, "");

        }

        public DataTable GetFinancialData()
        {
            string _query = "select top(1) Id,FinYears, convert(varchar, FDate, 101 ) as FDate,convert(varchar, TDate, 101) as TDate from Nva_FinancialYear order by Id desc";
            return Manag.GetViewData(_query, "");

        }

        #endregion

        public ActionResult AccountReceivableExcelReport(string PartyID, string ToDate, string ReportType,Boolean chkForeignCurrency, Boolean chkCreditDays, Boolean ChkCreditAmount, Boolean chkSalesPic)
        {
            AccountReceivable(PartyID,ToDate, ReportType, chkForeignCurrency, chkCreditDays, ChkCreditAmount, chkSalesPic);
            return View();

        }

        
        public void AccountReceivable(string PartyID, string ToDate, string ReportType, Boolean chkForeignCurrency, Boolean chkCreditDays, Boolean ChkCreditAmount, bool chkSalesPic)
        {

            string CompanyName = "";
            //DataTable _dtC = RegMang.GetViewData("select * from Nav_TblCompanyDetails", "");
            //if (_dtC.Rows.Count > 0)
            //    CompanyName = _dtC.Rows[0]["CompanyName"].ToString();

            string LeftAll = "<td align='center' style='background-color:#04208c; color:#ffffff; font-family:Arial; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; font-style: normal; font-weight: bold; font-size:13px;'>";
            string RightALm1Color = "<td align='Right' style='color: blue; background-color:#B6F2FC;  color:#FF0000; border-top: Black thin inset; border-bottom: Black thin inset; border-Right: Black thin inset; border-Left: Black thin inset; inset; font-family: Arial; font-style: normal;  font-size:16px;'>";
          
            if (ReportType == "2")
            {
                int RowCount = 0;
                int RowTowColumn = 0;
                if (chkSalesPic == true)
                {
                    RowCount += 1; RowTowColumn += 1;
                }
                if (chkForeignCurrency == true)
                    RowCount += 3;
                if (chkCreditDays == true)
                    RowCount += 1;
                if (ChkCreditAmount == true)
                    RowCount += 1;
               

                #region Details  
                sb.Append("<table>");
                sb.Append("<tr></tr><tr></tr>");

                sb.Append("<td colspan='" + (15 + RowCount) + "' align='left' style='color:#04208c;  border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family: Arial; font-style: normal; font-weight: bold; font-size:16px;'>ACCOUNT RECEIVABLE DETAILS</TD></TR>");
                sb.Append("<tr>");
                sb.Append("<td colspan='" + (12 + RowTowColumn) + "' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'></td>");

                sb.Append("<td colspan='3' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>Local Currency</td>");
                if (chkForeignCurrency == true)
                    sb.Append("<td colspan='3' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>Foreign  Currency</td>");
                if (ChkCreditAmount == true)
                    sb.Append("<td colspan='2' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>Credit Info</td>");
              
                    sb.Append("<td align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>JobNo</td>");
                sb.Append("</tr>");
                sb.Append("<tr>");
                sb.Append(LeftAll + " SL No</td>");
                sb.Append(LeftAll + " Location   </td>");
                sb.Append(LeftAll + " Party Name  </td>");
                sb.Append(LeftAll + " Effect   </td>");
                if (chkSalesPic == true)
                    sb.Append(LeftAll + "Sales person </td>");

                sb.Append(LeftAll + " Description</td>");
                sb.Append(LeftAll + "</td>");
                sb.Append(LeftAll + " BLNo</td>");
                sb.Append(LeftAll + " ContinerNo</td>");
                sb.Append(LeftAll + " Invoice No</td>");
                sb.Append(LeftAll + " Invoice Date </td>");

                sb.Append(LeftAll + " Outstanding Days</td>");
                sb.Append(LeftAll + " Invoice Currency</td>");
                sb.Append(LeftAll + " Invoice Amount</td>");
                sb.Append(LeftAll + " Settled Amount</td>");
                sb.Append(LeftAll + " Balance Amount</td>");
                if (chkForeignCurrency == true)
                {
                    sb.Append(LeftAll + " Invoice Amount</td>");
                    sb.Append(LeftAll + " Settled Amount</td>");
                    sb.Append(LeftAll + " Balance Amount</td>");
                }
                if (chkCreditDays == true)
                    sb.Append(LeftAll + " Credit days</td>");
                if (ChkCreditAmount == true)
                    sb.Append(LeftAll + " Credit Amount</td>");
              

                sb.Append("</tr>");

                DataTable _dtEx = GetAccountPayableDetails(ToDate,PartyID);
                string ContentLeft = "";
                string ContentRight = "";
                int RowAl = 1;
                int RolNo = 0;
                decimal DebitAmt = 0;
                decimal CreditAmt = 0;
                decimal FrDebitAmt = 0;
                decimal FRCreditAmt = 0;
                decimal BlanceDebitAmt = 0;
                decimal BlanceCreditAmt = 0;
                decimal BalancT = 0;

                decimal FBlanceDebitAmt = 0;
                decimal FBlanceCreditAmt = 0;
                decimal FBalancT = 0;
                for (int i = 0; i < _dtEx.Rows.Count; i++)
                {
                    if (RowAl == 1)
                    {
                        RowAl = 2;
                        ContentLeft = "<td align='Left' style= ' background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style= ' background-color: #FFFFFF;  color: #000000;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                    }
                    else if (RowAl == 2)
                    {
                        RowAl = 1;
                        ContentLeft = "<td align='Left' style='background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style='background-color: #FFFFFF;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";

                    }


                    sb.Append("<tr>");
                    RolNo++;

                    sb.Append(ContentRight + RolNo + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["Location"].ToString() + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["PartyName"].ToString() + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["Effect"].ToString() + "</td>");
                    if (chkSalesPic == true)
                        sb.Append(ContentLeft + _dtEx.Rows[i]["SalesPIC"].ToString() + "</td>");
                    //sb.Append(ContentLeft + _dtEx.Rows[i]["JobNo"].ToString() + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");
                    sb.Append(ContentLeft + "" + "</td>");

                    sb.Append(ContentLeft + _dtEx.Rows[i]["InvoiceNo"].ToString() + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["invDatev"].ToString() + "</td>");
                    sb.Append(ContentRight + _dtEx.Rows[i]["OutStandingDay"].ToString() + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["Currency"].ToString() + "</td>");

                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["InoiceAmount"].ToString()).ToString("#,#0.00")) + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["SettledAmount"].ToString()).ToString("#,#0.00")) + "</td>");
                    BlanceDebitAmt = decimal.Parse(_dtEx.Rows[i]["InoiceAmount"].ToString());
                    BlanceCreditAmt = decimal.Parse(_dtEx.Rows[i]["SettledAmount"].ToString());

                    if (BlanceDebitAmt != 0)
                        BalancT = BlanceDebitAmt;
                    if (BlanceCreditAmt != 0)
                        BalancT = BlanceCreditAmt;

                    //sb.Append(ContentRight + (decimal.Parse((BalancT).ToString()).ToString("#,#0.00")) + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["BalanceAmt"].ToString()).ToString("#,#0.00")) + "</td>");

                    DebitAmt += decimal.Parse(_dtEx.Rows[i]["InoiceAmount"].ToString());
                    CreditAmt += decimal.Parse(_dtEx.Rows[i]["SettledAmount"].ToString());

                    if (chkForeignCurrency == true)
                    {

                        sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["ForeignInoiceAmount"].ToString()).ToString("#,#0.00")) + "</td>");
                        sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["ForeignSettledAmount"].ToString()).ToString("#,#0.00")) + "</td>");

                        FBlanceDebitAmt = decimal.Parse(_dtEx.Rows[i]["ForeignInoiceAmount"].ToString());
                        FBlanceCreditAmt = decimal.Parse(_dtEx.Rows[i]["ForeignSettledAmount"].ToString());

                        if (FBlanceDebitAmt != 0)
                            FBalancT += FBlanceDebitAmt;
                        if (FBlanceCreditAmt != 0)
                            FBalancT -= FBlanceCreditAmt;

                        //  sb.Append(ContentRight + (decimal.Parse((FBalancT).ToString()).ToString("#,#0.00")) + "</td>");
                        sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["BalanceAmt"].ToString()).ToString("#,#0.00")) + "</td>");
                    }



                    FrDebitAmt += decimal.Parse(_dtEx.Rows[i]["ForeignInoiceAmount"].ToString());
                    FRCreditAmt += decimal.Parse(_dtEx.Rows[i]["ForeignSettledAmount"].ToString());
                    if (chkCreditDays == true)
                        sb.Append(ContentRight + _dtEx.Rows[i]["CreditDays"].ToString() + "</td>");
                    if (ChkCreditAmount == true)
                        sb.Append(ContentRight + _dtEx.Rows[i]["CreditAmt"].ToString() + "</td>");
                    


                    sb.Append("</tr>");
                }
                sb.Append("<tr>");
                sb.Append("<td colspan='" + (12 + RowTowColumn) + "' align='right' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>Grand Total Amount</td>");

                sb.Append(RightALm1Color + (decimal.Parse(DebitAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (decimal.Parse(CreditAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (DebitAmt - CreditAmt).ToString("#,#0.00") + " </td>");
                if (chkForeignCurrency == true)
                {
                    sb.Append(RightALm1Color + (decimal.Parse(FrDebitAmt.ToString()).ToString("#,#0.00")) + " </td>");
                    sb.Append(RightALm1Color + (decimal.Parse(FRCreditAmt.ToString()).ToString("#,#0.00")) + " </td>");
                    sb.Append(RightALm1Color + (FrDebitAmt - FRCreditAmt).ToString("#,#0.00") + " </td>");
                }
                if (chkCreditDays == true)
                    sb.Append("<td colspan='2' align='center' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'></td>");
               
                sb.Append("</tr>");

                sb.Append("</table>");
                Response.Write(sb.ToString());
                Response.Clear();
                Response.AddHeader("content-disposition", "attachment;filename=AccountReceviableDetails" + System.DateTime.Now + ".xls");
                Response.Charset = "";
                Response.ContentType = "application/vnd.xls";
                System.IO.StringWriter stringWrite = new System.IO.StringWriter();
                System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
                htmlWrite.Write(sb.ToString());
                Response.Write(stringWrite.ToString());
                Response.End();
                #endregion

            }
            if (ReportType == "1")
            {
                #region  Summary 

                int RowColumn = 0;
                if (chkSalesPic == false)
                    RowColumn += 1;
                if (chkCreditDays == false)
                    RowColumn += 1;
                if (ChkCreditAmount == false)
                    RowColumn += 1;

                sb.Append("<table>");
                sb.Append("<tr></tr><tr></tr>");
                sb.Append("<td colspan='" + (8 - RowColumn) + "' align ='left' style = 'color:#04208c;  border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family: Arial; font-style: normal; font-weight: bold; font-size:16px;'>ACCOUNT RECEIVABLE SUMMARY</TD></TR>");
                sb.Append("<tr>");
                sb.Append(LeftAll + " SL No</td>");
                sb.Append(LeftAll + " Party Name </td>");

                if (chkSalesPic == true)
                    sb.Append(LeftAll + " Sales PIC </td>");
                if (chkCreditDays == true)
                    sb.Append(LeftAll + " Credit Days   </td>");
                if (ChkCreditAmount == true)
                    sb.Append(LeftAll + " Credit Amount </td>");

                sb.Append(LeftAll + " Total Debit</td>");
                sb.Append(LeftAll + " Total Credit</td>");
                sb.Append(LeftAll + " Net Balance</td>");
                sb.Append("</tr>");

                DataTable _dtEx = GetAccountPayableSummary(ToDate,PartyID);
                string ContentLeft = "";
                string ContentRight = "";
                int RowAl = 1;
                int RolNo = 0;
                decimal DebitAmt = 0;
                decimal CreditAmt = 0;
                decimal BlanceDebitAmt = 0;
                decimal BlanceCreditAmt = 0;
                decimal BalancT = 0;

                for (int i = 0; i < _dtEx.Rows.Count; i++)
                {
                    if (RowAl == 1)
                    {
                        RowAl = 2;
                        ContentLeft = "<td align='Left' style= ' background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style= ' background-color: #FFFFFF;  color: #000000;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                    }
                    else if (RowAl == 2)
                    {
                        RowAl = 1;
                        ContentLeft = "<td align='Left' style='background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                        ContentRight = "<td align='right' style='background-color: #FFFFFF;black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset;  color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                    }

                    sb.Append("<tr>");
                    RolNo++;

                    sb.Append(ContentRight + RolNo + "</td>");
                    sb.Append(ContentLeft + _dtEx.Rows[i]["PartyName"].ToString() + "</td>");
                    if (chkSalesPic == true)
                        sb.Append(ContentLeft + _dtEx.Rows[i]["SalesPIC"].ToString() + "</td>");
                    if (chkCreditDays == true)
                        sb.Append(ContentRight + _dtEx.Rows[i]["CreditDays"].ToString() + "</td>");
                    if (ChkCreditAmount == true)
                        sb.Append(ContentRight + _dtEx.Rows[i]["CreditAmt"].ToString() + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["InvoiceAmt"].ToString()).ToString("#,#0.00")) + "</td>");
                    sb.Append(ContentRight + (decimal.Parse(_dtEx.Rows[i]["SettledAmount"].ToString()).ToString("#,#0.00")) + "</td>");
                    DebitAmt += decimal.Parse(_dtEx.Rows[i]["InvoiceAmt"].ToString());
                    CreditAmt += decimal.Parse(_dtEx.Rows[i]["SettledAmount"].ToString());

                    BlanceDebitAmt = decimal.Parse(_dtEx.Rows[i]["InvoiceAmt"].ToString());
                    BlanceCreditAmt = decimal.Parse(_dtEx.Rows[i]["SettledAmount"].ToString());

                    if (BlanceDebitAmt != 0)
                        BalancT = BlanceDebitAmt;
                    if (BlanceCreditAmt != 0)
                        BalancT = BlanceCreditAmt;

                    //sb.Append(ContentRight + (decimal.Parse((BalancT).ToString()).ToString("#,#0.00")) + "</td>");
                    sb.Append(ContentRight + (decimal.Parse((BlanceDebitAmt - BlanceCreditAmt).ToString()).ToString("#,#0.00")) + "</td>");


                    sb.Append("</tr>");
                }
                sb.Append("<tr>");

                sb.Append("<td colspan='" + (5 - RowColumn) + "' align='right' style= 'color: black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; font-family:Arial; font-style: normal; font-weight: bold; font-size:18px;'>Grand Total Amount</td>");
                sb.Append(RightALm1Color + (decimal.Parse(DebitAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (decimal.Parse(CreditAmt.ToString()).ToString("#,#0.00")) + " </td>");
                sb.Append(RightALm1Color + (DebitAmt - CreditAmt).ToString("#,#0.00") + " </td>");
                sb.Append("</tr>");

                sb.Append("</table>");
                Response.Write(sb.ToString());
                Response.Clear();
                Response.AddHeader("content-disposition", "attachment;filename=AccountReceviableSummary" + System.DateTime.Now + ".xls");
                Response.Charset = "";
                Response.ContentType = "application/vnd.xls";
                System.IO.StringWriter stringWrite = new System.IO.StringWriter();
                System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
                htmlWrite.Write(sb.ToString());
                Response.Write(stringWrite.ToString());
                Response.End();
                #endregion
            }
        }


        public DataTable GetAccountPayableDetails(string Date, string Party)
        {
            string fromdate = ""; string ToDate = "";
            if (Date != "")
                fromdate = (DateTime.Parse(Date)).ToString("MM/dd/yyyy");

            ToDate = ((System.DateTime.Now.Date)).ToString("MM/dd/yyyy");

            string strWhere = "";
            string _Query = " select * from NVO_V_AccountReceivableNew";
            if (Party != "?")
                if (strWhere == "")
                    strWhere += _Query + " where PartyID=" + Party;
                else
                    strWhere += " AND PartyID=" + Party;


            if (Date != "")
                if (strWhere == "")
                    strWhere += _Query + " where InvDate  <= '" + fromdate + "'";
                else
                    strWhere += " AND InvDate  <= '" + fromdate + "'";

            if (strWhere == "")
                strWhere = _Query;

         
            return Manag.GetViewData(strWhere + " order by Invdate desc", "");


        }

        public DataTable GetAccountPayableSummary(string Date, string Party)
        {
            string fromdate = ""; string ToDate = "";
            if (Date != "")
                fromdate = (DateTime.Parse(Date)).ToString("MM/dd/yyyy");

            ToDate = ((System.DateTime.Now.Date)).ToString("MM/dd/yyyy");

            string strWhere = "";
            string _Query = " select PartyName,SalesPIC,CreditDays,CreditAmt,sum(InoiceAmount) as InvoiceAmt, sum(SettledAmount) as SettledAmount " +
                            " from NVO_V_AccountReceivableNew";
            if (Party != "")
                if (strWhere == "?")
                    strWhere += _Query + " where PartyID=" + Party;
                else
                    strWhere += " AND PartyID=" + Party;

           

            if (Date != "")
                if (strWhere == "")
                    strWhere += _Query + " where InvDate  <= '" + fromdate + "'";
                else
                    strWhere += " AND InvDate  <= '" + fromdate + "'";
            if (strWhere == "")
                strWhere = _Query;

            return Manag.GetViewData(strWhere + "  group by PartyName,SalesPIC,CreditDays,CreditAmt", "");


        }

    }
}