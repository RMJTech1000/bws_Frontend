using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Configuration;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using DataManager;
using DataTier;

namespace NVOCShipping.api
{
    public class FinanceReportApiController : ApiController
    {
        #region FinanceReport



        [ActionName("OutstandingPartyReport")]
        public List<FinRoot> OutstandingPartyReport(FinRoot data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<FinRoot> st = cm.OutStandingView(data);
            return st;
        }

        [ActionName("OutstandingReport")]
        public List<MyFinanceReportData> OutstandingReport(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.OutStandingReport(Data);
            return st;
        }

        [ActionName("StatmentOfAccounts")]
        public List<MyFinanceReportData> StatmentOfAccounts(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.StatementOfAccReport(Data);
            return st;
        }

        [ActionName("CustomersDropDown")]
        public List<MyFinanceReportData> CustomersDropDown(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomersDetails(Data);
            return st;
        }

        [ActionName("CustomersSearch")]
        public List<MyFinanceReportData> CustomersSerach(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerWiseDetails(Data);
            return st;
        }

        [ActionName("CustomersFullDetails")]
        public List<MyFinanceReportData> CustomersFullDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerViewDetails(Data);
            return st;
        }
        [ActionName("CustomerInvDetails")]
        public List<MyFinanceReportData> CustomerInvDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerInvoiceDetails(Data);
            return st;
        }

        [ActionName("CustomerRecpDetails")]
        public List<MyFinanceReportData> CustomerRecpDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerReceiptDetails(Data);
            return st;
        }

        [ActionName("CustomerInvNoList")]
        public List<MyFinanceReportData> CustomerInvNoList(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerInvoiceListDetails(Data);
            return st;
        }

        [ActionName("CustomerInvNoChangeList")]
        public List<MyFinanceReportData> CustomerInvNoChangeList(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerInvoiceListChangeDetails(Data);
            return st;
        }

        [ActionName("CustomerRecNoList")]
        public List<MyFinanceReportData> CustomerRecNoList(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerReceiptListDetails(Data);
            return st;
        }


        //Acc Receivable

        [ActionName("AccReceivableRecords")]
        public List<MyFinanceReportData> AccReceivableRecords(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccountsReceivableReport(Data);
            return st;
        }
        [ActionName("AccReceivableFullDetails")]
        public List<MyFinanceReportData> AccReceivableFullDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccReceivableViewDetails(Data);
            return st;
        }

        [ActionName("AccRecInvDetails")]
        public List<MyFinanceReportData> AccRecInvDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccRecInvoiceDetails(Data);
            return st;
        }

        [ActionName("AccRecRecpDetails")]
        public List<MyFinanceReportData> AccRecRecpDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.ReceivableReceiptDetails(Data);
            return st;
        }

        [ActionName("AccRecInvNoList")]
        public List<MyFinanceReportData> AccRecInvNoList(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccRecInvoiceDetails(Data);
            return st;
        }

        [ActionName("AccRecRecNoList")]
        public List<MyFinanceReportData> AccRecRecNoList(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.ReceivableReceiptListDetails(Data);
            return st;
        }

        [ActionName("ReceivableRecpDetails")]
        public List<MyFinanceReportData> ReceivableRecpDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.ReceivableReceiptDetails(Data);
            return st;
        }

        [ActionName("CompanyDetails")]
        public List<MyFinanceReportData> CompanyDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.HeaderCompanyDetals(Data);
            return st;
        }

        [ActionName("ReceivableCustomersDropDown")]
        public List<MyFinanceReportData> ReceivableCustomersDropDown(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.ReceivableCustomersDetails(Data);
            return st;
        }

        [ActionName("ReceivableCustomersSearch")]
        public List<MyFinanceReportData> ReceivableCustomersSearch(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.ReceivableCustomerViewDetails(Data);
            return st;
        }

        //Acc Payable

        [ActionName("AccPayableRecords")]
        public List<MyFinanceReportData> AccPayableRecords(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccountsPayableReport(Data);
            return st;
        }
        [ActionName("AccPayableFullDetails")]
        public List<MyFinanceReportData> AccPayableFullDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccPayableViewDetails(Data);
            return st;
        }

        [ActionName("AccPayInvDetails")]
        public List<MyFinanceReportData> AccPayInvDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccPayInvoiceDetails(Data);
            return st;
        }

        [ActionName("AccPayRecpDetails")]
        public List<MyFinanceReportData> AccPayRecpDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccPayReceiptDetails(Data);
            return st;
        }

        [ActionName("AccPayInvNoList")]
        public List<MyFinanceReportData> AccPayInvNoList(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccPayInvoiceDetails(Data);
            return st;
        }

        [ActionName("PayableCustomersDropDown")]
        public List<MyFinanceReportData> PayableCustomersDropDown(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.PayableCustomersDetails(Data);
            return st;
        }

        [ActionName("PayableCustomersSearch")]
        public List<MyFinanceReportData> PayableCustomersSearch(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.PayableCustomerViewDetails(Data);
            return st;
        }
        #endregion

        #region OutStandingCustomer
        [ActionName("OutStandingCustomer")]
        public List<MyFinanceReportData> OutStandingCustomer(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.CustomerOutStandingDtls(Data);
            return st;
        }


        [ActionName("AccRecInvTotalDetails")]
        public List<MyFinanceReportData> AccRecInvTotalDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccRecInvoiceTotalDetails(Data);
            return st;
        }

        [ActionName("AccRecInvOutStandingDetails")]
        public List<MyFinanceReportData> AccRecInvOutStandingDetails(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.AccRecInvoiceOutStandingDetails(Data);
            return st;
        }

        [ActionName("Account_Recevable_LastsixmonthInvoice")]
        public List<MyFinanceReportData> Account_Recevable_LastsixmonthInvoice(MyFinanceReportData Data)
        {
            FinanceReportManager cm = new FinanceReportManager();
            List<MyFinanceReportData> st = cm.Chart_Account_Recevable_LastsixmonthInvoiceAll(Data);
            return st;
        }

        #endregion
    }
}
