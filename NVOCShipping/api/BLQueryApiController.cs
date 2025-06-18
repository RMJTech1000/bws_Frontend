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
    public class BLQueryApiController : ApiController
    {
        [ActionName("BLNumberByAgency")]
        public List<BLQueryData> BLNumberByAgency(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.BLNubmerAgencyWise(Data);
            return st;
        }

        [ActionName("CntrNoDropDown")]
        public List<BLQueryData> CntrNoDropDown(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.CntrDetailsView(Data);
            return st;
        }

        [ActionName("InvoiceNoDropDown")]
        public List<BLQueryData> InvoiceNoDropDown(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.InvoiceNoList(Data);
            return st;
        }

        [ActionName("BLQueryView")]
        public List<BLQueryData> BLQueryView(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.BLQueryView(Data);
            return st;
        }

        [ActionName("BLInvoiceDtls")]
        public List<BLQueryData> BLInvoiceDtls(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.BLInvoiceDtls(Data);
            return st;
        }

        [ActionName("BLRecepDtls")]
        public List<BLQueryData> BRLRecepDtls(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.BLRecpDtls(Data);
            return st;
        }

        [ActionName("CustomerInvNoList")]
        public List<BLQueryData> CustomerInvNoList(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.CustomerInvoiceListDetails(Data);
            return st;
        }
        [ActionName("BLRecepNoList")]
        public List<BLQueryData> BLRecepNoList(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.ReceiptNoList(Data);
            return st;
        }
        [ActionName("BLCrInvoiceDtls")]
        public List<BLQueryData> BLCrInvoiceDtls(BLQueryData Data)
        {
            BLQueryManager cm = new BLQueryManager();
            List<BLQueryData> st = cm.BLCrInvoiceDtls(Data);
            return st;
        }
        
    }
}
