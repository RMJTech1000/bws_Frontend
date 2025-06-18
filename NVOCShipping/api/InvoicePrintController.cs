using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using DataManager;
using DataTier;

namespace NVOCShipping.api
{
    public class InvoicePrintController : ApiController
    {
        [ActionName("InvDocumentNo")]
        public List<MyInvoicePrint> InvDocumentNo(MyInvoicePrint Data)
        {

            InvoicePrintManager AccMange = new InvoicePrintManager();
            List<MyInvoicePrint> st = AccMange.InvoiceDocumentNo(Data);
            return st;
        }

        [ActionName("SearchInvDocumentNo")]
        public List<MyInvoicePrint> SearchInvDocumentNo(MyInvoicePrint Data)
        {

            InvoicePrintManager AccMange = new InvoicePrintManager();
            List<MyInvoicePrint> st = AccMange.SearchInvoiceDocumentNo(Data);
            return st;
        }
    }
}