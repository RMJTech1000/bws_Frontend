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
    public class VendorApiController : ApiController
    {
        [ActionName("InvVendorCustomerMaster")]
        public List<MyInvoiceVendor> InvVendorCustomerMaster()
        {
            InvoiceVendorMaster invMange = new InvoiceVendorMaster();
            List<MyInvoiceVendor> st = invMange.InvVendorCustomerMaster();
            return st;
        }

        [ActionName("VoyageDetails")]
        public List<MyInvoiceVendor> VoyageDetails(MyInvoiceVendor Data)
        {
            InvoiceVendorMaster invMange = new InvoiceVendorMaster();
            List<MyInvoiceVendor> st = invMange.VoyageMaster(Data);
            return st;
        }

        [ActionName("VendorInvoiceInsert")]
        public List<MyInvoiceVendorInsert> VendorInvoiceInsert(MyInvoiceVendorInsert Data)
        {
            InvoiceVendorMaster invMange = new InvoiceVendorMaster();
            List<MyInvoiceVendorInsert> st = invMange.InsertVendorInvoiceMaster(Data);
            return st;
        }


        [ActionName("VendorInvoiceSearch")]
        public List<MyInvoiceVendorSearch> VendorInvoiceSearch(MyInvoiceVendorSearch Data)
        {
            InvoiceVendorMaster invMange = new InvoiceVendorMaster();
            List<MyInvoiceVendorSearch> st = invMange.VendorInvoiceSearch(Data);
            return st;
        }

        [ActionName("VendorInvoiceExisting")]
        public List<MyInvoiceVendorSearch> VendorInvoiceExisting(MyInvoiceVendorSearch Data)
        {
            InvoiceVendorMaster invMange = new InvoiceVendorMaster();
            List<MyInvoiceVendorSearch> st = invMange.VendorInvoiceExisting(Data);
            return st;
        }

        [ActionName("VendorInvoiceExistingDtls")]
        public List<MyInvoiceVendorSearch> VendorInvoiceExistingDtls(MyInvoiceVendorSearch Data)
        {
            InvoiceVendorMaster invMange = new InvoiceVendorMaster();
            List<MyInvoiceVendorSearch> st = invMange.VendorInvoiceExistingDtls(Data);
            return st;
        }

    }
}