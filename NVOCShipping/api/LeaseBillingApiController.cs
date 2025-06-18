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
    public class LeaseBillingApiController : ApiController
    {
        [ActionName("LeaseBillMonitringView")]
        public List<MyLeaseBilling> LeaseBillMonitringView(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseMonitoringBillingView(Data);
            return st;
        }

        [ActionName("LeaseBillMonitringExisting")]
        public List<MyLeaseBilling> LeaseBillMonitringExisting(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseMonitoringBillingExisting(Data);
            return st;
        }


        [ActionName("LeaseDetailsGirdExisting")]
        public List<MyLeaseBilling> LeaseDetailsGirdExisting(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseDetailsBillGridExisting(Data);
            return st;
        }


        [ActionName("LeaseDetailsGirdExistingPickRemind")]
        public List<MyLeaseBilling> LeaseDetailsGirdExistingPickRemind(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseDetailsBillGridExistingPickRemainer(Data);
            return st;
        }

        [ActionName("LeaseDetailsBillingRentGird")]
        public List<MyLeaseBilling> LeaseDetailsBillingRentGird(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseDetailsBillRentGrid(Data);
            return st;
        }

        [ActionName("ChargeCodeValue")]
        public List<MyCommonAccess> ChargeCodeValue()
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyCommonAccess> st = invMange.ChargeCodeMasterBind();
            return st;
        }


        [ActionName("LeaseDetailsRentGirdRetCharge")]
        public List<MyLeaseBilling> LeaseDetailsRentGirdRetCharge(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseDetailsRentGridRetCharges(Data);
            return st;
        }


        [ActionName("LeaseBillingInvoice")]
        public List<MyInvoice> LeaseBillingInvoice(MyInvoice Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyInvoice> st = invMange.LeaseBillingInvoiceInsertMaster(Data);
            return st;
        }


        [ActionName("LeaseDetailsBillingInvoice")]
        public List<MyLeaseBilling> LeaseDetailsBillingInvoice(MyLeaseBilling Data)
        {
            LeaseBillingManager invMange = new LeaseBillingManager();
            List<MyLeaseBilling> st = invMange.LeaseBillingInvoice(Data);
            return st;
        }
    }
}