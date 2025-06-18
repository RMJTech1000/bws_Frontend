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
    public class StandalonInvoiceController : ApiController
    {

        [ActionName("InvoiceInsert")]
        public List<MyInvoice> InvoiceInsert(MyInvoice Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MyInvoice> st = invMange.InvoiceInsertMaster(Data);
            return st;
        }

        [ActionName("InvoiceInsertCR")]
        public List<MyInvoice> InvoiceInsertCR(MyInvoice Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MyInvoice> st = invMange.InvoiceInsertMasterCR(Data);
            return st;
        }

        

        [ActionName("ExistingInvoiceDetails")]
        public List<MyInvoice> ExistingInvoiceDetails(MyInvoice Data)
        {
            InvoiceManager invMange = new InvoiceManager();
            List<MyInvoice> st = invMange.ExistingInvoiceView(Data);
            return st;
        }

        [ActionName("BLAgentLocation")]
        public List<MyInvoiceBL> BLAgentLocation(MyInvoiceBL Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MyInvoiceBL> st = invMange.AgentLocationValue(Data);
            return st;
        }

        [ActionName("ExistingInvoiceValue")]
        public List<MyInvoice> ExistingInvoicevalue(MyInvoice Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MyInvoice> st = invMange.ExistingInvoice(Data);
            return st;
        }

        [ActionName("BLInvoiceNo")]
        public List<MyInvoiceBL> BLInvoiceNo(MyInvoiceBL Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MyInvoiceBL> st = invMange.BLInvoiceNo(Data);
            return st;
        }
        [ActionName("InvBLChargesExistingValueCR")]
        public List<MYTariffInv> InvBLChargesExistingValueCR(MYTariffInv Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MYTariffInv> st = invMange.InvBLChargesExistingValueCR(Data);
            return st;
        }


        [ActionName("InvExistingTaxDetails")]
        public List<MyInvoiceBL> InvExistingTaxDetails(MyInvoiceBL Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MyInvoiceBL> st = invMange.InvoiceBLTaxChargeDetails(Data);
            return st;
        }
        [ActionName("InvBLChargesExistingValueBind")]
        public List<MYTariffInv> InvBLChargesExistingValueBind(MYTariffInv Data)
        {
            StandalonInvoiceManager invMange = new StandalonInvoiceManager();
            List<MYTariffInv> st = invMange.InvBLChargesExistingBind(Data);
            return st;
        }

    }
}
