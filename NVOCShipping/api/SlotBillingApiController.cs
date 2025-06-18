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
    public class SlotBillingApiController : ApiController
    {
        #region Slot Billing

        [ActionName("SlotBillingByVesVoy")]
        public List<MySlotBill> SlotBillingByVesVoy(MySlotBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MySlotBill> st = slot.SlotBillingByVesVoy(Data);
            return st;
        }

        [ActionName("SlotBillAttachments")]
        public List<MySlotBill> SlotBillAttachments(MySlotBill Data)
        {
            SlotBillingManager cm = new SlotBillingManager();
            List<MySlotBill> st = cm.InsertSlotBillAttachments(Data);
            return st;
        }

        [ActionName("SlotBillingInsert")]
        public List<MySlotBill> SlotBillingInsert(MySlotBill Data)
        {
            SlotBillingManager cm = new SlotBillingManager();
            List<MySlotBill> st = cm.VendorInvoiceSlotBillingInsert(Data);
            return st;
        }
        [ActionName("ViewSlotBilling")]
        public List<MySlotBill> SlotBillingView(MySlotBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MySlotBill> st = slot.SlotBillingView(Data);
            return st;
        }
        #endregion
        #region Port Billing

        [ActionName("PortBillingBySearch")]
        public List<MyPortBill> PortBillingBySearch(MyPortBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MyPortBill> st = slot.PortBillingBySearch(Data);
            return st;
        }
        [ActionName("PortBillAttachments")]
        public List<MyPortBill> PortBillAttachments(MyPortBill Data)
        {
            SlotBillingManager cm = new SlotBillingManager();
            List<MyPortBill> st = cm.InsertPortBillAttachments(Data);
            return st;
        }
        [ActionName("PortBillingInsert")]
        public List<MyPortBill> PortBillingInsert(MyPortBill Data)
        {
            SlotBillingManager cm = new SlotBillingManager();
            List<MyPortBill> st = cm.VendorInvoicePortBillingInsert(Data);
            return st;
        }
        [ActionName("ViewPortBilling")]
        public List<MyPortBill> PortBillingView(MyPortBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MyPortBill> st = slot.PortBillingView(Data);
            return st;
        }
        [ActionName("ViewUploadPortBilling")]
        public List<MyPortBill> ViewUploadPortBilling(MyPortBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MyPortBill> st = slot.ViewUploadPortBilling(Data);
            return st;
        }
        #endregion

        #region Vendor Invoice Approval
        [ActionName("BindVendorInvNoList")]
        public List<MyPortBill> ViewBindVendorInvNoList(MyPortBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MyPortBill> st = slot.ViewBindVendorInvNoList(Data);
            return st;
        }
     
        [ActionName("VendorInvRecordMaster")]
        public List<MyPortBill> ViewVendorInvRecordsList(MyPortBill Data)
        {
            SlotBillingManager slot = new SlotBillingManager();
            List<MyPortBill> st = slot.ViewVendorInvRecordsList(Data);
            return st;
        }
        [ActionName("UpdateVendorInvoiceApproval")]
        public List<MyPortBill> UpdateVendorInvoiceApproval(MyPortBill Data)
        {
            SlotBillingManager cm = new SlotBillingManager();
            List<MyPortBill> st = cm.UpdateVendorInvoiceApproval(Data);
            return st;
        }
        
        #endregion

    }
}
