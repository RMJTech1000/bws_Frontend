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
   
    public class OnLIneController : ApiController
    {
        
        [HttpPost,  ActionName("OnlineCntrTraking")]
        public List<OnlineTracking> OnlineCntrTraking(OnlineTracking Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<OnlineTracking> st = CompMang.OnlineCntrTraking_List(Data);
            return st;
        }

        [HttpPost, ActionName("OnlineCustomerCntrTraking")]
        public List<OnlineTracking> OnlineCustomerCntrTraking(OnlineTracking Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<OnlineTracking> st = CompMang.OnlineCustomerCntrTraking_List(Data);
            return st;
        }


        [HttpPost, ActionName("OnlineCntr_Mnr_Traking")]
        public List<OnlineCntrDamage> OnlineCntr_Mnr_Traking(OnlineCntrDamage Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<OnlineCntrDamage> st = CompMang.OnlineCntrTraking_Damage_List(Data);
            return st;
        }



        [ActionName("Online_Customer_registration")]
        public List<OnlineCustomer_Reg> Online_Customer_registration(OnlineCustomer_Reg Data)
        {
            OnlineManager Online_Mang = new OnlineManager();
            List<OnlineCustomer_Reg> st = Online_Mang.Insert_Online_Customer_registration(Data);
            return st;
        }

        [HttpPost, ActionName("PortValues")]
        public List<MyCommonAccess> PortValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.PortMaster();
            return st;
        }

        [HttpPost, ActionName("BindMainPorts")]
        public List<MyCommonAccess> BindMainPorts(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BindMainPortsList(Data);
            return st;
        }

        [HttpPost, ActionName("OnlineVesselVoyageTracking")]
        public List<Online_VesselVoyageTrack> OnlineVesselVoyageTracking(Online_VesselVoyageTrack Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<Online_VesselVoyageTrack> st = CompMang.Online_Vessel_voyage_Tranck_List(Data);
            return st;
        }
        //Booking
        [HttpPost, ActionName("Online_RateList_Quotation_List")]
        public List<Online_Rate_Quotation_view> Online_RateList_Quotation_List(Online_Rate_Quotation_view Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<Online_Rate_Quotation_view> st = CompMang.online_Rate_Quotation_List(Data);
            return st;
        }



        [HttpPost, ActionName("Online_RateList_Booking_List")]
        public List<Online_Rate_Quotation_view> Online_RateList_Booking_List(Online_Rate_Quotation_view Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<Online_Rate_Quotation_view> st = CompMang.online_Rate_Booking_List(Data);
            return st;
        }





        [HttpPost, ActionName("Online_Quotation_Rate_List")]
        public List<Online_Quotation_Rates> Online_Quotation_Rate_List(Online_Quotation_Rates Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<Online_Quotation_Rates> st = CompMang.online_Ratesheet_Rate_List(Data);
            return st;
        }


        [HttpPost, ActionName("RRNotificationModeview")]
        public List<MyRatesheet> RRNotificationModeview(MyRatesheet Data)
        {
            OnlineManager cm = new OnlineManager();
            List<MyRatesheet> st = cm.RRNotificationMode(Data.ID.ToString());
            return st;
        }

        [HttpPost, ActionName("RRCntrTypes_Listiew")]
        public List<Online_RatecntrTypes> RRCntrTypes_Listiew(Online_RatecntrTypes Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_RatecntrTypes> st = cm.RRCntrTypes_List(Data);
            return st;
        }



        [HttpPost, ActionName("CustomerMaster")]
        public List<MyCommonAccess> CustomerMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerMaster();
            return st;
        }


        [HttpPost, ActionName("OnlineCustomerMaster")]
        public List<MyCommonAccess> OnlineCustomerMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.OnlioneCustomerMaster(Data);
            return st;
        }


        [HttpPost, ActionName("Online_CustomerAddress_List")]
        public List<Online_CustomerAddress> Online_CustomerAddress_List(Online_CustomerAddress Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<Online_CustomerAddress> st = CompMang.online_Address_List(Data);
            return st;
        }

        [HttpPost, ActionName("CargoTypesMaster")]
        public List<MyCommonAccess> CargoTypesMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CargoTypesMaster();
            return st;
        }

        [HttpPost, ActionName("CROViewRecord")]
        public List<MyCROMaster> CROViewRecord(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROViewValus(Data);
            return st;
        }

        [HttpPost, ActionName("BookingMaster")]
        public List<MyCommonAccess> BookingMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BookingMaster();
            return st;
        }


        [HttpPost, ActionName("BOLCntrExistingViewRecord")]
        public List<Online_BOLCntrNo> BOLCntrExistingViewRecord(Online_BOLCntrNo Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BOLCntrNo> st = CompMang.BOLCntrExistingValus(Data);
            return st;
        }

        [HttpPost, ActionName("ModuleValues")]
        public List<MyCommonAccess> ModuleValues(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.Generalmaster(Data.ID.ToString());
            return st;
        }

        [HttpPost, ActionName("Online_BookingInsert")]
        public List<Online_Booking> Online_BookingInsert(Online_Booking Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_Booking> st = CompMang.Online_BookingInsert(Data);
            return st;
        }


        [HttpPost, ActionName("Online_UserDetails")]
        public List<UserDetails> Online_UserDetails(UserDetails Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<UserDetails> st = CompMang.Online_UserLogin(Data);
            return st;
        }

        [HttpPost, ActionName("Online_DashboardBkgview")]
        public List<DasboardBkgDetails> Online_DashboardBkgview(DasboardBkgDetails Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<DasboardBkgDetails> st = CompMang.Online_DashboardBkgDetails(Data);
            return st;
        }


        [HttpPost, ActionName("Online_ExportBookingView")]
        public List<DasboardBkgDetails> Online_ExportBookingView(DasboardBkgDetails Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<DasboardBkgDetails> st = CompMang.Online_ExportBkgDetails(Data);
            return st;
        }




        [HttpPost, ActionName("Online_BOLBookingSelectViewRecord")]
        public List<Online_BL> Online_BOLBookingSelectViewRecord(Online_BL Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_BL> st = cm.Online_BOLBkgSelectExistingValus(Data);
            return st;
        }

        [HttpPost, ActionName("Online_BOLCntrExistingViewRecord")]
        public List<Online_BOL> Online_BOLCntrExistingViewRecord(Online_BOL Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_BOL> st = cm.Online_BOLCntrExistingValus(Data);
            return st;
        }


        [HttpPost, ActionName("Online_BLInsert")]
        public List<Online_New_BOL> Online_BLInsert(Online_New_BOL Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_New_BOL> st = cm.Online_BOLInsert(Data);
            return st;
        }

        [HttpPost, ActionName("Online_Finance_OutStanding")]
        public List<Online_FinanceOutStanding> Online_Finance_OutStanding(Online_FinanceOutStanding Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_FinanceOutStanding> st = cm.Online_Finance_outStanding(Data);
            return st;
        }

        [HttpPost, ActionName("Online_Finance_InvoiceView")]
        public List<Online_FinanceInvoiceDetails> Online_Finance_InvoiceView(Online_FinanceInvoiceDetails Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_FinanceInvoiceDetails> st = cm.Online_Finance_InvoiceDetails(Data);
            return st;
        }

        [HttpPost, ActionName("Online_Finance_ReceiptView")]
        public List<Online_FinanceInvoiceDetails> Online_Finance_ReceiptView(Online_FinanceInvoiceDetails Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_FinanceInvoiceDetails> st = cm.Online_Finance_ReceiptDetails(Data);
            return st;
        }


        [HttpPost, ActionName("Online_Online_Dasboard_Status")]
        public List<Online_Dasboardview> Online_Online_Dasboard_Status(Online_Dasboardview Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_Dasboardview> st = cm.Online_DasboardBookingStatusDetails(Data);
            return st;
        }


        [HttpPost, ActionName("CustomerAddress")]
        public List<MyCommonAccess> CustomerAddress(MyCommonAccess Data)
        {
            OnlineManager cm = new OnlineManager();
            List<MyCommonAccess> st = cm.CustomerAddress(Data.ID.ToString());
            return st;
        }

        [HttpPost, ActionName("CountryMaster")]
        public List<MyCommonAccess> CountryMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CountryMaster();
            return st;
        }


        [HttpPost, ActionName("stateBind")]
        public List<MyState> stateBind(MyState Data)
        {
            MasterManager cm = new MasterManager();
            List<MyState> st = cm.GetCommonStateMaster(Data);
            return st;
        }

        [HttpPost, ActionName("BindCities")]
        public List<cityDD> BindCities(cityDD Data)
        {
            MasterManager cm = new MasterManager();
            List<cityDD> st = cm.ListCities(Data);
            return st;
        }


        [HttpPost, ActionName("ReceiptInvoiceDetls")]
        public List<MyAccount> ReceiptInvoiceDetls(MyAccount Data)
        {

            OnlineManager AccMange = new OnlineManager();
            List<MyAccount> st = AccMange.InvoiceDetailsMaster(Data);
            return st;
        }

        [HttpPost, ActionName("Online_PaymentInsert")]
        public List<Online_Payment_Confirm> Online_PaymentInsert(Online_Payment_Confirm Data)
        {

            OnlineManager AccMange = new OnlineManager();
            List<Online_Payment_Confirm> st = AccMange.Online_PaymentInsert(Data);
            return st;
        }

        [HttpPost, ActionName("PayMode")]
        public List<MyAccount> PayMode(MyAccount Data)
        {

            AccountMaster AccMange = new AccountMaster();
            List<MyAccount> st = AccMange.PayModeMaster();
            return st;
        }

        [HttpPost, ActionName("CurrencyValues")]
        public List<MyCommonAccess> CurrencyValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CurrencyMaster();
            return st;
        }

        [HttpPost, ActionName("OnlineDashboardList")]
        public List<Online_Dashboardview> OnlineDashboardList(Online_Dashboardview Data)
        {

            OnlineManager AccMange = new OnlineManager();
            List<Online_Dashboardview> st = AccMange.OnlineDashboardList(Data);
            return st;
        }


        [HttpPost, ActionName("Online_Userwise_Agency")]
        public List<UserDetails> Online_Userwise_Agency(UserDetails Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<UserDetails> st = CompMang.Online_UserWise_AgencyDetails(Data);
            return st;
        }


        [HttpPost, ActionName("Online_Customer_VelVoy")]
        public List<Online_VesselVoyageTrack> Online_Customer_VelVoy(Online_VesselVoyageTrack Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_VesselVoyageTrack> st = CompMang.Online_VesVoyValues(Data);
            return st;
        }

        [HttpPost, ActionName("OnlineVesselVoyageList")]
        public List<Online_VesselVoyageTrack> OnlineVesselVoyageList(Online_VesselVoyageTrack Data)
        {

            OnlineManager CompMang = new OnlineManager();
            List<Online_VesselVoyageTrack> st = CompMang.Online_Vessel_voyage_List(Data);
            return st;
        }


        [HttpPost, ActionName("Online_CRO_Email")]
        public List<Online_Booking> Online_CRO_Email(Online_Booking Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_Booking> st = CompMang.Online_CRO_Email(Data);
            return st;
        }


        [HttpPost, ActionName("Online_CRO_ID")]
        public List<Online_Booking> Online_CRO_ID(Online_Booking Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_Booking> st = CompMang.Online_CRO_ID(Data);
            return st;
        }

        public List<Online_RatesheetCharges> Online_Ratesheet_ChargesList(Online_RatesheetCharges Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_RatesheetCharges> st = CompMang.Online_Ratesheet_Charge_List(Data);
            return st;
        }


        public List<Online_BL> Online_BLPrintCheck(Online_BL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BL> st = CompMang.Online_BLPrint_Check(Data);
            return st;
        }

        public List<Online_BL> Online_BLDateupdate(Online_BL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BL> st = CompMang.Online_BLDateupdate(Data);
            return st;
        }

        public List<Online_BL> Online_MonthBL_Details(Online_BL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BL> st = CompMang.Online_MonthBL_Details(Data);
            return st;
        }

        public List<Online_BL> Online_SwitchBL_Insert(Online_BL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BL> st = CompMang.Online_SwitchBL_Insert(Data);
            return st;
        }

        public List<Online_BL> Online_SwitchBL_view(Online_BL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BL> st = CompMang.Online_SwitchBL_view(Data);
            return st;
        }

        public List<Online_BL> Online_ExistingSwitchBL_view(Online_BL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_BL> st = CompMang.Online_ExistingSwitchBL_view(Data);
            return st;
        }

        [HttpPost, ActionName("Online_SwitchBOLInsert")]
        public List<Online_New_BOL> Online_SwitchBOLInsert(Online_New_BOL Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_New_BOL> st = cm.Online_SwitchBOLInsert(Data);
            return st;
        }

        [HttpPost, ActionName("online_BLTypesView")]
        public List<MyCommonAccess> online_BLTypesView()
        {
            OnlineManager cm = new OnlineManager();
            List<MyCommonAccess> st = cm.Get_online_BLTypesView();
            return st;
        }


        [HttpPost, ActionName("Online_BOLAdmin_Insert")]
        public List<Online_New_BOL> Online_BOLAdmin_Insert(Online_New_BOL Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_New_BOL> st = CompMang.Online_BOLAdmin_Insert(Data);
            return st;
        }


        public List<Online_Payment_Confirm> Online_Payment_view(Online_Payment_Confirm Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_Payment_Confirm> st = CompMang.Online_Payment_view(Data);
            return st;
        }

        public List<Online_Payment_Confirm> Online_EditPayment_view(Online_Payment_Confirm Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_Payment_Confirm> st = CompMang.Online_Payment_Edit_view(Data);
            return st;
        }

        public List<Online_Payment_Confirm_dtls> Online_EditPayment_dlts_view(Online_Payment_Confirm_dtls Data)
        {
            OnlineManager CompMang = new OnlineManager();
            List<Online_Payment_Confirm_dtls> st = CompMang.Online_Payment_Edit_dtls_view(Data);
            return st;
        }

        [HttpPost, ActionName("Online_PaymentUpdate")]
        public List<Online_Payment_Confirm_Receipt> Online_PaymentUpdate(Online_Payment_Confirm_Receipt Data)
        {
            OnlineManager AccMange = new OnlineManager();
            List<Online_Payment_Confirm_Receipt> st = AccMange.Online_PaymentUpdate(Data);
            return st;
        }

        [HttpPost, ActionName("Online_BOLSurrender")]
        public List<Online_BL_Attached> Online_BOLSurrender(Online_BL_Attached Data)
        {
            OnlineManager AccMange = new OnlineManager();
            List<Online_BL_Attached> st = AccMange.Online_BOLSurrenderUpdate(Data);
            return st;
        }


        [HttpPost, ActionName("view_online_BLFileattached")]
        public List<Online_BL_Attached_view> view_online_BLFileattached(Online_BL_Attached_view Data)
        {
            OnlineManager AccMange = new OnlineManager();
            List<Online_BL_Attached_view> st = AccMange.online_BLFileattached(Data);
            return st;
        }

        [HttpPost, ActionName("List_online_BLprint")]
        public List<myBLPrint_onLine> List_online_BLprint(myBLPrint_onLine Data)
        {
            OnlineManager AccMange = new OnlineManager();
            List<myBLPrint_onLine> st = AccMange.online_BLprint(Data);
            return st;
        }


        [HttpPost, ActionName("Online_BOLLockUpdate")]
        public List<Online_New_BOL> Online_BL_LockUpdate(Online_New_BOL Data)
        {
            OnlineManager cm = new OnlineManager();
            List<Online_New_BOL> st = cm.Online_BOLLockUpdate(Data);
            return st;
        }


        [HttpPost, ActionName("List_online_BL_Lock_Check")]
        public List<myBLPrint_onLine> List_online_BL_Lock_Check(myBLPrint_onLine Data)
        {
            OnlineManager AccMange = new OnlineManager();
            List<myBLPrint_onLine> st = AccMange.online_BL_Lock_Check(Data);
            return st;
        }

    }

}

