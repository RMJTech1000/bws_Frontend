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
    public class LinerAgencyApiController : ApiController
    {
        #region LinerAgency



        #region Ratesheet


        [ActionName("LinerNameMaster")]
        public List<MyLinerName> LinerNameMaster(MyLinerName Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerName> st = cm.LinerNameDetails(Data);
            return st;
        }

        [ActionName("RatesheetInsert")]
        public List<MyLinerRatesheet> RatesheetInsert(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.InsertRatesheetMaster(Data);
            return st;
        }

        [ActionName("RatesheetInsertNew")]
        public List<MyLinerRatesheet> RatesheetInsertNew(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.InsertRatesheetMasterNew(Data);
            return st;
        }

        [ActionName("RRForeExpire")]
        public List<MyLinerRatesheet> RRForeExpire(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RRForeExpirevalues(Data);
            return st;
        }

        [ActionName("RateSheetViewRecord")]
        public List<MyLinerRatesheet> RateSheetViewRecord(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetRecordView(Data);
            return st;
        }


        [ActionName("RateSheetExistingMasterRecord")]
        public List<MyLinerRatesheet> RateSheetExistingMasterRecord(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetExistingMasterRecordView(Data.ID.ToString());
            return st;
        }

        [ActionName("RateSheetBookingCntrTypes")]
        public List<MyLinerRatesheet> RateSheetBookingCntrTypes(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetBkgCntrTypes(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetDashBordExisting")]
        public List<MyLinerRatesheet> RateSheetDashBordExisting(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetExistingDashBoard(Data.RRID.ToString());
            return st;
        }


        [ActionName("SalesRatesheetContainerInsert")]
        public List<MyLinerRatesheet> SalesRatesheetContainerInsert(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.InsertRatesheetContainerMaster(Data);
            return st;
        }

        [ActionName("RateSheetContainerExsisting")]
        public List<MyLinerRatesheet> RateSheetContainerExsisting(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetContainerExistingView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetModeExsisting")]
        public List<MyLinerRatesheet> RateSheetModeExsisting(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetModeExistingView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetTsPortExsisting")]
        public List<MyLinerRatesheet> RateSheetTsPortExsisting(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RatesheetTsPortExistingView(Data.RRID.ToString());
            return st;
        }


        [ActionName("SalesRatesheetChargesInsert")]
        public List<MyLinerRatesheet> SalesRatesheetChargesInsert(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.InsertRatesheetRateChargeMaster(Data);
            return st;
        }


        [ActionName("RateSheetRevRateExsisting")]
        public List<MyLinerRatesheetRates> RateSheetRevRateExsisting(MyLinerRatesheetRates Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheetRates> st = cm.RatesheetRevRateExistingView(Data);
            return st;
        }


        [ActionName("RateSheetCostRateExsisting")]
        public List<MyLinerRatesheetRates> RateSheetCostRateExsisting(MyLinerRatesheetRates Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheetRates> st = cm.RatesheetCostRateExistingView(Data);
            return st;
        }

        [ActionName("RateSheetCostRateTransExsisting")]
        public List<MyLinerRatesheetRates> RateSheetCostRateTransExsisting(MyLinerRatesheetRates Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheetRates> st = cm.RatesheetCostRateTransExistingView(Data);
            return st;
        }


        [ActionName("RateSheetMRG")]
        public List<MyLinerRRMRG> RateSheetMRG(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.RatesheetMRGView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetSLOTView")]
        public List<MyLinerRRMRG> RateSheetSLOTView(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.RatesheetSLOTView(Data.RRID.ToString());
            return st;
        }
        [ActionName("RateSheetSLOTDtlsView")]
        public List<MyLinerRRMRG> RateSheetSLOTDtlsView(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.RatesheetSLOTDtlsView(Data.RRID.ToString());
            return st;
        }


        [ActionName("SalesRatesheetSLTMRGInsert")]
        public List<MyLinerRRMRG> SalesRatesheetSLTMRGInsert(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.InsertRatesheetMRGSLOTMaster(Data);
            return st;
        }

        [ActionName("SalesRatesheetMRGSelect")]
        public List<MyLinerRRMRG> SalesRatesheetMRGSelect(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.RatesheetMRGSelectView(Data);
            return st;
        }

        [ActionName("SalesRatesheetSLOTSelect")]
        public List<MyLinerRRMRG> SalesRatesheetSLOTSelect(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.RatesheetSLOTSelectView(Data);
            return st;
        }

        [ActionName("SalesRatesheetSLOTDtlSelect")]
        public List<MyLinerRRMRG> SalesRatesheetSLOTDtlSelect(MyLinerRRMRG Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRMRG> st = cm.RatesheetSLOTDtlsSelectView(Data);
            return st;
        }


        [ActionName("RateSheetCheckTariffRateRevPOLExsisting")]
        public List<MyLinerRatesheetRates> RateSheetCheckTariffRateRevPOLExsisting(MyLinerRatesheetRates Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheetRates> st = cm.RatesheetCeckTariffRevRateExistingView(Data);
            return st;
        }


        [ActionName("SalesRatesheetRebateInsert")]
        public List<MyLinerRatesheet> SalesRatesheetRebateInsert(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.InsertRatesheetRebateMaster(Data);
            return st;
        }


        [ActionName("RateSheetRebateExsisting")]
        public List<MyLinerRatesheetRates> RateSheetRebateExsisting(MyLinerRatesheetRates Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheetRates> st = cm.RatesheetRebateExistingView(Data);
            return st;
        }
        #endregion

        [ActionName("SalesRRSendApproval")]
        public List<MyLinerRatesheet> SalesRRSendApproval(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.SendingApproval(Data);
            return st;
        }

        [ActionName("RateSheetNotification")]
        public List<MyLinerRatesheet> RateSheetNotification(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RRNotificationRecordView(Data);
            return st;
        }

        [ActionName("SalesRRFinalSubmitCheck")]
        public List<MyLinerRatesheet> SalesRRFinalSubmitCheck(MyLinerRatesheet Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RRFinalSubmitCheckRecordView(Data);
            return st;
        }

        [ActionName("RRDashBordCount")]
        public List<MyLinerRatesheet> RRDashBordCount()
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRatesheet> st = cm.RRDashBordCountRecordView();
            return st;
        }


        [ActionName("RRFreighTariff")]
        public List<MyLinerRRRate> RRFreighTariff(MyLinerRRRate Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRRate> st = cm.RRFreighTariff(Data);
            return st;
        }

        [ActionName("RRFreighTariffLocalAmt")]
        public List<MyLinerRRRate> RRFreighTariffLocalAmt(MyLinerRRRate Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRRate> st = cm.RRFreighTariffLocalAmt(Data);
            return st;
        }


        [ActionName("ExistRateFreightLocalChargeAmt")]
        public List<MyLinerRRRate> ExistRateFreightLocalChargeAmt(MyLinerRRRate Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRRate> st = cm.ExistingRateSheetFreighTariffLocalAmt(Data);
            return st;
        }

        [ActionName("RRTariffExistingValues")]
        public List<MyLinerRRRate> RRTariffExistingValues(MyLinerRRRate Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRRate> st = cm.RRTariffExisting(Data);
            return st;
        }

        [ActionName("RRBkgPartSales")]
        public List<MyLinerRRRate> RRBkgPartSales(MyLinerRRRate Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MyLinerRRRate> st = cm.RRBkgPartSales(Data);
            return st;
        }

        #endregion

        #region LinerBOL
        [ActionName("BOLViewRecord")]
        public List<MYLinerBOL> BOLViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLViewValus(Data);
            return st;
        }

        [ActionName("BOLCountRecord")]
        public List<MYLinerBOL> BOLCountRecord()
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLCountValus();
            return st;
        }

          [ActionName("BOLCntrNoSelectViewRecord")]
        public List<MYLinerBOL> BOLCntrNoSelectViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLCntrNoSelectExistingValus(Data);
            return st;
        }

        [ActionName("BOLVesselVoyage")]
        public List<MYLinerBOL> BOLVesselVoyage(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLVesselVoyageValus(Data);
            return st;
        }

        [ActionName("BOLInsert")]
        public List<MYLinerBOL> BOLInsert(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLInsert(Data);
            return st;
        }

        [ActionName("BOLExistingViewRecord")]
        public List<MYLinerBOL> BOLExistingViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLExistingViewValus(Data);
            return st;
        }
        [ActionName("BOLCustomerExistingViewRecord")]
        public List<MYLinerBOL> BOLCustomerExistingViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLCustomerExistingValus(Data);
            return st;
        }
        [ActionName("BOLCntrExistingViewRecord")]
        public List<MYLinerBOL> BOLCntrExistingViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLCntrExistingValus(Data);
            return st;
        }

        [ActionName("BkgVoyageDtlsExistingViewRecord")]
        public List<MYLinerBOL> BkgVoyageDtlsExistingViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BkgVoyagedtlsExistingValus(Data);
            return st;
        }

        [ActionName("BOLBookingSelectViewRecord")]
        public List<MYLinerBOL> BOLBookingSelectViewRecord(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLBkgSelectExistingValus(Data);
            return st;
        }
        [ActionName("BkgBOLVesselVoyage")]
        public List<MYLinerBOL> BkgBOLVesselVoyage(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BkgBOLVesselVoyageValus(Data);
            return st;
        }
        [ActionName("MRGExistingMasterRecord")]
        public List<MYLinerRRBooking> MRGExistingMasterRecord(MYLinerRRBooking Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerRRBooking> st = cm.BookingRRSearchBindValus(Data);
            return st;
        }

        [ActionName("BOLRRNumberUpdate")]
        public List<MYLinerBOL> BOLRRNumberUpdate(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLRRUpDate(Data);
            return st;
        }


        [ActionName("BOLStatusUpdate")]
        public List<MYLinerBOL> BOLStatusUpdate(MYLinerBOL Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBOL> st = cm.BOLStatusUpDate(Data);
            return st;
        }

        #endregion

        #region GANESH (LINER BL NUMBERING LOGICS)

        [ActionName("BLNoLogics")]
        public List<MYLinerBLNoLogics> BLNoLogicsData(MYLinerBLNoLogics Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBLNoLogics> st = cm.BLNoLogicsValues(Data);
            return st;
        }

        [ActionName("BLNoLogicsInsert")]
        public List<MYLinerBLNoLogics> BLNoLogicsInsert(MYLinerBLNoLogics Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBLNoLogics> st = cm.BLNoLogicsInsert(Data);
            return st;
        }

        [ActionName("BLNoLogicsView")]
        public List<MYLinerBLNoLogics> BLLogicsViewRecord(MYLinerBLNoLogics Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBLNoLogics> st = cm.BLLogicsViewRecordValues(Data);
            return st;
        }

        [ActionName("BLNoLogicsEdit")]
        public List<MYLinerBLNoLogics> BLNoLogicsEdit(MYLinerBLNoLogics Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBLNoLogics> st = cm.BLNoLogicsEditValues(Data);
            return st;
        }

        //[ActionName("BLReleaseExisCheckRecord")]
        //public List<MYLinerBLRelease> BLReleaseExisCheckRecord(MYLinerBLRelease Data)
        //{
        //    LinerAgencyManager cm = new LinerAgencyManager();
        //    List<MYLinerBLRelease> st = cm.BLReleaseExisCheckValus(Data);
        //    return st;
        //}

        [ActionName("BLReleaseExistingViewRecord")]
        public List<MYLinerBLRelease> BLReleaseExistingViewRecord(MYLinerBLRelease Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBLRelease> st = cm.BLReleaseExistingViewValus(Data);
            return st;
        }


        [ActionName("BLReleaseCntrViewRecord")]
        public List<MYLinerBLRelease> BLReleaseCntrViewRecord(MYLinerBLRelease Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();

            List<MYLinerBLRelease> st = cm.BLReleaseCntrExistingViewValus(Data);
            return st;
        }

        //[ActionName("BLReleaseViewValus")]
        //public List<MYLinerBLRelease> BLReleaseViewValus(MYLinerBLRelease Data)
        //{
        //    LinerAgencyManager cm = new LinerAgencyManager();
        //    List<MYLinerBLRelease> st = cm.BLReleaseFinalExistingViewValus(Data);
        //    return st;
        //}

        [ActionName("BLReleaseInsert")]
        public List<MYLinerBLRelease> BLReleaseInsert(MYLinerBLRelease Data)
        {
            LinerAgencyManager cm = new LinerAgencyManager();
            List<MYLinerBLRelease> st = cm.BLReleaseInsert(Data);
            return st;
        }

        [ActionName("ExtPickcntrBookingMaster")]
        public List<MyLinerCntrPickupdtls> ExtPickcntrBookingMaster(MyLinerCntrPickupdtls Data)
        {
            LinerAgencyManager Mange = new LinerAgencyManager();
            List<MyLinerCntrPickupdtls> st = Mange.ExtPickcntrBooking(Data);
            return st;
        }
        #endregion
    }
}
