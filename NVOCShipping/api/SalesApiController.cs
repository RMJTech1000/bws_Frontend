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
    public class SalesApiController : ApiController
    {
        #region Muthu

        [ActionName("SalesMRGInsert")]
        public List<MyMRG>SalesMRGInsert(MyMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyMRG> st = cm.InsertMRGMaster(Data);
            return st;
        }

        [ActionName("MRGViewRecord")]
        public List<MyMRG> MRGViewRecord(MyMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyMRG> st = cm.MRGRecordView(Data);
            return st;
        }

        [ActionName("MRGExistingMasterRecord")]
        public List<MyMRG> MRGExistingMasterRecord(MyMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyMRG> st = cm.MRGExistingMasterRecordView(Data.ID.ToString());
            return st;
        }


        [ActionName("SalesSLOTInsert")]
        public List<MySLOT> SalesSLOTInsert(MySLOT Data)
        {
            SalesManager cm = new SalesManager();
            List<MySLOT> st = cm.InsertSLOTMaster(Data);
            return st;
        }

        [ActionName("SLOTViewRecord")]
        public List<MySLOT> SLOTViewRecord(MySLOT Data)
        {
            SalesManager cm = new SalesManager();
            List<MySLOT> st = cm.SLOTRecordView(Data);
            return st;
        }
        [ActionName("SLOTExistingRecord")]
        public List<MySLOT> SLOTExistingRecord(MySLOT Data)
        {
            SalesManager cm = new SalesManager();
            List<MySLOT> st = cm.SLOTExistingRecordView(Data.ID.ToString());
            return st;
        }

        [ActionName("SLOTExistingMasterRecord")]
        public List<MySLOT> SLOTExistingMasterRecord(MySLOT Data)
        {
            SalesManager cm = new SalesManager();
            List<MySLOT> st = cm.SLOTExistingMasterRecordView(Data.ID.ToString());
            return st;
        }

        [ActionName("CheckTariffValidation")]
        public List<MYPortTariffMaster> CheckTariffValidation(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.ExistTariffValidation(Data);
            return st;
        }

        [ActionName("SalesPortTariffInsert")]
        public List<MYPortTariffMaster> SalesPortTariffInsert(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.InsertPortTariffMaster(Data);
            return st;
        }

        [ActionName("InsertPortTariffChargeInsert")]
        public List<MYPortTariffMaster> InsertPortTariffChargeInsert(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.InsertPortTariffChargeMaster(Data);
            return st;
        }

        [ActionName("InsertPortTariffIHCChargeInsert")]
        public List<MYPortTariffMaster> InsertPortTariffIHCChargeInsert(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.InsertPortTariffIHCChargeMaster(Data);
            return st;
        }



        [ActionName("PortTariffChargeViewRecord")]
        public List<MYPortTariffMaster> PortTariffChargeViewRecord(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.PortTariffChargeExistingRecordView(Data);
            return st;
        }

        [ActionName("PortTariffIHCChargeViewRecord")]
        public List<MYPortTariffMaster> PortTariffIHCChargeViewRecord(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.PortTariffIHCChargeExistingRecordView(Data);
            return st;
        }



        [ActionName("PortTariffViewRecord")]
        public List<MYPortTariffMaster> PortTariffViewRecord(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.PortTariffRecordView(Data);
            return st;
        }


        [ActionName("PortTariffExistingMasterRecord")]
        public List<MYPortTariffMaster> PortTariffExistingMasterRecord(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.PortTariffExistingMasterRecordView(Data.ID.ToString());
            return st;
        }

        [ActionName("PortTariffExistingRecord")]
        public List<MYPortTariffMaster> PortTariffExistingRecord(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.PortTraiffDtlsExistingView(Data.ID.ToString());
            return st;
        }

        [ActionName("PortTariffDelete")]
        public List<MYPortTariffMaster> PortTariffDelete(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.PortTariffDelete(Data);
            return st;
        }

        [ActionName("RatesheetSlotDelete")]
        public List<MYPortTariffMaster> RatesheetSlotDelete(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.RRSlotDeleteRecord(Data);
            return st;
        }

        [ActionName("RatesheetGridValueDelete")]
        public List<MYPortTariffMaster> RatesheetGridValueDelete(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.RRChargesDeleteRecord(Data);
            return st;
        }

        [ActionName("RRCntrTypesDelete")]
        public List<MYPortTariffMaster> RRCntrTypesDelete(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.RRCntrTypesDeleteRecord(Data);
            return st;
        }

        [ActionName("SlatTariffDelete")]
        public List<MYPortTariffMaster> SlatTariffDelete(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.DeleteSlotTraff(Data);
            return st;
        }


        [ActionName("PortTariffBrackupDelete")]
        public List<MYPortTariffMaster> PortTariffBrackupDelete(MYPortTariffMaster Data)
        {
            SalesManager cm = new SalesManager();
            List<MYPortTariffMaster> st = cm.DeleteTariffBrackupRecord(Data);
            return st;
        }


        #region Ratesheet

        [ActionName("RatesheetInsert")]
        public List<MyRatesheet> RatesheetInsert(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetMaster(Data);
            return st;
        }

        [ActionName("RatesheetInsertNew")]
        public List<MyRatesheet> RatesheetInsertNew(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetMasterNew(Data);
            return st;
        }

        [ActionName("RatesheetValidityInsert")]
        public List<MyRatesheet> RatesheetValidityInsert(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetMasterNewvalidity(Data);
            return st;
        }
        public List<MyRatesheet> RatesheetCastInsertNew(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetCastMasterNew(Data);
            return st;
        }

        [ActionName("RRForeExpire")]
        public List<MyRatesheet> RRForeExpire(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RRForeExpirevalues(Data);
            return st;
        }

        [ActionName("RateSheetViewRecord")]
        public List<MyRatesheet> RateSheetViewRecord(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetRecordView(Data);
            return st;
        }


        [ActionName("RateSheetExistingMasterRecord")]
        public List<MyRatesheet> RateSheetExistingMasterRecord(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetExistingMasterRecordView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetBookingCntrTypes")]
        public List<MyRatesheet> RateSheetBookingCntrTypes(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetBkgCntrTypes(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetDashBordExisting")]
        public List<MyRatesheet> RateSheetDashBordExisting(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetExistingDashBoard(Data.RRID.ToString());
            return st;
        }


        [ActionName("SalesRatesheetContainerInsert")]
        public List<MyRatesheet> SalesRatesheetContainerInsert(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetContainerMaster(Data);
            return st;
        }

        [ActionName("RateSheetContainerExsisting")]
        public List<MyRatesheet> RateSheetContainerExsisting(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetContainerExistingView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetModeExsisting")]
        public List<MyRatesheet> RateSheetModeExsisting(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetModeExistingView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetTsPortExsisting")]
        public List<MyRatesheet> RateSheetTsPortExsisting(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetTsPortExistingView(Data.RRID.ToString());
            return st;
        }


        [ActionName("SalesRatesheetChargesInsert")]
        public List<MyRatesheet> SalesRatesheetChargesInsert(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetRateChargeMaster(Data);
            return st;
        }


        [ActionName("RateSheetRevRateExsisting")]
        public List<MyRatesheetRates> RateSheetRevRateExsisting(MyRatesheetRates Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheetRates> st = cm.RatesheetRevRateExistingView(Data);
            return st;
        }


        [ActionName("RateSheetCostRateExsisting")]
        public List<MyRatesheetRates> RateSheetCostRateExsisting(MyRatesheetRates Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheetRates> st = cm.RatesheetCostRateExistingView(Data);
            return st;
        }

        [ActionName("RateSheetCostRateTransExsisting")]
        public List<MyRatesheetRates> RateSheetCostRateTransExsisting(MyRatesheetRates Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheetRates> st = cm.RatesheetCostRateTransExistingView(Data);
            return st;
        }


        [ActionName("RateSheetMRG")]
        public List<MyRRMRG> RateSheetMRG(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.RatesheetMRGView(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetSLOTView")]
        public List<MyRRMRG> RateSheetSLOTView(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.RatesheetSLOTView(Data.RRID.ToString());
            return st;
        }
        [ActionName("RateSheetSLOTDtlsView")]
        public List<MyRRMRG> RateSheetSLOTDtlsView(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.RatesheetSLOTDtlsView(Data.RRID.ToString());
            return st;
        }


        [ActionName("SalesRatesheetSLTMRGInsert")]
        public List<MyRRMRG> SalesRatesheetSLTMRGInsert(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.InsertRatesheetMRGSLOTMaster(Data);
            return st;
        }

        [ActionName("SalesRatesheetMRGSelect")]
        public List<MyRRMRG> SalesRatesheetMRGSelect(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.RatesheetMRGSelectView(Data);
            return st;
        }

        [ActionName("SalesRatesheetSLOTSelect")]
        public List<MyRRMRG> SalesRatesheetSLOTSelect(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.RatesheetSLOTSelectView(Data);
            return st;
        }

        [ActionName("SalesRatesheetSLOTDtlSelect")]
        public List<MyRRMRG> SalesRatesheetSLOTDtlSelect(MyRRMRG Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRMRG> st = cm.RatesheetSLOTDtlsSelectView(Data);
            return st;
        }


        [ActionName("RateSheetCheckTariffRateRevPOLExsisting")]
        public List<MyRatesheetRates> RateSheetCheckTariffRateRevPOLExsisting(MyRatesheetRates Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheetRates> st = cm.RatesheetCeckTariffRevRateExistingView(Data);
            return st;
        }


        [ActionName("SalesRatesheetRebateInsert")]
        public List<MyRatesheet> SalesRatesheetRebateInsert(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRatesheetRebateMaster(Data);
            return st;
        }


        [ActionName("RateSheetRebateExsisting")]
        public List<MyRatesheetRates> RateSheetRebateExsisting(MyRatesheetRates Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheetRates> st = cm.RatesheetRebateExistingView(Data);
            return st;
        }
        #endregion

        [ActionName("SalesRRSendApproval")]
        public List<MyRatesheet> SalesRRSendApproval(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.SendingApproval(Data);
            return st;
        }

        [ActionName("RateSheetNotification")]
        public List<MyRatesheet> RateSheetNotification(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RRNotificationRecordView(Data);
            return st;
        }

        [ActionName("SalesRRFinalSubmitCheck")]
        public List<MyRatesheet> SalesRRFinalSubmitCheck(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RRFinalSubmitCheckRecordView(Data);
            return st;
        }

        [ActionName("RRDashBordCount")]
        public List<MyRatesheet> RRDashBordCount(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RRDashBordCountRecordView(Data);
            return st;
        }


        [ActionName("RRFreighTariff")]
        public List<MyRRRate> RRFreighTariff(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRFreighTariff(Data);
            return st;
        }

        [ActionName("RRFreighTariffLocalAmt")]
        public List<MyRRRate> RRFreighTariffLocalAmt(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRFreighTariffLocalAmt(Data);
            return st;
        }


        [ActionName("ExistRateFreightLocalChargeAmt")]
        public List<MyRRRate> ExistRateFreightLocalChargeAmt(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.ExistingRateSheetFreighTariffLocalAmt(Data);
            return st;
        }

        [ActionName("ExistRateTranshipmentPort")]
        public List<MyRRRate> ExistRateTranshipmentPort(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.ExistingRRTranshipmentPort(Data);
            return st;
        }

        [ActionName("RRExistSlotAmt")]
        public List<MyRRRate> RRExistSlotAmt(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.ExistingRateSheetSlotAmt(Data);
            return st;
        }

        [ActionName("RRTariffExistingValues")]
        public List<MyRRRate> RRTariffExistingValues(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRTariffExisting(Data);
            return st;
        }
        [ActionName("RRTariffLocalExistingValues")]
        public List<MyRRRate> RRTariffLocalExistingValues(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRTariffExistingLocalcharge(Data);
            return st;
        }

        [ActionName("RRBkgPartSales")]
        public List<MyRRRate> RRBkgPartSales(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRBkgPartSales(Data);
            return st;
        }

        [ActionName("ListUsers")]
        public List<MyRRRate> RRListUsers(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRListUsers(Data);
            return st;
        }
        #endregion

        #region ganesh RR NOTIFICATION//Commission Contract //IHC TARIFF


        [ActionName("RateSheetNotificationPopup")]
        public List<MyRatesheet> RateSheetNotificationPopup(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetNotificationValues(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetNotificationSlotTerms")]
        public List<MyRatesheet> RateSheetNotificationSlotTerms(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RatesheetNotificationSlotValues(Data.RRID.ToString());
            return st;
        }

        [ActionName("RateSheetNotificationTHCandIHC")]
        public List<MyRatesheet> RateSheetNotificationTHCandIHC(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RateSheetNotificationTHCandIHCValues(Data.RRID.ToString());
            return st;
        }

        [ActionName("RRNotificationChargewiseRates")]
        public List<MyRatesheet> RRNotificationChargewiseRates(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RRNotificationChargewiseRates(Data.RRID.ToString());
            return st;
        }

        [ActionName("RRNotificationModeview")]
        public List<MyRatesheet> RRNotificationModeview(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.RRNotificationMode(Data.RRID.ToString());
            return st;
        }


        [ActionName("CommissionCharges")]
        public List<MyCommContract> CommissionCharges(MyCommContract Data)
        {
            SalesManager cm = new SalesManager();
            List<MyCommContract> st = cm.CommissionChargesValues(Data.ChargeCommTypes.ToString());
            return st;
        }

        [ActionName("ShipmentTypes")]
        public List<MyCommContract> ShipmentTypes(MyCommContract Data)
        {
            SalesManager cm = new SalesManager();
            List<MyCommContract> st = cm.ShipmentTypesValues(Data.ShipmentTypes.ToString());
            return st;
        }

        [ActionName("InsertCommissionContract")]
        public List<MyCommContract> InsertCommissionContract(MyCommContract Data)
        {
            SalesManager cm = new SalesManager();
            List<MyCommContract> st = cm.InsertCommissionContract(Data);
            return st;
        }
        [ActionName("CommContractView")]
        public List<MyCommContract> CommContractViewValues(MyCommContract Data)
        {
            SalesManager cm = new SalesManager();
            List<MyCommContract> st = cm.GetCommContractView(Data);
            return st;
        }

        [ActionName("CommContractEdit")]
        public List<MyCommContract> CommContractEdit(MyCommContract Data)
        {
            SalesManager cm = new SalesManager();
            List<MyCommContract> st = cm.CommContractEditValues(Data);
            return st;
        }

        [ActionName("CheckCCValidation")]
        public List<MyCommContract> CheckCCValidation(MyCommContract Data)
        {
            SalesManager cm = new SalesManager();
            List<MyCommContract> st = cm.ExistCheckCCValidation(Data);
            return st;
        }

        [ActionName("InsertIHCTariffMaster")]
        public List<MyIHCTariff> InsertIHCTariffMaster(MyIHCTariff Data)
        {
            SalesManager cm = new SalesManager();
            List<MyIHCTariff> st = cm.InsertIHCTariffMaster(Data);
            return st;
        }

        [ActionName("IHCHaulageTariffViewRecord")]
        public List<MyIHCTariff> IHCHaulageTariffViewValues(MyIHCTariff Data)
        {
            SalesManager cm = new SalesManager();
            List<MyIHCTariff> st = cm.IHCHaulageTariffViewValues(Data);
            return st;
        }

        [ActionName("IHCHaulageTariffVEdit")]
        public List<MyIHCTariff> IHCHaulageTariffVEdit(MyIHCTariff Data)
        {
            SalesManager cm = new SalesManager();
            List<MyIHCTariff> st = cm.IHCHaulageTariffVEditValues(Data);
            return st;
        }
        [ActionName("IHCHaulageTariffDtlsEdit")]
        public List<MyIHCTariff> IHCHaulageTariffDtlsEdit(MyIHCTariff Data)
        {
            SalesManager cm = new SalesManager();
            List<MyIHCTariff> st = cm.IHCHaulageTariffDtlsEditValues(Data);
            return st;
        }
        [ActionName("IHCHaulageTariffDtlsDelete")]
        public List<MyIHCTariff> IHCHaulageTariffDtlsDelete(MyIHCTariff Data)
        {
            SalesManager cm = new SalesManager();
            List<MyIHCTariff> st = cm.GetIHCHaulageTariffDtlsDelete(Data);
            return st;
        }
        [ActionName("IHCHaulageTariffValidation")]
        public List<MyIHCTariff> IHCHaulageTariffValidation(MyIHCTariff Data)
        {
            SalesManager cm = new SalesManager();
            List<MyIHCTariff> st = cm.IHCHaulageTariffValidation(Data);
            return st;
        }
        #endregion


        [ActionName("RRNotifyTariffExistingValues")]
        public List<MyRRRate> RRNotifyTariffExistingValues(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRNotifyTariffExisting(Data);
            return st;
        }

        [ActionName("RRNotifyCostTariffExistingValues")]
        public List<MyRRRate> RRNotifyCostTariffExistingValues(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.RRNotifyCostTariffExisting(Data);
            return st;
        }

        [ActionName("RRExistNotifySlotAmt")]
        public List<MyRRRate> RRExistNotifySlotAmt(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.ExistingRateSheetNotifySlotAmt(Data);
            return st;
        }

        [ActionName("RRExistNotifyCostRevenuTotalAmt")]
        public List<MyRRRate> RRExistNotifyCostRevenuTotalAmt(MyRRRate Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRRRate> st = cm.getExistingRateSheetNotifyTotalAmt(Data);
            return st;
        }


        

        #region Ganesh Empty Yardc cntr Cost

        [ActionName("GeoLocByCountry")]
        public List<MyEmptyYard> GeoLocByCountry(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.GeoLocByCountry(Data);
            return st;
        }

        [ActionName("InsertEmptyYardCosts")]
        public List<MyEmptyYard> InsertEmptyYardCosts(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.InsertEmptyYardCosts(Data);
            return st;
        }
        [ActionName("EmptyYardTariffViewRecord")]
        public List<MyEmptyYard> EmptyYardTariffViewRecord(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.EmptyYardTariffViewRecord(Data);
            return st;
        }
        [ActionName("EmptyYardTariffEdit")]
        public List<MyEmptyYard> EmptyYardTariffEdit(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.EmptyYardTariffEditValues(Data);
            return st;
        }

        [ActionName("EmptyYardTariffSlabEdit")]
        public List<MyEmptyYard> EmptyYardTariffSlabEdit(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.EmptyYardTariffSlabEdit(Data);
            return st;
        }
        [ActionName("EmptyYardTariffPortCostEdit")]
        public List<MyEmptyYard> EmptyYardTariffPortCostEdit(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.EmptyYardTariffPortCostEdit(Data);
            return st;
        }

        [ActionName("EmptyYardTariffValidation")]
        public List<MyEmptyYard> EmptyYardTariffValidation(MyEmptyYard Data)
        {
            SalesManager cm = new SalesManager();
            List<MyEmptyYard> st = cm.EmptyYardTariffValidation(Data);
            return st;
        }
        #endregion

    }
}
