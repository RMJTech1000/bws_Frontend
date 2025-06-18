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
    public class DocumentApiController : ApiController
    {
        [ActionName("MRGExistingMasterRecord")]
        public List<MYRRBooking> MRGExistingMasterRecord(MYRRBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYRRBooking> st = cm.BookingRRSearchBindValus(Data);
            return st;
        }

        [ActionName("BookingInsert")]
        public List<MyBooking> BookingInsert(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingInsert(Data);
            return st;
        }


        [ActionName("BookingConfirmInsert")]
        public List<MyBooking> BookingConfirmInsert(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingConfirmInsert(Data);
            return st;
        }

        [ActionName("BookingCancelled")]
        public List<MyBooking> BookingCancelled(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingConfirmCancelled(Data);
            return st;
        }


        [ActionName("Bookingupdate")]
        public List<MyBooking> Bookingupdate(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.Bookingupdate(Data);
            return st;
        }





        [ActionName("BookingViewRecord")]
        public List<MyBooking> BookingViewRecord(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingViewValus(Data);
            return st;
        }


        [ActionName("BookingCountRecord")]
        public List<MyBooking> BookingCountRecord(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingCountValus(Data);
            return st;
        }

        [ActionName("BookingExistingViewRecord")]
        public List<MyBooking> BookingExistingViewRecord(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingExistingViewValus(Data);
            return st;
        }

        [ActionName("BkgExistingCntrTypesViewRecord")]
        public List<MyBooking> BkgExistingCntrTypesViewRecord(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BkgCntrTypesExistingValus(Data);
            return st;
        }

        [ActionName("CROInsert")]
        public List<MyCROMaster> CROInsert(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROInsert(Data);
            return st;
        }

        [ActionName("CROUpdate")]
        public List<MyCROMaster> CROUpdate(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROUpdate(Data);
            return st;
        }

        [ActionName("CROCancelled")]
        public List<MyCROMaster> CROCancelled(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROCancelled(Data);
            return st;
        }


        [ActionName("CROViewRecord")]
        public List<MyCROMaster> CROViewRecord(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROViewValus(Data);
            return st;
        }

        [ActionName("CROViewRecordALL")]
        public List<MyCROMaster> CROViewRecordALL(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROViewValusAll(Data);
            return st;
        }


        [ActionName("CROBkgselectRecord")]
        public List<MyCROMaster> CROBkgSelectRecord(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROBkgSelectView(Data);
            return st;
        }


        [ActionName("CROExistingViewRecord")]
        public List<MyCROMaster> CROExistingViewRecord(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CROExistingViewValus(Data);
            return st;
        }

        [ActionName("CROExistingDtlsViewRecord")]
        public List<MyCROMaster> CROExistingDtlsViewRecord(MyCROMaster Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyCROMaster> st = cm.CRODetailsExistingValus(Data);
            return st;
        }


        [ActionName("BOLBookingSelectViewRecord")]
        public List<MYBOL> BOLBookingSelectViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLBkgSelectExistingValus(Data);
            return st;
        }

        [ActionName("BOLCntrNoSelectViewRecord")]
        public List<MYBOL> BOLCntrNoSelectViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLCntrNoSelectExistingValus(Data);
            return st;
        }


        [ActionName("BOLInsert")]
        public List<MYBOL> BOLInsert(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLInsert(Data);
            return st;
        }

        [ActionName("BOLExemptionUpdate")]
        public List<MYBOL> BOLExemptionUpdate(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLExemptionUpdate(Data);
            return st;
        }

        [ActionName("BOLCargoReleaseUpdate")]
        public List<MYBOL> BOLCargoReleaseUpdate(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLCargoReleaseUpdate(Data);
            return st;
        }


        [ActionName("BOLViewRecord")]
        public List<MYBOL> BOLViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLViewValus(Data);
            return st;
        }

        [ActionName("BOLExistingViewRecord")]
        public List<MYBOL> BOLExistingViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLExistingViewValus(Data);
            return st;
        }

        [ActionName("BOLCustomerExistingViewRecord")]
        public List<MYBOL> BOLCustomerExistingViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLCustomerExistingValus(Data);
            return st;
        }

        [ActionName("BOLCntrExistingViewRecord")]
        public List<MYBOL> BOLCntrExistingViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLCntrExistingValus(Data);
            return st;
        }

        [ActionName("BLReleaseExisCheckRecord")]
        public List<MYBLRelease> BLReleaseExisCheckRecord(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLRelease> st = cm.BLReleaseExisCheckValus(Data);
            return st;
        }

        [ActionName("BLReleaseExistingViewRecord")]
        public List<MYBLRelease> BLReleaseExistingViewRecord(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLRelease> st = cm.BLReleaseExistingViewValus(Data);
            return st;
        }


        [ActionName("BLReleaseCntrViewRecord")]
        public List<MYBLRelease> BLReleaseCntrViewRecord(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();

            List<MYBLRelease> st = cm.BLReleaseCntrExistingViewValus(Data);
            return st;
        }

        [ActionName("BLReleaseViewValus")]
        public List<MYBLRelease> BLReleaseViewValus(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLRelease> st = cm.BLReleaseFinalExistingViewValus(Data);
            return st;
        }


        [ActionName("BLReleaseSOBDateCheckExValuea")]
        public List<MYBLRelease> BLReleaseSOBDateCheckExValuea(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLRelease> st = cm.BLReleaseSOBDateCheckViewValus(Data);
            return st;
        }



        [ActionName("BLReleaseInsert")]
        public List<MYBLRelease> BLReleaseInsert(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLRelease> st = cm.BLReleaseInsert(Data);
            return st;
        }

        [ActionName("BOLVesselVoyage")]
        public List<MYBOL> BOLVesselVoyage(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLVesselVoyageValus(Data);
            return st;
        }

        [ActionName("BkgBOLVesselVoyage")]
        public List<MYBOL> BkgBOLVesselVoyage(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BkgBOLVesselVoyageValus(Data);
            return st;
        }

        [ActionName("BOLVoyageDtlsExistingViewRecord")]
        public List<MYBOL> BOLVoyageDtlsExistingViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLVoyagedtlsExistingValus(Data);
            return st;
        }

        [ActionName("BkgVoyageDtlsExistingViewRecord")]
        public List<MYBOL> BkgVoyageDtlsExistingViewRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BkgVoyagedtlsExistingValus(Data);
            return st;
        }


        [ActionName("BOLRRNumberUpdate")]
        public List<MYBOL> BOLRRNumberUpdate(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLRRUpDate(Data);
            return st;
        }

        [ActionName("BOLStatusUpdate")]
        public List<MYBOL> BOLStatusUpdate(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLStatusUpDate(Data);
            return st;
        }

        [ActionName("BOLCountRecord")]
        public List<MYBOL> BOLCountRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLCountValus(Data);
            return st;
        }

        [ActionName("BLUnlock")]
        public List<MYBOL> BLUnlock(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BLUnlock(Data);
            return st;
        }

        [ActionName("RRUNLink")]
        public List<MYBOL> RRUNLink(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.RRUnLink(Data);
            return st;
        }

        [ActionName("BookingSlotOperator")]
        public List<MyBooking> BookingSlotOperator(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingSlotValies(Data);
            return st;
        }

        [ActionName("VoyageBookingSlotOperator")]
        public List<MyBooking> VoyageBookingSlotOperator(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.VoyageBookingSlotValies(Data);
            return st;
        }

        

        [ActionName("BookingSlotRefNoOperator")]
        public List<MyBooking> BookingSlotRefNoOperator(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingSlotRefNoValies(Data);
            return st;
        }

        [ActionName("RRBookingSlotAmountfill")]
        public List<MyBooking> RRBookingSlotAmountfill(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.RRBookingSlotAmount(Data);
            return st;
        }


        

        [ActionName("BookingSlotContractNoOperator")]
        public List<MyBooking> BookingSlotContractNoOperator(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingSlotContactNoValies(Data);
            return st;
        }

        [ActionName("BookingTransPortOperator")]
        public List<MyBooking> BookingTransPortOperator(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingTransPortValues(Data);
            return st;
        }

        [ActionName("BOLBtnViewDtlsRecord")]
        public List<MYBOL> BOLBtnViewDtlsRecord(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLBtnViewDtlsRecordValus(Data);
            return st;
        }

        [ActionName("ListDepotByPort")]
        public List<MyBooking> ListDepotByPort(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.ListDepotByPortRecordValus(Data);
            return st;
        }

        [ActionName("ModuleValues")]
        public List<MyBooking> ModuleValues(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.ListModuleValues(Data);
            return st;
        }

        #region anand
        [ActionName("BLNoDropDown")]
        public List<MYBOL> BLNoDropDown()
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BLNoBind();
            return st;
        }

        [ActionName("BLNoAgentwise")]
        public List<MYBOL> BLNoAgentwise(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BLNoAgentwiseBind(Data);
            return st;
        }

        


        [ActionName("CorrectorSlotOpVessel")]
        public List<ChargeCorrector> CorrectorSlotOpVessel(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.ExportSlotOpVesselManeger(Data);
            return st;
        }

        [ActionName("PrincipalChargeCorrector")]
        public List<ChargeCorrector> PrincipalChargeCorrector(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.PrincipalChargeCorrectorMaster(Data);
            return st;
        }



        [ActionName("CorrectorSlotOpVesselAmd")]
        public List<ChargeCorrector> CorrectorSlotOpVesselAmd(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.ExportSlotOpVesselManegerAmd(Data);
            return st;
        }

        [ActionName("CorrectorSlotOpVesselAmdNew")]
        public List<ChargeCorrector> CorrectorSlotOpVesselAmdNew(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.ExportSlotOpVesselManegerAmdNew(Data);
            return st;
        }


        [ActionName("CorrectionExsitingValues")]
        public List<ChargeCorrector> CorrectionExsitingValues(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.BLCorrectionExistingValues(Data);
            return st;
        }

        [ActionName("CorrectorSlotChargeChaged")]
        public List<ChargeCorrector> CorrectorSlotChargeChaged(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.ExportSlotChargesChange(Data);
            return st;
        }

        [ActionName("SlotCorrectionExsitingValues")]
        public List<ChargeCorrector> SlotCorrectionExsitingValues(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.BLSlotCorrectorExistingValues(Data);
            return st;
        }

        [ActionName("BLCorrectionExsitingChargeValues")]
        public List<ChargeCorrector> BLCorrectionExsitingChargeValues(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.BLCorrectionExistingChargeValues(Data);
            return st;
        }


        [ActionName("BLCorrectionApproval")]
        public List<ChargeCorrectorInsert> BLCorrectionApproval(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLCorrectionUpdateApproval(Data);
            return st;
        }

        [ActionName("BLCorrectionReject")]
        public List<ChargeCorrectorInsert> BLCorrectionReject(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLCorrectionUpdateReject(Data);
            return st;
        }

        


        [ActionName("AgencyChargeCorrector")]
        public List<ChargeCorrector> AgencyChargeCorrector(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.AgencyChargeCorrectorMaster(Data);
            return st;
        }
        [ActionName("ChargeCorrectorExiting")]
        public List<MyRatesheetRates> ChargeCorrectorExiting(MyRatesheetRates Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyRatesheetRates> st = cm.ChargeCorrectorRecordMaster(Data);
            return st;
        }
        [ActionName("ChargeCorrectorInsert")]
        public List<ChargeCorrectorInsert> ChargeCorrectorInsert(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.InsertChargeCorrector(Data);
            return st;
        }
        [ActionName("ChargeCorrectorUpdated")]
        public List<ChargeCorrectorInsert> ChargeCorrectorUpdated(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.ChargeCorrectorMasterDetails(Data);
            return st;
        }

        [ActionName("BLCorrectionApprovalAmd")]
        public List<ChargeCorrectorInsert> BLCorrectionApprovalAmd(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLCorrectionUpdateApprovalAmd(Data);
            return st;
        }


        [ActionName("BLCorrectionRejectAmd")]
        public List<ChargeCorrectorInsert> BLCorrectionRejectAmd(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLCorrectionSlotReject(Data);
            return st;
        }

        [ActionName("ChargeSlotInsert")]

        public List<ChargeCorrector> ChargeSlotInsert(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.BLCorrectionSlotChargeInsert(Data);
            return st;
        }





        [ActionName("BLChargeCorrectorSearch")]
        public List<ChargeCorrectorInsert> BLChargeCorrectorSearch(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLChargeCorrectorSearch(Data);
            return st;
        }


        [ActionName("ChargeDetailsDelete")]
        public List<ChargeCorrectorInsert> ChargeDetailsDelete(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.ChargeCorrectorDelete(Data);
            return st;
        }

        [ActionName("BLChargeCorrcetiorDelete")]
        public List<ChargeCorrectorInsert> BLChargeCorrcetiorDelete(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLChargeCorrcetionDelete(Data);
            return st;
        }


        
        #endregion


        [ActionName("BLSLOTCorrectorView")]
        public List<ChargeCorrectorInsert> BLSLOTCorrectorView(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.ExpSlotCorrectorValues(Data);
            return st;
        }

        [ActionName("BLPrintLogInsert")]
        public List<MYBLPrintLog> BLPrintLogInsert(MYBLPrintLog Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLPrintLog> st = cm.BLPrintLogInsert(Data);
            return st;
        }


        [ActionName("BLPrintLogView")]
        public List<MYBLPrintLog> BLPrintLogView(MYBLPrintLog Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLPrintLog> st = cm.BLPrintLogViewValus(Data);
            return st;
        }

        [ActionName("BLPrintLogDelete")]
        public List<MYBLPrintLog> BLPrintLogDelete(MYBLPrintLog Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLPrintLog> st = cm.BLPrintLogDelete(Data);
            return st;
        }
        [ActionName("BLCorrectorView")]
        public List<ChargeCorrectorInsert> BLCorrectorView(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLCorrectorViewValues(Data);
            return st;
        }
        [ActionName("BLCorrectorCountRecord")]
        public List<ChargeCorrectorInsert> BLCorrectorCountRecord(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.BLCorrectorCountRecordValues(Data);
            return st;
        }
        [ActionName("SlotCorrectorCountRecord")]
        public List<ChargeCorrectorInsert> SlotCorrectorCountRecord(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.SlotCorrectorCountRecordValues(Data);
            return st;
        }
        [ActionName("BOLCusPartyDelete")]
        public List<MYBOL> BOLCusPartyDelete(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLCustomerPartyDelete(Data);
            return st;
        }

        [ActionName("BOLBLNoVlaue")]
        public List<MYBOL> BOLBLNoVlaue()
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.getBOLBLNoValues();
            return st;
        }


        //[ActionName("BOLReleseCheck")]
        //public List<MYBOL> BOLReleseCheck(MYBOL Data)
        //{
        //    DocumentManager cm = new DocumentManager();
        //    List<MYBOL> st = cm.BOLReleasecheckValitationValue(Data);
        //    return st;
        //}

        [ActionName("BOLSurrenderUpdate")]
        public List<MYBOL> BOLSurrenderUpdate(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLSurrederTariffValue(Data);
            return st;
        }

        [ActionName("BOLSurrenderValuesUpdate")]
        public List<MYBOL> BOLSurrenderValuesUpdate(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.BOLSurrenderUpdate(Data);
            return st;
        }

        [ActionName("OnlineBLUnlock")]
        public List<MYBOL> OnlineBLUnlock(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.Online_BL_Unlock_Update(Data);
            return st;
        }

        [ActionName("BLPrintLayout")]
        public List<MYBLRelease> BLPrintLayout(MYBLRelease Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBLRelease> st = cm.BLPrintLayoutMaster(Data);
            return st;
        }

        [ActionName("BkgIDGetvalue")]
        public List<ChargeCorrector> BkgIDGetvalue(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.BkgIDGetMaster(Data);
            return st;
        }

        [ActionName("AgencyChargeCorrectorBrackup")]
        public List<ChargeCorrector> AgencyChargeCorrectorBrackup(ChargeCorrector Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrector> st = cm.AgencyCorrectorChargeBrackupMaster(Data);
            return st;
        }
        [ActionName("ChargeCorrectorBrackupSave")]
        public List<ChargeCorrectorInsert> ChargeCorrectorBrackupSave(ChargeCorrectorInsert Data)
        {
            DocumentManager cm = new DocumentManager();
            List<ChargeCorrectorInsert> st = cm.InsertBrackupChargeCorrector(Data);
            return st;
        }


        [ActionName("TsVoyageSlotContractRate")]
        public List<MyBooking> TsVoyageSlotContractRate(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.TSPortSlotContactRate(Data);
            return st;
        }

        [ActionName("CommisionCorrection")]
        public List<MYBOL> CommisionCorrection(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.CommisionCorrection(Data);
            return st;
        }
        [ActionName("BOLExcelCntrsUpload")]
        public List<MYBOL> BOLExcelCntrsUpload(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.InsertExcelContainersPush(Data);
            return st;
        }

        [ActionName("BOLExcelCustomerUpload")]
        public List<MYBOL> BOLExcelCustomerUpload(MYBOL Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MYBOL> st = cm.InsertExcelCustomerPush(Data);
            return st;
        }


        [ActionName("BookingVesselVoy")]
        public List<MyBooking> BookingVesselVoy(MyBooking Data)
        {
            DocumentManager cm = new DocumentManager();
            List<MyBooking> st = cm.BookingVesselVoyViewValus(Data);
            return st;
        }

    }
}
