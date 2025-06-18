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
    public class ImportApiController : ApiController
    {
        #region anand
        [ActionName("ImportView")]
        public List<MyImport> ImportView(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.GetImportMaster(Data);
            return st;
        }

        [ActionName("BkgNoDropDown")]
        public List<MyImport> BkgNoDropDown(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindBookingNo(Data);
            return st;
        }
        [ActionName("BkgExistingDtls")]
        public List<MyImport> BkgExistingDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindBookingValues(Data);
            return st;
        }

        [ActionName("BkgExistingModeDtls")]
        public List<MyImport> BkgExistingModeDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindBookingModeValues(Data);
            return st;
        }


        [ActionName("BOLCntrImportViewRecord")]
        public List<MYBOL> BOLCntrImportViewRecord(MYBOL Data)
        {
            ImportManager cm = new ImportManager();
            List<MYBOL> st = cm.ImpBOLCntrExistingValus(Data);
            return st;
        }

        [ActionName("ImportCntrPickupRecord")]
        public List<MYBOL> ImportCntrPickupRecord(MYBOL Data)
        {
            ImportManager cm = new ImportManager();
            List<MYBOL> st = cm.ImportCntrPickupDetailsValus(Data);
            return st;
        }



        [ActionName("ImpVesVoyDtls")]
        public List<MyImport> ImpVesVoyDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindVesVoyMaster(Data);
            return st;
        }
        [ActionName("ImpCntrDtls")]
        public List<MyImport> ImpCntrDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindImpCntrMaster(Data);
            return st;
        }

        [ActionName("CntrDtls")]
        public List<MyImport> CntrDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindCntrMaster(Data);
            return st;
        }
        [ActionName("PreAlertDtls")]
        public List<MyImport> PreAlertDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindPreAlertMaster(Data);
            return st;
        }
        [ActionName("CANDtls")]
        public List<MyImport> CANDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindCANMaster(Data);
            return st;
        }

        [ActionName("DeliveryOrderDtls")]
        public List<MyImport> DeliveryOrderDtls(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindDeliveryOrderMaster(Data);
            return st;
        }

        [ActionName("InsertImpBooking")]
        public List<MyImpBooking> InsertImpBooking(MyImpBooking Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpBooking> st = cm.InsertImpBooking(Data);
            return st;
        }
        [ActionName("ImpBookingDetails")]
        public List<MyImpBooking> ImpBookingDetails(MyImpBooking Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpBooking> st = cm.BindImpBooking(Data);
            return st;
        }


        [ActionName("ExistingDOGridValues")]
        public List<MyImpBooking> ExistingDOGridValues(MyImpBooking Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpBooking> st = cm.ExistingDOGirdBooking(Data);
            return st;
        }

        [ActionName("ExisDOValues")]
        public List<MYImpDeliveryOrder> ExisDOValues(MYImpDeliveryOrder Data)
        {
            ImportManager cm = new ImportManager();
            List<MYImpDeliveryOrder> st = cm.ExisImpDOvalues(Data);
            return st;
        }


        [ActionName("InsertImpCAN")]
        public List<MyImpCAN> InsertImpCAN(MyImpCAN Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpCAN> st = cm.InsertImpCAN(Data);
            return st;
        }
        [ActionName("ImpCANDetails")]
        public List<MyImpCAN> ImpCANDetails(MyImpCAN Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpCAN> st = cm.BindImpCAN(Data);
            return st;
        }

        [ActionName("ImpContainerDetails")]
        public List<MyImpCAN> ImpContainerDetails(MyImpCAN Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpCAN> st = cm.BindImpContainerDetails(Data);
            return st;
        }

        [ActionName("ImpDOContainerDetails")]
        public List<MyImpCAN> ImpDOContainerDetails(MyImpCAN Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpCAN> st = cm.BindImpDOContainerDetails(Data);
            return st;
        }

        [ActionName("ExistingDocheckValuesInsert")]
        public List<MYImpDeliveryOrder> ExistingDocheckValuesInsert(MYImpDeliveryOrder Data)
        {
            ImportManager cm = new ImportManager();
            List<MYImpDeliveryOrder> st = cm.ExisingDOCheckBeforeInsert(Data);
            return st;
        }


       

        [ActionName("ImpDOExistContainerDetails")]
        public List<MyImpCAN> ImpDOExistContainerDetails(MyImpCAN Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImpCAN> st = cm.BindImpExistDOContainerDetails(Data);
            return st;
        }



        [ActionName("InsertImpDO")]
        public List<MYImpDeliveryOrder> InsertImpDO(MYImpDeliveryOrder Data)
        {
            ImportManager cm = new ImportManager();
            List<MYImpDeliveryOrder> st = cm.InsertImpDO(Data);
            return st;
        }
        [ActionName("ImpDOExtUpdate")]
        public List<MYImpDeliveryOrder> ImpDOExtUpdate(MYImpDeliveryOrder Data)
        {
            ImportManager cm = new ImportManager();
            List<MYImpDeliveryOrder> st = cm.UpdateDOExistingImpDO(Data);
            return st;
        }


        [ActionName("ImpDODetails")]
        public List<MYImpDeliveryOrder> ImpDODetails(MYImpDeliveryOrder Data)
        {
            ImportManager cm = new ImportManager();
            List<MYImpDeliveryOrder> st = cm.BindImpDO(Data);
            return st;
        }


        [ActionName("ImpPickcntrBookingMaster")]
        public List<MyCntrPickupdtls> ImpPickcntrBookingMaster(MyCntrPickupdtls Data)
        {
            ImportManager Mange = new ImportManager();
            List<MyCntrPickupdtls> st = Mange.ImpPickcntrBooking(Data);
            return st;
        }
        #endregion anand

        #region Ganesh

        [ActionName("BindNominationCFSTypes")]
        public List<MyImport> BindNominationCFSTypes(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindNominationCFSValues(Data);
            return st;
        }



        [ActionName("CustomerCFSMaster")]
        public List<MyImport> CustomerCFSMaster(MyImport Data)
        {
            ImportManager cm = new ImportManager();
            List<MyImport> st = cm.BindCustomerCFSValues(Data);
            return st;
        }

        #endregion



        [ActionName("ImpDETContainerWise")]
        public List<MYDETCalculation> ImpDETContainerWise(MYDETCalculation Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYDETCalculation> st = Mange.ImpDETContainer(Data);
            return st;
        }


        [ActionName("ImpDETCalculationSlabe")]
        public List<MYDETCalculation> ImpDETCalculationSlabe(MYDETCalculation Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYDETCalculation> st = Mange.ImpDETCalculationSlab(Data);
            return st;
        }


        [ActionName("ImpDETContainerWiseSearch")]
        public List<MYDETCalculation> ImpDETContainerWiseSearch(MYDETCalculation Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYDETCalculation> st = Mange.ImpDETContainerSearch(Data);
            return st;
        }

        [ActionName("ImpDDGInsert")]
        public List<MYImpDDG> ImpDDGInsert(MYImpDDG Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYImpDDG> st = Mange.ImportDDGInsert(Data);
            return st;
        }

        [ActionName("ImportImpDDGInsert")]
        public List<MYImpDDG> ImportImpDDGInsert(MYImpDDG Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYImpDDG> st = Mange.ImportCalculationDDGInsert(Data);
            return st;
        }

        

        [ActionName("ImpDDGvalueInsertInvoiceTable")]
        public List<MYImpDDG> ImpDDGvalueInsertInvoiceTable(MYImpDDG Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYImpDDG> st = Mange.ImportInsertDTValueInvoice(Data);
            return st;
        }

        [ActionName("ExpEstimateexsitingValues")]
        public List<MYImpDDG> ExpEstimateexsitingValues(MYImpDDG Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYImpDDG> st = Mange.ExportExistingEstimatedisplay(Data);
            return st;
        }


        [ActionName("DisplayEstimateexsitingValues")]
        public List<DDGRoot> DisplayEstimateexsitingValues(DDGRoot Data)
        {
            ImportManager Mange = new ImportManager();
            // List<MYDETCalculation> st = Mange.DisplayExistingEstimateValues(Data);
            List<DDGRoot> st = Mange.NewDisplayExistingEstimateValues(Data);
            return st;
        }



        //[ActionName("ExportDDGInsert")]
        //public List<MYImpDDG> ExportDDGInsert(MYImpDDG Data)
        //{
        //    ImportManager Mange = new ImportManager();
        //    List<MYImpDDG> st = Mange.InsertExpDetDmmGrd(Data);
        //    return st;
        //}

        [ActionName("IMPDDGBilledAmount")]
        public List<MYDETCalculation> IMPDDGBilledAmount(MYDETCalculation Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYDETCalculation> st = Mange.ImpDDGBillingAmount(Data);
            return st;
        }

        [ActionName("ImpExistingVesVoyDetails")]
        public List<MyImpBooking>ImpExistingVesVoyDetails(MyImpBooking Data)
        {
            ImportManager Mange = new ImportManager();
            List<MyImpBooking> st = Mange.BOLVesselVoyageValus(Data);
            return st;
        }



        [ActionName("ImpDestination")]
        public List<MyRRRate> ImpDestination(MyRRRate Data)
        {
            ImportManager Mange = new ImportManager();
            List<MyRRRate> st = Mange.ImpDestinationExisting(Data);
            return st;
        }


        [ActionName("ImpBookingCntrPickup")]
        public List<MyCntrPickup> ImpBookingCntrPickup(MyCntrPickup Data)
        {
            ImportManager Mange = new ImportManager();
            List<MyCntrPickup> st = Mange.ImpBkgPickupContainers(Data);
            return st;
        }


        [ActionName("ImportCheckDtMovmentLast")]
        public List<MyCntrPickup> ImportCheckDtMovmentLast(MyCntrPickup Data)
        {
            ImportManager Mange = new ImportManager();
            List<MyCntrPickup> st = Mange.DtImpLastMovment(Data);
            return st;
        }

        [ActionName("ImpInsertCntrLastMovmentMaster")]
        public List<MyCntrPickup> ImpInsertCntrLastMovmentMaster(MyCntrPickup Data)
        {
            ImportManager Mange = new ImportManager();
            List<MyCntrPickup> st = Mange.InsertImpCntrLastMovment(Data);
            return st;
        }


        [ActionName("ImpDDGDelete")]
        public List<MYImpDDG> ImpDDGDelete(MYImpDDG Data)
        {
            ImportManager Mange = new ImportManager();
            List<MYImpDDG> st = Mange.ImportDDGDelete(Data);
            return st;
        }

    }
}
