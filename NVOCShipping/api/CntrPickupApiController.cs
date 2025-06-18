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
    public class CntrPickupApiController : ApiController
    {

        
        [ActionName("Bookingdtls")]
        public List<MyCntrPickup> Bookingdtls(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.BookingDtlsValues(Data);
            return st;
        }

        [ActionName("BookingCntrTypes")]
        public List<MyCntrPickup> BookingCntrTypes(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.BookingCntrTypesValues(Data);
            return st;
        }

        [ActionName("BookingCntrPickup")]
        public List<MyCntrPickup> BookingCntrPickup(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.BkgPickupContainers(Data);
            return st;
        }

        [ActionName("BookingCntrPickupSearch")]
        public List<MyCntrPickup> BookingCntrPickupSearch(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.BkgPickupContainersSearch(Data);
            return st;
        }

        [ActionName("InsertCntrPickupMaster")]
        public List<MyCntrPickup> InsertCntrPickupMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrPickup(Data);
            return st;
        }

        [ActionName("InsertCntrPortInMaster")]
        public List<MyCntrPickup> InsertCntrPortInMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrPortIN(Data);
            return st;
        }

        [ActionName("InsertCntrRoadToPortMaster")]
        public List<MyCntrPickup> InsertCntrRoadToPortMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrRoadToPort(Data);
            return st;
        }


        [ActionName("InsertCntrRailToPortMaster")]
        public List<MyCntrPickup> InsertCntrRailToPortMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrRailToPort(Data);
            return st;
        }

        [ActionName("InsertCntrDepaturePortMaster")]
        public List<MyCntrPickup> InsertCntrDepaturePortMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrDepaturePort(Data);
            return st;
        }

        [ActionName("InsertCntrDepaturePortMasterMB")]
        public List<MyCntrPickup> InsertCntrDepaturePortMasterMB(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrDepaturePortMB(Data);
            return st;
        }

        [ActionName("ExtPickcntrBookingMaster")]
        public List<MyCntrPickupdtls> ExtPickcntrBookingMaster(MyCntrPickupdtls Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickupdtls> st = Mange.ExtPickcntrBooking(Data);
            return st;
        }

        [ActionName("ExtPickcntrBlReleaseMaster")]
        public List<MyCntrPickupdtls> ExtPickcntrBlReleaseMaster(MyCntrPickupdtls Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickupdtls> st = Mange.ExtPickcntrBLRelease(Data);
            return st;
        }

        [ActionName("FileDeleteCntrPickupdtls")]
        public List<MyCntrPickup> FileDeleteCntrPickupdtls(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.DeleteCntrPickup(Data);
            return st;
        }


        [ActionName("BkgCROReleaseOrderNo")]
        public List<MyCntrPickup> BkgCROReleaseOrderNo(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.BkgCROReleaseOrderNo(Data);
            return st;
        }

        [ActionName("InsertTransitFIMaster")]
        public List<MyCntrPickup> InsertTransitFIMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCMTransitFI(Data);
            return st;
        }

        [ActionName("InsertCMTransitFVMaster")]
        public List<MyCntrPickup> InsertCMTransitFVMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCMTransitFV(Data);
            return st;
        }

        [ActionName("InsertCMPortOutFUMaster")]
        public List<MyCntrPickup> InsertCMPortOutFUMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCMPortOutFU(Data);
            return st;
        }

        [ActionName("InsertCMDepoInMAMaster")]
        public List<MyCntrPickup> InsertCMDepoInMAMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCMDepoInMA(Data);
            return st;
        }


        [ActionName("InsertCMDischargeFVMaster")]
        public List<MyCntrPickup> InsertCMDischargeFVMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCMDischargeFV(Data);
            return st;
        }

        [ActionName("InsertCntrLastMovmentMaster")]
        public List<MyCntrPickup> InsertCntrLastMovmentMaster(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrLastMovment(Data);
            return st;
        }

        [ActionName("InsertCntrLastMovmentMADL")]
        public List<MyCntrPickup> InsertCntrLastMovmentMADL(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrLastMovmentMADL(Data);
            return st;
        }


        [ActionName("ImportCheckDtMovmentLast")]
        public List<MyCntrPickup> ImportCheckDtMovmentLast(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.DtImpLastMovment(Data);
            return st;
        }


        [ActionName("InsertFLICDMovement")]
        public List<MyCntrPickup> InsertFLICDMovement(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertFLICDMovement(Data);
            return st;
        }

        [ActionName("InsertCntrPLMovement")]
        public List<MyCntrPickup> InsertCntrPLMovement(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertCntrPLMovement(Data);
            return st;
        }


        [ActionName("InsertCntrFVCFSMovement")]
        public List<MyCntrPickup> InsertCntrFVCFSMovement(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertFVCFSMovement(Data);
            return st;
        }



        [ActionName("InsertCntrDVMovement")]
        public List<MyCntrPickup> InsertCntrDVMovement(MyCntrPickup Data)
        {
            ContainerPickupManager Mange = new ContainerPickupManager();
            List<MyCntrPickup> st = Mange.InsertDVMovement(Data);
            return st;
        }

    }
}