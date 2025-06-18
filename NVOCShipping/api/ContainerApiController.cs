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
    public class ContainerApiController : ApiController
    {

        #region Ganesh 

        #region Containers
        [ActionName("CntrTypeSize")]
        public List<MyContainerData> CntrTypeSizeList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListCntrTypeSize(Data);
            return st;
        }

        [ActionName("ISOCodes")]
        public List<MyContainerData> ISOCodesList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListISOCodes(Data);
            return st;
        }

        [ActionName("Grade")]
        public List<MyContainerData> GradeList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListGrade(Data);
            return st;
        }
        [ActionName("LeaseTerm")]
        public List<MyContainerData> LeaseTermList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListLeaseTerm(Data);
            return st;
        }

        [ActionName("BoxOwner")]
        public List<MyContainerData> BoxOwnerList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListBoxOwner(Data);
            return st;
        }
        [ActionName("LeasingPartner")]
        public List<MyContainerData> LeasingPartnerList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListLeasingPartner(Data);
            return st;
        }

        [ActionName("ContainerStatus")]
        public List<MyContainerData> ContainerStatusList(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ListContainerStatus(Data);
            return st;
        }
        [ActionName("InsertContainers")]
        public List<MyContainerData> InsertContainers(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.InsertContainerMaster(Data);
            return st;
        }
        [ActionName("ContainersValidation")]
        public List<MyContainerData> ContainersValidation(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ContainersValidation(Data);
            return st;
        }
        [ActionName("ViewContainers")]
        public List<MyContainerData> ViewContainers(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.ContainersMaster(Data);
            return st;
        }

        [ActionName("ContainerEdit")]
        public List<MyContainerData> ContainerEdit(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.GetContainerMasterRecord(Data);
            return st;
        }

        #endregion

        #region Lease Contract
        [ActionName("DepotValues")]
        public List<MyLease> DepotValues()
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.DepotMaster();
            return st;
        }

        [ActionName("DepotByPort")]
        public List<MyLease> DepotByPortValues(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.DepotByPortMaster(Data);
            return st;
        }
        [ActionName("CityValues")]
        public List<MyLease> CityValues()
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.CityMaster();
            return st;
        }

        [ActionName("PickUpCriteria")]
        public List<MyLease> PickUpCriteriaList(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.ListPickUpCriteria(Data);
            return st;
        }
        [ActionName("InsertLeasingContract")]
        public List<MyLease> InsertLeasingContract(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.InsertLeasingContractMaster(Data);
            return st;
        }

        [ActionName("InsertLeaseDetails")]
        public List<MyLeaseDetails> InsertLeaseDetails(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = cm.InsertLeaseDetails(Data);
            return st;
        }


        [ActionName("ViewLeaseContract")]
        public List<MyLease> ViewLeaseContract(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.LeaseContractMaster(Data);
            return st;
        }
        [ActionName("LeaseContractEdit")]
        public List<MyLease> LeaseContractEdit(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.GetLeaseContractRecord(Data);
            return st;
        }

        [ActionName("ViewLeaseDetails")]
        public List<MyLeaseDetails> ViewLeaseDetails(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = cm.LeaseDetails(Data);
            return st;
        }
        [ActionName("LeaseDetailsEdit")]
        public List<MyLeaseDetails> LeaseDetailsEdit(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = cm.GetLeaseDetailsEdit(Data);
            return st;
        }

        [HttpPost, ActionName("LeaseDetailsDelete")]
        public List<MyLeaseDetails> LeaseDetailsDelete(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = new List<MyLeaseDetails>();
            DataTable dt = cm.DeleteLeaseDetails(Data);
            return st;
        }


        [ActionName("LeasePickUpDtlsEdit")]
        public List<MyLease> LeasePickUpDtls(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.GetLeasePickUpDtls(Data.LeaseContractID.ToString());
            return st;
        }
        [ActionName("LeaseDropUpDtlsEdit")]
        public List<MyLease> LeaseDropUpDtls(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.GetLeaseDropUpDtls(Data.LeaseContractID.ToString());
            return st;
        }

        [ActionName("InsertDropoffLocs")]
        public List<MyLeaseDetails> InsertDropoffLocs(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = cm.InsertDropoffLocs(Data);
            return st;
        }
        [ActionName("LeaseDropoffLocsEdit")]
        public List<MyLeaseDetails> LeaseDropUpDtls(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = cm.GetLeaseDropoffLocsEdit(Data);
            return st;
        }


        [ActionName("DropOffLocsDelete")]
        public List<MyLeaseDetails> DropOffLocsDelete(MyLeaseDetails Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLeaseDetails> st = cm.DropOffLocsDelete(Data);
            return st;
        }
        
        #endregion

        #region ON-HIRE REQUEST

        [ActionName("PickUpReference")]
        public List<MyOnHire> PickUpReferenceList(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = cm.ListPickUpReference(Data);
            return st;
        }

        [ActionName("CntrTypesFromPickUp")]
        public List<MyLease> CntrTypesFromPickUp(MyLease Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyLease> st = cm.GetCntrTypesFromPickUp(Data.LeaseContractID.ToString());
            return st;
        }



        [ActionName("InsertOnHireRequest")]
        public List<MyOnHire> InsertOnHireRequest(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = cm.InsertOnHireRequestMaster(Data);
            return st;
        }

        [ActionName("ApplicableAtList")]
        public List<MyOnHire> ApplicableAtList(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = cm.ListApplicableAtList(Data);
            return st;
        }
        [ActionName("ViewOnHireRequest")]
        public List<MyOnHire> ViewOnHireRequest(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = cm.OnHireRequest(Data);
            return st;
        }

        [ActionName("OnHireRequestEdit")]
        public List<MyOnHire> OnHireRequestEdit(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = cm.OnHireRequestEditValues(Data);
            return st;
        }

        [ActionName("LeaseExistingContainersBind")]
        public List<MyContainerData> LeaseExistingContainersEdit(MyContainerData Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerData> st = cm.LeaseExistingContainersValues(Data);
            return st;
        }

        [ActionName("ApproveOnHire")]
        public List<MyOnHire> ApproveOnHire(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = new List<MyOnHire>();
            DataTable dt = cm.ApproveOnHire(Data);
            return st;
        }

        [ActionName("CheckISApproveOnHire")]
        public List<MyOnHire> CheckISApproveOnHire(MyOnHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOnHire> st = cm.CheckISApproveOnHire(Data);
            return st;
        }


        #endregion

        #region  OFF-Hire Request
        [ActionName("BindCntrNos")]
        public List<MyOffHire> BindCntrNoslist(MyOffHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOffHire> st = cm.ListBindCntrNos(Data);
            return st;
        }

        [ActionName("BindOffHireCntrValues")]
        public List<MyOffHire> BindOffHireCntrValues(MyOffHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOffHire> st = cm.ListBindOffHireCntrValues(Data);
            return st;
        }

        [ActionName("BindDropOffCurrAmtFromLease")]
        public List<MyOffHire> BindDropOffCurrAmtFromLease(MyOffHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOffHire> st = cm.ListBindDropOffCurrAmtFromLease(Data);
            return st;
        }

        [ActionName("InsertOffHireRequest")]
        public List<MyOffHire> InsertOffHireRequest(MyOffHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOffHire> st = cm.InsertOffHireRequest(Data);
            return st;
        }
        [ActionName("ViewOffHireRequest")]
        public List<MyOffHire> ViewOffHireRequest(MyOffHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOffHire> st = cm.OffHireRequest(Data);
            return st;
        }
        [ActionName("OffHireEdit")]
        public List<MyOffHire> OffHireEdit(MyOffHire Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyOffHire> st = cm.GetOffHireRecord(Data);
            return st;
        }


        #endregion

        #region Container Movement Entry

        [ActionName("ListDepotByGeoLoc")]
        public List<MyCntrMoveMent> ListDepotByGeoLoc(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ListDepotByGeoLocValus(Data);
            return st;
        }

        [ActionName("BookingBLList")]
        public List<MyCntrMoveMent> BookingBLList(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.BookingBLList(Data);
            return st;
        }
        [ActionName("BookingBLListAgentwise")]
        public List<MyCntrMoveMent> BookingBLListAgentwise(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.BookingBLListAgentwise(Data);
            return st;
        }
        [ActionName("CntrStatusCodes")]
        public List<MyCntrMoveMent> CntrStatusCodes(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ListCntrStatusCodes(Data);
            return st;
        }
        [ActionName("TransitModeValues")]
        public List<MyCntrMoveMent> TransitModes(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ListTransitModes(Data);
            return st;
        }
        [ActionName("PartyList")]
        public List<MyCntrMoveMent> CustomerMaster()
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.CustomerMaster();
            return st;
        }
        [ActionName("StatusCodesPossibleMoves")]
        public List<MyCntrMoveMent> StatusCodesPossibleMoves(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ListCntrStatusCodesPossibleMoves(Data);
            return st;
        }

        [ActionName("InsertContainersTxns")]
        public List<MyCntrMoveMent> InsertContainersTxns(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.InsertContainerTxnsData(Data);
            return st;
        }

        [ActionName("ContainersTxnsView")]
        public List<MyCntrMoveMent> ContainersTxnsView(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ContainersTxnsViewValues(Data);
            return st;
        }

        [ActionName("ContainersTxnEntryView")]
        public List<MyCntrMoveMent> ContainersTxnEntryView(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ContainersTxnEntryViewValues(Data);
            return st;
        }

        [ActionName("CheckStatusCodeValidation")]
        public List<MyStatusCode> CheckStatusCodeValidation(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.GetCheckStatusCodeValidation(Data);
            return st;
        }
        #endregion

        #region Container Tracking
        [ActionName("ContainersTrackingView")]
        public List<MyCntrMoveMent> ContainersTrackingView(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.ContainersTrackingViewValues(Data);
            return st;
        }

        [ActionName("AmendmentEdit")]
        public List<MyCntrMoveMent> AmendmentEditValues(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.InvAmendmentEditValues(Data);
            return st;
        }

        [ActionName("ValidateAgencyLastMove")]
        public List<MyCntrMoveMent> ValidateAgencyLastMove(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.InventoryValidateAgencyLastMove(Data);
            return st;
        }

        [ActionName("AmendmentUpdateCntrTxns")]
        public List<MyCntrMoveMent> AmendmentUpdateCntrTxns(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.InventoryAmendmentUpdateCntrTxns(Data);
            return st;
        }


        [ActionName("LastMoveRemoveandUpdate")]
        public List<MyCntrMoveMent> LastMoveRemoveandUpdate(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.LastMoveRemoveandUpdate(Data);
            return st;
        }

        #endregion

        #region Invent statuscode creation



        [ActionName("BindStatusCodeList")]
        public List<MyStatusCode> BindStatusCodeList(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.GetStatusCodeList(Data);
            return st;
        }

        [ActionName("StatusDescriptionByStatusCode")]
        public List<MyStatusCode> BindStatusDescriptionByStatusCode(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.GetStatusDescriptionByStatusCode(Data);
            return st;
        }

        [ActionName("ViewInventoryStatusCodes")]
        public List<MyStatusCode> ViewInventoryStatusCodes(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.ViewInventoryStatusCodes(Data);
            return st;
        }

        [ActionName("StatusCodesEdit")]
        public List<MyStatusCode> StatusCodesEdit(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.GetStatusCodesEdit(Data);
            return st;
        }

        [ActionName("BindStatusCodePossiblemoves")]
        public List<MyStatusCode> BindStatusCodePossiblemoves(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.GetBindStatusCodePossiblemoves(Data);
            return st;
        }
        [ActionName("InsertContainerStatusCreation")]
        public List<MyStatusCode> InsertContainerStatusCreation(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.InsertContainerStatusCreation(Data);
            return st;
        }
        [ActionName("BindStatusCodeDesc")]
        public List<MyStatusCode> BindStatusCodeDesc(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.GetBindStatusCodeDesc(Data);
            return st;
        }
        [ActionName("NextMoveDelete")]
        public List<MyStatusCode> NextMoveDelete(MyStatusCode Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyStatusCode> st = cm.NextMoveDelete(Data);
            return st;
        }
        #endregion

        #region Single Movement Entry

        [ActionName("SingleContainersTrackingView")]
        public List<MyCntrMoveMent> SingleContainersTrackingView(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.SingleContainersTrackingViewValues(Data);
            return st;
        }

        [ActionName("CntrStatusandTypeByCntrNo")]
        public List<MyCntrMoveMent> CntrStatusandTypeByCntrNo(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.CntrStatusandTypeByCntrNo(Data);
            return st;
        }
        [ActionName("CntrDtMovementValidation")]
        public List<MyCntrMoveMent> CntrDtMovementValidation(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.CntrDtMovementValidation(Data);
            return st;
        }
        [ActionName("SingleInsertContainersTxns")]
        public List<MyCntrMoveMent> SingleInsertContainersTxns(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.SingleInsertContainerTxnsData(Data);
            return st;
        }
        [ActionName("TrackingCntrDtMovementValidation")]
        public List<MyCntrMoveMent> TrackingCntrDtMovementValidation(MyCntrMoveMent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyCntrMoveMent> st = cm.TrackingCntrDtMovementValidation(Data);
            return st;
        }
        
        #endregion

        #endregion

        #region anand
        [ActionName("ContainerRent")]
        public List<MyContainerRent> ContainerRent(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.InsertContainerRentContract(Data);
            return st;
        }

        [ActionName("StorageContract")]
        public List<MyContainerRent> StorageContract(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.InsertStorageContract(Data);
            return st;
        }

        [ActionName("ContainerRentView")]
        public List<MyContainerRent> ContainerRentView(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.ContainerRentViewMaster(Data);
            return st;
        }
        [ActionName("ContainerRentParticular")]
        public List<MyContainerRent> ContainerRentParticular(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.GetContainerRentRecordMaster(Data.ID.ToString());
            return st;
        }
        [ActionName("ContainerRentTariffDtls")]
        public List<MyContainerRent> ContainerRentTariffDtls(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.GetContainerRentTariffRecordMaster(Data);
            return st;
        }
        [ActionName("RentTariffDelete")]
        public List<MyContainerRent> RentTariffDelete(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.GetRentTariffDelete(Data);
            return st;
        }


        [ActionName("StorageContractEdit")]
        public List<MyContainerRent> StorageContractEdit(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.GetStorageContractEditMaster(Data.ID.ToString());
            return st;
        }

        [ActionName("StorageContractTariffDtls")]
        public List<MyContainerRent> StorageContractTariffDtls(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.GetStorageContractEditTariffMaster(Data);
            return st;
        }

        [ActionName("StorageTariffDelete")]
        public List<MyContainerRent> StorageTariffDelete(MyContainerRent Data)
        {
            ContainerManager cm = new ContainerManager();
            List<MyContainerRent> st = cm.GetStorageTariffDelete(Data);
            return st;
        }
        #endregion

    }
}
