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
    public class MNRApiController : ApiController
    {
        #region MNR MASTERS (GANESH)

        [ActionName("DamageMaster")]
        public List<MyDamage> DamageMaster(MyDamage Data)
        {
            MNRManager cm = new MNRManager();
            List<MyDamage> st = cm.InsertDamageMaster(Data);
            return st;
        }

        [ActionName("DamageView")]
        public List<MyDamage> DamageView(MyDamage Data)
        {
            MNRManager cm = new MNRManager();
            List<MyDamage> st = cm.GetDamageMaster(Data);
            return st;
        }


        [ActionName("DamageViewEdit")]
        public List<MyDamage> DamageViewEdit(MyDamage Data)
        {
            MNRManager cm = new MNRManager();
            List<MyDamage> st = cm.GetDamageRecord(Data);
            return st;
        }

        [ActionName("RepairMaster")]
        public List<MyRepair> RepairMaster(MyRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyRepair> st = cm.InsertRepairMaster(Data);
            return st;
        }


        [ActionName("RepairView")]
        public List<MyRepair> RepairView(MyRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyRepair> st = cm.GetRepairMaster(Data);
            return st;
        }


        [ActionName("RepairViewEdit")]
        public List<MyRepair> RepairViewEdit(MyRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyRepair> st = cm.GetRepairRecord(Data);
            return st;
        }

        [ActionName("MNRLocationMaster")]
        public List<MyMNRLoc> MNRLocationMaster(MyMNRLoc Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRLoc> st = cm.InsertMNRLocMaster(Data);
            return st;
        }


        [ActionName("MNRLocationView")]
        public List<MyMNRLoc> MNRLocationView(MyMNRLoc Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRLoc> st = cm.MNRLocationMaster(Data);
            return st;
        }


        [ActionName("MNRLocationViewEdit")]
        public List<MyMNRLoc> MNRLocationViewEdit(MyMNRLoc Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRLoc> st = cm.GetMNRLocationRecord(Data);
            return st;
        }


        [ActionName("BindAssembly")]
        public List<MyComponent> BindAssembly(MyComponent Data)
        {
            MNRManager cm = new MNRManager();
            List<MyComponent> st = cm.GetBindAssembly(Data);
            return st;
        }
        [ActionName("MNRComponentMaster")]
        public List<MyComponent> MNRComponentMaster(MyComponent Data)
        {
            MNRManager cm = new MNRManager();
            List<MyComponent> st = cm.InsertMNRComponentMaster(Data);
            return st;
        }

        [ActionName("MNRComponentView")]
        public List<MyComponent> MNRComponentView(MyComponent Data)
        {
            MNRManager cm = new MNRManager();
            List<MyComponent> st = cm.MNRComponentMaster(Data);
            return st;
        }


        [ActionName("MNRComponentViewEdit")]
        public List<MyComponent> MNRComponentViewEdit(MyComponent Data)
        {
            MNRManager cm = new MNRManager();
            List<MyComponent> st = cm.GetMNRComponentRecord(Data);
            return st;
        }
        [ActionName("MaterialMaster")]
        public List<MyMaterial> MaterialMaster(MyMaterial Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMaterial> st = cm.InsertMaterialMaster(Data);
            return st;
        }

        [ActionName("MaterialMasterView")]
        public List<MyMaterial> MaterialMasterView(MyMaterial Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMaterial> st = cm.GetMaterialMasterView(Data);
            return st;
        }


        [ActionName("MaterialViewEdit")]
        public List<MyMaterial> MaterialViewEdit(MyMaterial Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMaterial> st = cm.GetMaterialRecord(Data);
            return st;
        }
        #endregion

        #region MNR Tariff

        [ActionName("MNRTariffView")]
        public List<MyMNRTariff> MNRTariffView(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.MyMNRTariffMaster(Data);
            return st;
        }

        [ActionName("VendorValues")]
        public List<MyMNRTariff> BindVendorValues(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.GetVendorValues(Data);
            return st;
        }

        [ActionName("MNRTariff")]
        public List<MyMNRTariff> MNRTariff(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.InsertMNRTariff(Data);
            return st;
        }

        [ActionName("MNRTariffEdit")]
        public List<MyMNRTariff> MNRTariffEdit(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.GetMNRTariffValuesData(Data);
            return st;
        }

        [ActionName("MNRTariffDtlsEdit")]
        public List<MyMNRTariff> MNRTariffDtlsEdit(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.GetMNRTariffDtls(Data);
            return st;
        }

        [ActionName("MNRTariffDtlsSearch")]
        public List<MyMNRTariff> MNRTariffDtlsSearch(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.MNRTariffDtlsSearchData(Data);
            return st;
        }

        #endregion

        #region MNR CONTAINER Maintainence Repair

        [ActionName("DepotDropDownByCity")]
        public List<MyDamage> DepotDropDownByCity(MyDamage Data)
        {
            MNRManager cm = new MNRManager();
            List<MyDamage> st = cm.GetDepotDropDownByCity(Data);
            return st;
        }

        [ActionName("BindMNRStatus")]
        public List<MyComponent> BindMNRStatus(MyComponent Data)
        {
            MNRManager cm = new MNRManager();
            List<MyComponent> st = cm.GetMNRStatus(Data);
            return st;
        }
        [ActionName("BindDamage")]
        public List<MyDamage> BindDamage(MyDamage Data)
        {
            MNRManager cm = new MNRManager();
            List<MyDamage> st = cm.GetDamageList(Data);
            return st;
        }

        [ActionName("BindRepair")]
        public List<MyRepair> BindRepair(MyRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyRepair> st = cm.GetRepairList(Data);
            return st;
        }

        [ActionName("BindComponent")]
        public List<MyComponent> BindComponent(MyComponent Data)
        {
            MNRManager cm = new MNRManager();
            List<MyComponent> st = cm.GetComponentList(Data);
            return st;
        }

        [ActionName("BindMaterial")]
        public List<MyMaterial> BindMaterial(MyMaterial Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMaterial> st = cm.GetMaterialList(Data);
            return st;
        }

        [ActionName("BindMNRLocation")]
        public List<MyMNRLoc> BindMNRLocation(MyMNRLoc Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRLoc> st = cm.GetMNRLocationList(Data);
            return st;
        }

        [ActionName("DamageDescriptionByCode")]
        public List<MyDamage> DamageDescriptionByCode(MyDamage Data)
        {
            MNRManager cm = new MNRManager();
            List<MyDamage> st = cm.GetDamageDescriptionByCode(Data);
            return st;
        }

        [ActionName("BindDLCntrNos")]
        public List<MyMNRRepair> BindDLCntrNos(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetBindDLCntrNos(Data);
            return st;
        }

        [ActionName("BindCostToDropdown")]
        public List<MyMNRRepair> BindCostToDropdown(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetCostTo(Data);
            return st;
        }

        [ActionName("BindApproveRejectDropdown")]
        public List<MyMNRRepair> BindApproveRejectDropdown(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetApproveReject(Data);
            return st;
        }

        [ActionName("CurrencyByVendor")]
        public List<MyMNRRepair> CurrencyByVendor(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetCurrencyByVendorValues(Data);
            return st;
        }

        [ActionName("BindCntrNoOnChangeValues")]
        public List<MyMNRRepair> BindCntrNoOnChangeValues(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetCntrNoOnChangeValues(Data);
            return st;
        }
        [ActionName("BindMNRTariffDtlsValuesByDepot")]
        public List<MyMNRTariff> BindMNRTariffDtlsValuesByDepot(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.GetMNRTariffDtlsValuesByDepot(Data);
            return st;
        }

        [ActionName("CheckMNRTariff")]
        public List<MyMNRTariff> CheckMNRTariff(MyMNRTariff Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRTariff> st = cm.CheckMNRTariffValues(Data);
            return st;
        }

        [ActionName("MNRRepairReq")]
        public List<MyMNRRepair> MNRRepairReq(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.InsertMNRRepairReq(Data);
            return st;
        }

        [ActionName("MNRRepairReqView")]
        public List<MyMNRRepair> MNRRepairReqView(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRRepairReqViewData(Data);
            return st;
        }

        [ActionName("MNRRepairReqPreviewGrid1")]
        public List<MyMNRRepair> MNRRepairReqPreviewGrid1(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRRepairReqPreviewGrid1Data(Data);
            return st;
        }
        

        [ActionName("MNRRepairReqPreviewCostDtls")]
        public List<MyMNRRepair> MNRRepairReqPreviewCostDtls(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRRepairReqPreviewCostDtlsData(Data);
            return st;
        }
        [ActionName("MNRRepairReqPreviewApprCostDtls")]
        public List<MyMNRRepair> MNRRepairReqPreviewApprCostDtls(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRRepairReqPreviewApprCostDtlsData(Data);
            return st;
        }
        [ActionName("MNRRepairReqApprovCostDtls")]
        public List<MyMNRRepair> MNRRepairReqApprovCostDtls(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.InsertMNRRepairReqApprovCostDtls(Data);
            return st;
        }

        [ActionName("MNREstApprTotalView")]
        public List<MyMNRRepair> MNREstApprTotalView(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNREstApprTotalViewData(Data);
            return st;
        }

        [ActionName("BindMNRAttachTypeDropdown")]
        public List<MyMNRRepair> BindMNRAttachTypeDropdown(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetMNRAttachType(Data);
            return st;
        }

        [ActionName("MNRAttachments")]
        public List<MyMNRRepair> MNRAttachments(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.InsertMNRAttachments(Data);
            return st;
        }

        [ActionName("MNRAttachmentsView")]
        public List<MyMNRRepair> MNRAttachmentsView(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetMNRAttachmentsView(Data);
            return st;
        }

        [ActionName("MNRApprovedDtlsView")]
        public List<MyMNRRepair> MNRApprovedDtlsView(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.GetMNRApprovedDtlsView(Data);
            return st;
        }

        [ActionName("MNREORConfirmUpdateApprDtls")]
        public List<MyMNRRepair> MNREORConfirmUpdateApprDtls(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();

            List<MyMNRRepair> st = new List<MyMNRRepair>();
            DataTable dt = cm.UpdateMNREORConfirmApprDtls(Data);
            return st;
        }

        [ActionName("MNRApprove")]
        public List<MyMNRRepair> MNRApprove(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager(); 

            List<MyMNRRepair> st = cm.UpdateMNRApprove(Data);

            return st;
        }

        [ActionName("MNRApprovedBLCargoDetails")]
        public List<MyMNRRepair> MNRApprovedBLCargoDetails(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRApprovedBLCargoDetailsData(Data);
            return st;
        }

        [ActionName("MNRSurveyorView")]
        public List<MyMNRRepair> MNRSurveyorView(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRSurveyorViewData(Data);
            return st;
        }

        [ActionName("MNRCntrRepairHistoryDetails")]
        public List<MyMNRRepair> MNRCntrRepairHistoryDetails(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRCntrRepairHistoryDetailsData(Data);
            return st;
        }

        [ActionName("MNRCntrCurrentStatusView")]
        public List<MyMNRRepair> MNRCntrCurrentStatusView(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRCntrCurrentStatusViewData(Data);
            return st;
        }
        [ActionName("MNRDashBoardCount")]
        public List<MyMNRRepair> MNRDashBoardCount(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();
            List<MyMNRRepair> st = cm.MNRDashBoardCountData(Data);
            return st;
        }

        [ActionName("MNRStatusChangeToAV")]
        public List<MyMNRRepair> MNRStatusChangeToAV(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();

            List<MyMNRRepair> st = cm.UpdateMNRStatusChangeToAV(Data);

            return st;
        }

        [ActionName("MNRRepair")]
        public List<MyMNRRepair> MNRRepair(MyMNRRepair Data)
        {
            MNRManager cm = new MNRManager();

            List<MyMNRRepair> st = cm.UpdateMNRRepair(Data);

            return st;
        }
        #endregion
    }
}
