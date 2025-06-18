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
    public class PartyApiController : ApiController
    {
        #region anand
        [ActionName("Customer")]
        public List<MyCustomer> Customer(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertCustomerMaster(Data);
            return st;
        }

        [ActionName("BOLCustomer")]
        public List<MyCustomer> BOLCustomer(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertBOLCustomerMaster(Data);
            return st;
        }


        [ActionName("CustomerSubmitAll")]
        public List<MyCustomer> CustomerSubmitAll(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.CheckCustomerValues(Data);
            return st;
        }
        [ActionName("CustomerView")]
        public List<MyCustomer> CustomerView(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCustMaster(Data);
            return st;
        }
        [ActionName("CustomerViewParticular")]
        public List<MyCustomer> CustomerViewParticular(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCustMasterRecord(Data.ID.ToString());
            return st;
        }

        [ActionName("CusBranchUpdate")]
        public List<MyCustomer> CusBranchUpdate(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusBranchMasterUpdate(Data);
            return st;
        }

        [ActionName("BusinessTypes")]
        public List<MyCustomer> BusinessTypes(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetBusinessTypeMaster(Data);
            return st;
        }

        [ActionName("BindBusinessTypes")]
        public List<MYCustomerBuss> BindBusinessTypes(MYCustomerBuss Data)
        {
            PartyManager cm = new PartyManager();
            List<MYCustomerBuss> st = cm.GetBusinessMasterRecord(Data);
            return st;
        }

        [ActionName("CustomerBranch")]
        public List<MyCustomer> CustomerBranch(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertCustomerLocation(Data);
            return st;
        }

        [ActionName("BranchDropDown")]
        public List<MyCustomer> BranchDropDown(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.BindBranchDropDown(Data);
            return st;
        }

        [ActionName("CustomerBranchView")]
        public List<MyCustomer> CustomerBranchValues(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusBranchMaster(Data);
            return st;
        }

        [ActionName("CusBranchDelete")]
        public List<MyCustomer> CusBranchDelete(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.CusBranchDeleteMaster(Data);
            return st;
        }

        [ActionName("CustomerEmailAlert")]
        public List<MyCustomer> CustomerEmailAlert(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertCustomerEmailAlerts(Data);
            return st;
        }

        [ActionName("CustomerEmailAlertView")]
        public List<MyCustomer> CustomerEmailAlertView(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusEmailAlertsMaster(Data);
            return st;
        }



        [ActionName("CusEmailAlertUpdate")]
        public List<MyCustomer> CusEmailAlertUpdate(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusEmailAltMasterUpdate(Data);
            return st;
        }

        [ActionName("CusEmailAlertDelete")]
        public List<MyCustomer> CusEmailAlertDelete(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.CusEmailAlertsDeleteMaster(Data);
            return st;
        }

        [ActionName("CustomerAddInfo")]
        public List<MyCustomer> CustomerAddInfo(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertCusAddInfoMaster(Data);
            return st;
        }

        [ActionName("CustomerAddInfoView")]
        public List<MyCustomer> CustomerAddInfoView(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusAddInfoMaster(Data);
            return st;
        }

        [ActionName("CusAddInfoUpdate")]
        public List<MyCustomer> CusAddInfoUpdate(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusAddInfoUpdateMaster(Data);
            return st;
        }

        [ActionName("CusAddInfoDelete")]
        public List<MyCustomer> CusAddInfoDelete(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.CusAddInfoDeleteMaster(Data);
            return st;
        }

        [ActionName("AgencyState")]
        public List<MyCustomer> AgencyState(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetAgencyStateMaster(Data);
            return st;
        }
        [ActionName("AgencyMaster")]
        public List<MyCustomer> AgencyMaster(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetAgencyMaster(Data);
            return st;
        }
        [ActionName("CustomerSalesLink")]
        public List<MyCustomer> CustomerSalesLink(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertCusSalesLinkMaster(Data);
            return st;
        }


        [ActionName("CustomerSalesLinkView")]
        public List<MyCustomer> CustomerSalesLinkView(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCustomerSalesLinkMaster(Data);
            return st;
        }
        [ActionName("CusSalesLinkUpdate")]
        public List<MyCustomer> CusSalesLinkUpdate(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCustomerSalesLinkMasterUpdate(Data);
            return st;
        }
        [ActionName("CusSalesLinkDelete")]
        public List<MyCustomer> CusSalesLinkDelete(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCustomerSalesLinkMasterDelete(Data);
            return st;
        }
        [ActionName("CustomerAttachments")]
        public List<MyCustomer> CustomerAttachments(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.InsertCustomerAttachmentsMaster(Data);
            return st;
        }

        [ActionName("CustomerAttachDetails")]
        public List<MyCustomer> CustomerAttachDetails(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusAttachmentMaster(Data);
            return st;
        }

        [ActionName("CusAttachDelete")]
        public List<MyCustomer> CusAttachDelete(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusAttachDelete(Data);
            return st;
        }

        [ActionName("CusSalPersonDtls")]
        public List<MyCustomer> CusSalPersonDtls(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCusSalesUserList(Data);
            return st;
        }

        #endregion

        #region Ganesh
        #region Agency Master
        [ActionName("Agency")]
        public List<MyAgency> Agency(MyAgency Data)
        {
            PartyManager pm = new PartyManager();
            List<MyAgency> st = pm.InsertAgencyMaster(Data);
            return st;
        }

        [ActionName("Agencyview")]
        public List<MyAgency> Agencyview(MyAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAgency> st = cm.GetAgencyMaster(Data);
            return st;
        }

        [ActionName("Agencyviewedit")]
        public List<MyAgency> AgencyViewEdit(MyAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAgency> st = cm.GetAgencyMasterRecord(Data);
            return st;
        }
        [ActionName("AgencyBind")]
        public List<MyAgency> AgencyBind(MyAgency Data)
        {
            PartyManager pm = new PartyManager();
            List<MyAgency> st = pm.BindAgencyDropDown(Data);
            return st;
        }

        [ActionName("GetPortCodes")]
        public List<MyPortAgency> GetPortCode()
        {
            PartyManager cm = new PartyManager();
            List<MyPortAgency> st = cm.PortList();
            return st;
        }

        [ActionName("AgencyPortCodes")]
        public List<MyPortAgency> AgencyPortCodes(MyPortAgency Data)
        {
            PartyManager pm = new PartyManager();
            List<MyPortAgency> st = pm.InsertAgencyPortCodes(Data);
            return st;
        }


        [ActionName("AgencyPortView")]
        public List<MyPortAgency> AgencyPortView(MyPortAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyPortAgency> st = cm.GetAgencyPortView(Data);
            return st;
        }
        [HttpPost, ActionName("DeleteAgencyPortDtls")]
        public List<MyPortAgency> DeleteAgencyPortDtls(MyPortAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyPortAgency> st = new List<MyPortAgency>();
            DataTable dt = cm.DelAgencyPortDtls(Data);
            return st;
        }


        [ActionName("GetCityCodes")]
        public List<MyCityAgency> GetCitCode()
        {
            PartyManager cm = new PartyManager();
            List<MyCityAgency> st = cm.CityList();
            return st;
        }

        [ActionName("AgencyCityCodes")]
        public List<MyCityAgency> AgencyCityCodes(MyCityAgency Data)
        {
            PartyManager pm = new PartyManager();
            List<MyCityAgency> st = pm.InsertAgencyCityCodes(Data);
            return st;
        }


        [ActionName("AgencyCityView")]
        public List<MyCityAgency> AgencyCityView(MyCityAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCityAgency> st = cm.GetAgencyCityView(Data);
            return st;
        }

        [HttpPost, ActionName("DeleteAgencyCityDtls")]
        public List<MyCityAgency> DeleteAgencyCityDtls(MyCityAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCityAgency> st = new List<MyCityAgency>();
            DataTable dt = cm.DelAgencyCityDtls(Data);
            return st;
        }


        [ActionName("AgencyAccDtls")]
        public List<MyAccAgency> InsertAgencyAccDtls(MyAccAgency Data)
        {
            PartyManager pm = new PartyManager();
            List<MyAccAgency> st = pm.InsertAgencyAccDtls(Data);
            return st;
        }

        [ActionName("CurrencyValues")]
        public List<MyAccCurrency> CurrencyValues()
        {
            PartyManager cm = new PartyManager();
            List<MyAccCurrency> st = cm.CurrencyMaster();
            return st;
        }
        [ActionName("AgencyAccEdit")]
        public List<MyAccAgency> AgencyAccEdit(MyAccAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAccAgency> st = cm.GetAgencyAccDtls(Data);
            return st;
        }

        [ActionName("AgencyAlertEmailDtls")]
        public List<MyAlertEmailAgency> InsertAgencyAlertEmailDtls(MyAlertEmailAgency Data)
        {
            PartyManager pm = new PartyManager();
            List<MyAlertEmailAgency> st = pm.InsertAgencyEmailDtls(Data);
            return st;
        }

        [ActionName("AgencyAlertEmailView")]
        public List<MyAlertEmailAgency> AgencyAlertEmailView(MyAlertEmailAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAlertEmailAgency> st = cm.GetAgencyAlertEmailView(Data);
            return st;
        }

        [ActionName("AgencyEmailDtlsEdit")]
        public List<MyAlertEmailAgency> AgencyEmailDtlsEdit(MyAlertEmailAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAlertEmailAgency> st = cm.GetAgencyEmailDetails(Data);
            return st;
        }


        [HttpPost, ActionName("DeleteEmailDtls")]
        public List<MyAlertEmailAgency> DeleteAgencyCityDtls(MyAlertEmailAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAlertEmailAgency> st = new List<MyAlertEmailAgency>();
            DataTable dt = cm.DelAgencyEmailDtls(Data);
            return st;
        }
        [ActionName("OrganizationType")]
        public List<MyAgency> OrganizationType()
        {
            PartyManager cm = new PartyManager();

            List<MyAgency> st = cm.ListOrganizationType();
            return st;
        }
        [ActionName("BindGeoLocations")]
        public List<MyAgency> BindGeoLocations()
        {
            PartyManager cm = new PartyManager();
            List<MyAgency> st = cm.GeoLocationsList();
            return st;
        }

        [ActionName("CheckAgencyValidation")]
        public List<MyAgency> CheckAgencyValidation(MyAgency Data)
        {
            PartyManager cm = new PartyManager();
            List<MyAgency> st = cm.ExistAgencyValidation(Data);
            return st;
        }
        #endregion

        #region Online Portal Creation

        [ActionName("BindPartyWiseByRole")]
        public List<MyOnlinePortal> BindPartyWiseByRole(MyOnlinePortal Data)
        {
            PartyManager pm = new PartyManager();
            List<MyOnlinePortal> st = pm.BindPartyWiseByRole(Data);
            return st;
        }

        [ActionName("InsertOnlinePortal")]
        public List<MyOnlinePortal> InsertOnlinePortal(MyOnlinePortal Data)
        {
            PartyManager cm = new PartyManager();
            List<MyOnlinePortal> st = cm.InsertOnlinePortal(Data);
            return st;
        }
        [ActionName("OnlinePortalView")]
        public List<MyOnlinePortal> OnlinePortalView(MyOnlinePortal Data)
        {
            PartyManager pm = new PartyManager();
            List<MyOnlinePortal> st = pm.BindOnlinePortalView(Data);
            return st;
        }
        [ActionName("OnlinePortalEdit")]
        public List<MyOnlinePortal> OnlinePortalEdit(MyOnlinePortal Data)
        {
            PartyManager pm = new PartyManager();
            List<MyOnlinePortal> st = pm.BindOnlinePortalEdit(Data);
            return st;
        }

        [ActionName("OnlinePortalDtlsEdit")]
        public List<MyOnlinePortal> OnlinePortalDtlsEdit(MyOnlinePortal Data)
        {
            PartyManager pm = new PartyManager();
            List<MyOnlinePortal> st = pm.BindOnlinePortalDtlsEdit(Data);
            return st;
        }
        #endregion




        [ActionName("BindRoleMaster")]
        public List<MyAgency> BindRoleMaster()
        {
            PartyManager cm = new PartyManager();
            List<MyAgency> st = cm.RoleMasterList();
            return st;
        }
        #endregion

    }
}
