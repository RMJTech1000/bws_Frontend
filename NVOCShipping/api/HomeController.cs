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
    public class HomeController : ApiController
    {
        #region muthu
        [ActionName("mainMenus")]
        public List<MyMenu> mainMenus(MyMenu Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyMenu> st = cm.MenusMaster(Data);
            return st;
        }
        [ActionName("FinancialYearList")]
        public List<MyDataBusinessLogic> FinancialYearList(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.ListFinancialYear(Data);
            return st;
        }

        [ActionName("BindCurrentYear")]
        public List<MyDataBusinessLogic> BindCurrentFinancialYear(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.BindCurrentFinancialYear(Data);
            return st;
        }
        [ActionName("GeoLocationsMaster")]
        public List<MyDataBusinessLogic> BindGeoLocations()
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.GeoLocationsList();
            return st;
        }


        [ActionName("GeoLocationsUser")]
        public List<MyDataBusinessLogic> GeoLocationsUser(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.GeoLocationsUserList(Data);
            return st;
        }


        [ActionName("AgencyMasterByGeoLoc")]
        public List<MyDataBusinessLogic> AgencyMasterByGeoLoc(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.AgencyMasterByGeoloc(Data);
            return st;
        }

        [ActionName("AgencyMasterByGeoLocUser")]
        public List<MyDataBusinessLogic> AgencyMasterByGeoLocUser(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.AgencyMasterByGeolocUser(Data);
            return st;
        }



        [ActionName("AgencyDocumentSuffix")]
        public List<MyDataBusinessLogic> AgencyDocumentSuffix(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.ListAgencyDocumentSuffix(Data);
            return st;
        }

        [ActionName("MenuListDD")]
        public List<MyDataBusinessLogic> MenuListDD(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.ListMenuListDD(Data);
            return st;
        }
        [ActionName("InsertControlParameter")]
        public List<MyDataBusinessLogic> ControlParameter(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.InsertControlParameter(Data);
            return st;
        }

        [ActionName("ControlParameterView")]
        public List<MyDataBusinessLogic> ControlParameterView(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.ListControlParameterView(Data);
            return st;
        }
        [ActionName("ControlParameterEdit")]
        public List<MyDataBusinessLogic> ControlParameterEdit(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.ListControlParameterEdit(Data);
            return st;
        }

        #endregion

        #region anand
        [ActionName("Login")]
        public List<MyDataBusinessLogic> LogIn(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.LoginValues(Data);
            return st;
        }


        [ActionName("Register")]
        public List<MyDataBusinessLogic> Register(MyDataBusinessLogic Data)
        {
            RegistrationManager cm = new RegistrationManager();
            List<MyDataBusinessLogic> st = cm.InsertUserMaster(Data);
            return st;
        }

        [ActionName("Country")]
        public List<MyCountry> Country(MyCountry Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCountry> st = cm.InsertCountryMaster(Data);
            return st;
        }
        [ActionName("Countryview")]
        public List<MyCountry> Countryview(MyCountry Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCountry> st = cm.GetCountryMaster(Data);
            return st;
        }
        [ActionName("Countryviewparticular")]
        public List<MyCountry> Countryviewparticular(MyCountry Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCountry> st = cm.GetCountryMasterRecord(Data);
            return st;
        }

        [ActionName("Currency")]
        public List<MyCurrency> Currency(MyCurrency Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCurrency> st = cm.InsertCurrencyMaster(Data);
            return st;
        }
        [ActionName("Currencyview")]
        public List<MyCurrency> Currencyview(MyCurrency Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCurrency> st = cm.GetCurrrencyMaster(Data);
            return st;
        }
        [ActionName("Currencyviewparticular")]
        public List<MyCurrency> Currencyviewparticular(MyCurrency Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCurrency> st = cm.GetCurrencyMasterRecord(Data);
            return st;
        }

        [ActionName("Depot")]
        public List<MyDepot> Depot(MyDepot Data)
        {
            MasterManager cm = new MasterManager();
            List<MyDepot> st = cm.InsertDepotMaster(Data);
            return st;
        }

        [ActionName("countryBind")]
        public List<MyCountry> countryBind(MyCountry Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCountry> st = cm.GetCommonCountryMaster(Data);
            return st;
        }

        [ActionName("cityBind")]
        public List<MyCity> cityBind(MyCity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCity> st = cm.GetCommonCityMaster(Data);
            return st;
        }

        [ActionName("stateBind")]
        public List<MyState> stateBind(MyState Data)
        {
            MasterManager cm = new MasterManager();
            List<MyState> st = cm.GetCommonStateMaster(Data);
            return st;
        }
        [ActionName("PortByCountry")]
        public List<MyDepot> PortByCountry(MyDepot Data)
        {
            MasterManager cm = new MasterManager();
            List<MyDepot> st = cm.GetPortByCountry(Data);
            return st;
        }
        [ActionName("DepotView")]
        public List<MyDepot> DepotView(MyDepot Data)
        {
            MasterManager cm = new MasterManager();
            List<MyDepot> st = cm.GetDepotMaster(Data);
            return st;
        }

        [ActionName("DepotRecord")]
        public List<MyDepot> DepotRecord(MyDepot Data)
        {
            MasterManager cm = new MasterManager();
            List<MyDepot> st = cm.GetDepotMasterRecord(Data);
            return st;
        }
        [ActionName("BindDepoMasterPortDtls")]
        public List<MyDepot> BindDepoMasterPortDtls(MyDepot Data)
        {
            MasterManager cm = new MasterManager();
            List<MyDepot> st = cm.GetDepoMasterPortDtls(Data);
            return st;
        }
        [ActionName("DepotApplicablePortDelete")]
        public List<MyDepot> DepotApplicablePortDelete(MyDepot Data)
        {
            MasterManager cm = new MasterManager();
            List<MyDepot> st = cm.DepotApplicablePortDelete(Data);
            return st;
        }

        [ActionName("Terminal")]
        public List<MyTerminal> Terminal(MyTerminal Data)
        {
            MasterManager cm = new MasterManager();
            List<MyTerminal> st = cm.InsertTerminalMaster(Data);
            return st;
        }

        [ActionName("TerminalView")]
        public List<MyTerminal> TerminalView(MyTerminal Data)
        {
            MasterManager cm = new MasterManager();
            List<MyTerminal> st = cm.GetTerminalMaster(Data);
            return st;
        }
        [ActionName("TerminalRecord")]
        public List<MyTerminal> TerminalRecord(MyTerminal Data)
        {
            MasterManager cm = new MasterManager();
            List<MyTerminal> st = cm.GetTerminalMasterRecord(Data);
            return st;
        }
        [ActionName("portBind")]
        public List<MyPort> portBind(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.GetCommonPortMaster(Data);
            return st;
        }

        [ActionName("State")]
        public List<MyState> State(MyState Data)
        {
            MasterManager cm = new MasterManager();
            List<MyState> st = cm.InsertStateMaster(Data);
            return st;
        }

        [ActionName("StateView")]
        public List<MyState> StateView(MyState Data)
        {
            MasterManager cm = new MasterManager();
            List<MyState> st = cm.GetStateMaster(Data);
            return st;
        }
        [ActionName("StateRecord")]
        public List<MyState> StateRecord(MyState Data)
        {
            MasterManager cm = new MasterManager();
            List<MyState> st = cm.GetStateMasterRecord(Data);
            return st;
        }

        [ActionName("Vessel")]
        public List<MyVessel> Vessel(MyVessel Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVessel> st = cm.InsertVesselMaster(Data);
            return st;
        }
        [ActionName("Vesselview")]
        public List<MyVessel> Vesselview(MyVessel Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVessel> st = cm.GetVesselMaster(Data);
            return st;
        }
        [ActionName("Vesselviewparticular")]
        public List<MyVessel> Vesselviewparticular(MyVessel Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVessel> st = cm.GetVesselMasterRecord(Data);
            return st;
        }
        [ActionName("VesselDropDown")]
        public List<MyVessel> VesselDropDown(MyVessel Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVessel> st = cm.GetVesselMasterDropDown(Data);
            return st;
        }
        [ActionName("VoyageDropDown")]
        public List<MyVoyage> VoyageDropDown(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.GetVoyageMasterDropDown(Data);
            return st;
        }
        [ActionName("RotNoDropDown")]
        public List<MyVoyage> RotNoDropDown(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.GetRotNoMasterDropDown(Data);
            return st;
        }
        [ActionName("Voyage")]
        public List<MyVoyage> Voyage(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.InsertVoyageMaster(Data);
            return st;
        }

        [ActionName("VoyageView")]
        public List<MyVoyage> VoyageView(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.GetVoyageMaster(Data);
            return st;
        }

        [ActionName("Voyageviewparticular")]
        public List<MyVoyage> Voyageviewparticular(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.GetVoyageMasterRecord(Data);
            return st;
        }
        [ActionName("VoyageDropDownChange")]
        public List<MyVoyage> VoyageDropDownChange(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.GetVoyDropDownChangeMaster(Data);
            return st;
        }
        [ActionName("RotNoDropDownChange")]
        public List<MyVoyage> RotNoDropDownChange(MyVoyage Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyage> st = cm.GetRotNoDropDownChangeMaster(Data);
            return st;
        }

        [ActionName("CustDropDown")]
        public List<MyCustomer> CustDropDown(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetCustDropDownMaster(Data);
            return st;
        }

        [ActionName("PortDropDown")]
        public List<MyPort> PortDropDown(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.GetPortDropDownMaster(Data);
            return st;
        }

        [ActionName("TerminalDropDown")]
        public List<MyTerminal> TerminalDropDown(MyTerminal Data)
        {
            MasterManager cm = new MasterManager();
            List<MyTerminal> st = cm.GetTerminalDropDownMaster(Data);
            return st;
        }

        [ActionName("VoyageDetails")]
        public List<MyVoyageDetails> VoyageDetails(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.InsertVoyageDetails(Data);
            return st;
        }

        [ActionName("VoyDtlsView")]
        public List<MyVoyageDetails> VoyDtlsView(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.GetVoyDtlsViewMaster(Data);
            return st;
        }
        [ActionName("VoyDtlsViewParticular")]
        public List<MyVoyageDetails> VoyDtlsViewParticular(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.GetVoyDtlsPartRecordMaster(Data.ID.ToString());
            return st;
        }
        [ActionName("VoyPortDtlsParticular")]
        public List<MyVoyageDetails> VoyPortDtlsParticular(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.GetVoyPortDtlsMaster(Data.ID.ToString());
            return st;
        }
        #endregion

        #region Ganesh

        [ActionName("City")]
        public List<MyCity> City(MyCity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCity> st = cm.InsertCityMaster(Data);
            return st;
        }


        [ActionName("Cityview")]
        public List<MyCity> Cityview(MyCity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCity> st = cm.GetCityMaster(Data);
            return st;
        }


        [ActionName("BindCountries")]
        public List<countryDD> BindCountries()
        {
            MasterManager cm = new MasterManager();
            List<countryDD> st = cm.Listcountry();
            return st;
        }

        [ActionName("BindStates")]
        public List<StateDD> BindStates()
        {
            MasterManager cm = new MasterManager();
            List<StateDD> st = cm.ListStates();
            return st;
        }


        [ActionName("BindCities")]
        public List<cityDD> BindCities(cityDD Data)
        {
            MasterManager cm = new MasterManager();
            List<cityDD> st = cm.ListCities(Data);
            return st;
        }

        [ActionName("Cityviewedit")]
        public List<MyCity> CitViewEdit(MyCity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCity> st = cm.GetCityMasterRecord(Data);
            return st;
        }

        [ActionName("Port")]
        public List<MyPort> Port(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.InsertPortMaster(Data);
            return st;
        }


        [ActionName("Portview")]
        public List<MyPort> Portview(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.GetPortMaster(Data);
            return st;
        }

        [ActionName("MainPortview")]
        public List<MyPort> MainPortview(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.GetMainPortMaster(Data);
            return st;
        }

        [ActionName("Portviewedit")]
        public List<MyPort> PortViewEdit(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.GetPortMasterRecord(Data);
            return st;
        }

        [ActionName("MainPortviewedit")]
        public List<MyPort> MainPortviewedit(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.GetMainPortRecord(Data);
            return st;
        }

        [ActionName("MainPort")]
        public List<MyPort> MainPort(MyPort Data)
        {
            MasterManager cm = new MasterManager();
            List<MyPort> st = cm.InsertMainPortMaster(Data);
            return st;
        }
        [ActionName("CargoPackage")]
        public List<MyCargo> CargoPackage(MyCargo Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCargo> st = cm.InsertCargoMaster(Data);
            return st;
        }

        [ActionName("CargoPkgviewedit")]
        public List<MyCargo> CargoPkgViewEdit(MyCargo Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCargo> st = cm.GetCargoPkgMasterRecord(Data);
            return st;
        }


        [ActionName("CargoPkgview")]
        public List<MyCargo> CargoPkgView(MyCargo Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCargo> st = cm.GetCargoPkgMaster(Data);
            return st;
        }



        [ActionName("Commodity")]
        public List<MyCommodity> Commodity(MyCommodity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCommodity> st = cm.InsertCommodityMaster(Data);
            return st;
        }

        [ActionName("Commodityviewedit")]
        public List<MyCommodity> CommodityViewEdit(MyCommodity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCommodity> st = cm.GetCommodityMasterRecord(Data);
            return st;
        }


        [ActionName("Commodityview")]
        public List<MyCommodity> CommodityView(MyCommodity Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCommodity> st = cm.GetCommodityMaster(Data);
            return st;
        }

        [ActionName("CommodityTypes")]
        public List<MyCommodityTypes> CommodityTypes()
        {
            MasterManager cm = new MasterManager();
            List<MyCommodityTypes> st = cm.CommodityTypeValues();
            return st;
        }

        [ActionName("ExchangeRate")]
        public List<MyExRate> ExchangeRate(MyExRate Data)
        {
            MasterManager cm = new MasterManager();
            List<MyExRate> st = cm.InsertExRateMaster(Data);
            return st;
        }

        [ActionName("FromCurrency")]
        public List<CurrencyDD> FromCurrency()
        {
            MasterManager cm = new MasterManager();

            List<CurrencyDD> st = cm.ListFromCurrencyDD();
            return st;
        }

        [ActionName("ToCurrency")]
        public List<CurrencyDD> ToCurrency()
        {
            MasterManager cm = new MasterManager();

            List<CurrencyDD> st = cm.ListToCurrencyDD();
            return st;
        }


        [ActionName("ExRateView")]
        public List<MyExRate> ExChangeRateView(MyExRate Data)
        {
            MasterManager cm = new MasterManager();
            List<MyExRate> st = cm.ExRateMaster(Data);
            return st;
        }

        [ActionName("ExRateviewedit")]
        public List<MyExRate> ExChangeRateViewEdit(MyExRate Data)
        {
            MasterManager cm = new MasterManager();
            List<MyExRate> st = cm.GetExRateMasterRecord(Data);
            return st;
        }
        [ActionName("InsertGeoLocation")]
        public List<MyGeoLocation> InsertGeoLocation(MyGeoLocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyGeoLocation> st = cm.InsertGeoLocationMaster(Data);
            return st;
        }

        [ActionName("GeoLocationView")]
        public List<MyGeoLocation> GeoLocationView(MyGeoLocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyGeoLocation> st = cm.GeoLocationViewMaster(Data);
            return st;
        }

        [ActionName("GeoLocationEdit")]
        public List<MyGeoLocation> GeoLocationEdit(MyGeoLocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyGeoLocation> st = cm.GeoLocationEditMaster(Data);
            return st;
        }

        [ActionName("BindGeoLocDepotDtls")]
        public List<MyGeoLocation> BindGeoLocDepotDtls(MyGeoLocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyGeoLocation> st = cm.BindGeoLocDepotDtls(Data);
            return st;
        }
        [ActionName("GeoLocApplicableDepotDelete")]
        public List<MyGeoLocation> GeoLocApplicableDepotDelete(MyGeoLocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyGeoLocation> st = cm.GeoLocApplicableDepotDelete(Data);
            return st;
        }

        #region voyage allocation
        [ActionName("VoyageTypes")]
        public List<MyVoyageAllocation> VoyageTypes(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.ListVoyageTypes(Data);
            return st;
        }


        [ActionName("LegInfoTypes")]
        public List<MyVoyageAllocation> LegInfoTypes(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.ListLegInfoTypes(Data);
            return st;
        }
        [ActionName("BLListAgentwise")]
        public List<MyVoyageAllocation> BLListAgentwise(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.ListBLListAgentwise(Data);
            return st;
        }
        [ActionName("EXPandTSBLListAgentwise")]
        public List<MyVoyageAllocation> EXPandTSBLListAgentwise(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.ListEXPandTSBLListAgentwise(Data);
            return st;
        }
        [ActionName("VoyageAllocationInsert")]
        public List<MyVoyageAllocation> VoyageAllocationInsert(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageAllocation> st = cm.InsertVoyageAllocation(Data);
            return st;
        }

        [ActionName("BindLegVslVoyDetails")]
        public List<MyVoyageAllocation> BindLegVslVoyDetails(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageAllocation> st = cm.ListBindLegVslVoyDetails(Data);
            return st;
        }

        [ActionName("ExistingTSVoyageAllication")]
        public List<MyVoyageAllocation> ExistingTSVoyageAllication(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageAllocation> st = cm.ExistingTsVoyageAllocationDetails(Data);
            return st;
        }


        [ActionName("VoyAllocationView")]
        public List<MyVoyageAllocation> VoyAllocationView(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.ListVoyAllocationView(Data);
            return st;
        }

        [ActionName("BindVoyAllocationEdit")]
        public List<MyVoyageAllocation> BindVoyAllocationEdit(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.BindVoyAllocationEditValues(Data);
            return st;
        }


        [ActionName("BindVoyAllocationDtlsEdit")]
        public List<MyVoyageAllocation> BindVoyAllocationDtlsEdit(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageAllocation> st = cm.BindVoyAllocationDtlsEditValues(Data);
            return st;
        }

        [ActionName("BindVoyAllocationDtlsModify")]
        public List<MyVoyageAllocation> BindVoyAllocationDtlsModify(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageAllocation> st = cm.VoyAllocationDtlsModifyValues(Data);
            return st;
        }



        [ActionName("BLVoyAllocationDelete")]
        public List<MyVoyageAllocation> BLVoyAllocationDelete(MyVoyageAllocation Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageAllocation> st = new List<MyVoyageAllocation>();
            DataTable dt = cm.BLVoyAllocationDelete(Data);
            return st;
        }

        #endregion

        #region ServiceSetup

        [ActionName("SlotOperatorByServices")]
        public List<MyServiceSetup> BindSlotOperatorByServices(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();

            List<MyServiceSetup> st = cm.ListSlotOperatorByServices(Data);
            return st;
        }

        [ActionName("BindSlotRefByOperator")]
        public List<MyServiceSetup> BindSlotRefByOperator(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();

            List<MyServiceSetup> st = cm.ListSlotRefByOperator(Data);
            return st;
        }

        [ActionName("ServiceValidation")]
        public List<MyServiceSetup> ServiceValidation(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ExistServiceValidation(Data);
            return st;
        }
        [ActionName("InsertServiceSetup")]
        public List<MyServiceSetup> InsertServiceSetup(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.InsertServiceSetup(Data);
            return st;
        }

        [ActionName("ServiceSetupView")]
        public List<MyServiceSetup> ServiceSetupView(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ServiceSetupViewMaster(Data);
            return st;
        }

        [ActionName("ServiceSetupEdit")]
        public List<MyServiceSetup> ServiceSetupEdit(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ServiceSetupEditMaster(Data);
            return st;
        }

        [ActionName("ServiceRouteEdit")]
        public List<MyServiceSetup> ServiceRouteEdit(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ServiceRouteEditMaster(Data);
            return st;
        }

        [ActionName("ServiceOperatorsEdit")]
        public List<MyServiceSetup> ServiceOperatorsEdit(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ServiceOperatorsEditMaster(Data);
            return st;
        }

        [ActionName("ServiceOperatorsDelete")]
        public List<MyServiceSetup> ServiceOperatorsDelete(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ServiceOperatorsDeleteValues(Data);
            return st;
        }
        [ActionName("ServicePortsDelete")]
        public List<MyServiceSetup> ServicePortsDelete(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.ServicePortsDeleteValues(Data);
            return st;
        }
        #endregion

        #region Voyage Details New (TPS)

        [ActionName("BindServices")]
        public List<MyServiceSetup> BindServices(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.BindServicesMaster(Data);
            return st;

        }

        [ActionName("BindTerminalByPort")]
        public List<MyCommonAccess> BindTerminalByPort(MyCommonAccess Data)
        {
            MasterManager cm = new MasterManager();
            List<MyCommonAccess> st = cm.TerminalMasterByPort(Data);
            return st;

        }

        [ActionName("BindServiceSchedule")]
        public List<MyServiceSetup> BindServiceSchedule(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.BindServiceScheduleMaster(Data);
            return st;

        }

        [ActionName("VoyageOperatorsEdit")]
        public List<MyServiceSetup> VoyageOperatorsEdit(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.VoyageOperatorsEditMaster(Data);
            return st;
        }

        [ActionName("InsertVoyageFirstTab")]
        public List<MyVoyageDetails> InsertVoyageFirstTab(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.InsertVoyageFirstTab(Data);
            return st;
        }

        [ActionName("VoyageDetailsView")]
        public List<MyVoyageDetails> VoyageDetailsView(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.VoyageDetailsView(Data);
            return st;
        }

        [ActionName("VoyageDetailsEdit")]
        public List<MyVoyageDetails> VoyageDetailsEdit(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.VoyageDetailsEditMaster(Data);
            return st;
        }

        [ActionName("VoyageSailingDetailsEdit")]
        public List<MyVoyageDetails> VoyageSailingDetailsEdit(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.VoyageSailingDetailsEditMaster(Data);
            return st;
        }
        [ActionName("InsertOperatorServices")]
        public List<MyVoyageDetails> InsertOperatorServices(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.InsertOperatorServiceValues(Data);
            return st;
        }
        [ActionName("VoyageOperatorsDelete")]
        public List<MyServiceSetup> VoyageOperatorsDelete(MyServiceSetup Data)
        {
            MasterManager cm = new MasterManager();
            List<MyServiceSetup> st = cm.VoyageOperatorsDeleteValues(Data);
            return st;
        }
        [ActionName("BerthingPortDropdown")]
        public List<MyVoyageDetails> BerthingPortDropdown(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.BerthingPortDropdownList(Data);
            return st;
        }

        [ActionName("InsertBerthingDtls")]
        public List<MyVoyageDetails> InsertBerthingDtls(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.InsertBerthingDtlsValues(Data);
            return st;
        }


        [ActionName("BindBerthingDetails")]
        public List<MyVoyageDetails> BindBerthingDetails(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.BindBerthingDetailsValues(Data);
            return st;
        }

        [ActionName("InsertManifestDtls")]
        public List<MyVoyageDetails> InsertManifestDtls(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.InsertManifestDtlsValues(Data);
            return st;
        }
        [ActionName("ViewManifestDtls")]
        public List<MyVoyageDetails> ViewManifestDtls(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.ViewManifestDtlsValues(Data);
            return st;
        }
        [ActionName("BkgVoyageOperatorValidation")]
        public List<MyVoyageDetails> BkgVoyageOperatorValidation(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.BkgVoyageOperatorValidation(Data);
            return st;
        }

        [ActionName("SendEmailPopUpVoyDtls")]
        public List<MyVoyageDetails> SendEmailPopUpVoyDtls(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.SendEmailPopUpVoyDtls(Data);
            return st;
        }
        [ActionName("VoyBookingPartyEmailDtls")]
        public List<MyVoyageDetails> VoyBookingPartyEmailDtls(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.VoyBookingPartyEmailDtls(Data);
            return st;
        }
        [ActionName("BindNoteTypesList")]
        public List<MyVoyageDetails> BindNoteTypesList(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.BindNoteTypesListValues(Data);
            return st;
        }
        [ActionName("InsertNotes")]
        public List<MyVoyageDetails> InsertNotes(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.InsertNotesValues(Data);
            return st;
        }
        [ActionName("ViewNotesDtls")]
        public List<MyVoyageDetails> ViewNotesDtls(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.ViewNotesDtlsEdit(Data);
            return st;
        }
        [ActionName("VoyageNotesDelete")]
        public List<MyVoyageDetails> VoyageNotesDelete(MyVoyageDetails Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageDetails> st = cm.VoyageNotesDeleteValues(Data);
            return st;
        }

        #endregion

        #region Voyage opening

        [ActionName("VoyageOpeningView")]
        public List<MyVoyageOpening> VoyageOpeningView(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.ListVoyageOpening(Data);
            return st;
        }

        [ActionName("VoyageOpeningEdit")]
        public List<MyVoyageOpening> VoyageOpeningEdit(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.ListVoyageOpeningEdit(Data);
            return st;
        }

        [ActionName("PortByGeoLoc")]
        public List<MyVoyageOpening> PortByGeoLoc(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.PortByGeoLoc(Data);
            return st;
        }
        [ActionName("TerminalDropDownByPortGeoLoc")]
        public List<MyVoyageOpening> TerminalDropDownByPortGeoLoc(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.TerminalDropDownByPortGeoLoc(Data);
            return st;
        }


        [ActionName("VoyageOpeningEditBLDetails")]
        public List<MyVoyageOpening> VoyageOpeningEditBLDetails(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.VoyageOpeningEditBLDetails(Data);
            return st;
        }
        [ActionName("VoyageUnOpenedBLDetails")]
        public List<MyVoyageOpening> VoyageUnOpenedBLDetails(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.VoyageUnOpenedBLDetails(Data);
            return st;
        }
        [ActionName("InsertVoyageOpening")]
        public List<MyVoyageOpening> InsertVoyageOpening(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.InsertVoyageOpening(Data);
            return st;
        }
        #endregion

        #region Voyage locking

        [ActionName("VoyageLockingRecView")]
        public List<MyVoyageOpening> VoyageLockingRecView(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.VoyageLockingRecView(Data);
            return st;
        }

        [ActionName("VoyageLockingDetailsEdit")]
        public List<MyVoyageOpening> VoyageLockingDetailsEdit(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.VoyageLockingDetailsEdit(Data);
            return st;
        }


        [HttpPost, ActionName("VesVoyMasterByAgency")]
        public List<MyVoyageOpening> VesVoyMaster(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.VesVoyMaster(Data);
            return st;
        }


        [ActionName("VesVoyWithOutAgenctMaster")]
        public List<MyVoyageOpening> VesVoyWithOutAgenctMaster(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.VesVoy_WithoutAgentMaster(Data);
            return st;
        }




        [ActionName("VoyageLockingBLDetails")]
        public List<MyVoyageOpening> VoyageLockingBLDetails(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.VoyageLockingBLDetails(Data);
            return st;
        }
        [ActionName("VoyageLockedBLDetails")]
        public List<MyVoyageOpening> VoyageLockedBLDetails(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();

            List<MyVoyageOpening> st = cm.VoyageLockedBLDetails(Data);
            return st;
        }
        [ActionName("VoyageLockingUpdate")]
        public List<MyVoyageOpening> VoyageLockingUpdate(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.VoyageLockingUpdate(Data);
            return st;
        }

        [ActionName("NotesandClausesInsert")]
        public List<MyNotes> NotesandClausesInsert(MyNotes Data)
        {
            MasterManager cm = new MasterManager();
            List<MyNotes> st = cm.NotesandClausesMaster(Data);
            return st;
        }

        [ActionName("NotesandClausesView")]
        public List<MyNotes> NotesandClausesView(MyNotes Data)
        {
            MasterManager cm = new MasterManager();
            List<MyNotes> st = cm.NotesandClausesView(Data);
            return st;
        }


        [ActionName("ExistingNotesandClauses")]
        public List<MyNotes> ExistingNotesandClauses(MyNotes Data)
        {
            MasterManager cm = new MasterManager();
            List<MyNotes> st = cm.NotesandClausesDetails(Data);
            return st;
        }

        [ActionName("ExistingDocType")]
        public List<MyNotes> ExistingDocType(MyNotes Data)
        {
            MasterManager cm = new MasterManager();
            List<MyNotes> st = cm.DocumentTypeValues(Data);
            return st;
        }

        [ActionName("NotesandClausesDelete")]
        public List<MyNotes> NotesandClausesDelete(MyNotes Data)
        {
            MasterManager cm = new MasterManager();
            List<MyNotes> st = cm.NotesandClausesDelete(Data);
            return st;
        }


        #endregion

        #endregion

        [ActionName("Gatewayview")]
        public List<MyVoyageOpening> Gatewayview(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.GatewayViewValues(Data);
            return st;
        }
        [ActionName("VesVoyOpeningViewPage")]
        public List<MyVoyageOpening> VesVoyOpeningViewPage(MyVoyageOpening Data)
        {
            MasterManager cm = new MasterManager();
            List<MyVoyageOpening> st = cm.VesVoyOpeningViewPage(Data);
            return st;
        }
    }
}