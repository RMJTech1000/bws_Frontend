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
    public class CommonAccessApiController : ApiController
    {
        [ActionName("PortValues")]
        public List<MyCommonAccess> PortValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.PortMaster();
            return st;
        }
        [ActionName("CntrTypeValues")]
        public List<MyCommonAccess> CntrTypeValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CntrTypesMaster();
            return st;
        }

        [ActionName("CostTypeValues")]
        public List<MyCommonAccess> CostTypeValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CostTypesMaster();
            return st;
        }

        [ActionName("BookingCntrTypeValues")]
        public List<MyCommonAccess> BookingCntrTypeValues(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BookingCntrTypesMaster(Data);
            return st;
        }

        [ActionName("BookingCntrTypeQtyValues")]
        public List<MyCommonAccess> BookingCntrTypeQtyValues(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BookingCntrQtyTypesMaster(Data);
            return st;
        }

        [ActionName("CntrTypeValuesPass")]
        public List<MyCommonAccess> CntrTypeValuesPass(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CntrTypesMasterValuePass(Data);
            return st;
        }



        [ActionName("CommodityValues")]
        public List<MyCommonAccess> CommodityValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CommodityMaster();
            return st;
        }

        [ActionName("ServiceTypesValues")]
        public List<MyCommonAccess> ServiceTypesValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.ServiceTypesMaster();
            return st;
        }

        [ActionName("CurrencyValues")]
        public List<MyCommonAccess> CurrencyValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CurrencyMaster();
            return st;
        }

        [ActionName("ChargeCode")]
        public List<MyCommonAccess> ChargeCode()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.ChargecodeMaster();
            return st;
        }
        [ActionName("ChargeCodeValue")]
        public List<MyCommonAccess> ChargeCodeValue()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.ChargeCodeMasterBind();
            return st;
        }


        [ActionName("ChargeCodeAgent")]
        public List<MyCommonAccess> ChargeCodeAgent()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.ChargeCodeAgentMaster();
            return st;
        }




        [ActionName("ChargeCodeTypes")]
        public List<MyCommonAccess> ChargeCodeTypes(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.ChargecodeMasterTypes(Data);
            return st;
        }

        [ActionName("MUCChargeCodeTypes")]
        public List<MyCommonAccess> MUCChargeCodeTypes(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.MUCChargecodeMasterTypes(Data);
            return st;
        }

        [ActionName("IFCChargeCodeTypes")]
        public List<MyCommonAccess> IFCChargeCodeTypes(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.IFSChargecodeMasterTypes(Data);
            return st;
        }



        [ActionName("ModuleValues")]
        public List<MyCommonAccess> ModuleValues(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.Generalmaster(Data.ID.ToString());
            return st;
        }

        [ActionName("ModuleValuesTraiffMaster")]
        public List<MyCommonAccess> ModuleValuesTraiffMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.TariffMasterGeneralmaster(Data.ID.ToString());
            return st;
        }

        [ActionName("AgencyMaster")]
        public List<MyCommonAccess> AgencyMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.AgencyMaster();
            return st;
        }



        [ActionName("PortAgencyMaster")]
        public List<MyCommonAccess> PortAgencyMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.PortAgencyMaster(Data);
            return st;
        }

        [ActionName("CustomerMaster")]
        public List<MyCommonAccess> CustomerMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerMaster();
            return st;
        }

        [ActionName("CustomerMasterParameterPass")]
        public List<MyCommonAccess> CustomerMasterParameterPass(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerMasterValuesPass(Data);
            return st;
        }

        


        [ActionName("VendorMaster")]
        public List<MyCommonAccess> VendorMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.VendorMaster();
            return st;
        }

        [ActionName("CustomerMasterNew")]
        public List<MyCommonAccessNew> CustomerMasterNew()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccessNew> st = cm.CustomerMasterNew();
            return st;
        }

        [ActionName("CountryMaster")]
        public List<MyCommonAccess> CountryMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CountryMaster();
            return st;
        }

        [ActionName("UserMaster")]
        public List<MyCommonAccess> UserMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.UserMaster();
            return st;
        }


        [ActionName("RRMaster")]
        public List<MyCommonAccess> RRMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.RRNumberMaster(Data);
            return st;
        }


        [ActionName("RRMasterAccessList")]
        public List<MyCommonAccess> RRMasterAccessList(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.RRNumberAccessListMaster(Data);
            return st;
        }

        [ActionName("RRMasterAccess")]
        public List<MyCommonAccess> RRMasterAccess(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.RRNumberAccessMaster(Data);
            return st;
        }

        [ActionName("DepotMaster")]
        public List<MyCommonAccess> DepotMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.DepotMaster();
            return st;
        }

        [ActionName("TerminalMaster")]
        public List<MyCommonAccess> TerminalMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.TerminalMaster();
            return st;
        }

        [ActionName("VesselMaster")]
        public List<MyCommonAccess> VesselMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.VesselMaster();
            return st;
        }

        [ActionName("VesVoyMaster")]
        public List<MyCommonAccess> VesVoyMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.VesVoyMaster();
            return st;
        }
        [ActionName("VesVoyMasterAgencywise")]
        public List<MyCommonAccess> VesVoyAgencywiseMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.VesVoyAgencywiseMaster(Data);
            return st;
        }

        [ActionName("ImportBLVesVoy")]
        public List<MyCommonAccess> ImportBLVesVoy(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.ImportBLVesVoy(Data);
            return st;
        }

        

        [ActionName("BookingMaster")]
        public List<MyCommonAccess> BookingMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BookingMaster();
            return st;
        }

        [ActionName("BLMaster")]
        public List<MyCommonAccess> BLMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BLMaster();
            return st;
        }

        [ActionName("CntrNoValues")]
        public List<MyCommonAccess> CntrNoValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CntrNoMaster();
            return st;
        }


        [ActionName("BusinessTypes")]
        public List<MyCommonAccess> BusinessTypes()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BusinessTypesMaster();
            return st;
        }


        [ActionName("CustomerAddress")]
        public List<MyCommonAccess> CustomerAddress(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerAddress(Data.ID.ToString());
            return st;
        }


        [ActionName("CargoTypesMaster")]
        public List<MyCommonAccess> CargoTypesMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CargoTypesMaster();
            return st;
        }

        [ActionName("SlotTremsValues")]
        public List<MyCommonAccess> SlotTremsValues()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.SlotTermsMaster();
            return st;
        }

        [ActionName("CustomerBussTypesMaster")]
        public List<MyCommonAccess> CustomerBussTypesMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerBussTypesmaster(Data.BussTypes.ToString());
            return st;
        }

        [ActionName("CustomerRRSlotoperatorMaster")]
        public List<MyCommonAccess> CustomerRRSlotoperatorMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerRRSlotOperatorsmaster(Data.BkgID.ToString());
            return st;
        }


        


        [ActionName("CntrMovementMaster")]
        public List<MyCommonAccess> CntrMovementMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CntrMovementMaster();
            return st;
        }

        [ActionName("StatusMaster")]
        public List<MyCommonAccess> StatusMaster()
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.StatusMaster();
            return st;
        }

        [ActionName("VesVoyByAgency")]
        public List<MyCommonAccess> VesVoyByAgency(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.VesVoyByAgency(Data);
            return st;
        }
        [ActionName("BindMonthsList")]
        public List<MyCommonAccess> BindMonthsList(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BindMonthsList(Data);
            return st;
        }
        [ActionName("BindYearList")]
        public List<MyCommonAccess> BindYearList(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BindYearList(Data);
            return st;
        }
        [ActionName("BindMainPorts")]
        public List<MyCommonAccess> BindMainPorts(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BindMainPortsList(Data);
            return st;
        }
        [ActionName("BindAlertTypesAgency")]
        public List<MyCommonAccess> BindAlertTypesAgency(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BindAlertTypesAgency(Data);
            return st;
        }


        [ActionName("BindNotsClauses")]
        public List<MyCommonAccess> BindNotsClauses(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.BindNotsclauses(Data);
            return st;
        }

        [ActionName("GeoLocByCountry")]
        public List<MyCommonAccess> GeoLocByCountry(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.GeoLocByCountryValues(Data);
            return st;
        }
        [ActionName("PortValuesByGeoLoc")]
        public List<MyCommonAccess> PortValuesByGeoLoc(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.PortMasterValuesByGeoLoc(Data);
            return st;
        }


        [ActionName("Customer_EDI_VesVoy")]
        public List<MyCommonAccess> Customer_EDI_VesVoy(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.Customer_EDI_VesVoy(Data);
            return st;
        }

    }
}
