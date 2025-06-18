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
    public class ExpTraffViewController : ApiController
    {

        [ActionName("ExpTariffView")]
        public List<MyRRRate> ExpTariffView(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.ExpTarfiiExisting(Data);
            return st;
        }


        [ActionName("ExpTariffViewImp")]
        public List<MyRRRate> ExpTariffViewImp(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.ExpTarfiiExistingImp(Data);
            return st;
        }

        [ActionName("ExpSlotCost")]
        public List<MyRRRate> ExpSlotCost(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.ExpSlotCostExisting(Data);
            return st;
        }

        [ActionName("ExpTariffCost")]
        public List<MyRRRate> ExpTariffCost(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.ExpTariffCostExisting(Data);
            return st;
        }

        [ActionName("ExpTariffCommisionCost")]
        public List<MyRRRate> ExpTariffCommisionCost(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.ExpTariffCommisionCostExisting(Data);
            return st;
        }

        [ActionName("VendorCostInsert")]
        public List<MyRRRate> VendorCostInsert(MyRRRate Data)
        {
            ExpTraffMaster cm = new ExpTraffMaster();
            List<MyRRRate> st = cm.VendorCostInsert(Data);
            return st;
        }


        [ActionName("OffLineImportTariffPortCharges")]
        public List<MyRRRate> OffLineImportTariffPortCharges(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.InsertOffLineImportTariffCharges(Data);
            return st;
        }

        [ActionName("UpdateCustomerCost")]
        public List<MyRRRate> UpdateCustomerCost(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.CustomerCostInsert(Data);
            return st;
        }

        [ActionName("ImportBLChargedelete")]
        public List<MyRRRate> ImportBLChargedelete(MyRRRate Data)
        {
            ExpTraffMaster Mange = new ExpTraffMaster();
            List<MyRRRate> st = Mange.DeleteCustomerBLCharges(Data);
            return st;
        }




    }
}