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
    public class PortStorageController : ApiController
    {


        [ActionName("Imp_PortStorageContainerWise")]
        public List<MYDETCalculation> Imp_PortStorageContainerWise(MYDETCalculation Data)
        {
            StorageManager Mange = new StorageManager();
            List<MYDETCalculation> st = Mange.Imp_PortStorageContainer(Data);
            return st;
        }

        [ActionName("Import_PortStorageInsert")]
        public List<MYImpDDG> Import_PortStorageInsert(MYImpDDG Data)
        {
            StorageManager Mange = new StorageManager();
            List<MYImpDDG> st = Mange.Import_PortStorage_CalculationInsert(Data);
            return st;
        }

        [ActionName("PortStorageView")]
        public List<MYDETCalculation> PortStorageView(MYDETCalculation Data)
        {
            StorageManager Mange = new StorageManager();
            List<MYDETCalculation> st = Mange.PortStorageView(Data);
            return st;
        }



        [ActionName("Port_EstimateexsitingValues")]
        public List<myPort_Storage> Port_EstimateexsitingValues(myPort_Storage Data)
        {
            StorageManager Mange = new StorageManager();
            List<myPort_Storage> st = Mange.Port_ExistingEstimatedisplay(Data);
            return st;
        }


        [ActionName("DisplayPort_EstimateexsitingValues")]
        public List<DDGRoot> DisplayPort_EstimateexsitingValues(DDGRoot Data)
        {
            StorageManager Mange = new StorageManager();
            List<DDGRoot> st = Mange.NewPort_DisplayExistingEstimateValues(Data);
            return st;
        }

        [ActionName("Port_DDGBilledAmount")]
        public List<MYDETCalculation> Port_DDGBilledAmount(MYDETCalculation Data)
        {
            StorageManager Mange = new StorageManager();
            List<MYDETCalculation> st = Mange.Port_DDGBillingAmount(Data);
            return st;
        }


        [ActionName("Port_DDGDelete")]
        public List<MYImpDDG> Port_DDGDelete(MYImpDDG Data)
        {
            StorageManager Mange = new StorageManager();
            List<MYImpDDG> st = Mange.Port_DDGDelete(Data);
            return st;
        }
    }
}
