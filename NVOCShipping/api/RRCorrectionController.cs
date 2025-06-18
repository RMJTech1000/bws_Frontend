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
    public class RRCorrectionController : ApiController
    {

        [ActionName("RRCorrcetionUpdate")]
        public List<MyRatesheet> RRCorrcetionUpdate(MyRatesheet Data)
        {
            SalesManager cm = new SalesManager();
            List<MyRatesheet> st = cm.InsertRRCorrectionRatesheetMasterNew(Data);
            return st;
        }

        [ActionName("RRPricingView")]
        public List<RRPricing> RRPricingView(RRPricing Data)
        {
            SalesManager cm = new SalesManager();
            List<RRPricing> st = cm.RRPricingView(Data);
            return st;
        }

        [ActionName("RRPricingViewPOD")]
        public List<RRPricing> RRPricingViewPOD(RRPricing Data)
        {
            SalesManager cm = new SalesManager();
            List<RRPricing> st = cm.RRPricingViewPOD(Data);
            return st;
        }

        [ActionName("RRPricingSlotView")]
        public List<RRPricing> RRPricingSlotView(RRPricing Data)
        {
            SalesManager cm = new SalesManager();
            List<RRPricing> st = cm.RRPricingSlotView(Data);
            return st;
        }
    }
}
