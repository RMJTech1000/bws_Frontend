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
    public class DashboardApiController : ApiController
    {
        [ActionName("LocationMaster")]
        public List<MyDashboarData> LocationMaster(MyDashboarData Data)
        {
            DashboardManager cm = new DashboardManager();
            List<MyDashboarData> st = cm.GeoLocationDtls(Data);
            return st;
        }

        [ActionName("DashboardInventoryCount")]
        public List<MyDashboarData> DashboardInventoryCount(MyDashboarData Data)
        {
            DashboardManager cm = new DashboardManager();
            List<MyDashboarData> st = cm.DashboardInventoryCount(Data);
            return st;
        }

        [ActionName("AgencyWiseBL")]
        public List<MyDashboarData> AgencyWiseBL(MyDashboarData Data)
        {
            DashboardManager cm = new DashboardManager();
            List<MyDashboarData> st = cm.AgencyWiseBL(Data);
            return st;
        }


        [ActionName("DashboardInventorySearch")]
        public List<MyDashboarData> DashboardInventorySearch(MyDashboarData Data)
        {
            DashboardManager cm = new DashboardManager();
            List<MyDashboarData> st = cm.DashboardInventorySearch(Data);
            return st;
        }
    }
}
