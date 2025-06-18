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
    public class ReportApiController : ApiController
    {
        [ActionName("AgencyByVesVoy")]
        public List<MyReportData> AgencyByVesVoy(MyReportData Data)
        {
            ReportManager cm = new ReportManager();
            List<MyReportData> st = cm.DesAgencyByVesVoy(Data);
            return st;
        }

        [ActionName("OperatorByVesVoy")]
        public List<MyReportData> OperatorByVesVoy(MyReportData Data)
        {
            ReportManager cm = new ReportManager();
            List<MyReportData> st = cm.SlotOperatorByVesVoy(Data);
            return st;
        }

        [ActionName("BLByVesVoy")]
        public List<MyReportData> BLByVesVoy(MyReportData Data)
        {
            ReportManager cm = new ReportManager();
            List<MyReportData> st = cm.BLNumberByVesVoy(Data);
            return st;
        }

       
    }

}
