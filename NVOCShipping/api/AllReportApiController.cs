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
    public class AllReportApiController : ApiController
    {
      

        #region Muthu
        [ActionName("MISReportView")]
        public List<AllReportData> MISReportView(AllReportData Data)
        {
            AllReportManager cm = new AllReportManager();
            List<AllReportData> st = cm.MISReportView(Data);
            return st;
        }

        [ActionName("MISReportViewALL")]
        public List<AllReportData> MISReportViewALL(AllReportData Data)
        {
            AllReportManager cm = new AllReportManager();
            List<AllReportData> st = cm.MISReportViewALL(Data);
            return st;
        }
        [ActionName("MISReportSurcharges")]
        public List<AllReportData> MISReportSurcharges(AllReportData Data)
        {
            AllReportManager cm = new AllReportManager();
            List<AllReportData> st = cm.MISReportSurcharge(Data);
            return st;
        }


        [ActionName("MIS_Misc_RevenuValues")]
        public List<AllReportData> MIS_Misc_RevenuValues(AllReportData Data)
        {
            AllReportManager cm = new AllReportManager();
            List<AllReportData> st = cm.MIS_Misc_RevenuReport(Data);
            return st;
        }

        #endregion


    }
}
