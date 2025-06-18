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
    public class SOAReportApiController : ApiController
    {

        [ActionName("SOAExportReportView")]
        public List<MySOAData> SOAExportReportView(MySOAData Data)
        {
            SOAManager cm = new SOAManager();
            List<MySOAData> st = cm.SOAExportReportView(Data);
            return st;
        }

        [ActionName("SOAImportReportView")]
        public List<MySOAData> SOAImportReportView(MySOAData Data)
        {
            SOAManager cm = new SOAManager();
            List<MySOAData> st = cm.SOAImportReportView(Data);
            return st;
        }
    }
}
