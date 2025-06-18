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
    public class ExportReportApiController : ApiController
    {
        #region anand
        [ActionName("TerminalDepReport")]
        public List<MyExportReport> TerminalDepReport(MyExportReport Data)
        {
            ExportReportManager cm = new ExportReportManager();
            List<MyExportReport> st = cm.TerminalDepReportMaster(Data);
            return st;
        }

        [ActionName("BLNoValues")]
        public List<MyExportReport> BLNoValues(MyExportReport Data)
        {
            ExportReportManager cm = new ExportReportManager();
            List<MyExportReport> st = cm.BLNoDropdownMaster(Data);
            return st;
        }
        #endregion

    }
}
