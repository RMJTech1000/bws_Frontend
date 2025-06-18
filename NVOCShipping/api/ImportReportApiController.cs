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
    public class ImportReportApiController : ApiController
    {
        #region DO REPORT
   
        [ActionName("DoReportView")]
        public List<MyImportReport> DoReportView(MyImportReport Data)
        {
            ImportReportManager cm = new ImportReportManager();
            List<MyImportReport> st = cm.DoReportViewValues(Data);
            return st;
        }
        #endregion
    }
}
