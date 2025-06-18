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
using DataSourceManager;
using DataTier;

namespace NVOCShipping.api
{
    public class DataSourceController : ApiController
    {

        [ActionName("BankMaster")]
        public List<MyAccount> CostTypeValues(MyAccount Data)
        {
            DataSource cm = new DataSource();
            List<MyAccount> st = cm.BankMasterDetails(Data);
            return st;
        }
    }
}