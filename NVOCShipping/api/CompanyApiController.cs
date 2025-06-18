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
    public class CompanyApiController : ApiController
    {
        [ActionName("CompanyMaster")]
        public List<MyCompany> CompanyMaster(MyCompany Data)
        {

            CompanyManager CompMang = new CompanyManager();
            List<MyCompany> st = CompMang.GetCompanyDetails(Data);
            return st;
        }
    }
}
