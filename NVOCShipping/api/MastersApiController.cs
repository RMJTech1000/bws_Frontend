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
    public class MastersApiController : ApiController
    {

        [ActionName("GeneralMaster")]
        public List<MyGeneralMaster> GeneralMaster(MyGeneralMaster Data)
        {

            GeneralMasterManager GenMang = new GeneralMasterManager();
            List<MyGeneralMaster> st = GenMang.GeneralMasterView(Data);
            return st;
        }

        [ActionName("GeneralMasteredit")]
        public List<MyGeneralMaster> GeneralMasteredit(MyGeneralMaster Data)
        {
            GeneralMasterManager GenMang = new GeneralMasterManager();
            List<MyGeneralMaster> st = GenMang.GetGeneralMasterEditRecord(Data);
            return st;
        }

        [ActionName("GeneralMasterInsert")]
        public List<MyGeneralMaster> GeneralMasterInsert(MyGeneralMaster Data)
        {

            GeneralMasterManager GenMang = new GeneralMasterManager();
            List<MyGeneralMaster> st = GenMang.InsertGeneralMaster(Data);
            return st;
        }
    }
}
