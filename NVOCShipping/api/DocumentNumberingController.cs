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
    public class DocumentNumberingController : ApiController
    {

        [ActionName("CustomerBussTypesMaster")]
        public List<MyCommonAccess> CustomerBussTypesMaster(MyCommonAccess Data)
        {
            CommonAccessManager cm = new CommonAccessManager();
            List<MyCommonAccess> st = cm.CustomerBussTypesmaster(Data.BussTypes.ToString());
            return st;
        }

        [ActionName("BLNoLogicsInsert")]
        public List<MyDOCNumbering> BLNoLogicsInsert(MyDOCNumbering Data)
        {
            DocumentNumberingManager cm = new DocumentNumberingManager();
            List<MyDOCNumbering> st = cm.BLNoLogicsInsert(Data);
            return st;
        }

        [ActionName("BLNoLogicsView")]
        public List<MyDOCNumbering> BLLogicsViewRecord(MyDOCNumbering Data)
        {
            DocumentNumberingManager cm = new DocumentNumberingManager();
            List<MyDOCNumbering> st = cm.BLLogicsViewRecordValues(Data);
            return st;
        }

        [ActionName("BLNoLogicsEdit")]
        public List<MyDOCNumbering> BLNoLogicsEdit(MyDOCNumbering Data)
        {
            DocumentNumberingManager cm = new DocumentNumberingManager();
            List<MyDOCNumbering> st = cm.BLNoLogicsEditValues(Data);
            return st;
        }


        [ActionName("BLNoLogics")]
        public List<MyDOCNumbering> BLNoLogics(MyDOCNumbering Data)
        {
            DocumentNumberingManager cm = new DocumentNumberingManager();
            List<MyDOCNumbering> st = cm.BLNoLogicsValues(Data);
            return st;
        }


    }
}