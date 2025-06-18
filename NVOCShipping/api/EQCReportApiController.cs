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
    public class EQCReportApiController : ApiController
    {
        [ActionName("LongStayReportView")]
        public List<MyEQCReport> LongStayReportView(MyEQCReport Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCReport> st = cm.LongStayReportView(Data);
            return st;
        }
        [ActionName("CntrTurnAroundStatusList")]
        public List<MyEQCReport> CntrTurnAroundStatusList(MyEQCReport Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCReport> st = cm.CntrTurnAroundStatusList(Data);
            return st;
        }

        [ActionName("CntrTurnAroundView")]
        public List<MyEQCReport> CntrTurnAroundView(MyEQCReport Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCReport> st = cm.CntrTurnAroundView(Data);
            return st;
        }

        [ActionName("CntrStatusWiseReportView")]
        public List<MyEQCReport> CntrStatusWiseReportView(MyEQCReport Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCReport> st = cm.CntrStatusWiseReportView(Data);
            return st;
        }

        [ActionName("DCMRGlobalReportView")]
        public List<MyEQCDCMR> DCMRGlobalReportView(MyEQCDCMR Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCDCMR> st = cm.DCMRGlobalReportView(Data);
            return st;
        }

        [ActionName("DCMRLocationWiseReportView")]
        public List<MyEQCDCMR> DCMRLocationWiseReportView(MyEQCDCMR Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCDCMR> st = cm.DCMRLocationWiseReportView(Data);
            return st;
        }
        [ActionName("EQCStockReport")]
        public List<MyEQCStock> EQCStockReport(MyEQCStock Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCStock> st = cm.EQCStockReport(Data);
            return st;
        }
        [ActionName("EQCAgewiseReportView")]
        public List<MyEQCAgewise> EQCAgewiseReportView(MyEQCAgewise Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCAgewise> st = cm.EQCAgewiseReportView(Data);
            return st;
        }
        [ActionName("GeoLocByCountry")]
        public List<MyEQCAgewise> GeoLocByCountry(MyEQCAgewise Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCAgewise> st = cm.GeoLocByCountry(Data);
            return st;
        }
        [ActionName("CntrStatusCodes")]
        public List<MyCntrMoveMent> CntrStatusCodes(MyCntrMoveMent Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyCntrMoveMent> st = cm.ListCntrStatusCodes(Data);
            return st;
        }
        [ActionName("ContainerTypes")]
        public List<MyEQCAgewise> ContainerTypes(MyEQCAgewise Data)
        {
            EQCReportManager cm = new EQCReportManager();
            List<MyEQCAgewise> st = cm.ContainerTypes(Data);
            return st;
        }
    }
}
