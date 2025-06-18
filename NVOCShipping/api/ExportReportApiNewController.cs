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
    public class ExportReportApiNewController : ApiController
    {
        #region Anand

        #region ExportReport
        [ActionName("NextPortByVesVoy")]
        public List<MyExportReportNew> NextPortByVesVoy(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.NextPortByVesVoyDtls(Data);
            return st;
        }
        [ActionName("TerminalDepDetails")]
        public List<MyExportReportNew> TerminalDepDetails(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TerminaDepBind(Data);
            return st;
        }

        [ActionName("DestinationAgentDtls")]
        public List<MyExportReportNew> DestinationAgentDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.DestinationAgentMaster(Data);
            return st;
        }
        [ActionName("SlotOperatorDtls")]
        public List<MyExportReportNew> SlotOperatorDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.SlotOperatorMaster(Data);
            return st;
        }
        [ActionName("TerminalDepMainDtls")]
        public List<MyExportReportNew> TerminalDepMainDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TerminalDepReportMain(Data);
            return st;
        }

        [ActionName("TerminalDepMainBLDtls")]
        public List<MyExportReportNew> TerminalDepMainBLDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TerminalDepReportBLMain(Data);
            return st;
        }


        [ActionName("BLNumberByVesVoy")]
        public List<MyExportReportNew> BLNumberByVesVoy(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.BLNumberByVesVoy(Data);
            return st;
        }

        [ActionName("FreightManifestMain")]
        public List<MyExportReportNew> FreightManifestMain(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.FreightManifestDtls(Data);
            return st;
        }

        [ActionName("FreightManifestMainBLDtls")]
        public List<MyExportReportNew> FreightManifestMainBLDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.FreightManifestBLDtls(Data);
            return st;
        }

        [ActionName("FreightManifestMainBLDtlsByBL")]
        public List<MyExportReportNew> FreightManifestMainBLDtlsByBL(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.FreightManifestBLDtlsByBL(Data);
            return st;
        }

        [ActionName("FreightManifestMainChargeDtls")]
        public List<MyExportReportNew> FreightManifestMainChargeDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.FreightManifestChargeDtls(Data);
            return st;
        }

        [ActionName("BLNumberChange")]
        public List<MyExportReportNew> BLNumberChange(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.BLNumberChange(Data);
            return st;
        }

        [ActionName("TranshipmentDropDown")]
        public List<MyExportReportNew> TranshipmentDropDown(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TranshipmentDetails(Data);
            return st;
        }

        [ActionName("PODDropDown")]
        public List<MyExportReportNew> PODDropDown(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.PODDropDown(Data);
            return st;
        }

        [ActionName("BLNumberByVesVoyGlobal")]
        public List<MyExportReportNew> BLNumberByVesVoyGlobal(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.BLNumberByVesVoyGlobal(Data);
            return st;
        }
        #endregion ExportReport



        #endregion Anand

        #region multi port TDR
        [ActionName("TerminalMultiDepDetails")]
        public List<MyExportReportNew> TerminalMultiDepDetails(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TerminalMultiDepDetailsBind(Data);
            return st;
        }

        [ActionName("TerminalDepMultiPortMainDtls")]
        public List<MyExportReportNew> TerminalMultiDepMainDtls(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TerminalMultiDepMainValues(Data);
            return st;
        }

        [ActionName("NextPortDropDown")]
        public List<MyExportReportNew> NextPortDropDown(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.NextPortDropDownData(Data);
            return st;
        }

        [ActionName("VesVoyMultiportsByAgency")]
        public List<MyExportReportNew> VesVoyMultiportsByAgency(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.VesVoyMultiportsByAgencyMaster(Data);
            return st;
        }
        #endregion

        #region  FRIEGHT GLOBAL 

        [ActionName("TerminalALLDepDetails")]
        public List<MyExportReportNew> TerminalALLDepDetails(MyExportReportNew Data)
        {
            ExportReportNewManager cm = new ExportReportNewManager();
            List<MyExportReportNew> st = cm.TerminalALLDepDetailsVALUES(Data);
            return st;
        }

        #endregion
    }
}
