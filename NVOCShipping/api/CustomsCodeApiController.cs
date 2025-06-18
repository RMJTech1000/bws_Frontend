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
    public class CustomsCodeApiController : ApiController
    {

        [ActionName("AgentCodeView")]
        public List<MyAgentCode> AgentCodeView(MyAgentCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyAgentCode> st = AgentMng.GetCustomsAgenctCode(Data);
            return st;
        }

        [ActionName("AgentInsert")]
        public List<MyAgentCode> AgentInsert(MyAgentCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyAgentCode> st = AgentMng.InsertCustoms_Agent(Data);
            return st;
        }

        [ActionName("AgentDelete")]
        public List<MyAgentCode> AgentDelete(MyAgentCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyAgentCode> st = AgentMng.CustomsAgentDeleteMaster(Data);
            return st;
        }

        [ActionName("PortCodeView")]
        public List<MyPortCode> PortCodeView(MyPortCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyPortCode> st = AgentMng.GetCustomsPortCode(Data);
            return st;
        }

        [ActionName("PortInsert")]
        public List<MyPortCode> PortInsert(MyPortCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyPortCode> st = AgentMng.InsertCustoms_Port(Data);
            return st;
        }
        [ActionName("PortDelete")]
        public List<MyPortCode> PortDelete(MyPortCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyPortCode> st = AgentMng.CustomsPortDeleteMaster(Data);
            return st;
        }

        [ActionName("CarrierCodeView")]
        public List<MyCarrierCode> CarrierCodeView(MyCarrierCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyCarrierCode> st = AgentMng.GetCustomsCarrierCode(Data);
            return st;
        }

        [ActionName("CarrierInsert")]
        public List<MyCarrierCode> CarrierInsert(MyCarrierCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyCarrierCode> st = AgentMng.InsertCustoms_Carrier(Data);
            return st;
        }
        [ActionName("CarrierDelete")]
        public List<MyCarrierCode> CarrierDelete(MyCarrierCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyCarrierCode> st = AgentMng.CustomsCarrierDeleteMaster(Data);
            return st;
        }

        [ActionName("TerminalCodeView")]
        public List<MyTerminalCode> TerminalCodeView(MyTerminalCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyTerminalCode> st = AgentMng.GetCustomsTerminalCode(Data);
            return st;
        }

        [ActionName("TerminalInsert")]
        public List<MyTerminalCode> TerminalInsert(MyTerminalCode Data)
        {
            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyTerminalCode> st = AgentMng.InsertCustoms_Terminal(Data);
            return st;
        }
        [ActionName("TerminalDelete")]
        public List<MyTerminalCode> TerminalDelete(MyTerminalCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MyTerminalCode> st = AgentMng.CustomsTerminalDeleteMaster(Data);
            return st;
        }

        [ActionName("CFSDropDown")]
        public List<MYCFS> CFSDropDown(MYCFS Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYCFS> st = AgentMng.GetCustomsCFSName(Data);
            return st;
        }

        [ActionName("CFSView")]
        public List<MYCFS> CFSView(MYCFS Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYCFS> st = AgentMng.GetCustomsCFS(Data);
            return st;
        }

        [ActionName("CFSInsert")]
        public List<MYCFS> CFSInsert(MYCFS Data)
        {
            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYCFS> st = AgentMng.InsertCustoms_CFS(Data);
            return st;
        }
        [ActionName("CFSDelete")]
        public List<MYCFS> CFSDelete(MYCFS Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYCFS> st = AgentMng.CustomsCFSDeleteMaster(Data);
            return st;
        }

        [ActionName("LocationDropDown")]
        public List<MYPackageCode> LocationDropDown(MYPackageCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYPackageCode> st = AgentMng.GetPackageLocation(Data);
            return st;
        }

        [ActionName("PackageCodeView")]
        public List<MYPackageCode> PackageCodeView(MYPackageCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYPackageCode> st = AgentMng.GetCustomsPackage(Data.LocID.ToString());
            return st;
        }

        [ActionName("PackageCodeInsert")]
        public List<MYPackageCode> PackageCodeInsert(MYPackageCode Data)
        {
            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYPackageCode> st = AgentMng.InsertCustoms_PackageCode(Data);
            return st;
        }
        [ActionName("PackageDelete")]
        public List<MYPackageCode> PackageDelete(MYPackageCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYPackageCode> st = AgentMng.CustomsPackageCodeDeleteMaster(Data);
            return st;
        }

        [ActionName("TransporterDropDown")]
        public List<MYTransporter> TransporterDropDown(MYTransporter Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYTransporter> st = AgentMng.GetCustomsTransporterName(Data);
            return st;
        }

        [ActionName("TransporterView")]
        public List<MYTransporter> TransporterView(MYTransporter Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYTransporter> st = AgentMng.GetCustomsTransporter(Data);
            return st;
        }

        [ActionName("TransporterInsert")]
        public List<MYTransporter> TransporterInsert(MYTransporter Data)
        {
            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYTransporter> st = AgentMng.InsertCustoms_Transporter(Data);
            return st;
        }
        [ActionName("TransporterDelete")]
        public List<MYTransporter> TransporterDelete(MYTransporter Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYTransporter> st = AgentMng.CustomsTransporterDeleteMaster(Data);
            return st;
        }

        [ActionName("EqTypeDropDown")]
        public List<MYISOCode> EqTypeDropDown(MYISOCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYISOCode> st = AgentMng.GetCustomsEqTypeName(Data);
            return st;
        }
        [ActionName("ISOCodeView")]
        public List<MYISOCode> ISOCodeView(MYISOCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYISOCode> st = AgentMng.GetCustomsISOCode(Data);
            return st;
        }

        [ActionName("ISOCodeInsert")]
        public List<MYISOCode> ISOCodeInsert(MYISOCode Data)
        {
            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYISOCode> st = AgentMng.InsertCustoms_ISOCode(Data);
            return st;
        }
        [ActionName("ISOCodeDelete")]
        public List<MYISOCode> ISOCodeDelete(MYISOCode Data)
        {

            CustomsCodeManager AgentMng = new CustomsCodeManager();
            List<MYISOCode> st = AgentMng.CustomsISOCodeDeleteMaster(Data);
            return st;
        }
    }
}
