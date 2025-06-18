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
    public class CorrectionApiController : ApiController
    {
        [ActionName("ListBLListAgentwise")]
        public List<MYBOL> ListBLListAgentwise(MYBOL Data)
        {
            CorrectionManager cm = new CorrectionManager();

            List<MYBOL> st = cm.ListBLListAgentwise(Data);
            return st;
        }
        [ActionName("ListBookingIDByBL")]
        public List<MYBOL> ListBookingIDByBL(MYBOL Data)
        {
            CorrectionManager cm = new CorrectionManager();

            List<MYBOL> st = cm.ListBookingIDByBL(Data);
            return st;
        }
        [ActionName("BOLPartyDetailsRecord")]
        public List<MYCorrMemo> BOLPartyDetailsRecord(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.BOLPartyDetailsRecord(Data);
            return st;
        }
        [ActionName("BOLConsigneeRecord")]
        public List<MYBOL> BOLConsigneeRecord(MYBOL Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYBOL> st = cm.BOLConsigneeRecord(Data);
            return st;
        }
        [ActionName("BOLNotifyRecord")]
        public List<MYBOL> BOLNotifyRecord(MYBOL Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYBOL> st = cm.BOLNotifyRecord(Data);
            return st;
        }
        [ActionName("BOLCntrExistingRecord")]
        public List<MYCorrMemo> BOLCntrExistingRecord(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.BOLCntrExistingValus(Data);
            return st;
        }
        [ActionName("PartyByAddress")]
        public List<MYCorrMemo> PartyByAddress(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.PartyByAddress(Data);
            return st;
        }

        [ActionName("CorrectionMemoInsert")]
        public List<MYCorrMemo> ChargeCorrectorInsert(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.CorrectionMemoInsert(Data);
            return st;
        }

        [ActionName("CorrectionMemoView")]
        public List<MYCorrMemo> CorrectionMemoView(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.CorrectionMemoViewValues(Data);
            return st;
        }
        [ActionName("CorrectionMemoCountRecord")]
        public List<ChargeCorrectorInsert> CorrectionMemoCountRecord(ChargeCorrectorInsert Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<ChargeCorrectorInsert> st = cm.CorrectionMemoCountRecord(Data);
            return st;
        }
        [ActionName("CorrectionMemoExistingValues")]
        public List<MYCorrMemo> CorrectionMemoExistingValues(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.CorrectionMemoExistingValues(Data);
            return st;
        }
        [ActionName("CorrectionMemoCntrExistingValues")]
        public List<MYCorrMemo> CorrectionMemoCntrExistingValues(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.CorrectionMemoCntrExistingValues(Data);
            return st;
        }
        [ActionName("BOLReleaseExistingViewRecord")]
        public List<MYCorrMemo> BOLReleaseExistingViewRecord(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.BOLReleaseExistingViewRecord(Data);
            return st;
        }

        [ActionName("CorrectionMemoUpdate")]
        public List<MYCorrMemo> ChargeCorrectorUpdate(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.CorrectionMemoupdate(Data);
            return st;
        }

        [ActionName("CorrectionMemoRejected")]
        public List<MYCorrMemo> CorrectionMemoRejected(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.CorrectionMemoReject(Data);
            return st;
        }
        [ActionName("AgencyPartyByAddress")]
        public List<MYCorrMemo> AgencyPartyByAddress(MYCorrMemo Data)
        {
            CorrectionManager cm = new CorrectionManager();
            List<MYCorrMemo> st = cm.AgencyPartyByAddress(Data);
            return st;
        }
    }
}
