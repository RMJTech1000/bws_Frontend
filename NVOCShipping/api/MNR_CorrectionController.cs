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
using System;
using System.Net.Mail;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.IO;

namespace NVOCShipping.api
{
    public class MNR_CorrectionController : ApiController
    {

        [ActionName("MNRHistoryView_correction")]
        public List<MyMNRCorrector> MNRHistoryView_correction(MyMNRCorrector Data)
        {
            MNR_Correction_Manager cm = new MNR_Correction_Manager();
            List<MyMNRCorrector> st = cm.GetMNRHistoryCorrection_List(Data);
            return st;
        }

        [ActionName("MNRReferenceDropdown")]
        public List<MyMNRNew> MNRReferenceDropdown(MyMNRNew Data)
        {
            MNR_Correction_Manager cm = new MNR_Correction_Manager();
            List<MyMNRNew> st = cm.MNRReferenceDropdownData(Data);
            return st;
        }
        [ActionName("MNRReferenceDropdownExisting")]
        public List<MyMNRNew> MNRReferenceDropdownExisting(MyMNRNew Data)
        {
            MNR_Correction_Manager cm = new MNR_Correction_Manager();
            List<MyMNRNew> st = cm.MNRReferenceDropdownExistingData(Data);
            return st;
        }

        [ActionName("MNRStatusByPrevStatus")]
        public List<MyMNRNew> MNRStatusByPrevStatus(MyMNRNew Data)
        {
            MNR_Correction_Manager cm = new MNR_Correction_Manager();
            List<MyMNRNew> st = cm.MNRStatusByPrevStatusDropdownData(Data);
            return st;
        }



        [ActionName("InsertMNRCorrectorSave")]
        public List<MyMNRCorrector> InsertMNRCorrectorSave(MyMNRCorrector Data)
        {
            MNR_Correction_Manager cm = new MNR_Correction_Manager();
            List<MyMNRCorrector> st = cm.InsertMNRCorrectorSave(Data);
            return st;
        }

        [ActionName("MNRCorrectorEdit")]
        public List<MyMNRCorrector> MNRCorrectorEdit(MyMNRCorrector Data)
        {
            MNR_Correction_Manager cm = new MNR_Correction_Manager();
            List<MyMNRCorrector> st = cm.MNRCorrectorEdit(Data);
            return st;
        }

    }
}