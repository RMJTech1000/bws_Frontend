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
    public class SubLeaseApiController : ApiController
    {
        #region SI IN
        [ActionName("SubLeasePickUpReference")]
        public List<MyOnHire> SubLeasePickUpReferenceList(MyOnHire Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MyOnHire> st = cm.ListSubLeasePickUpReference(Data);
            return st;
        }

        [ActionName("BindSOContainers")]
        public List<MyCntrPickup> BindSOContainers(MyCntrPickup Data)
        {
            SubLeaseManager Mange = new SubLeaseManager();
            List<MyCntrPickup> st = Mange.BindSOContainers(Data);
            return st;
        }
        
        [ActionName("InsertSOContainers")]
        public List<MySubLeaseDtls> InsertSOContainers(MySubLeaseDtls Data)
        {
            SubLeaseManager cm =  new SubLeaseManager();

            List<MySubLeaseDtls> st = cm.InsertSOContainers(Data);
            return st;
        }
        [ActionName("SubLeaseInView")]
        public List<MySubLeaseDtls> SubLeaseInView(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseInView(Data);
            return st;
        }
        [ActionName("SubLeaseInEdit")]
        public List<MySubLeaseDtls> SubLeaseInEdit(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseInEditRecord(Data);
            return st;
        }
        [ActionName("SubLeaseInContainersEdit")]
        public List<MySubLeaseDtls> SubLeaseInContainersEdit(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseInContainersEditRecord(Data);
            return st;
        }
        [ActionName("SubLeaseInCntrCheck")]
        public List<MySubLeaseDtls> SubLeaseInCntrCheck(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseInCntrCheckRecord(Data);
            return st;
        }
        
        #endregion
        #region SO OUT
        [ActionName("BindAVContainers")]
        public List<MyCntrPickup> BindAVContainers(MyCntrPickup Data)
        {
            SubLeaseManager Mange = new SubLeaseManager();
            List<MyCntrPickup> st = Mange.BindAVContainers(Data);
            return st;
        }


        [ActionName("InsertAVContainers")]
        public List<MySubLeaseDtls> InsertAVContainers(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();

            List<MySubLeaseDtls> st = cm.InsertAVContainers(Data);
            return st;
        }
        [ActionName("SubLeaseOutView")]
        public List<MySubLeaseDtls> SubLeaseOutView(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseOutView(Data);
            return st;
        }
        [ActionName("SubLeaseOutEdit")]
        public List<MySubLeaseDtls> SubLeaseOutEdit(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseOutEditRecord(Data);
            return st;
        }
        [ActionName("SubLeaseOutContainersEdit")]
        public List<MySubLeaseDtls> SubLeaseOutContainersEdit(MySubLeaseDtls Data)
        {
            SubLeaseManager cm = new SubLeaseManager();
            List<MySubLeaseDtls> st = cm.GetSubLeaseOutContainersEditRecord(Data);
            return st;
        }
        #endregion 
    }
}
