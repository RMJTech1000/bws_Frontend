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
    public class BookingApiController : ApiController
    {
        #region mergebl
        [ActionName("BookingBLList")]
        public List<MyMergeBooking> BookingBLList(MyMergeBooking Data)
        {
           BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.GetBookingBLList(Data);
            return st;
        }

        [ActionName("ConfirmBookingBLList")]
        public List<MyMergeBooking> ConfirmBookingBLList(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.GetConfirmBookingBLList(Data);
            return st;
        }

        [ActionName("BLOpenList")]
        public List<MyMergeBooking> BLOpenList(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.GetBLOpenMaster(Data);
            return st;
        }


        [ActionName("ViewMergeBLGrid1")]
        public List<MyMergeBooking> ViewMergeBLGrid1(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.GetMergeBLGrid1(Data);
            return st;
        }

   
        [ActionName("MoveBL1toBL2")]
        public List<MyMergeBooking> MoveBL1toBL2(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.UpdateBOLCntrsMoveBL1toBL2(Data);
            return st;
        }

        [ActionName("MoveBL2toBL1")]
        public List<MyMergeBooking> MoveBL2toBL1(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.UpdateBOLCntrsMoveBL2toBL1(Data);
            return st;
        }
        #endregion

        #region splitbl
        [ActionName("ViewSplitBLGrid")]
        public List<MyMergeBooking> ViewSplitBLGrid(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.GetSplitBLGrid(Data);
            return st;
        }
        [ActionName("VslVoyByBL")]
        public List<MySplitBooking> VslVoyByBL(MySplitBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MySplitBooking> st = cm.GetVslVoyByBL(Data);
            return st;
        }
        [ActionName("SplitBL")]
        public List<MySplitBooking> SplitBL(MySplitBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MySplitBooking> st = cm.InsertSplitBL(Data);
            return st;
        }
        #endregion


        [ActionName("ViewMergeBLUpdate")]
        public List<MyMergeBooking> ViewMergeBLUpdate(MyMergeBooking Data)
        {
            BookingManager cm = new BookingManager();
            List<MyMergeBooking> st = cm.MergeUpdateBL(Data);
            return st;
        }

    }
}
