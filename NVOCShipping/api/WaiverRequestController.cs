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
    public class WaiverRequestController : ApiController
    {

        [ActionName("WaierBookingNo")]
        public List<MyWaiver> WaierBookingNo(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.WaiverBLNumber();
            return st;
        }

        [ActionName("WaierBookingCntrNo")]
        public List<MyWaiver> WaierBookingCntrNo(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.WaiverBkgCntrNo(Data);
            return st;
        }

        [ActionName("InsertWaier")]
        public List<MyWaiver> InsertWaier(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.InsertWaiverMaster(Data);
            return st;
        }

        [ActionName("UpdateStatusWaier")]
        public List<MyWaiver> UpdateStatusWaier(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.UpdateWaiverStatus(Data);
            return st;
        }

        [ActionName("WaierView")]
        public List<MyWaiver> WaierView(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.WaiverView(Data);
            return st;
        }

        [ActionName("WaierViewValues")]
        public List<MyWaiver> WaierViewValues(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.WaiverViewValues(Data);
            return st;
        }

        [ActionName("WaiverChargeList")]
        public List<MyWaiver> WaiverChargeList(MyWaiver Data)
        {

            WaiverManager Mange = new WaiverManager();
            List<MyWaiver> st = Mange.ChargCodeList(Data);
            return st;
        }


    }
}
