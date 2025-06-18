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
    public class VoyageGateWayController : ApiController
    {
        [ActionName("VoyageETD")]
        public List<MyVoyageOpening>VoyageETD(MyVoyageOpening Data)
        {
            GatewayvoyageManager cm = new GatewayvoyageManager();
            List<MyVoyageOpening> st = cm.getwayvoyageETD(Data);
            return st;
        }

        [ActionName("BookingValuedisplay")]
        public List<MyVoyageOpening> BookingValuedisplay(MyVoyageOpening Data)
        {
            GatewayvoyageManager cm = new GatewayvoyageManager();
            List<MyVoyageOpening> st = cm.BookingdisplayVoyageGateway(Data);
            return st;
        }

        [ActionName("insertGateway_VoyageBL")]
        public List<MyVoyageOpening> insertGateway_VoyageBL(MyVoyageOpening Data)
        {
            GatewayvoyageManager cm = new GatewayvoyageManager();
            List<MyVoyageOpening> st = cm.InsertGateWayVoyageBL(Data);
            return st;
        }
        [ActionName("ExisBookingValuedisplay")]
        public List<MyVoyageOpening> ExisBookingValuedisplay(MyVoyageOpening Data)
        {
            GatewayvoyageManager cm = new GatewayvoyageManager();
            List<MyVoyageOpening> st = cm.ExistingBookingVoyageGateway(Data);
            return st;
        }

        [ActionName("ExisBkgValuedisplay")]
        public List<MyVoyageOpening> ExisBkgValuedisplay(MyVoyageOpening Data)
        {
            GatewayvoyageManager cm = new GatewayvoyageManager();
            List<MyVoyageOpening> st = cm.ExistBkgVoyageGateway(Data);
            return st;
        }

        [ActionName("ExisBkgValuedisplayMain")]
        public List<MyVoyageOpening> ExisBkgValuedisplayMain(MyVoyageOpening Data)
        {
            GatewayvoyageManager cm = new GatewayvoyageManager();
            List<MyVoyageOpening> st = cm.ExistBkgVoyageGatewayMain(Data);
            return st;
        }





    }
}