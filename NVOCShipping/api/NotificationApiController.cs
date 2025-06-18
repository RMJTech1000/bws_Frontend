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
    public class NotificationApiController : ApiController
    {
        //GET: NotificationApi
       [ActionName("NotificationExemption")]
        public List<MyNotification> NotificationExemption(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.ExemptionNotificationRecordView(Data);
            return st;
        }

        [ActionName("NotificationExemptionValue")]
        public List<MyNotification> NotificationExemptionValue(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.ExemptionNotificationValue(Data);
            return st;
        }

        [ActionName("NotificationBLOpenValues")]
        public List<MyNotification>NotificationBLOpenValues(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.BLOpenValues(Data);
            return st;
        }

        [ActionName("NotificationBLUpdate")]
        public List<MyNotification> NotificationBLUpdate(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.BLStatusCheck(Data);
            return st;
        }

        [ActionName("NotificationSlotAssign")]
        public List<MyNotification> NotificationSlotAssign(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.NotificationSlotAssignValues(Data);
            return st;
        }

        [ActionName("NotificationSlotAssignExisting")]
        public List<MyNotification> NotificationSlotAssignExisting(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.NotificationSlotAssignExistingValues(Data);
            return st;
        }

        [ActionName("NotificationAssignUpdate")]
        public List<MyNotification> NotificationAssignUpdate(MyNotification Data)
        {
            NotificationManager cm = new NotificationManager();
            List<MyNotification> st = cm.SlotAssignUpdate(Data);
            return st;
        }


        

    }
}