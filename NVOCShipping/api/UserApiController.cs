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
    public class UserApiController : ApiController
    {
        #region anand
        [ActionName("InsertUser")]
        public List<MyUser> InsertUser(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.InsertUserMaster(Data);
            return st;
        }

        [ActionName("InsertUserRole")]
        public List<MyUser> InsertUserRole(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.InsertUserRoleMaster(Data);
            return st;
        }

        [ActionName("UserView")]
        public List<MyUser> UserView(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.GetUserViewMaster(Data);
            return st;
        }

        [ActionName("UserViewRecord")]
        public List<MyUser> UserViewRecord(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.GetUserMasterRecord(Data.ID.ToString());
            return st;
        }


        [ActionName("UserViewExstingRole")]
        public List<MyUser> UserViewExstingRole(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.GetUserMasterRole(Data.ID.ToString());
            return st;
        }



        [ActionName("Division")]
        public List<MyDivision> Division(MyDivision Data)
        {
            UserManager cm = new UserManager();
            List<MyDivision> st = cm.GetDivisionMaster(Data);
            return st;
        }
        [ActionName("Department")]
        public List<MyDepartment> Department(MyDepartment Data)
        {
            UserManager cm = new UserManager();
            List<MyDepartment> st = cm.GetDepartementMaster(Data);
            return st;
        }
        [ActionName("Designation")]
        public List<MyDesignation> Designation(MyDesignation Data)
        {
            UserManager cm = new UserManager();
            List<MyDesignation> st = cm.GetDesignationMaster(Data);
            return st;
        }


        [HttpPost,ActionName("UserDetails")]
        public List<MyUser> UserDetails(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.GetUserDetails();
            return st;
        }



        [ActionName("NotificationMaster")]
        public List<MyUser> NotificationMaster(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.NotificationMaster();
            return st;
        }


        [ActionName("InsertNotification")]
        public List<MyUser> InsertNotification(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.InsertNotificationMaster(Data);
            return st;
        }


        [ActionName("NotificationMasterUser")]
        public List<MyUser> NotificationMasterUser(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.NotificationMasterUser(Data);
            return st;
        }

        [ActionName("ExistingNotificationMaster")]
        public List<MyUser> ExistingNotificationMaster(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.NotificationExistingMasterUser(Data);
            return st;
        }
        [ActionName("UserAndUserIDValidation")]
        public List<MyUser> UserAndUserIDValidation(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.UserAndUserIDValidation(Data);
            return st;
        }


        [ActionName("InsertMISExRate")]
        public List<MyUser> InsertMISExRate(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.InsertMISExRateMaster(Data);
            return st;
        }



        [ActionName("ExistingMISExRateMaster")]
        public List<MyUser> ExistingMISExRateMaster(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.ExistingExRateMaster(Data);
            return st;
        }

        [ActionName("ExtMISExRateRecord")]
        public List<MyUser> ExtMISExRateRecord(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.ExtExRateMaster(Data);
            return st;
        }


        [ActionName("ExtMISExRateRecorddtls")]
        public List<MyUser> ExtMISExRateRecorddtls(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.ExtExMisRatedtls(Data);
            return st;
        }

        [ActionName("ExtMISExRate")]
        public List<MyUser> ExtMISExRate(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.ExtExMisRatedtlsMis(Data);
            return st;
        }


        [ActionName("UserRoledeleteMaster")]
        public List<MyUser> UserRoledeleteMaster(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.DeleteUserRoleMaster(Data);
            return st;
        }

        [ActionName("UserRoleNotification")]
        public List<MyUser> UserRoleNotification(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.UserAccessRoleNotification(Data);
            return st;
        }

        [ActionName("ExtMISExRateRecordDelete")]
        public List<MyUser> ExtMISExRateRecordDelete(MyUser Data)
        {
            UserManager cm = new UserManager();
            List<MyUser> st = cm.DeleteMISExRateMaster(Data);
            return st;
        }

        #endregion
    }
}
