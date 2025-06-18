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
    public class StaffApiController : ApiController
    {
        [ActionName("InsertStaffUser")]
        public List<MyStaffData> InsertStaffUser(MyStaffData Data)
        {
            StaffManager cm = new StaffManager();
            List<MyStaffData> st = cm.InsertStaffMaster(Data);
            return st;
        }

        [ActionName("LocationMaster")]
        public List<MyLocation> LocationMaster(MyLocation Data)
        {
            StaffManager cm = new StaffManager();
            List<MyLocation> st = cm.LocationMasterDtls(Data);
            return st;
        }
        [ActionName("StaffMasterView")]
        public List<MyStaffData> StaffMasterView(MyStaffData Data)
        {
            StaffManager cm = new StaffManager();
            List<MyStaffData> st = cm.GetStaffViewMaster(Data);
            return st;
        }
        [ActionName("StaffMasterRecordView")]
        public List<MyStaffData> StaffMasterRecordView(MyStaffData Data)
        {
            StaffManager cm = new StaffManager();
            List<MyStaffData> st = cm.GetStaffRecordDtls(Data);
            return st;
        }

        [ActionName("StaffRoleMenu")]
        public List<MyRoleMenu> StaffRoleMenu(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetMenuDtls(Data);
            return st;
        }

        [ActionName("StaffRoleInsert")]
        public List<MyRoleMenu> StaffRoleInsert(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.InsertStaffRoleMaster(Data);
            return st;
        }

        [ActionName("StaffRoleView")]
        public List<MyRoleMenu> StaffRoleView(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetStaffRoleMasterDtls(Data);
            return st;
        }

        [ActionName("StaffRoleViewRecord")]
        public List<MyRoleMenu> StaffRoleViewRecord(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetStaffRoleMasterRecord(Data);
            return st;
        }
        [ActionName("StaffRoleBindRecord")]
        public List<MyRoleMenu> StaffRoleBindRecord(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetStaffRoleMasterBindRecord(Data);
            return st;
        }
        [ActionName("StaffRoleSubMenuView")]
        public List<MyRoleMenu> StaffRoleSubMenuView(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetStaffRoleSubmenuMaster(Data);
            return st;
        }
        //[ActionName("ExistingRolesByMenuID")]
        //public List<MyRoleMenu> ExistingRolesByMenuID(MyRoleMenu Data)
        //{
        //    StaffManager cm = new StaffManager();
        //    List<MyRoleMenu> st = cm.ExistingRolesByMenuID(Data);
        //    return st;
        //}
        [ActionName("StaffRoleExistingView")]
        public List<MyRoleMenu> StaffRoleExistingView(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetStaffRoleExistingMaster(Data);
            return st;
        }


        [ActionName("StaffRoleMenuModule")]
        public List<MyRoleMenu> StaffRoleMenuModule(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetMenuModuleDtls(Data);
            return st;
        }

        [ActionName("StaffRoleMenuAccess")]
        public List<MyRoleMenu> StaffRoleMenuAccess(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetMenuAccessDtls(Data);
            return st;
        }



        [ActionName("RoleMenuExistingRecordBindGrid")]
        public List<MyRoleMenu> RoleMenuExistingRecordBindGrid(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetExistingGridProgramName(Data);
            return st;
        }

        [ActionName("RoleMenuExistingRecordBindGridAccess")]
        public List<MyRoleMenu> RoleMenuExistingRecordBindGridAccess(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.GetExistingGridAcess(Data);
            return st;
        }




        [ActionName("RRBLAccessInsert")]
        public List<MyRoleMenu> RRBLAccessInsert(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.RRBLAccessTableInsert(Data);
            return st;
        }


        [ActionName("RRBLAccessExistingRecord")]
        public List<MyRoleMenu> RRBLAccessExistingRecord(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.RRBLAccessExstingTable(Data);
            return st;
        }

        [ActionName("RRBLAccessDelete")]
        public List<MyRoleMenu> RRBLAccessDelete(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.RRBLAccessTableDelete(Data);
            return st;
        }

        [ActionName("StaffRoleAccess")]
        public List<MyRoleMenu> StaffRoleAccess(MyRoleMenu Data)
        {
            StaffManager cm = new StaffManager();
            List<MyRoleMenu> st = cm.StaffRoleMenuDelete(Data);
            return st;
        }


        

    }
}
