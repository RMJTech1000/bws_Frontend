using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DataManager;
using System.Data;
using System.Text;
using DataTier;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Security.Cryptography;

namespace NVOCShipping
{
    public partial class UserManagement : System.Web.UI.Page
    {
        DynaGrid Grd = new DynaGrid();
        RegistrationManager RegMng = new RegistrationManager();
        UserManager Mng = new UserManager();
        MyUser Data = new MyUser();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                InitalBind();
                if (Request.QueryString["RegId"].ToString() != "")
                {
                    
                    BindUserDetails(Request.QueryString["RegId"].ToString());
                }
            }
        }
        public void InitalBind()
        {
            DataTable dtmain = Grd.FillGrid(1);
            mygrid.DataSource = dtmain;
            mygrid.DataBind();
            ImageButton btInsert = (ImageButton)mygrid.Rows[mygrid.Rows.Count - 1].FindControl("btInsert");
            btInsert.Visible = true;
            ImageButton btnDelete = (ImageButton)mygrid.Rows[mygrid.Rows.Count - 1].FindControl("btnDelete");
            btnDelete.Visible = false;
            ViewState["Vue0"] = (DataTable)mygrid.DataSource;

            DataTable dtLoc = GetLocation();
            if (dtLoc.Rows.Count > 0)
            {
                multipleCheckBoxLoc.DataSource = dtLoc;
                multipleCheckBoxLoc.DataTextField = "GeoLocation";
                multipleCheckBoxLoc.DataValueField = "ID";
                multipleCheckBoxLoc.DataBind();

            }
        }

        public void BindUserDetails(string UserID)
        {
            DataTable _dtv = GetUserDetails(UserID);
            if (_dtv.Rows.Count > 0)
            {
                hdId.Value = _dtv.Rows[0]["ID"].ToString();
                lblAgentName.Text = _dtv.Rows[0]["AgencyName"].ToString();
                lblBranch.Text = _dtv.Rows[0]["CityName"].ToString();
                lblUserName.Text = _dtv.Rows[0]["UserName"].ToString();
                lblOldpwd.Text = _dtv.Rows[0]["Password"].ToString();
                lblEmailID.Text = _dtv.Rows[0]["EmailID"].ToString();
                if (_dtv.Rows[0]["Active"].ToString() == "True")
                    chkActive.Checked = true;
                else
                    chkActive.Checked = false;

                var RoleLocation = _dtv.Rows[0]["RoleLocation"].ToString().Split(',');
                for (int i = 0; i < RoleLocation.Length; i++)
                {
                    for (int j = 0; j < multipleCheckBoxLoc.Items.Count; j++)
                    {
                        if (multipleCheckBoxLoc.Items[j].Value == RoleLocation[i].ToString())
                            multipleCheckBoxLoc.Items[j].Selected = true;
                    }
                }

            }
            DataTable dtRolEx = GetRoleValues(UserID);
            if (dtRolEx.Rows.Count > 0)
            {
                mygrid.DataSource = dtRolEx;
                mygrid.DataBind();
                ImageButton btInsert = (ImageButton)mygrid.Rows[mygrid.Rows.Count - 1].FindControl("btInsert");
                btInsert.Visible = true;
                ImageButton btnDelete = (ImageButton)mygrid.Rows[mygrid.Rows.Count - 1].FindControl("btnDelete");
                btnDelete.Visible = false;

                ViewState["Vue0"] = (DataTable)mygrid.DataSource;
            }
        }
        public DataTable GetRoleValues(string Id)
        {
            string _Query = "select RoleID as Field1 from NVO_UserRoleDetails where StaffID =" + Id;
            return RegMng.GetViewData(_Query, "");
        }
        protected void btnReset_Click(object sender, EventArgs e)
        {
            txtpassword.Text = "";
            string disp_name = string.Empty;
            Random rd = new Random();
            txtpassword.Text = "Nerida" + rd.Next(10000000).ToString();
        }
        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            DataTable _DTint = new DataTable();
            _DTint.Columns.Add("RoleID");

            Data.ID = Int32.Parse(hdId.Value);
            if (txtpassword.Text != "")
                Data.Password = txtpassword.Text;
            else
                Data.Password = lblOldpwd.Text.Trim();

            if (chkActive.Checked == true)
                Data.IsActive = true;
            else
                Data.IsActive = false;

            foreach (GridViewRow gRow in mygrid.Rows)
            {

                DropDownList ddlRole = (DropDownList)gRow.FindControl("ddVal1");
                if (ddlRole.SelectedValue != "")
                {
                    _DTint.Rows.Add(_DTint.NewRow());
                    _DTint.Rows[_DTint.Rows.Count - 1]["RoleID"] = ddlRole.SelectedValue;
                }
            }

            string RoleLocation = "";
            foreach (ListItem item in multipleCheckBoxLoc.Items)
            {
                if (item.Selected) RoleLocation += (item.Value) + ",";
            }
            if (RoleLocation != "")
                RoleLocation = RoleLocation.Substring(0, RoleLocation.Length - 1);
            Data.Description = RoleLocation;

            if (Mng.UpdateUserMaster(Data, _DTint) >0)
            {
                string message1 = "Record Saved Successfully";
                ScriptManager.RegisterStartupScript((sender as Control), this.GetType(), "Popup", "ShowSave('" + message1 + "');", true);
            }
        }

        protected void btnNewonClick(object sender, EventArgs e)
        {
            UpdateTableGrid(mygrid);

            ImageButton ib1 = (ImageButton)sender;
            GridViewRow gRow = (GridViewRow)ib1.NamingContainer;
            InsertRowGrid(mygrid, gRow.RowIndex);

        }

        private DataTable UpdateTableGrid(GridView GridView1)
        {

            DataTable dt = (DataTable)ViewState["Vue0"];
            foreach (GridViewRow gRow in GridView1.Rows)
            {
                #region UpdateValue


                DropDownList ddVal1 = (DropDownList)gRow.FindControl("ddVal1");
                dt.Rows[gRow.RowIndex]["Field1"] = ddVal1.SelectedValue;




                #endregion
            }
            return dt;
        }

        private void InsertRowGrid(GridView GridView1, int gridRowIndex)
        {
            DataTable dt = (DataTable)ViewState["Vue0"];
            dt.Rows.InsertAt(dt.NewRow(), gridRowIndex + 1);
            mygrid.DataSource = dt;
            mygrid.DataBind();

            ImageButton btnDelete = (ImageButton)mygrid.Rows[mygrid.Rows.Count - 1].FindControl("btnDelete");
            ImageButton btInsert = (ImageButton)mygrid.Rows[mygrid.Rows.Count - 1].FindControl("btInsert");
            btInsert.Visible = true;
            btnDelete.Visible = false;
        }
        protected void mygrid_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataTable _dtvRole = GetRolevalues();
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddVal1 = (DropDownList)e.Row.FindControl("ddVal1");
                Label lblField0 = (Label)e.Row.FindControl("lblField0");
                ddVal1.DataSource = _dtvRole;
                ddVal1.DataTextField = "RoleName";
                ddVal1.DataValueField = "RID";
                ddVal1.DataBind();
                ddVal1.Items.Insert(0, new ListItem("", ""));
                ddVal1.SelectedValue = lblField0.Text;
            }
        }
        public DataTable GetRolevalues()
        {
            string _Query = "select * from Nav_UserAccessRoleMaster";
            return RegMng.GetViewData(_Query, "");
        }
        public DataTable GetUserDetails(string UserID)
        {
            string _Query = " select NVO_UserDetails.ID, UserName,AgencyName,CityName,EmailID,Password,Active,RoleLocation from NVO_UserDetails " +
                            " inner join NVO_AgencyMaster On NVO_AgencyMaster.ID = NVO_UserDetails.AgentID " +
                            " inner join NVO_CityMaster On NVO_CityMaster.ID = NVO_UserDetails.BranchID where NVO_UserDetails.ID=" + UserID;
            return RegMng.GetViewData(_Query, "");
        }
        public DataTable GetLocation()
        {
            string _query = "select * from NVO_GeoLocations";
            return RegMng.GetViewData(_query,"");
        }

        protected void btnDelete_Click(object sender, ImageClickEventArgs e)
        {
            string message = "Are you sure  want to Delete!!(Y/N)";
            ScriptManager.RegisterStartupScript((sender as Control), this.GetType(), "Popup", "ShowPopupDelete('" + message + "');", true);
            ImageButton btnedit = (ImageButton)sender;
            GridViewRow gRow = (GridViewRow)btnedit.NamingContainer;
            Label lbl = (Label)gRow.FindControl("lblField0");
            ConfirmBit.Value = lbl.Text;
        }

        public DataTable DeleteValue(string Id)
        {
            string _Query = "delete from Nav_UserRoleDetails where RoleID =" + Id + " and StaffID=" + hdId.Value;
            return RegMng.GetViewData(_Query, "");
        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            string Id = ConfirmBit.Value;
            DeleteValue(Id);
            BindUserDetails(Request.QueryString["RegId"].ToString());
            ConfirmBit.Value = "";
            return;
        }

        //public DataTable GetRoleLocation(string RoleID)
        //{
        //    string _query = "select * from NVO_GeoLocations where Id in(" + RoleID + ")";
        //    return RegMng.GetViewData(_query, "");
        //}
    }
}
