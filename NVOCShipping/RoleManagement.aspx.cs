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
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Security.Cryptography;


namespace NVOCShipping
{
    public partial class RoleManagement : System.Web.UI.Page
    {
        RegistrationManager RegMng = new RegistrationManager();
        UserManager Mng = new UserManager();
        MyRole BussData = new MyRole();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                BindDropdown();
                //if (Request.QueryString["RegId"] != null)
                //    ExistingValues(Request.QueryString["RegId"].ToString());
            }
        }
        protected void ddlRol_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlRol.SelectedValue != "")
                BindGrid(ddlRol.SelectedValue);
        }
        public void BindDropdown()
        {
            DataTable _dts = GetFirstMenu();
            if (_dts.Rows.Count > 0)
            {
                ddlRol.DataSource = _dts;
                ddlRol.DataTextField = "FileName";
                ddlRol.DataValueField = "ID";
                ddlRol.DataBind();
                ddlRol.Items.Insert(0, new System.Web.UI.WebControls.ListItem("---Select---", ""));
            }
        }
        public DataTable GetFirstMenu()
        {
            string _query = "select * from NVO_Menu  where  MenuID =0";
            return Mng.GetViewData(_query, "");

        }

        public void BindGrid(string Id)
        {
            DataTable _dtExtV = new DataTable();
            DataTable _dtBind = GetSecoudMenu(Id);
            mygridRole.DataSource = _dtBind;
            mygridRole.DataBind();


            if (HDID.Value != "")
                _dtExtV = GetExitingRoleData(HDID.Value);

            foreach (GridViewRow gRow in mygridRole.Rows)
            {
                GridView GridInvoice = (GridView)gRow.FindControl("GridInvoice");
                Label lblScoundMenu = (Label)gRow.FindControl("lblScoundMenu");
                Label lblUrl = (Label)gRow.FindControl("lblUrl");
                CheckBox chkMSearch = (CheckBox)gRow.FindControl("chkMSearch");
                CheckBox chkMEdit = (CheckBox)gRow.FindControl("chkMEdit");
                CheckBox chkMdelete = (CheckBox)gRow.FindControl("chkMdelete");
                CheckBox chkMPrint = (CheckBox)gRow.FindControl("chkMPrint");
                CheckBox chkScoundchk = (CheckBox)gRow.FindControl("chkScoundchk");


                if (lblUrl.Text == "")
                {
                    chkMSearch.Visible = false; chkMEdit.Visible = false;
                    chkMdelete.Visible = false; chkMPrint.Visible = false;
                    chkScoundchk.Visible = false;
                }
                if (HDID.Value != "")
                {
                    DataRow[] DrRow = _dtExtV.Select("MenuID=" + lblScoundMenu.Text);
                    if (DrRow.Length > 0)
                    {
                        chkMSearch.Checked = true; chkMEdit.Checked = true;
                        chkMdelete.Checked = true; chkMPrint.Checked = true;
                        chkScoundchk.Checked = true;
                    }
                    else
                    {
                        chkMSearch.Checked = false; chkMEdit.Checked = false;
                        chkMdelete.Checked = false; chkMPrint.Checked = false;
                        chkScoundchk.Checked = false;
                    }
                }

                GridInvoice.DataSource = GetThiredMenu(lblScoundMenu.Text);
                GridInvoice.DataBind();
                foreach (GridViewRow gRow1 in GridInvoice.Rows)
                {

                    CheckBox chkTried = (CheckBox)gRow1.FindControl("chkTriedchk");
                    Label lblThiredMenu = (Label)gRow1.FindControl("lblThiredMenu");
                    CheckBox chkTHSearch = (CheckBox)gRow1.FindControl("chkThSearch");
                    CheckBox chkThEdit = (CheckBox)gRow1.FindControl("chkThEdit");
                    CheckBox chkThdelete = (CheckBox)gRow1.FindControl("chkThdelete");
                    CheckBox chkThPrint = (CheckBox)gRow1.FindControl("chkThPrint");
                    if (HDID.Value != "")
                    {
                        DataRow[] DrRow = _dtExtV.Select("MenuID=" + lblThiredMenu.Text);
                        if (DrRow.Length > 0)
                        {
                            chkTried.Checked = true; chkTHSearch.Checked = true;
                            chkThEdit.Checked = true; chkThdelete.Checked = true;
                            chkThPrint.Checked = true;
                        }
                        else
                        {
                            chkTried.Checked = false; chkTHSearch.Checked = false;
                            chkThEdit.Checked = false; chkThdelete.Checked = false;
                            chkThPrint.Checked = false;
                        }
                    }
                }



            }

        }
        public DataTable GetSecoudMenu(string Id)
        {
            string _query = "select * from NVO_Menu  where  MenuID =" + Id;
            return Mng.GetViewData(_query, "");

        }

        public DataTable GetThiredMenu(string Id)
        {
            string _query = "select * from NVO_Menu  where  MenuID =" + Id;
            return Mng.GetViewData(_query, "");

        }

        protected void mygridRole_RowDataBound(object sender, GridViewRowEventArgs e)
        {

            mygridRole.Columns[5].Visible = false;
        }

        protected void GridInvoice_RowDataBound(object sender, GridViewRowEventArgs e)
        {

        }
        protected void chkTriedchk_CheckedChanged(object sender, EventArgs e)
        {

        }
        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            DataTable _DTint = new DataTable();
            _DTint.Columns.Add("Menulevel");
            _DTint.Columns.Add("MenuID");
            _DTint.Columns.Add("MainMenuID");
            _DTint.Columns.Add("Deleteboolean");
            _DTint.Columns.Add("Edit");
            _DTint.Columns.Add("Search");
            _DTint.Columns.Add("PrintBoolean");


            if (HDID.Value != "")
                BussData.RId = Int32.Parse(HDID.Value);
            else
                BussData.RId = 0;

            BussData.RoleName = txtRole.Text;
            BussData.Models = Int32.Parse(ddlRol.SelectedValue);

            foreach (GridViewRow gRow in mygridRole.Rows)
            {
                GridView GridInvoice = (GridView)gRow.FindControl("GridInvoice");
                CheckBox chkScoundchk = (CheckBox)gRow.FindControl("chkScoundchk");
                Label lblScoundMenu = (Label)gRow.FindControl("lblScoundMenu");
                Label lblUrl = (Label)gRow.FindControl("lblUrl");
                CheckBox chkMSearch = (CheckBox)gRow.FindControl("chkMSearch");
                CheckBox chkMEdit = (CheckBox)gRow.FindControl("chkMEdit");
                CheckBox chkMdelete = (CheckBox)gRow.FindControl("chkMdelete");
                CheckBox chkMPrint = (CheckBox)gRow.FindControl("chkMPrint");
                Label lblMenuID = (Label)gRow.FindControl("lblMenuID");


                if (chkScoundchk.Checked == true)
                {
                    _DTint.Rows.Add(_DTint.NewRow());

                    _DTint.Rows[_DTint.Rows.Count - 1]["Menulevel"] = "2";
                    _DTint.Rows[_DTint.Rows.Count - 1]["MenuID"] = lblScoundMenu.Text;
                    _DTint.Rows[_DTint.Rows.Count - 1]["Deleteboolean"] = chkMdelete.Checked;
                    _DTint.Rows[_DTint.Rows.Count - 1]["Edit"] = chkMEdit.Checked;
                    _DTint.Rows[_DTint.Rows.Count - 1]["Search"] = chkMSearch.Checked;
                    _DTint.Rows[_DTint.Rows.Count - 1]["PrintBoolean"] = chkMPrint.Checked;
                    _DTint.Rows[_DTint.Rows.Count - 1]["MainMenuID"] = lblScoundMenu.Text;
                }
                foreach (GridViewRow gRow1 in GridInvoice.Rows)
                {

                    CheckBox chkTried = (CheckBox)gRow1.FindControl("chkTriedchk");
                    Label lblThiredMenu = (Label)gRow1.FindControl("lblThiredMenu");
                    CheckBox chkTHSearch = (CheckBox)gRow1.FindControl("chkThSearch");
                    CheckBox chkThEdit = (CheckBox)gRow1.FindControl("chkThEdit");
                    CheckBox chkThdelete = (CheckBox)gRow1.FindControl("chkThdelete");
                    CheckBox chkThPrint = (CheckBox)gRow1.FindControl("chkThPrint");
                    Label lblThiredMenuID = (Label)gRow1.FindControl("lblThiredMenuID");

                    if (chkTried.Checked == true)
                    {

                        _DTint.Rows.Add(_DTint.NewRow());
                        _DTint.Rows[_DTint.Rows.Count - 1]["Menulevel"] = "3";
                        _DTint.Rows[_DTint.Rows.Count - 1]["MenuID"] = lblThiredMenu.Text;
                        _DTint.Rows[_DTint.Rows.Count - 1]["Deleteboolean"] = chkThdelete.Checked;
                        _DTint.Rows[_DTint.Rows.Count - 1]["Edit"] = chkThEdit.Checked;
                        _DTint.Rows[_DTint.Rows.Count - 1]["Search"] = chkTHSearch.Checked;
                        _DTint.Rows[_DTint.Rows.Count - 1]["PrintBoolean"] = chkThdelete.Checked;
                        _DTint.Rows[_DTint.Rows.Count - 1]["MainMenuID"] = lblThiredMenuID.Text;

                    }
                }

            }
            if (Mng.InsertUserRollMaster(BussData, _DTint) > 0)
            {
                HDID.Value = BussData.RId.ToString();
                string message1 = "Record Saved Successfully";
                ScriptManager.RegisterStartupScript((sender as Control), this.GetType(), "Popup", "ShowPopup('" + message1 + "');", true);
            }



        }

        public void ExistingValues(string Id)
        {
            DataTable _dtEx = GetExitingData(Id);
            if (_dtEx.Rows.Count > 0)
            {

                HDID.Value = _dtEx.Rows[0]["RoleId"].ToString();
                ddlRol.SelectedValue = _dtEx.Rows[0]["Models"].ToString();
                txtRole.Text = _dtEx.Rows[0]["RoleName"].ToString().ToUpper();
                BindGrid(ddlRol.SelectedValue);
            }
        }

        public DataTable GetExitingData(string Id)
        {
            string _query = " Select distinct RoleId,Models,RoleName from Nav_UserAccessRoleMaster inner join Nav_UserAccessRoleMasterdtls on Nav_UserAccessRoleMasterdtls.RoleID = Nav_UserAccessRoleMaster.RID where RID=" + Id;
            return Mng.GetViewData(_query, "");

        }

        public DataTable GetExitingRoleData(string Id)
        {
            string _query = " select * from Nav_UserAccessRoleMasterdtls where RoleId = " + Id;
            return Mng.GetViewData(_query, "");

        }
    }
}