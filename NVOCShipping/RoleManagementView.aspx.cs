using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data.Common;
using System.Data;
using DataManager;
using System.IO;
using System.Security.Cryptography;

namespace NVOCShipping
{
    public partial class RoleManagementView : System.Web.UI.Page
    {
        RegistrationManager RegMng = new RegistrationManager();
        UserManager Mng = new UserManager();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
               
                    BindItemsList();
              
            }
        }
        public DataTable viewImageValue()
        {
            string strWhere = "";
            string _Query = "";
            _Query = " select distinct RId as ID,RoleName from Nav_UserAccessRoleMaster inner join Nav_UserAccessRoleMasterdtls on Nav_UserAccessRoleMasterdtls.RoleId = Nav_UserAccessRoleMaster.RID inner join NVO_Menu on NVO_Menu.Id = NVO_UserAccessRoleMasterdtls.Models";
            if (strWhere == "")
                strWhere = _Query;


            return Mng.GetViewData(strWhere, "");
        }

        #region Funcation Adding


        #region Private Properties
        private int CurrentPage
        {
            get
            {
                object objPage = ViewState["_CurrentPage"];
                int _CurrentPage = 0;
                if (objPage == null)
                {
                    _CurrentPage = 0;
                }
                else
                {
                    _CurrentPage = (int)objPage;
                }
                return _CurrentPage;
            }
            set { ViewState["_CurrentPage"] = value; }
        }
        private int fistIndex
        {
            get
            {

                int _FirstIndex = 0;
                if (ViewState["_FirstIndex"] == null)
                {
                    _FirstIndex = 0;
                }
                else
                {
                    _FirstIndex = Convert.ToInt32(ViewState["_FirstIndex"]);
                }
                return _FirstIndex;
            }
            set { ViewState["_FirstIndex"] = value; }
        }
        private int lastIndex
        {
            get
            {

                int _LastIndex = 0;
                if (ViewState["_LastIndex"] == null)
                {
                    _LastIndex = 0;
                }
                else
                {
                    _LastIndex = Convert.ToInt32(ViewState["_LastIndex"]);
                }
                return _LastIndex;
            }
            set { ViewState["_LastIndex"] = value; }
        }
        #endregion

        #region PagedDataSource
        PagedDataSource _PageDataSource = new PagedDataSource();
        #endregion


        private void BindItemsList()
        {

            try
            {
                DataTable dataTable = viewImageValue();
                if (dataTable.Rows.Count > 0)
                {

                    _PageDataSource.DataSource = dataTable.DefaultView;
                    _PageDataSource.AllowPaging = true;
                    _PageDataSource.PageSize = 8;
                    _PageDataSource.CurrentPageIndex = CurrentPage;
                    ViewState["TotalPages"] = _PageDataSource.PageCount;


                    lblPageInfo.Text = "( Page " + (CurrentPage + 1) + " of  " + _PageDataSource.PageCount + " )";
                    lblTotalRecord.Text = " Total Records  * [ " + dataTable.Rows.Count + " ]  Founds";

                    if (_PageDataSource.PageCount > 1)
                        lblPageInfo.Visible = lbtnFirst.Visible = lbtnPrevious.Visible = lbtnNext.Visible = lbtnLast.Visible = true;

                    else
                        lblPageInfo.Visible = lbtnFirst.Visible = lbtnPrevious.Visible = lbtnNext.Visible = lbtnLast.Visible = false;


                    lbtnPrevious.Enabled = !_PageDataSource.IsFirstPage;
                    lbtnNext.Enabled = !_PageDataSource.IsLastPage;
                    lbtnFirst.Enabled = !_PageDataSource.IsFirstPage;
                    lbtnLast.Enabled = !_PageDataSource.IsLastPage;

                    ddlProductView.DataSource = _PageDataSource;
                    ddlProductView.DataBind();
                    ddlPaging();


                    ddlPager.Visible = ddlProductView.Visible = true;
                }
                else
                {
                    ddlPager.Visible = ddlProductView.Visible = lbtnPrevious.Visible = lbtnNext.Visible = lbtnFirst.Visible = lbtnLast.Visible = lblPageInfo.Visible = false;
                    // displayValues.InnerHtml = "There is no Records";
                    // mpeCourtError.Show();
                    //  return;
                }

            }
            catch (Exception ex)
            {
                //displayValues.InnerHtml = ex.ToString();
                // mpeCourtError.Show();
                // return;
                //string str = ex.Message.Replace("'", "");
                //string Strs = str.Replace("\r\n", "-");
                //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "Window", "<script>alert('" + Strs + "');</script>", false);

            }
        }
        protected void lbtnNext_Click(object sender, EventArgs e)
        {

            CurrentPage += 1;
            this.BindItemsList();

        }
        protected void lbtnPrevious_Click(object sender, EventArgs e)
        {
            CurrentPage -= 1;
            BindItemsList();

        }
        protected void ddlPager_ItemCommand(object source, DataListCommandEventArgs e)
        {
            if (e.CommandName.Equals("Paging"))
            {
                CurrentPage = Convert.ToInt16(e.CommandArgument.ToString());
                BindItemsList();
            }
        }
        protected void ddlPager_ItemDataBound(object sender, DataListItemEventArgs e)
        {
            LinkButton lnkbtnPage = (LinkButton)e.Item.FindControl("lnkbtnPaging");
            if (lnkbtnPage.CommandArgument.ToString() == CurrentPage.ToString())
            {
                lnkbtnPage.Enabled = false;
                lnkbtnPage.Style.Add("fone-size", "14px");
                lnkbtnPage.Font.Bold = true;

            }
        }
        protected void lbtnLast_Click(object sender, EventArgs e)
        {

            CurrentPage = (Convert.ToInt32(ViewState["TotalPages"]) - 1);
            this.BindItemsList();

        }
        protected void lbtnFirst_Click(object sender, EventArgs e)
        {

            CurrentPage = 0;
            BindItemsList();


        }
        private void ddlPaging()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("PageIndex");
            dt.Columns.Add("PageText");

            fistIndex = CurrentPage - 5;


            if (CurrentPage > 5)
            {
                lastIndex = CurrentPage + 5;
            }
            else
            {
                lastIndex = 10;
            }
            if (lastIndex > Convert.ToInt32(ViewState["TotalPages"]))
            {
                lastIndex = Convert.ToInt32(ViewState["TotalPages"]);
                fistIndex = lastIndex - 10;
            }

            if (fistIndex < 0)
            {
                fistIndex = 0;
            }

            for (int i = fistIndex; i < lastIndex; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = i;
                dr[1] = i + 1;
                dt.Rows.Add(dr);
            }
            if (dt.Rows.Count > 1)
            {
                ddlPager.DataSource = dt;
                ddlPager.DataBind();
            }
            else
            {
                ddlPager.DataSource = null;
                ddlPager.DataBind();
            }
        }




        #endregion




        protected void btnsearch_Click(object sender, EventArgs e)
        {
            BindItemsList();
        }
    }
}