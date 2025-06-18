using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;
using DataManager;
using Common.Types;
using System.Net.Mail;
using System.Net;


    public partial class SICLQuotationForm : System.Web.UI.Page
    {
        DynaGrid Grd = new DynaGrid();
        CommonFunctionManager SQLFun = new CommonFunctionManager();
        SeqNumberGenManager SeqNoGen = new SeqNumberGenManager();
        MyRateSheet Data = new MyRateSheet();
        RateQutationManager RateManag = new RateQutationManager();

        RateSheetManager RSM = new RateSheetManager();
        public enum MessageType { Success, Error, Info, Warning };
        protected void Page_Load(object sender, EventArgs e)
        {

            txtFreeDays.Attributes.Add("onkeypress", "return numericvalidation();");


            //if (Session["LocName"].ToString().ToUpper() == "MUMBAI" || Session["LocName"].ToString().ToUpper() == "COCHIN" || Session["LocName"].ToString().ToUpper() == "CHENNAI" || Session["LocName"].ToString().ToUpper() == "DUBAI")
            //{
            // string em = "hari";
            //}
            //else
            //{
            //    string ems = "hari1";
            //}

            if (!IsPostBack)
            {
                InitalBind();
                BindDropDown();

                if (Request.QueryString["R_RSheetID"] != null)
                    ExistingData(Request.QueryString["R_RSheetID"].ToString());

            }

             if(Request.QueryString["Status"] != null)
                {
                  if(Request.QueryString["Status"].ToString() == "3".Trim())
                  {
                      btnApproval.Visible = false;
                      btnSave.Visible = false;
                
                  }
                }

        }

        public void InitalBind()
        {
            DataTable dtmain = Grd.FillGrid(12);
            GView_BuyingRate.DataSource = dtmain;
            GView_BuyingRate.DataBind();
            ImageButton btInsert = (ImageButton)GView_BuyingRate.Rows[GView_BuyingRate.Rows.Count - 1].FindControl("btnBRInsert");
            btInsert.Visible = true;
            ViewState["Vue0"] = (DataTable)GView_BuyingRate.DataSource;

            DataTable dtmain1 = Grd.FillGrid(12);
            GView_SellingRate.DataSource = dtmain1;
            GView_SellingRate.DataBind();
            ImageButton btInsert1 = (ImageButton)GView_SellingRate.Rows[GView_SellingRate.Rows.Count - 1].FindControl("btnSRInsert");
            btInsert1.Visible = true;
            ViewState["Vue1"] = (DataTable)GView_SellingRate.DataSource;

            DataTable dtmain2 = Grd.FillGrid(13);
            GView_RR.DataSource = dtmain2;
            GView_RR.DataBind();
            ImageButton btInsert2 = (ImageButton)GView_RR.Rows[GView_RR.Rows.Count - 1].FindControl("btnRRInsert");
            btInsert2.Visible = true;
            ViewState["Vue2"] = (DataTable)GView_RR.DataSource;

            DataTable dtmain3 = Grd.FillGrid(3);
            GView_CntrTypes.DataSource = dtmain3;
            GView_CntrTypes.DataBind();
            ImageButton btInsert3 = (ImageButton)GView_CntrTypes.Rows[GView_CntrTypes.Rows.Count - 1].FindControl("btnCTInsert");
            btInsert3.Visible = true;
            ViewState["Vue3"] = (DataTable)GView_CntrTypes.DataSource;







        }

        #region "Trade Selected Changed"
        protected void ddTrade_SelectedIndexChanged(object sender, EventArgs e)
        {
            optFH.Checked = false;
            optAN.Checked = false;
            optSG.Checked = false;
            optBN.Checked = false;

            if (ddTrade.SelectedValue.ToString() == "-1")
            {
                ImpOpt.Visible = false;
                ImpOpt.Visible = false;
                ACHolder.Visible = false;
                FreeDays.Visible = false;
                Carrier_Freedays.Visible = false;

            }

            if (ddTrade.SelectedValue.ToString() == "1")
            {
                ImpOpt.Visible = true;
                FreeDays.Visible = true;
                Carrier_Freedays.Visible = true;
            }
            else
            {
                ImpOpt.Visible = false;
                FreeDays.Visible = false;
                Carrier_Freedays.Visible = false;
            }

            if (ddTrade.SelectedValue.ToString() == "2")
                ExpOpt.Visible = true;
            else
                ExpOpt.Visible = false;

            if (ddTrade.SelectedValue.ToString() == "3")
                ACHolder.Visible = true;
            else
                ACHolder.Visible = false;
        }
        #endregion


        #region "Terms Of Shipment Selected Changed"
        protected void ddTermsOfShipment_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddTermsOfShipment.SelectedValue.ToString() == "-1")
            {
                FPayableAt.Visible = false;
                CollectingAgent.Visible = false;
            }
            if (ddTermsOfShipment.SelectedValue.ToString() == "1")
            {
                FPayableAt.Visible = false;
                CollectingAgent.Visible = false;
            }
            else
            {
                FPayableAt.Visible = true;
                CollectingAgent.Visible = true;
            }

            if (ddTermsOfShipment.SelectedValue.ToString() == "2")
            {
                FPayableAt.Visible = true;
                CollectingAgent.Visible = true;
            }
            else
            {
                FPayableAt.Visible = false;
                CollectingAgent.Visible = false;
            }
        }
        #endregion


        #region "Checkbox changes (Collection Agent)"
        protected void ChkCollectAgent_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkCollectAgent.Checked == true)
                DivCollectAgent.Visible = true;
            else
                DivCollectAgent.Visible = false;
        }
        #endregion
        #region "Business Type Changed"
        protected void ddBusinessTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
        //if (ddBusinessTypes.SelectedItem.Text.ToString() == "CONSOLE")
        //    FS_BuyRate.Visible = false;
        //else
        //    FS_BuyRate.Visible = true;

        //if (ddBusinessTypes.SelectedItem.Text.ToString() == "FCL BROKING/CLEARANCE" || ddBusinessTypes.SelectedItem.Text.ToString() == "LCL BROKING/CLEARANCE")
        //{
        //    Carrier_PLD.Visible = false;
        //    Carrier_OverseasAgent.Visible = false;
        //}
        //else
        //{
        //    Carrier_PLD.Visible = true;
        //    Carrier_OverseasAgent.Visible = true;
        //}

        //if (ddBusinessTypes.SelectedItem.Text.ToString() == "CONSOLE")
        //    FS_BuyRate.Visible = false;
        //else
        //    FS_BuyRate.Visible = true;

        if (ddBusinessTypes.SelectedItem.Text.ToString() == "FCL BROKING" || ddBusinessTypes.SelectedItem.Text.ToString() == "LCL BROKING")
        {
            Carrier_PLD.Visible = false;
            Carrier_OverseasAgent.Visible = false;
        }
        else
        {
            Carrier_PLD.Visible = true;
            Carrier_OverseasAgent.Visible = true;
        }
    }
        #endregion

        #region "Option Buttons Changed"
        protected void optFH_CheckedChanged(object sender, EventArgs e)
        {
            if (optAN.Checked == true)
                ACHolder.Visible = true;
            else
                ACHolder.Visible = false;
        }
        protected void optAN_CheckedChanged(object sender, EventArgs e)
        {
            if (optAN.Checked == true)
                ACHolder.Visible = true;
            else
                ACHolder.Visible = false;
        }
        protected void optSG_CheckedChanged(object sender, EventArgs e)
        {
            if (optSG.Checked == true)
                ACHolder.Visible = true;
            else
                ACHolder.Visible = false;
        }
        protected void optBN_CheckedChanged(object sender, EventArgs e)
        {
            if (optSG.Checked == true)
                ACHolder.Visible = true;
            else
                ACHolder.Visible = false;
        }
        #endregion


       

        protected void ddBusinessTypes_SelectedIndexChanged1(object sender, EventArgs e)
        {

        }
        protected void GView_CntrTypes_RowCommand(object sender, GridViewCommandEventArgs e)
        {

        }
        protected void GView_CntrTypes_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataTable _dtCntrTypes = RSM.Get_CntrType();
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddCT = (DropDownList)e.Row.FindControl("ddlVal0");

                Label lblsno = (Label)e.Row.FindControl("lblSNo");
                lblsno.Visible = false;

                HiddenField HdCnVal1 = (HiddenField)e.Row.FindControl("HdCnVal1");
                ddCT.DataSource = _dtCntrTypes;
                ddCT.DataValueField = "ID";
                ddCT.DataTextField = "Type";
                ddCT.DataBind();
                //ddCT.Items.Insert(0, new ListItem("", "-1"));

                if (HdCnVal1 != null)
                    ddCT.SelectedValue = HdCnVal1.Value;
            }
        }
        protected void GView_CntrTypes_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {

        }

        protected void lnkBuyRate_Click(object sender, EventArgs e)
        {

        }
        protected void btnBRInsert_Click(object sender, ImageClickEventArgs e)
        {
            UpdateTableGrid(GView_BuyingRate);
            ImageButton ib1 = (ImageButton)sender;
            GridViewRow gRow = (GridViewRow)ib1.NamingContainer;
            InsertRowGrid(GView_BuyingRate, gRow.RowIndex);

        }
        private DataTable UpdateTableGrid(GridView GridView1)
        {
            DataTable dt = (DataTable)ViewState["Vue0"];

            foreach (GridViewRow gRow in GridView1.Rows)
            {
                HiddenField HdVal0 = (HiddenField)gRow.FindControl("HdVal0");
                TextBox txtVal0 = (TextBox)gRow.FindControl("txtVal0");

                HiddenField HdVal1 = (HiddenField)gRow.FindControl("HdVal1");
                DropDownList ddlVal1 = (DropDownList)gRow.FindControl("ddlVal1");

                HiddenField HdVal2 = (HiddenField)gRow.FindControl("HdVal2");
                DropDownList ddlVal2 = (DropDownList)gRow.FindControl("ddlVal2");

                HiddenField HdVal3 = (HiddenField)gRow.FindControl("HdVal3");
                DropDownList ddlVal3 = (DropDownList)gRow.FindControl("ddlVal3");

                HiddenField HdVal4 = (HiddenField)gRow.FindControl("HdVal4");
                DropDownList ddlVal4 = (DropDownList)gRow.FindControl("ddlVal4");

                TextBox txtVal5 = (TextBox)gRow.FindControl("txtVal5");


                if (HdVal0.Value != "")
                    dt.Rows[gRow.RowIndex]["Field1"] = HdVal0.Value;
                else
                    dt.Rows[gRow.RowIndex]["Field1"] = 0;
                dt.Rows[gRow.RowIndex]["Field2"] = txtVal0.Text;

                dt.Rows[gRow.RowIndex]["Field3"] = ddlVal1.SelectedValue;

                dt.Rows[gRow.RowIndex]["Field5"] = ddlVal2.SelectedValue;
                dt.Rows[gRow.RowIndex]["Field7"] = ddlVal3.SelectedValue;
                dt.Rows[gRow.RowIndex]["Field9"] = ddlVal4.SelectedValue;

                if (txtVal5.Text != "")
                    dt.Rows[gRow.RowIndex]["Field11"] = txtVal5.Text;
                else
                    dt.Rows[gRow.RowIndex]["Field11"] = "0.00";




            }
            return dt;
        }
        private void InsertRowGrid(GridView GridView1, int gridRowIndex)
        {

            DataTable dt = (DataTable)ViewState["Vue0"];
            dt.Rows.InsertAt(dt.NewRow(), gridRowIndex + 1);
            GView_BuyingRate.DataSource = dt;
            GView_BuyingRate.DataBind();
            ImageButton btInsert = (ImageButton)GView_BuyingRate.Rows[GView_BuyingRate.Rows.Count - 1].FindControl("btnBRInsert");
            btInsert.Visible = true;



        }


        protected void GView_BuyingRate_RowDataBound(object sender, GridViewRowEventArgs e)
        {


            DataTable _dtRateRcvdFm = RSM.Get_RateRcvdFrom();
            DataTable _dtBRCntrTypes = RSM.Get_CntrType();
            DataTable _dtBRBasis = RSM.Get_Basis();
            DataTable _dtBRCurrency = Get_Currency();
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddRRF = (DropDownList)e.Row.FindControl("ddlVal1");
                HiddenField HdVal1 = (HiddenField)e.Row.FindControl("HdVal1");
                ddRRF.DataSource = _dtRateRcvdFm;
                ddRRF.DataValueField = "ID";
                ddRRF.DataTextField = "RateReceivedFrom";
                ddRRF.DataBind();
                ddRRF.Items.Insert(0, new ListItem("", "-1"));

                if (HdVal1 != null)
                    ddRRF.SelectedValue = HdVal1.Value;

                DropDownList ddBRCT = (DropDownList)e.Row.FindControl("ddlVal2");
                HiddenField HdVal2 = (HiddenField)e.Row.FindControl("HdVal2");
                ddBRCT.DataSource = _dtBRCntrTypes;
                ddBRCT.DataValueField = "ID";
                ddBRCT.DataTextField = "Type";
                ddBRCT.DataBind();
                ddBRCT.Items.Insert(0, new ListItem("", "-1"));

                if (HdVal2 != null)
                    ddBRCT.SelectedValue = HdVal2.Value;

                DropDownList ddBRBS = (DropDownList)e.Row.FindControl("ddlVal3");
                HiddenField HdVal3 = (HiddenField)e.Row.FindControl("HdVal3");
                ddBRBS.DataSource = _dtBRBasis;
                ddBRBS.DataValueField = "ID";
                ddBRBS.DataTextField = "Basis";
                ddBRBS.DataBind();
                ddBRBS.Items.Insert(0, new ListItem("", "-1"));

                if (HdVal3 != null)
                    ddBRBS.SelectedValue = HdVal3.Value;

                DropDownList ddBRCR = (DropDownList)e.Row.FindControl("ddlVal4");
                HiddenField HdVal4 = (HiddenField)e.Row.FindControl("HdVal4");
                ddBRCR.DataSource = _dtBRCurrency;
                ddBRCR.DataValueField = "ID";
                ddBRCR.DataTextField = "Code";
                ddBRCR.DataBind();
                ddBRCR.Items.Insert(0, new ListItem("", "-1"));

                if (HdVal4 != null)
                    ddBRCR.SelectedValue = HdVal4.Value;
            }
        }

        public DataTable Get_Currency()
        {
            string _Query = "SELECT * FROM F_TblCurrencies WHERE ID IN(28,76) ORDER BY Code";
            return SQLFun.GetSQLFunction(_Query);
        }

        protected void btnSRInsert_Click(object sender, ImageClickEventArgs e)
        {
            UpdateTableGridSR(GView_SellingRate);
            ImageButton ib1 = (ImageButton)sender;
            GridViewRow gRow = (GridViewRow)ib1.NamingContainer;
            InsertRowGridSR(GView_SellingRate, gRow.RowIndex);
        }
        private void InsertRowGridSR(GridView GridView1, int gridRowIndex)
        {
            DataTable dt = (DataTable)ViewState["Vue1"];
            dt.Rows.InsertAt(dt.NewRow(), gridRowIndex + 1);
            GView_SellingRate.DataSource = dt;
            GView_SellingRate.DataBind();
            ImageButton btInsert = (ImageButton)GView_SellingRate.Rows[GView_SellingRate.Rows.Count - 1].FindControl("btnSRInsert");
            btInsert.Visible = true;



        }
        private DataTable UpdateTableGridSR(GridView GridView1)
        {
            DataTable dt = (DataTable)ViewState["Vue1"];

            foreach (GridViewRow gRow in GridView1.Rows)
            {
                HiddenField HdVal0 = (HiddenField)gRow.FindControl("HdSRVal0");
                TextBox txtVal0 = (TextBox)gRow.FindControl("txtSRVal0");

                HiddenField HdVal1 = (HiddenField)gRow.FindControl("HdSRVal1");
                DropDownList ddlVal1 = (DropDownList)gRow.FindControl("ddlVal1");

                HiddenField HdVal2 = (HiddenField)gRow.FindControl("HdSRVal2");
                DropDownList ddlVal2 = (DropDownList)gRow.FindControl("ddlVal2");

                HiddenField HdVal3 = (HiddenField)gRow.FindControl("HdSRVal3");
                DropDownList ddlVal3 = (DropDownList)gRow.FindControl("ddlVal3");

                HiddenField HdVal4 = (HiddenField)gRow.FindControl("HdSRVal4");
                DropDownList ddlVal4 = (DropDownList)gRow.FindControl("ddlVal4");

                TextBox txtVal5 = (TextBox)gRow.FindControl("txtSRVal5");

                if (HdVal0.Value != "")
                    dt.Rows[gRow.RowIndex]["Field1"] = HdVal0.Value;
                else
                    dt.Rows[gRow.RowIndex]["Field1"] = 0;
                dt.Rows[gRow.RowIndex]["Field2"] = txtVal0.Text;

                dt.Rows[gRow.RowIndex]["Field3"] = ddlVal1.SelectedValue;

                dt.Rows[gRow.RowIndex]["Field5"] = ddlVal2.SelectedValue;
                dt.Rows[gRow.RowIndex]["Field7"] = ddlVal3.SelectedValue;
                dt.Rows[gRow.RowIndex]["Field9"] = ddlVal4.SelectedValue;

                if (txtVal5.Text != "")
                    dt.Rows[gRow.RowIndex]["Field11"] = txtVal5.Text;
                else
                    dt.Rows[gRow.RowIndex]["Field11"] = "0.00";
            }
            return dt;
        }


        protected void GView_SellingRate_RowDataBound(object sender, GridViewRowEventArgs e)
        {


            DataTable _dtRateQuotedTo = RSM.Get_RateQuotedTo();
            DataTable _dtSRCntrTypes = RSM.Get_CntrType();
            DataTable _dtSRBasis = RSM.Get_Basis();
            DataTable _dtSRCurrency = Get_Currency();
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddRQT = (DropDownList)e.Row.FindControl("ddlVal1");
                HiddenField HdSRVal1 = (HiddenField)e.Row.FindControl("HdSRVal1");
                ddRQT.DataSource = _dtRateQuotedTo;
                ddRQT.DataValueField = "ID";
                ddRQT.DataTextField = "RateQuotedTo";
                ddRQT.DataBind();
                ddRQT.Items.Insert(0, new ListItem("", "-1"));

                if (HdSRVal1 != null)
                    ddRQT.SelectedValue = HdSRVal1.Value;

                DropDownList ddSRCT = (DropDownList)e.Row.FindControl("ddlVal2");
                HiddenField HdSRVal2 = (HiddenField)e.Row.FindControl("HdSRVal2");
                ddSRCT.DataSource = _dtSRCntrTypes;
                ddSRCT.DataValueField = "ID";
                ddSRCT.DataTextField = "Type";
                ddSRCT.DataBind();
                ddSRCT.Items.Insert(0, new ListItem("", "-1"));

                if (HdSRVal2 != null)
                    ddSRCT.SelectedValue = HdSRVal2.Value;

                DropDownList ddSRBS = (DropDownList)e.Row.FindControl("ddlVal3");
                HiddenField HdSRVal3 = (HiddenField)e.Row.FindControl("HdSRVal3");
                ddSRBS.DataSource = _dtSRBasis;
                ddSRBS.DataValueField = "ID";
                ddSRBS.DataTextField = "Basis";
                ddSRBS.DataBind();
                ddSRBS.Items.Insert(0, new ListItem("", "-1"));

                if (HdSRVal3 != null)
                    ddSRBS.SelectedValue = HdSRVal3.Value;

                DropDownList ddSRCR = (DropDownList)e.Row.FindControl("ddlVal4");
                HiddenField HdSRVal4 = (HiddenField)e.Row.FindControl("HdSRVal4");
                ddSRCR.DataSource = _dtSRCurrency;
                ddSRCR.DataValueField = "ID";
                ddSRCR.DataTextField = "Code";
                ddSRCR.DataBind();
                ddSRCR.Items.Insert(0, new ListItem("", "-1"));

                if (HdSRVal4 != null)
                    ddSRCR.SelectedValue = HdSRVal4.Value;
            }



        }


        protected void btnRRInsert_Click(object sender, ImageClickEventArgs e)
        {
            UpdateTableGridRR(GView_RR);
            ImageButton ib1 = (ImageButton)sender;
            GridViewRow gRow = (GridViewRow)ib1.NamingContainer;
            InsertRowGridRR(GView_RR, gRow.RowIndex);
        }

        private void InsertRowGridRR(GridView GridView1, int gridRowIndex)
        {
            DataTable dt = (DataTable)ViewState["Vue2"];
            dt.Rows.InsertAt(dt.NewRow(), gridRowIndex + 1);
            GView_RR.DataSource = dt;
            GView_RR.DataBind();
            ImageButton btInsert = (ImageButton)GView_RR.Rows[GView_RR.Rows.Count - 1].FindControl("btnRRInsert");
            btInsert.Visible = true;



        }

        private DataTable UpdateTableGridRR(GridView GridView1)
        {
            DataTable dt = (DataTable)ViewState["Vue2"];

            foreach (GridViewRow gRow in GridView1.Rows)
            {
                HiddenField HdVal0 = (HiddenField)gRow.FindControl("HdRRVal0");
                DropDownList ddlVal0 = (DropDownList)gRow.FindControl("ddlVal0");

                HiddenField HdVal1 = (HiddenField)gRow.FindControl("HdRRVal1");
                DropDownList ddlVal1 = (DropDownList)gRow.FindControl("ddlVal1");


                HiddenField HdVal2 = (HiddenField)gRow.FindControl("HdRRVal2");
                DropDownList ddlVal2 = (DropDownList)gRow.FindControl("ddlVal2");

                HiddenField HdVal3 = (HiddenField)gRow.FindControl("HdRRVal3");
                DropDownList ddlVal3 = (DropDownList)gRow.FindControl("ddlVal3");

                HiddenField HdVal4 = (HiddenField)gRow.FindControl("HdRRVal4");
                DropDownList ddlVal4 = (DropDownList)gRow.FindControl("ddlVal4");

                TextBox txtVal5 = (TextBox)gRow.FindControl("txtRRVal5");

                dt.Rows[gRow.RowIndex]["Field1"] = ddlVal0.SelectedValue;
                //  dt.Rows[gRow.RowIndex]["Field2"] = txtVal0.Text;

                dt.Rows[gRow.RowIndex]["Field3"] = ddlVal1.SelectedValue;
                // dt.Rows[gRow.RowIndex]["Field4"] = txtVal1.Text;



                dt.Rows[gRow.RowIndex]["Field5"] = ddlVal2.SelectedValue;
                // dt.Rows[gRow.RowIndex]["Field6"] = txtVal2.Text;

                dt.Rows[gRow.RowIndex]["Field7"] = ddlVal3.SelectedValue;
                //dt.Rows[gRow.RowIndex]["Field8"] = txtVal3.Text;

                dt.Rows[gRow.RowIndex]["Field9"] = ddlVal4.SelectedValue;

                // dt.Rows[gRow.RowIndex]["Field10"] = txtVal4.Text;

                dt.Rows[gRow.RowIndex]["Field11"] = txtVal5.Text;

            }
            return dt;
        }


        protected void GView_RR_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            DataTable _dtTransaction = RSM.Get_Transaction();
            DataTable _dtRRCntrTypes = RSM.Get_CntrType();
            DataTable _dtBasis = RSM.Get_Basis();
            DataTable _dtCurrency = Get_Currency();
            DataTable _dtRateQuotedTo = RSM.Get_RateQuotedTo();
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddTrans = (DropDownList)e.Row.FindControl("ddlVal0");
                HiddenField HdRRVal0 = (HiddenField)e.Row.FindControl("HdRRVal0");
                ddTrans.DataSource = _dtTransaction;
                ddTrans.DataValueField = "Value";
                ddTrans.DataTextField = "Description";
                ddTrans.DataBind();
                ddTrans.Items.Insert(0, new ListItem("", "-1"));

                if (HdRRVal0 != null)
                    ddTrans.SelectedValue = HdRRVal0.Value;

                DropDownList ddRRCT = (DropDownList)e.Row.FindControl("ddlVal2");
                HiddenField HdRRVal2 = (HiddenField)e.Row.FindControl("HdRRVal2");
                ddRRCT.DataSource = _dtRRCntrTypes;
                ddRRCT.DataValueField = "ID";
                ddRRCT.DataTextField = "Type";
                ddRRCT.DataBind();
                ddRRCT.Items.Insert(0, new ListItem("", "-1"));

                if (HdRRVal2 != null)
                    ddRRCT.SelectedValue = HdRRVal2.Value;

                DropDownList ddBS = (DropDownList)e.Row.FindControl("ddlVal3");
                HiddenField HdRRVal3 = (HiddenField)e.Row.FindControl("HdRRVal3");
                ddBS.DataSource = _dtBasis;
                ddBS.DataValueField = "ID";
                ddBS.DataTextField = "Basis";
                ddBS.DataBind();
                ddBS.Items.Insert(0, new ListItem("", "-1"));

                if (HdRRVal3 != null)
                    ddBS.SelectedValue = HdRRVal3.Value;

                DropDownList ddCR = (DropDownList)e.Row.FindControl("ddlVal4");
                HiddenField HdRRVal4 = (HiddenField)e.Row.FindControl("HdRRVal4");
                ddCR.DataSource = _dtCurrency;
                ddCR.DataValueField = "ID";
                ddCR.DataTextField = "Code";
                ddCR.DataBind();
                ddCR.Items.Insert(0, new ListItem("", "-1"));

                if (HdRRVal4 != null)
                    ddCR.SelectedValue = HdRRVal4.Value;


                HiddenField txtTrans = (HiddenField)e.Row.FindControl("HdRRVal0");
                if (txtTrans != null)
                {
                    if (txtTrans.Value.ToString() == "1")
                    {
                        DataTable _dtRateRcvdFm = RSM.Get_RateRcvdFrom();
                        if (e.Row.RowType == DataControlRowType.DataRow)
                        {
                            DropDownList ddFT = (DropDownList)e.Row.FindControl("ddlVal1");
                            HiddenField HdRRVal1 = (HiddenField)e.Row.FindControl("HdRRVal1");
                            ddFT.DataSource = _dtRateRcvdFm;
                            ddFT.DataValueField = "ID";
                            ddFT.DataTextField = "RateReceivedFrom";
                            ddFT.DataBind();
                            ddFT.Items.Insert(0, new ListItem("", "-1"));

                            if (HdRRVal1 != null)
                                ddFT.SelectedValue = HdRRVal1.Value;
                        }
                    }
                    else
                    {

                        if (e.Row.RowType == DataControlRowType.DataRow)
                        {
                            DropDownList ddFT = (DropDownList)e.Row.FindControl("ddlVal1");
                            HiddenField HdRRVal1 = (HiddenField)e.Row.FindControl("HdRRVal1");
                            ddFT.DataSource = _dtRateQuotedTo;
                            ddFT.DataValueField = "ID";
                            ddFT.DataTextField = "RateQuotedTo";
                            ddFT.DataBind();
                            ddFT.Items.Insert(0, new ListItem("", "-1"));

                            if (HdRRVal1 != null)
                                ddFT.SelectedValue = HdRRVal1.Value;
                        }
                    }
                }
            }
        }

        protected void ddlVal0_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList dd = (DropDownList)sender;
            GridViewRow gr = (GridViewRow)dd.Parent.Parent;

            if (dd.SelectedValue.ToString() == "1")
            {
                DataTable _dtRateRcvdFm = RSM.Get_RateRcvdFrom();
                DropDownList ddFT = (DropDownList)gr.FindControl("ddlVal1");
                ddFT.DataSource = _dtRateRcvdFm;
                ddFT.DataValueField = "ID";
                ddFT.DataTextField = "RateReceivedFrom";
                ddFT.DataBind();
                ddFT.Items.Insert(0, new ListItem("", "-1"));
            }
            else
            {
                DataTable _dtRateQuotedTo = RSM.Get_RateQuotedTo();
                DropDownList ddFT = (DropDownList)gr.FindControl("ddlVal1");
                ddFT.DataSource = _dtRateQuotedTo;
                ddFT.DataValueField = "ID";
                ddFT.DataTextField = "RateQuotedTo";
                ddFT.DataBind();
                ddFT.Items.Insert(0, new ListItem("", "-1"));
            }
        }


        protected void btnCTInsert_Click(object sender, ImageClickEventArgs e)
        {
            UpdateTableGridCN(GView_CntrTypes);
            ImageButton ib1 = (ImageButton)sender;
            GridViewRow gRow = (GridViewRow)ib1.NamingContainer;
            InsertRowGridCN(GView_CntrTypes, gRow.RowIndex);
        }


        private void InsertRowGridCN(GridView GridView1, int gridRowIndex)
        {
            DataTable dt = (DataTable)ViewState["Vue3"];

            

            //if (dt.Rows[gridRowIndex]["Field1"].ToString() != "")
            //{
                dt.Rows.InsertAt(dt.NewRow(), gridRowIndex + 1);
                GView_CntrTypes.DataSource = dt;
                GView_CntrTypes.DataBind();
                ImageButton btInsert = (ImageButton)GView_CntrTypes.Rows[GView_CntrTypes.Rows.Count - 1].FindControl("btnCTInsert");
                btInsert.Visible = true;
           // }



        }

        private DataTable UpdateTableGridCN(GridView GridView1)
        {
            DataTable dt = (DataTable)ViewState["Vue3"];

            foreach (GridViewRow gRow in GridView1.Rows)
            {


                HiddenField HdVal1 = (HiddenField)gRow.FindControl("HdCnVal1");
                TextBox txtVal1 = (TextBox)gRow.FindControl("txtCnVal1");
                TextBox txtCNVal2 = (TextBox)gRow.FindControl("txtCNVal2");
                DropDownList ddlVal0 = (DropDownList)gRow.FindControl("ddlVal0");
                Label lblID = (Label)gRow.FindControl("lblID");

                if (ddlVal0.SelectedValue != "" && txtCNVal2.Text != "")
                {
                    dt.Rows[gRow.RowIndex]["Field1"] = ddlVal0.SelectedValue;
                    //dt.Rows[gRow.RowIndex]["Field2"] = txtVal1.Text;
                    if (txtCNVal2.Text != "")
                        dt.Rows[gRow.RowIndex]["Field3"] = txtCNVal2.Text;

                    if (lblID.Text == "")
                        dt.Rows[gRow.RowIndex]["ID"] = 0;
                    else
                        dt.Rows[gRow.RowIndex]["ID"] = lblID.Text;
                    //else
                    //    dt.Rows[gRow.RowIndex]["Field3"] = "";
                }


                
               
               
            }
            return dt;
        }



        protected void lnkSellRate_Click(object sender, EventArgs e)
        {

        }

        protected void ddRR_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddRR.SelectedValue.ToString() == "1")
                FSet_RR.Visible = true;
            else
                FSet_RR.Visible = false;
        }

        protected void chkSingleUse_CheckedChanged(object sender, EventArgs e)
        {

            if (chkSingleUse.Checked.ToString() == "True") // Added by Prithivi 6-April-2016 
            {
                chkSingleUse.Checked = true;
                divValidity.Visible = false;
            }
            else
            {
                chkSingleUse.Checked = false;
                divValidity.Visible = true;
               
            }


        }


        #region Funcation
        public void BindDropDown()
        {
            DataTable _dtTd = BindTrid();
            ddTrade.DataSource = _dtTd;
            ddTrade.DataTextField = "Trade";
            ddTrade.DataValueField = "ID";
            ddTrade.DataBind();
            ddTrade.Items.Insert(0, new ListItem("", ""));

            DataTable _dtGeo = BindGeoLocation();
            ddHandleLocation.DataSource = _dtGeo;
            ddHandleLocation.DataTextField = "Location";
            ddHandleLocation.DataValueField = "ID";
            ddHandleLocation.DataBind();
            ddHandleLocation.Items.Insert(0, new ListItem("", ""));

            DataTable _dtBuType = BindBusinessTypes();
            ddBusinessTypes.DataSource = _dtBuType;
            ddBusinessTypes.DataTextField = "BusinessType";
            ddBusinessTypes.DataValueField = "ID";
            ddBusinessTypes.DataBind();
            ddBusinessTypes.Items.Insert(0, new ListItem("", ""));

            DataTable _dtservice = BindGetService();
            ddServiceTypes.DataSource = _dtservice;
            ddServiceTypes.DataTextField = "ServiceType";
            ddServiceTypes.DataValueField = "ID";
            ddServiceTypes.DataBind();
            ddServiceTypes.Items.Insert(0, new ListItem("", ""));

            DataTable _dtserviceMode = BindGetServiceMode();
            ddServiceMode.DataSource = _dtserviceMode;
            ddServiceMode.DataTextField = "ServiceMode";
            ddServiceMode.DataValueField = "ID";
            ddServiceMode.DataBind();
            ddServiceMode.Items.Insert(0, new ListItem("", ""));

            DataTable _dtAcchold = BindAccountHold();
            ddACHolder.DataSource = _dtAcchold;
            ddACHolder.DataTextField = "ContactPerson";
            ddACHolder.DataValueField = "ID";
            ddACHolder.DataBind();
            ddACHolder.Items.Insert(0, new ListItem("", ""));

            DataTable _dtTems = BindTermsShip();
            ddTermsOfShipment.DataSource = _dtTems;
            ddTermsOfShipment.DataTextField = "Description";
            ddTermsOfShipment.DataValueField = "Value";
            ddTermsOfShipment.DataBind();
            ddTermsOfShipment.Items.Insert(0, new ListItem("", ""));


            ddTermsOfShipment_Carrier.DataSource = _dtTems;
            ddTermsOfShipment_Carrier.DataTextField = "Description";
            ddTermsOfShipment_Carrier.DataValueField = "Value";
            ddTermsOfShipment_Carrier.DataBind();
            ddTermsOfShipment_Carrier.Items.Insert(0, new ListItem("", ""));

            ddRR.Items.Insert(0, new ListItem("NO", "0"));
            ddRR.Items.Insert(1, new ListItem("YES", "1"));

        }

        DataTable BindTrid()
        {
            string _Query = "SELECT ID,Trade FROM F_TblTrade where Id != 3 ";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindGeoLocation()
        {
            string _Query = "SELECT ID,Location FROM F_TblGeolocations ORDER BY Location";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindBusinessTypes()
        {
            string _Query = "SELECT ID,BusinessType FROM F_TblBusinessTypes ORDER BY BusinessType";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindGetService()
        {
            string _Query = "SELECT ID,ServiceType FROM F_TblServiceTypes ORDER BY ServiceType";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindGetServiceMode()
        {
            string _Query = "SELECT ID,ServiceMode FROM F_TblServiceMode ORDER BY ServiceMode";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindAccountHold()
        {
            string _Query = "SELECT ID,ContactPerson FROM F_Users WHERE IsActive=1";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindTermsShip()
        {
            string _Query = "SELECT Value,Description FROM F_TblDLValues WHERE DLTypeID = 9";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindPortCodeValue(string GeoLocId)
        {
            string _Query = "SELECT LocCode FROM F_TblGeoLocations where ID= " + GeoLocId;
            return SQLFun.GetSQLFunction(_Query);
        }


        DataTable BindExisting(string RsID)
        {
            string _Query = "SELECT RS.ID, RS.RateSheetNo, CONVERT(varchar, RS.DtRateSheet,103) AS DtRateSheet, RS.Trade, CM.ID AS BkgPartyID, CM.CustomerName AS BkgPartyName, "
                                + "P1.ID AS PLRID, P1.Name as PLRName, P2.ID AS POLID, P2.Name as POLName, P3.ID AS PODID, P3.Name as PODName, "
                                + "P4.ID AS PLDID,  P4.Name as PLDName, RS.BusinessTypeID, RS.ServiceTypeID, RS.ServiceModeID, RS.TermsOfShipment, "
                                + "P5.ID AS FreightPayableID, P5.Name AS FreightPayableLoc, RS.IsCollectingAgentInvolved, "
                                + "(SELECT ID FROM FF_CustomerMaster WHERE ID = RS.CollectingAgentID AND Types = 8) AS CollectingAgentID, "
                                + "(SELECT CustomerName FROM FF_CustomerMaster WHERE ID = RS.CollectingAgentID AND Types = 8) AS CollectingAgentLoc, "
                                + "RS.Freedays, RS.FreehandOrNomination, RS.SelfGenaratedOrNomination, RS.ACHolderID, "
                                + "RS.ShippingLine AS ShpLnID, VM.VendorName AS ShpLnName, RS.TermsOfShipmentID_Carrier, "
                                + "P6.ID AS PLDID_Carrier, P6.Name AS PLDName_Carrier, RS.Freedays_Carrier, "
                                + "(SELECT ID FROM FF_CustomerMaster WHERE ID = RS.OverseasAgentID AND Types = 8) AS OverseasAgentID, "
                                + "(SELECT CustomerName FROM FF_CustomerMaster WHERE ID = RS.OverseasAgentID AND Types = 8) AS OverseasAgentLoc, "
                                + "RS.IsRR, RS.ExpireAfterSingleUse, CONVERT(varchar, RS.DtValidity,103) AS DtValidity, RS.HandleLocationID, RS.Remarks, "
                                + "(SELECT ID FROM F_TblCommodity WHERE ID = RS.CommodityID) AS CommodityID, "
                                + "(SELECT Commodity FROM F_TblCommodity WHERE ID = RS.CommodityID) AS CommodityName ,RS.Status,RS.DtApproval,RS.Comments,(SELECT contactperson FROM f_users WHERE ID = RS.approvedby) as ApprovedBy " //,FileAttachment "
                                + " FROM F_RateSheet RS "
                                + "INNER JOIN FF_CustomerMaster CM ON RS.BkgPartyID = CM.ID "
                                + "INNER JOIN FF_VendorMaster VM ON RS.ShippingLine = VM.ID "
                                + "INNER JOIN F_Ports P1 ON RS.PLRID = P1.ID "
                                + "INNER JOIN F_Ports P2 ON RS.POLID = P2.ID "
                                + "INNER JOIN F_Ports P3 ON RS.PODID = P3.ID "
                                + "INNER JOIN F_Ports P4 ON RS.PLDID = P4.ID "
                                + "LEFT OUTER JOIN F_Ports P5 ON RS.FreightPayableAt = P5.ID "
                                + "LEFT OUTER JOIN F_Ports P6 ON RS.PlaceofDeliveryID_Carrier = P6.ID "
                                + "WHERE RS.ID = " + RsID;
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindExistingCntr(string RsID)
        {
            string _Query = " select CntrTypeID as Field1,Type + ' ' + Size AS Field2,ApproxUnits as Field3, F_RatesheetCntrTypes.ID  from F_RatesheetCntrTypes  " +
                            " inner join F_TblCntrTypes on F_TblCntrTypes.ID = F_RatesheetCntrTypes.CntrTypeID where Rsheetid = " + RsID;
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindExistingSelling(string RsID)
        {
            string _Query = " select SrChargesId as Field1,ChgCodedes as Field2,SRRateQuotedTo as Field3, RateQuotedTo as Field4, " +
                            " SRCntrTypeID as Field5,Type + ' ' + Size AS Field6,SRBasis as Field7, " +
                            " Basis as Field8,SRCurrID as Field9,Code as Field10,SRAmount  as Field11,F_SellingRateDtls.ID as Field12 from F_SellingRateDtls " +
                            " inner join ChargeTB on ChargeTB.ID = F_SellingRateDtls.SrChargesId " +
                            " inner join F_TblRateQutdTo on F_TblRateQutdTo.ID=F_SellingRateDtls.SRRateQuotedTo " +
                            " inner join F_TblCntrTypes on F_TblCntrTypes.ID =F_SellingRateDtls.SRCntrTypeID " +
                            " inner join F_TblBasis on F_TblBasis.Id =F_SellingRateDtls.SRBasis " +
                            " inner join F_TblCurrencies on F_TblCurrencies.Id = F_SellingRateDtls.SRCurrID " +
                            " where RatesheetId = " + RsID + " order by F_SellingRateDtls.ID";
            return SQLFun.GetSQLFunction(_Query);
        }

        DataTable BindExistingBuying(string RsID)
        {
            string _Query = " select BRChargesID as Field1,ChgCodedes as Field2,BRRateRcvdFrom as Field3, RateReceivedFrom as Field4," +
                            " BRCntrTypeID as Field5,Type + ' ' + Size AS Field6,BRBasis as Field7, " +
                            " Basis as Field8,BRCurrID as Field9,Code as Field10,BRAmount  as Field11,buyID  as Field12 from F_BuyingRateDtls  " +
                            " inner join ChargeTB on ChargeTB.ID = F_BuyingRateDtls.BrChargesId " +
                            " inner join F_TblRateRcvdFrom on F_TblRateRcvdFrom.ID=F_BuyingRateDtls.BRRateRcvdFrom " +
                            " inner join F_TblCntrTypes on F_TblCntrTypes.ID =F_BuyingRateDtls.BRCntrTypeID " +
                            " inner join F_TblBasis on F_TblBasis.Id =F_BuyingRateDtls.BRBasis " +
                            " inner join F_TblCurrencies on F_TblCurrencies.Id = F_BuyingRateDtls.BRCurrID " +
                            " where RatesheetId = " + RsID + " order by buyID";
            return SQLFun.GetSQLFunction(_Query);
        }


        DataTable BindExistingRebat(string RsID)
        {
            string _Query = " select TransTypeID as Field1,Description as Field2,RateRFromQTo as Field3,case when TransTypeID = 1 then (select top(1) RateReceivedFrom from  F_TblRateRcvdFrom  " +
                            " where F_TblRateRcvdFrom.ID = F_RateSheetRebateRefund.RateRFromQTo) else (select top(1) RateQuotedTo from  F_TblRateQutdTo  " +
                            " where F_TblRateQutdTo.ID = F_RateSheetRebateRefund.RateRFromQTo) end as Field4,CntrTypeID as Field5, " +
                            " Type + ' ' + Size AS Field6,F_RateSheetRebateRefund.Basis as Field7,F_TblBasis.Basis as Field8,CurrID as Field9,Code as Field10,Amount as Field11, RebateID as Field12 " +
                            " from F_RateSheetRebateRefund " +
                            " inner join F_TblDLValues on F_TblDLValues.Value = F_RateSheetRebateRefund.TransTypeID  and  DLTypeID=217 " +
                            " inner join F_TblCntrTypes on F_TblCntrTypes.ID =F_RateSheetRebateRefund.CntrTypeID " +
                            " inner join F_TblBasis on F_TblBasis.Id =F_RateSheetRebateRefund.Basis " +
                            " inner join F_TblCurrencies on F_TblCurrencies.Id = F_RateSheetRebateRefund.CurrID  " +
                            " where RatesheetId = " + RsID;
            return SQLFun.GetSQLFunction(_Query);
        }

        #endregion

        public void ExistingData(string ReID)
        {
            DataTable _dtExt = BindExisting(ReID);
            #region ContrilsBind
            if (_dtExt.Rows.Count > 0)
            {
                if (Request.QueryString["Resue"] == "1")
                {
                    HdRateID.Value = "";
                    lblRateSheetNo.Text = "";
                    txtRSDate.Text = "";
                    txtValidityDt.Text = System.DateTime.Today.Date.ToString();
                    displayApproval.Visible = false;
                  
                }
                else
                {
                    HdRateID.Value = ReID;
                    lblRateSheetNo.Text = _dtExt.Rows[0]["RateSheetNo"].ToString();
                    txtRSDate.Text = _dtExt.Rows[0]["DtRateSheet"].ToString();
                    displayApproval.Visible = true;
                }


                ddTrade.SelectedValue = _dtExt.Rows[0]["Trade"].ToString();
                txtBkgParty.Text = _dtExt.Rows[0]["BkgPartyName"].ToString();
                hddBkgParty.Value = _dtExt.Rows[0]["BkgPartyID"].ToString();
                txtPLR.Text = _dtExt.Rows[0]["PLRName"].ToString();
                hddPLR.Value = _dtExt.Rows[0]["PLRID"].ToString();
                txtPOL.Text = _dtExt.Rows[0]["POLName"].ToString();
                hddPOL.Value = _dtExt.Rows[0]["POLID"].ToString();
                txtPOD.Text = _dtExt.Rows[0]["PODName"].ToString();
                hddPOD.Value = _dtExt.Rows[0]["PODID"].ToString();
                txtPLD.Text = _dtExt.Rows[0]["PLDName"].ToString();
                hddPLD.Value = _dtExt.Rows[0]["PLDID"].ToString();
                ddBusinessTypes.SelectedValue = _dtExt.Rows[0]["BusinessTypeID"].ToString();

                ddServiceTypes.SelectedValue = _dtExt.Rows[0]["ServiceTypeID"].ToString();
                ddServiceMode.SelectedValue = _dtExt.Rows[0]["ServiceModeID"].ToString();

                ddTermsOfShipment.SelectedValue = _dtExt.Rows[0]["TermsOfShipment"].ToString();

                if (_dtExt.Rows[0]["TermsOfShipment"].ToString() == "1")
                {
                    FPayableAt.Visible = false;
                    CollectingAgent.Visible = false;
                }
                else
                {
                    FPayableAt.Visible = true;
                    CollectingAgent.Visible = true;

                    txtPayAt.Text = _dtExt.Rows[0]["FreightPayableLoc"].ToString();
                    hddPayAt.Value = _dtExt.Rows[0]["FreightPayableID"].ToString();

                    if (_dtExt.Rows[0]["IsCollectingAgentInvolved"].ToString() == "True")
                    {
                        DivCollectAgent.Visible = true;
                        ChkCollectAgent.Checked = true;
                        txtOverseas.Text = _dtExt.Rows[0]["CollectingAgentLoc"].ToString();
                        hddOverseas.Value = _dtExt.Rows[0]["CollectingAgentID"].ToString();
                    }
                    else
                    {
                        DivCollectAgent.Visible = false;
                        ChkCollectAgent.Checked = false;
                    }

                }

                if (ddTrade.SelectedValue.ToString() == "1")
                {
                    ImpOpt.Visible = true;

                    if (_dtExt.Rows[0]["FreehandOrNomination"].ToString() == "True")
                        optAN.Checked = true;
                    else
                        optFH.Checked = true;

                    if (optAN.Checked == true)
                    {
                        ACHolder.Visible = true;
                        ddACHolder.SelectedValue = _dtExt.Rows[0]["ACHolderID"].ToString();
                    }
                    else
                    {
                        ACHolder.Visible = false;
                        ddACHolder.SelectedValue = "-1";
                    }

                    FreeDays.Visible = true;
                    txtFreeDays.Text = _dtExt.Rows[0]["Freedays"].ToString();

                    Carrier_Freedays.Visible = true;
                    txtFreeDays_Carrier.Text = _dtExt.Rows[0]["Freedays_Carrier"].ToString();
                }

                if (ddTrade.SelectedValue.ToString() == "2")
                {
                    ExpOpt.Visible = true;

                    if (_dtExt.Rows[0]["SelfGenaratedOrNomination"].ToString() == "True")
                        optSG.Checked = true;
                    else
                        optBN.Checked = true;

                    if (optSG.Checked == true)
                    {
                        ACHolder.Visible = true;
                        ddACHolder.SelectedValue = _dtExt.Rows[0]["ACHolderID"].ToString();
                    }
                    else
                    {
                        ACHolder.Visible = false;
                        ddACHolder.SelectedValue = "-1";
                    }

                    FreeDays.Visible = false;
                    Carrier_Freedays.Visible = false;
                }


                if (ddTrade.SelectedValue.ToString() == "3")
                {
                    ACHolder.Visible = true;
                    ddACHolder.SelectedValue = _dtExt.Rows[0]["ACHolderID"].ToString();

                    FreeDays.Visible = false;
                    Carrier_Freedays.Visible = false;
                }


                txtShippingLine.Text = _dtExt.Rows[0]["ShpLnName"].ToString();
                hddShpLine.Value = _dtExt.Rows[0]["ShpLnID"].ToString();

                ddTermsOfShipment_Carrier.SelectedValue = _dtExt.Rows[0]["TermsOfShipmentID_Carrier"].ToString();

                if (_dtExt.Rows[0]["BusinessTypeID"].ToString() != "2")
                {
                    if (_dtExt.Rows[0]["BusinessTypeID"].ToString() != "5")
                    {
                        Carrier_PLD.Visible = true;
                        txtPLD_Carrier.Text = _dtExt.Rows[0]["PLDName_Carrier"].ToString(); ;
                        hddPLD_Carrier.Value = _dtExt.Rows[0]["PLDID_Carrier"].ToString(); ;

                        Carrier_OverseasAgent.Visible = true;
                        txtOverseas_Carrier.Text = _dtExt.Rows[0]["OverseasAgentLoc"].ToString(); ;
                        hddOverseas_Carrier.Value = _dtExt.Rows[0]["OverseasAgentID"].ToString(); ;
                    }
                    else
                    {
                        Carrier_PLD.Visible = false;
                        Carrier_OverseasAgent.Visible = false;
                    }
                }
                else
                {
                    Carrier_PLD.Visible = false;
                    Carrier_OverseasAgent.Visible = false;
                }

                if (_dtExt.Rows[0]["ExpireAfterSingleUse"].ToString() == "True")
                {
                    chkSingleUse.Checked = true;
                    divValidity.Visible = false;
                }
                else
                {
                    chkSingleUse.Checked = false;
                    divValidity.Visible = true;
                    if (Request.QueryString["Resue"] == "1")
                    {
                        txtValidityDt.Text = System.DateTime.Today.Date.ToString();
                    }
                    else
                    {
                        txtValidityDt.Text = _dtExt.Rows[0]["DtValidity"].ToString();
                    }
                }

                ddHandleLocation.SelectedValue = _dtExt.Rows[0]["HandleLocationID"].ToString();
                txtRemarks.Text = _dtExt.Rows[0]["Remarks"].ToString();

                txtCommodity.Text = _dtExt.Rows[0]["CommodityName"].ToString();
                hddCommodity.Value = _dtExt.Rows[0]["CommodityID"].ToString();

                ddRR.SelectedValue = _dtExt.Rows[0]["IsRR"].ToString();

                if (_dtExt.Rows[0]["IsRR"].ToString() == "True")
                {
                    ddRR.SelectedValue = "1";
                    FSet_RR.Visible = true;
                }
                else
                {
                    ddRR.SelectedValue = "0";
                    FSet_RR.Visible = false;
                }


                //if (_dtExt.Rows[0]["BusinessTypeID"].ToString() != "3")
                //    btnPDF.Visible = true;
                //else
                //    btnPDF.Visible = false;

                //end   RS.Status,RS.DtApproval,RS.Comments,RS.ApprovedBy

                lbl_ARBy.Text = _dtExt.Rows[0]["ApprovedBy"].ToString();
                lbl_Status.Text = _dtExt.Rows[0]["Status"].ToString();
                lbl_Comments.Text = _dtExt.Rows[0]["Comments"].ToString();
                lbl_ARDate.Text = _dtExt.Rows[0]["DtApproval"].ToString();



                if (_dtExt.Rows[0]["Status"].ToString() == "3".Trim())
                {
                    lbl_Status.Text = "Approved";
                }
                else
                {
                    lbl_Status.Text = "Waiting for Approval/Rejected";
                }

                if (_dtExt.Rows[0]["Status"].ToString() == "3".Trim())
                {
                    btnApproval.Visible = false;
                    btnSave.Visible = false;

                }

                if (Request.QueryString["Resue"] == "1")
                {
                    btnApproval.Visible = true;
                    btnSave.Visible = true;
                }

            }
            #endregion
            BindGridValue(ReID);
          
        }


        public void BindGridValue(string ReID)
        {
            DataTable _dtExtGridCntr = BindExistingCntr(ReID);
            if (_dtExtGridCntr.Rows.Count > 0)
            {
                GView_CntrTypes.DataSource = _dtExtGridCntr;
                GView_CntrTypes.DataBind();
                ImageButton btInsert3 = (ImageButton)GView_CntrTypes.Rows[GView_CntrTypes.Rows.Count - 1].FindControl("btnCTInsert");
                btInsert3.Visible = true;
                ViewState["Vue3"] = (DataTable)GView_CntrTypes.DataSource;
            }
            DataTable _dtExtGridSelling = BindExistingSelling(ReID);
            if (_dtExtGridSelling.Rows.Count > 0)
            {
                GView_SellingRate.DataSource = _dtExtGridSelling;
                GView_SellingRate.DataBind();
                ImageButton btInsert1 = (ImageButton)GView_SellingRate.Rows[GView_SellingRate.Rows.Count - 1].FindControl("btnSRInsert");
                btInsert1.Visible = true;
                ViewState["Vue1"] = (DataTable)GView_SellingRate.DataSource;
            }

            DataTable _dtExtGridBuying = BindExistingBuying(ReID);
            if (_dtExtGridBuying.Rows.Count > 0)
            {
                GView_BuyingRate.DataSource = _dtExtGridBuying;
                GView_BuyingRate.DataBind();
                ImageButton btInsert = (ImageButton)GView_BuyingRate.Rows[GView_BuyingRate.Rows.Count - 1].FindControl("btnBRInsert");
                btInsert.Visible = true;
                ViewState["Vue0"] = (DataTable)GView_BuyingRate.DataSource;
            }

            DataTable _dtExtGridRebat = BindExistingRebat(ReID);
            if (_dtExtGridRebat.Rows.Count > 0)
            {
                GView_RR.DataSource = _dtExtGridRebat;
                GView_RR.DataBind();
                ImageButton btInsert2 = (ImageButton)GView_RR.Rows[GView_RR.Rows.Count - 1].FindControl("btnRRInsert");
                btInsert2.Visible = true;
                ViewState["Vue2"] = (DataTable)GView_RR.DataSource;
            }
        }
        protected void btnSave_Click(object sender, EventArgs e)
        {
            try
            {

                #region Validation for container type mismatch
                var StrCntrValue = ""; bool ReturnValue = true;
                for (int j = 0; j < GView_CntrTypes.Rows.Count; j++)
                {
                    DropDownList ddlCntVal0 = (DropDownList)GView_CntrTypes.Rows[j].FindControl("ddlVal0");
                    StrCntrValue += ddlCntVal0.SelectedValue + ",";
                }
                var VCntrV = StrCntrValue.Split(',');



                if (ddBusinessTypes.SelectedItem.Text.ToString() != "CONSOLE") //Console check
                {
                    for (int x = 0; x < GView_BuyingRate.Rows.Count; x++)
                    {
                        DropDownList ddCntrcode = (DropDownList)GView_BuyingRate.Rows[x].FindControl("ddlVal2");
                        for (int y = 0; y < VCntrV.Length; y++)
                        {
                            if (VCntrV[y] != "")
                            {
                                if (ddCntrcode.SelectedValue.ToString() == VCntrV[y].ToString())
                                {
                                    ReturnValue = true;
                                    break;
                                }
                                else
                                    ReturnValue = false;
                            }
                        }
                        if (ReturnValue == false)
                        {
                            ShowMessage("Buy rate container type is mismatch", MessageType.Info);
                            return;
                        }
                    }
                }


                for (int x = 0; x < GView_SellingRate.Rows.Count; x++)
                {
                    DropDownList ddCntrcode = (DropDownList)GView_SellingRate.Rows[x].FindControl("ddlVal2");
                    for (int y = 0; y < VCntrV.Length; y++)
                    {
                        if (VCntrV[y] != "")
                        {
                            if (ddCntrcode.SelectedValue.ToString() == VCntrV[y].ToString())
                            {
                                ReturnValue = true;
                                break;
                            }
                            else
                                ReturnValue = false;
                        }
                    }
                    if (ReturnValue == false)
                    {
                        ShowMessage("Selling rate container type is mismatch", MessageType.Info);
                        return;
                    }
                }


                for (int x = 0; x < GView_RR.Rows.Count; x++)
                {
                    DropDownList ddCntrcode = (DropDownList)GView_RR.Rows[x].FindControl("ddlVal2");
                    if (ddCntrcode.SelectedValue == "")
                    {
                        for (int y = 0; y < VCntrV.Length; y++)
                        {
                            if (VCntrV[y] != "")
                            {
                                if (ddCntrcode.SelectedValue.ToString() == VCntrV[y].ToString())
                                {
                                    ReturnValue = true;
                                    break;
                                }
                                else
                                    ReturnValue = false;
                            }
                        }
                        if (ReturnValue == false)
                        {
                            ShowMessage("Rebate rate container type is mismatch", MessageType.Info);
                            return;
                        }
                    }
                }

                #endregion



                #region RatesheetValue

                string PortCode = ""; string AutoGenId = "";
                if (HdRateID.Value == "")
                {
                    if (Request.QueryString["Id"] != null)
                        Data.ID = Int32.Parse(Request.QueryString["Id"].ToString());
                    else
                    {
                        string Datecode = (DateTime.Now.Year % 100).ToString();
                        string month = (System.DateTime.Now.Month).ToString("00");
                        DataTable _dtportCode = BindPortCodeValue(Session["GeoLocID"].ToString());
                        // DataTable _dtportCode = BindPortCodeValue("2");
                        if (_dtportCode.Rows.Count > 0)
                            PortCode = _dtportCode.Rows[0]["LocCode"].ToString();
                        string LocCode = PortCode + "Q";
                        AutoGenId = SeqNoGen.GetMaxseqNumber("Ratesheet", Session["GeoLocID"].ToString());
                        // AutoGenId = SeqNoGen.GetMaxseqNumber("ExpBkg", "2");
                        if (AutoGenId == "")
                            Data.RateSheetNo = LocCode + Datecode + month + (Int32.Parse("1")).ToString("0000");
                        else
                            Data.RateSheetNo = LocCode + Datecode + month + (Int32.Parse(AutoGenId) + 1).ToString("0000");

                        lblRateSheetNo.Text = Data.RateSheetNo;
                        SeqNoGen.InsertSeqNuumber("Ratesheet", Session["GeoLocID"].ToString());
                    }
                }
                else
                {
                    Data.ID = Int32.Parse(HdRateID.Value);
                    Data.RateSheetNo = lblRateSheetNo.Text;
                }

                Data.PLRID = Int32.Parse(hddPLR.Value);
                Data.DtRateSheet = System.Convert.ToDateTime(txtRSDate.Text.Trim());
                Data.DtModified = System.Convert.ToDateTime(txtRSDate.Text.Trim());
                Data.Trade = Int32.Parse(ddTrade.SelectedValue);
                Data.BkgPartyID = Int32.Parse(hddBkgParty.Value);

                Data.POLID = Int32.Parse(hddPOL.Value);
                Data.PODID = Int32.Parse(hddPOD.Value);
                Data.PLDID = Int32.Parse(hddPLD.Value);
                Data.BusinessTypeID = Int32.Parse(ddBusinessTypes.SelectedValue);
                Data.ServiceTypeID = Int32.Parse(ddServiceTypes.SelectedValue);
                Data.ServiceModeID = Int32.Parse(ddServiceMode.SelectedValue);
                Data.TermsOfShipment = Int32.Parse(ddTermsOfShipment.SelectedValue);

            //if (ddTermsOfShipment.SelectedValue.ToString() == "2")
            //{
            //    Data.FreightPayableAt = int.Parse(hddPayAt.Value);
            //    if (ChkCollectAgent.Checked == true)
            //    {
            //        Data.IsCollectingAgentInvolved = 1;
            //        Data.CollectingAgentID = int.Parse(hddOverseas.Value);
            //    }
            //    else
            //    {
            //        Data.IsCollectingAgentInvolved = 0;
            //        Data.CollectingAgentID = 0;
            //    }
            //}
            //else
            //{
            //    Data.FreightPayableAt = 0;
            //    Data.IsCollectingAgentInvolved = 0;
            //    Data.CollectingAgentID = 0;
            //}

            //if (ddTrade.SelectedValue.ToString() == "1")
            //{
            //    if (optAN.Checked == true)
            //    {
            //        Data.FreehandOrNomination = 1;
            //        Data.ACHolderID = int.Parse(ddACHolder.SelectedValue);
            //    }
            //    else
            //    {
            //        Data.FreehandOrNomination = 0;
            //        Data.ACHolderID = 0;
            //    }
            //}
            //else
            //    Data.FreehandOrNomination = 0;

            //if (ddTrade.SelectedValue.ToString() == "2")
            //{
            //    if (optSG.Checked == true)
            //    {
            //        Data.SelfGenaratedOrNomination = 1;
            //        if (ddACHolder.SelectedValue != "")
            //            Data.ACHolderID = int.Parse(ddACHolder.SelectedValue);
            //        else
            //            Data.ACHolderID = 0;
            //    }
            //    else
            //    {
            //        Data.SelfGenaratedOrNomination = 0;
            //        Data.ACHolderID = 0;
            //    }
            //}
            //else
            //    Data.SelfGenaratedOrNomination = 0;

            //if (ddTrade.SelectedValue.ToString() == "3")
            //{
            //    Data.FreehandOrNomination = 0;
            //    Data.SelfGenaratedOrNomination = 0;
            //    Data.ACHolderID = int.Parse(ddACHolder.SelectedValue);

            //}

            //if (ddTrade.SelectedValue.ToString() == "1")
            //{
            //    Data.Freedays = int.Parse(txtFreeDays.Text.Trim());
            //    Data.Freedays_Carrier = int.Parse(txtFreeDays_Carrier.Text.Trim());
            //}
            //else
            //{
            //    Data.Freedays = 0;
            //    Data.Freedays_Carrier = 0;
            //}

            //Data.ShippingLine = int.Parse(hddShpLine.Value);
            //Data.TermsOfShipmentID_Carrier = int.Parse(ddTermsOfShipment_Carrier.SelectedValue);

            //if (ddBusinessTypes.SelectedValue.ToString() == "2" || ddBusinessTypes.SelectedValue.ToString() == "5")
            //{
            //    Data.PlaceofDeliveryID_Carrier = 0;
            //    Data.OverseasAgentID = 0;
            //}
            //else
            //{
            //    Data.PlaceofDeliveryID_Carrier = int.Parse(hddPLD_Carrier.Value);
            //    Data.OverseasAgentID = int.Parse(hddOverseas_Carrier.Value);
            //}

            //if (chkSingleUse.Checked == false)
            //{
            //    Data.ExpireAfterSingleUse = 0;
            //    Data.DtValidity = System.Convert.ToDateTime(txtValidityDt.Text.Trim());
            //}
            //else
            //{
            //    Data.ExpireAfterSingleUse = 1;

            //}
            if (ddTermsOfShipment.SelectedValue.ToString() == "2")
            {
                Data.FreightPayableAt = int.Parse(hddPayAt.Value);

                if (ChkCollectAgent.Checked == true)
                {
                    Data.IsCollectingAgentInvolved = 1;
                    Data.CollectingAgentID = int.Parse(hddOverseas.Value);
                }
                else
                {
                    Data.IsCollectingAgentInvolved = 0;
                    Data.CollectingAgentID = 0;
                }
            }
            else
            {
                Data.FreightPayableAt = 0;
                Data.IsCollectingAgentInvolved = 0;
                Data.CollectingAgentID = 0;
            }

            //--------END--------

            if (ddTrade.SelectedValue.ToString() == "1")
            {
                if (optAN.Checked == true)
                {
                    Data.FreehandOrNomination = 1;
                    Data.ACHolderID = int.Parse(ddACHolder.SelectedValue);
                }
                else
                {
                    Data.FreehandOrNomination = 0;
                    Data.ACHolderID = 0;
                }
            }
            else
                Data.FreehandOrNomination = 0;


            if (ddTrade.SelectedValue.ToString() == "2")
            {
                if (optSG.Checked == true)
                {
                    Data.SelfGenaratedOrNomination = 1;
                    Data.ACHolderID = int.Parse(ddACHolder.SelectedValue);
                }
                else
                {
                    Data.SelfGenaratedOrNomination = 0;
                    Data.ACHolderID = 0;
                }
            }
            else
                Data.SelfGenaratedOrNomination = 0;

            if (ddTrade.SelectedValue.ToString() == "3")
            {
                Data.FreehandOrNomination = 0;
                Data.SelfGenaratedOrNomination = 0;
                Data.ACHolderID = int.Parse(ddACHolder.SelectedValue);

            }
            //-------BEGIN---------

            if (ddTrade.SelectedValue.ToString() == "1")
            {
                Data.Freedays = int.Parse(txtFreeDays.Text.Trim());
                Data.Freedays_Carrier = int.Parse(txtFreeDays_Carrier.Text.Trim());
            }
            else
            {
                Data.Freedays = 0;
                Data.Freedays_Carrier = 0;
            }

            Data.ShippingLine = int.Parse(hddShpLine.Value);
            Data.TermsOfShipmentID_Carrier = int.Parse(ddTermsOfShipment_Carrier.SelectedValue);

            if (ddBusinessTypes.SelectedValue.ToString() == "2" || ddBusinessTypes.SelectedValue.ToString() == "5")
            {
                Data.PlaceofDeliveryID_Carrier = 0;
                Data.OverseasAgentID = 0;
            }
            else
            {
                Data.PlaceofDeliveryID_Carrier = int.Parse(hddPLD_Carrier.Value);
                Data.OverseasAgentID = int.Parse(hddOverseas_Carrier.Value);
            }

            if (chkSingleUse.Checked == false)
            {
                Data.ExpireAfterSingleUse = 0;
                Data.DtValidity = System.Convert.ToDateTime(txtValidityDt.Text.Trim());
            }
            else
            {
                Data.ExpireAfterSingleUse = 1;
                //Obj_MyRateSheet.DtValidity = "";
            }

            Data.HandleLocationID = int.Parse(ddHandleLocation.SelectedValue);
                Data.GeoLocID = int.Parse(Session["GeoLocID"].ToString());
                // Data.GeoLocID = int.Parse("");
                Data.IsRR = int.Parse(ddRR.SelectedValue);
                Data.UserID = int.Parse(Session["UserID"].ToString());
                //Data.UserID = int.Parse("1");

                if (txtRemarks.Text != "")
                    Data.Remarks = txtRemarks.Text.ToString();
                else
                    Data.Remarks = "";

                //if (ddBusinessTypes.SelectedValue.ToString() == "3")
                //    Data.Status = 3;
                //else
                //    Data.Status = 1;
                Data.Status = 1;

                Data.CommodityID = int.Parse(hddCommodity.Value);
                #endregion

                #region CntrAdding

                DataTable _dtCntr = new DataTable();
                _dtCntr.Columns.Add("CntrTypeID");
                _dtCntr.Columns.Add("ApproxUnits");

                foreach (GridViewRow gRow in GView_CntrTypes.Rows)
                {
                    DropDownList ddlCntVal0 = (DropDownList)gRow.FindControl("ddlVal0");
                    TextBox txtCNVal2 = (TextBox)gRow.FindControl("txtCNVal2");
                    if (ddlCntVal0.SelectedValue != "" && txtCNVal2.Text != "0".ToString())
                    {

                        _dtCntr.Rows.Add(_dtCntr.NewRow());
                        _dtCntr.Rows[_dtCntr.Rows.Count - 1]["CntrTypeID"] = ddlCntVal0.SelectedValue;
                        _dtCntr.Rows[_dtCntr.Rows.Count - 1]["ApproxUnits"] = txtCNVal2.Text;

                    }
                }
                #endregion

                #region RateBuying

                DataTable _dtBuying = new DataTable();

                _dtBuying.Columns.Add("BRChargesID");
                _dtBuying.Columns.Add("BRRateRcvdFrom");
                _dtBuying.Columns.Add("BRCntrTypeID");
                _dtBuying.Columns.Add("BRBasis");
                _dtBuying.Columns.Add("BRCurrID");
                _dtBuying.Columns.Add("BRAmount");
                _dtBuying.Columns.Add("BuyRateFromRateSheet");

                foreach (GridViewRow gRow in GView_BuyingRate.Rows)
                {
                    TextBox txtVal5 = (TextBox)gRow.FindControl("txtVal5");
                    if (txtVal5.Text != "" && txtVal5.Text != "0" && txtVal5.Text != "0.00")
                    {
                        HiddenField HdCnargeCode = (HiddenField)gRow.FindControl("HdVal0");
                        DropDownList ddpaycode = (DropDownList)gRow.FindControl("ddlVal1");
                        DropDownList ddCntrcode = (DropDownList)gRow.FindControl("ddlVal2");
                        DropDownList ddBasic = (DropDownList)gRow.FindControl("ddlVal3");
                        DropDownList ddCurr = (DropDownList)gRow.FindControl("ddlVal4");

                        _dtBuying.Rows.Add(_dtBuying.NewRow());

                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BRChargesID"] = HdCnargeCode.Value;
                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BRRateRcvdFrom"] = ddpaycode.SelectedValue;
                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BRCntrTypeID"] = ddCntrcode.SelectedValue;
                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BRBasis"] = ddBasic.SelectedValue;
                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BRCurrID"] = ddCurr.SelectedValue;
                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BRAmount"] = txtVal5.Text;
                        _dtBuying.Rows[_dtBuying.Rows.Count - 1]["BuyRateFromRateSheet"] = 1;

                    }

                }
                #endregion

                #region RateSelling

                DataTable _dtSelling = new DataTable();

                _dtSelling.Columns.Add("SRChargesID");
                _dtSelling.Columns.Add("SRRateQuotedTo");
                _dtSelling.Columns.Add("SRCntrTypeID");
                _dtSelling.Columns.Add("SRBasis");
                _dtSelling.Columns.Add("SRCurrID");
                _dtSelling.Columns.Add("SRAmount");
                _dtSelling.Columns.Add("SellRateFromRateSheet");

                foreach (GridViewRow gRow in GView_SellingRate.Rows)
                {
                    TextBox txtVal5 = (TextBox)gRow.FindControl("txtSRVal5");
                    if (txtVal5.Text != "")
                    {
                        HiddenField HdCnargeCode = (HiddenField)gRow.FindControl("HdSRVal0");

                        DropDownList ddpaycode = (DropDownList)gRow.FindControl("ddlVal1");
                        DropDownList ddCntrcode = (DropDownList)gRow.FindControl("ddlVal2");
                        DropDownList ddBasic = (DropDownList)gRow.FindControl("ddlVal3");
                        DropDownList ddCurr = (DropDownList)gRow.FindControl("ddlVal4");



                        _dtSelling.Rows.Add(_dtSelling.NewRow());

                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SRChargesID"] = HdCnargeCode.Value;
                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SRRateQuotedTo"] = ddpaycode.SelectedValue;
                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SRCntrTypeID"] = ddCntrcode.SelectedItem.Value;
                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SRBasis"] = ddBasic.SelectedValue;
                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SRCurrID"] = ddCurr.SelectedValue;
                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SRAmount"] = txtVal5.Text;
                        _dtSelling.Rows[_dtSelling.Rows.Count - 1]["SellRateFromRateSheet"] = 1;
                        //error

                    }

                }

                #endregion

                #region Rebate

                DataTable _dtRebate = new DataTable();

                _dtRebate.Columns.Add("TransTypeID");
                _dtRebate.Columns.Add("RateRFromQTo");
                _dtRebate.Columns.Add("CntrTypeID");
                _dtRebate.Columns.Add("Basis");
                _dtRebate.Columns.Add("CurrID");
                _dtRebate.Columns.Add("Amount");

                foreach (GridViewRow gRow in GView_RR.Rows)
                {
                    TextBox txtVal5 = (TextBox)gRow.FindControl("txtRRVal5");
                    if (txtVal5.Text != "")
                    {
                        DropDownList HdTrans = (DropDownList)gRow.FindControl("ddlVal0");
                        DropDownList HdfromTo = (DropDownList)gRow.FindControl("ddlVal1");
                        DropDownList Hdcntr = (DropDownList)gRow.FindControl("ddlVal2");
                        DropDownList HdBasic = (DropDownList)gRow.FindControl("ddlVal3");
                        DropDownList HdCurr = (DropDownList)gRow.FindControl("ddlVal4");

                        _dtRebate.Rows.Add(_dtRebate.NewRow());

                        _dtRebate.Rows[_dtRebate.Rows.Count - 1]["TransTypeID"] = HdTrans.SelectedValue;
                        _dtRebate.Rows[_dtRebate.Rows.Count - 1]["RateRFromQTo"] = HdfromTo.SelectedValue;
                        _dtRebate.Rows[_dtRebate.Rows.Count - 1]["CntrTypeID"] = Hdcntr.SelectedValue;
                        _dtRebate.Rows[_dtRebate.Rows.Count - 1]["Basis"] = HdBasic.SelectedValue;
                        _dtRebate.Rows[_dtRebate.Rows.Count - 1]["CurrID"] = HdCurr.SelectedValue;
                        _dtRebate.Rows[_dtRebate.Rows.Count - 1]["Amount"] = txtVal5.Text;

                    }

                }

                #endregion

                if (RateManag.InsertQutationRates(Data, _dtCntr, _dtBuying, _dtSelling, _dtRebate) == -1)
                {
                    HdRateID.Value = Data.ID.ToString();

                    foreach (GridViewRow gRow in GView_BuyingRate.Rows)
                    {
                        TextBox txtVal5 = (TextBox)gRow.FindControl("txtVal5");
                        if (txtVal5.Text == "0" || txtVal5.Text == "0.00")
                        {
                            Label lblField14 = (Label)gRow.FindControl("lblField14");
                            BuyingDelete(lblField14.Text);
                        }
                    }

                    foreach (GridViewRow gRow in GView_SellingRate.Rows)
                    {
                        TextBox txtVal5 = (TextBox)gRow.FindControl("txtSRVal5");
                        if (txtVal5.Text == "0" || txtVal5.Text == "0.00")
                        {
                            Label lblField14 = (Label)gRow.FindControl("lblField14");
                            SellingDelete(lblField14.Text);
                        }
                    }
                    foreach (GridViewRow gRow in GView_RR.Rows)
                    {
                        TextBox txtVal5 = (TextBox)gRow.FindControl("txtRRVal5");
                        if (txtVal5.Text == "0" || txtVal5.Text == "0.00")
                        {
                            Label lblField14 = (Label)gRow.FindControl("lblField14");
                            RebateDelete(lblField14.Text);
                        }
                    }

                    foreach (GridViewRow gRow in GView_CntrTypes.Rows)
                    {
                        TextBox txtCNVal2 = (TextBox)gRow.FindControl("txtCNVal2");
                        if (txtCNVal2.Text == "0")
                        {
                            Label lblID = (Label)gRow.FindControl("lblID");
                            CntryTable(lblID.Text);
                        }
                    }

                    BindGridValue(HdRateID.Value);

                    ShowMessage("Record saved  successfully!", MessageType.Success);
                    return;
                }
            }
            catch(Exception e1)
            {
                ShowMessage(e1.Message.ToString() + "system could not accept the data", MessageType.Error);
                
                return;
            }

        }
        protected void ShowMessage(string Message, MessageType type)
        {
            ScriptManager.RegisterStartupScript(this, this.GetType(), System.Guid.NewGuid().ToString(), "ShowMessage('" + Message + "','" + type + "'); ", true);
        }
        //protected void btnContinues_Click(object sender, EventArgs e)
      
        // //   multiTabs.ActiveViewIndex = Int32.Parse("1");
        //}


        protected void btnPDF_Click(object sender, EventArgs e)
        {
            string Rec_Count;
            Rec_Count = RSM.IsExist(HdRateID.Value);
            if (Rec_Count.ToString() == "1")
                btnApproval.Visible = false;
            else
                btnApproval.Visible = true;

            string str = "";
            str = "SICLQuotationPDF.aspx?RId=" + HdRateID.Value;
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "_POPUP_Window", "<script>window.open('" + str + "','new','left=10,top=10,width=1200,height=600,scrollbars=yes')</script>", false);

        }

    protected void btnApproval_Click(object sender, EventArgs e)
    {

        try
        {
            RSM.Update_Ratesheet_Status(Data, HdRateID.Value, "2");
            string sHtml = string.Empty;
            sHtml = "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
            sHtml += "</table>";
            sHtml += "</br>";

            sHtml += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";

            sHtml += "<tr style='font-family:Verdana; font-size:12px; font-weight:bold; text-align:center; background-color:#333366; color=#FFFFFF'>";
            sHtml += "<td>QUOTATION APPROVAL #</td>";
            sHtml += "<td>REQUESTED BY</td>";
            sHtml += "<td>REQUESTED DATE & TIME</td>";
            sHtml += "</tr>";

            sHtml += "<tr style='font-family:Verdana; font-size:12px; text-align:center; color=#0033FF'>";
            sHtml += "<td style=' border-right:1px solid #336699 ; border-left:1px solid #336699; border-top: 0px solid #336699; border-bottom: 1px solid #336699;'>" + lblRateSheetNo.Text + "</td>";
            sHtml += "<td style=' border-right:1px solid #336699 ; border-left:1px solid #336699; border-top: 0px solid #336699; border-bottom: 1px solid #336699;'>" + "LOGINNAME" + "</td>";
            sHtml += "<td style=' border-right:1px solid #336699 ; border-left:1px solid #336699; border-top: 0px solid #336699; border-bottom: 1px solid #336699;'>" + System.DateTime.Now + "</td>";
            sHtml += "</tr>";
            sHtml += "</table>";
            sHtml += "</br>";
            string _Remarks = "";
            _Remarks = RSM.Get_Remarks(HdRateID.Value);
            if (_Remarks.ToString() != "")
            {
                sHtml += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr>";
                sHtml += "<td style='font-family:Arial; font-size:12px; font-style:italic' text-align:left; color:#0033FF'><B>Note : </B>" + _Remarks.ToString() + "</td>";
                sHtml += "</tr>";
                sHtml += "</table>";
                sHtml += "</br>"; 
            }

            DataTable _dtsv = Get_GeoLoction(Session["GeoLocID"].ToString());
            if (_dtsv.Rows.Count > 0)
            {

                sHtml = "<table border='0' cellpadding='0' cellspacing='0'>";
                sHtml += "<tr><td style='font-family:Verdana; font-size:11px; color=#336699'>NOTE :  THIS IS AN AUTO-GENERATED MAIL . PLEASE DO NOT REPLY  </td></tr>";
                sHtml += "<tr><td style='font-family:Verdana; font-size:12px; font-weight:bold; color=#336699'> </td></tr>";
                sHtml += "<tr><td style='font-family:Verdana; font-size:11px; color=#336699'>" + _dtsv.Rows[0]["Address"].ToString() + "," + "</td></tr>";
                sHtml += "<tr><td style='font-family:Verdana; font-size:11px; color=#336699'>" + _dtsv.Rows[0]["Location"].ToString() + " - " + _dtsv.Rows[0]["Pincode"].ToString() + "." + "</td></tr>";
                sHtml += "<tr><td style='font-family:Verdana; font-size:11px; color=#336699'> Tel # " + _dtsv.Rows[0]["Areacode"].ToString() + "  " + _dtsv.Rows[0]["Phone1"].ToString() + "   Fax # " + _dtsv.Rows[0]["Areacode"].ToString() + "  " + _dtsv.Rows[0]["Fax"].ToString() + "</td></tr>";
            }
            string l_mailSubject = "";

            MailMessage EmailObject = new MailMessage();
            string stremail = Session["EmailID"].ToString().Trim();

            EmailObject.From = new MailAddress("finance@maxxpress.com", "QUOTATION APPROVAL");
            EmailObject.CC.Add(new MailAddress("sunil_paul@maxxpress.com"));
            EmailObject.CC.Add(new MailAddress("sanju_b@maxxpress.com"));
            EmailObject.To.Add(new MailAddress("bose_s@maxxpress.com"));
            EmailObject.Subject = "QUOTATION APPROVAL REQUEST" + lblRateSheetNo.Text;
            string RatesheetValue = "Quotation" + lblRateSheetNo.Text;

            EmailObject.Body = sHtml;
            EmailObject.IsBodyHtml = true;
            EmailObject.Priority = MailPriority.Normal;
            EmailObject.Attachments.Add(new Attachment(Server.MapPath("~/PDFForms\\" + RatesheetValue + ".pdf")));
            SmtpClient SMTPServer = new SmtpClient();
            SMTPServer.UseDefaultCredentials = false;
            SMTPServer.Credentials = new NetworkCredential("finance@maxxpress.com", "(HOmIvr3");
            SMTPServer.Host = "smtp.maxxpress.com";
            SMTPServer.ServicePoint.MaxIdleTime = 1;
            SMTPServer.Port = 587;
            SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
            SMTPServer.Send(EmailObject);

            string Rec_Count;
            Rec_Count = RSM.IsExist(HdRateID.Value);

            if (Rec_Count.ToString() == "1")
                btnSave.Enabled = false;
            else
                btnApproval.Enabled = true;
            ShowMessage("Quotation Sent for Approval successfully!", MessageType.Info);
            return;
        }
        catch (Exception ex)
        {
            ShowMessage(ex.Message + "Please Contact Your Admin", MessageType.Info);
        }
    } 

    

        public DataTable Get_GeoLoction(string GeoID)
        {
            string _Query = "select * from F_TblGeoLocations where Id =" + GeoID;
            return SQLFun.GetSQLFunction(_Query);
        }

        public DataTable BuyingDelete(string ID)
        {
            string _Query = "delete from F_BuyingRateDtls where BUYID =" + ID;
            return SQLFun.GetSQLFunction(_Query);
        }

        public DataTable SellingDelete(string ID)
        {
            string _Query = "delete from F_SellingRateDtls where ID = " + ID;
            return SQLFun.GetSQLFunction(_Query);
        }

        public DataTable RebateDelete(string ID)
        {
            string _Query = "delete from F_RateSheetRebateRefund where RebateID =" + ID;
            return SQLFun.GetSQLFunction(_Query);
        }

        public DataTable CntryTable(string ID)
        {
            string _Query = "delete from F_RatesheetCntrTypes where ID =" + ID;
            return SQLFun.GetSQLFunction(_Query);
        }
    }
