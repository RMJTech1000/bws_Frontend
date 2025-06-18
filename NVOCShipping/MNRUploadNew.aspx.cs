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
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Data.Common;
using DataBaseFactory;
using Infrastructure;

namespace NVOCShipping
{
    public partial class MNRUploadNew : System.Web.UI.Page
    {
        #region Membervariable
        private IDataBaseFactory _dbFactory = null;
        #endregion

        #region Constructor Method
        public MNRUploadNew()
        {
            _dbFactory = new SQLFactory();

        }
        #endregion

        static string FolderPath = HttpRuntime.AppDomainAppPath + "//UploadFolder//MNRTariffUpload//";
        MNRUploadManager Manag = new MNRUploadManager();
        MyMRG Datav = new MyMRG();
        string FileName = "";
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                //btnMoveData.Enabled = false;
            }
        }
        protected void btnfileUploading_Click(object sender, EventArgs e)
        {
            try
            {
                string UserID = Request.QueryString["UserID"].ToString();
                Random rm = new Random();
                int IDNo = rm.Next(1, 100);

                if (ExcelFileUploading.FileName == "")
                {
                    string message1 = "MNR Tariff Upload filing is missing !!!";
                    lblError.Text = message1;
                    //ScriptManager.RegisterStartupScript(this, GetType(), "Popup", "ShowPopup('" + message1 + "');", true);
                    return;
                }

                string disp_name = string.Empty;
                Random rd = new Random();
                disp_name += rd.Next(1000).ToString();

                if (ExcelFileUploading.HasFile)
                {
                    string FileName = Path.GetFileName(disp_name + "_" + ExcelFileUploading.PostedFile.FileName);
                    string Extension = Path.GetExtension(disp_name + "_" + ExcelFileUploading.PostedFile.FileName);
                    //FolderPath = HttpRuntime.AppDomainAppPath + "//BulkEmailExcel//";
                    ExcelFileUploading.SaveAs(FolderPath + FileName);
                    Import_To_Grid(FolderPath + FileName, Extension, UserID);


                    //ScriptManager.RegisterStartupScript((sender as Control), this.GetType(), "Popup", "ShowPopup('" + message + "');", true);

                }

                //else if (HDSave.Value == "false")
                //{

                //    // ExistingSaveValues(hdId.Value);

                //    string message = "Check your input file";
                //    ScriptManager.RegisterStartupScript((sender as Control), this.GetType(), "Popup", "ShowPopup('" + message + "');", true);
                //}
            }
            catch (Exception ex)
            {

                string message = Regex.Replace(ex.Message, @"[^0-9a-zA-Z]+", " ");
                lblError.Text = message.ToString();
            }
        }


        private void Import_To_Grid(string FilePath, string Extension, string UserID)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".xls": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    break;
            }
            conStr = String.Format(conStr, FilePath);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            cmdExcel.Connection = connExcel;

            //Get the name of First Sheet
            connExcel.Open();
            DataTable dtExcelSchema;
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            connExcel.Close();

            //Read Data from First Sheet
            connExcel.Open();
            cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
            oda.SelectCommand = cmdExcel;
            oda.Fill(dt);
            //DataTable dtv = dt.Copy();
            DataTable dtv = dt.Rows
                 .Cast<DataRow>()
                 .Where(row => !row.ItemArray.All(f => f is DBNull))
                 .CopyToDataTable();
            connExcel.Close();

            int ONhirecount = 0;
            int QTY = 0;
            int PickUpRefID = 0;
            int CntrCount = 0;
            DbConnection con = null;
            DbTransaction trans;

            try
            {
                con = _dbFactory.GetConnection();
                con.Open();
                trans = _dbFactory.GetTransaction(con);
                DbCommand Cmd = _dbFactory.GetCommand();
                Cmd.Connection = con;
                Cmd.Transaction = trans;


                //Cmd.CommandText = "  select COUNT(*) as OnHireCount  from NVO_ContainerOnHire INNER JOIN NVO_LeaseContract On NVO_LeaseContract.ID = NVO_ContainerOnHire.LeasePickUpRefID  where NVO_LeaseContract.ContractRefNo = '" + dtv.Rows[0]["PICKUPREF"].ToString() + "' ";
                //ONhirecount = Int32.Parse(Cmd.ExecuteScalar().ToString());

                //Cmd.CommandText = " select ID from NVO_LeaseContract where ContractRefNo = '" + dtv.Rows[0]["PICKUPREF"].ToString() + "'";
                //PickUpRefID = Int32.Parse(Cmd.ExecuteScalar().ToString());

                ////Cmd.CommandText = " select LC.QTY from NVO_LeaseContract L INNER JOIN NVO_LeaseDetails LC ON LC.LeaseContractID = L.ID where ContractRefNo = '" + dtv.Rows[0]["PICKUPREF"].ToString() + "'";

                //Cmd.CommandText = "select (Sum(LC.QTY) - ISNULL((SELECT count (CntrNo)  FROM NVO_Containers where PickUpRefID=" + PickUpRefID + "),0 ) ) as Qty from NVO_LeaseContract L INNER JOIN NVO_LeaseDetails LC ON LC.LeaseContractID = L.ID where L.ID = " + PickUpRefID + "";

                //QTY = Int32.Parse(Cmd.ExecuteScalar().ToString());
                for (int x = 0; x < dtv.Rows.Count; x++)
                {
                    //    Cmd.CommandText = "select COUNT(*) as CntrCount  from NVO_Containers where CntrNo ='" + dtv.Rows[x]["CONTAINER_NO"].ToString() + "'";
                    //    CntrCount = Int32.Parse(Cmd.ExecuteScalar().ToString());

                    //    if (dtv.Rows.Count > QTY)
                    //    {
                    //        lblError.Text = "Lease Contract Container Quantity already Exceeded For this Reference...";
                    //    }

                    //    else if (CntrCount >= 1)
                    //    {
                    //        lblError.Text = dtv.Rows[x]["CONTAINER_NO"].ToString() + "-" + " Container No Already Exists..Please modify";
                    //    }
                    //    //if (ONhirecount >= 1)
                    //    //{
                    //    //    lblError.Text = "Already Lease Ref Exists..Please modify";
                    //    //    //return msg;
                    //    //}

                    //    else
                    //    {
                    string Str = "";
                        dtv.Columns.Add("Result");
                        dtv.Columns.Add("Status");
                        lblError.Text = "Process";
                        DataTable dtx = Manag.InsertMNRTariffExcelUploading(dtv, UserID);
                        string message = "Record Saved successfully";
                        lblError.Text = message;
                        //btnMoveData.Enabled = true;
                    //}

                }

               
            }
            catch (Exception ex)
            {
                string ss = ex.ToString();

            }
        }
        protected void btnMoveData_Click(object sender, EventArgs e)
        {
           // btnMoveData.Enabled = false;
            DataTable dtx = Manag.InsertMNRTariffDataMove();

            string message = "Record Moved Successfully";
            lblMovSucess.Text = message;


        }
    }
}