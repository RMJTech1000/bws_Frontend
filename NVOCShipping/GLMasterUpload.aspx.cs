using System;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
//using DataBaseFactory;
using DataManager;
using DataTier;

namespace NVOCShipping
{
    public partial class GLMasterUpload : System.Web.UI.Page
    {
        static string FolderPath = HttpRuntime.AppDomainAppPath + "//UploadFolder//FinanceUpload//";
        FinanceUploadManager Manag = new FinanceUploadManager();
        MyMRG Datav = new MyMRG();
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btnfileUploading_Click(object sender, EventArgs e)
        {
            try
            {
                Random rm = new Random();
                int IDNo = rm.Next(1, 100);

                if (ExcelFileUploading.FileName == "")
                {
                    string message1 = "GL Master file is missing !!!";
                    ScriptManager.RegisterStartupScript(this, GetType(), "Popup", "ShowPopup('" + message1 + "');", true);
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
                    Import_To_Grid(FolderPath + FileName, Extension);

                    string message = "Record Saved successfully";
                    lblError.Text = message;
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


        private void Import_To_Grid(string FilePath, string Extension)
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
            DataTable dtv = dt.Copy();
            connExcel.Close();
            string Str = "";
            dtv.Columns.Add("Result");
            dtv.Columns.Add("Status");
            lblError.Text = "Process";
            DataTable dtx = Manag.InsertGLmasterExcelUploading(dtv);
            lblError.Text = "Sucesss";
            StringBuilder sb = new StringBuilder();
            string ContentLeft = "<td align='Left' style= ' background-color: #FFFFFF; black; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right: Black thin inset; inset; color: #000000;  font-family: Arial, Helvetica, sans-serif; font-style: normal; font-size:11px;font-weight: normal;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
            string LeftAll = "<td align='center' style='background-color:#04208c; color:#ffffff; font-family:Arial; border-top: Black thin inset; border-bottom: Black thin inset; border-Left: Black thin inset; border-right:Black thin inset; font-style:normal; font-weight:bold;font-size:13px;'>";
            sb.Append("<table>");
            sb.Append("<tr>");
            sb.Append(LeftAll + " ID</td>");
            sb.Append(LeftAll + " CODE  </td>");
            sb.Append(LeftAll + " GLDESCRIPTION </td>");
            sb.Append(LeftAll + " COMPANY  </ td>");
            sb.Append(LeftAll + " MAINACCOUNT   </ td>");
            sb.Append(LeftAll + " NATURE   </ td>");
            sb.Append(LeftAll + " CATEGORY   </ td>");
            sb.Append(LeftAll + " GLMATCHING   </ td>");
            sb.Append(LeftAll + " GLSTATUS   </ td>");
            //sb.Append(LeftAll + " PRODUCTTYPE </ td>");
            sb.Append(LeftAll + " RESULT  </td>");
            sb.Append(LeftAll + " STATUS</td>");
            sb.Append("</tr>");
            lblError.Text = "Sucesss 1";
            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                lblError.Text = "Sucesss 2";
                sb.Append("<tr>");
                sb.Append(ContentLeft + dtx.Rows[i]["ID"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["CODE"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["GLDESCRIPTION"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["COMPANY"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["MAINACCOUNT"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["NATURE"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["CATEGORY"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["GLMATCHING"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["GLSTATUS"].ToString() + "</td>");
               // sb.Append(ContentLeft + dtx.Rows[i]["PRODUCTTYPE"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["RESULT"].ToString() + "</td>");
                sb.Append(ContentLeft + dtx.Rows[i]["STATUS"].ToString() + "</td>");
            }
            sb.Append("</table>");

            lblError.Text = "Sucesss Final";
            Response.Write(sb.ToString());
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=GLMasterUploadReport" + System.DateTime.Now + ".xls");
            Response.Charset = "";
            Response.ContentType = "application/vnd.xls";
            System.IO.StringWriter stringWrite = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
            htmlWrite.Write(sb.ToString());
            Response.Write(stringWrite.ToString());
            Response.End();




        }
    }
}