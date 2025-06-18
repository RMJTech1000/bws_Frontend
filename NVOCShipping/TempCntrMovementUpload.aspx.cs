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
using System.Data.OleDb;
using System.Configuration;
using System.Data.Common;
using DataBaseFactory;
using Infrastructure;
using DataManager;

namespace NVOCShipping
{
    public partial class TempCntrMovementUpload : System.Web.UI.Page
    {
        UserManager Mng = new UserManager();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

            }

        }
        protected void btnUploadEmail_Click1(object sender, EventArgs e)
        {
            if (ExcelFileUploading.FileName == "")
            {
                string message1 = "Please Upload Excel File";
                ScriptManager.RegisterStartupScript(this, GetType(), "Popup", "ShowSave('" + message1 + "');", true);
                return;
            }
            if (ExcelFileUploading.HasFile)
            {
                string FileName = Path.GetFileName(ExcelFileUploading.PostedFile.FileName);
                string Extension = Path.GetExtension(ExcelFileUploading.PostedFile.FileName);
                string FolderPath = HttpRuntime.AppDomainAppPath + "//UploadFolder//ContainerMasterUpload//";
                ExcelFileUploading.SaveAs(FolderPath + FileName);
                Import_To_Grid(FolderPath + FileName, Extension);
            }

        }
        private IDataBaseFactory _dbFactory = new SQLFactory();
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
            InsertEmail(dt);
            string message = "Record Saved successfully";
            lblError.Text = message;

        }

        protected void btnMove_Click(object sender, EventArgs e)
        {
            MoveData();
            string message = "Record Moved Successfully";
            lblMovSucess.Text = message;
        }

        public int MoveData()
        {
            DataTable dt = new DataTable();

            int indexrow = 0;
            int result = 0;
            int IDv = 0;

            DbConnection con = null;
            DbTransaction trans;
            int rowdt = 0;
            int rowdTCN = 0;
            
            try
            {
                con = _dbFactory.GetConnection();
                con.Open();
                trans = _dbFactory.GetTransaction(con);
                DbCommand Cmd = _dbFactory.GetCommand();
                Cmd.Connection = con;
                Cmd.Transaction = trans;
                try
                {
                    string LastCntrID="",LastCntrTrnsID = "0", LastCntrportID="0", LastStatusCode="0", LastDepotID = "0", LastAgencyID = "0", LastMvDate="0", LastTransitID="0", LastVesVoyID = "0", LastCustomerID = "0";
                    dt = GetTempCntr();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        rowdTCN = i;
                        DataTable dtV = new DataTable();
                        dtV = GetCntrListDistinct(dt.Rows[i]["ID"].ToString());

                        for (int k = 0; k < dtV.Rows.Count; k++)
                        {
                            rowdt = k;
                            Cmd.CommandText = "INSERT INTO  NVO_ContainerTxns(ContainerID,LocationID,StatusCode,DtMovement,NextPortID,DepotID,AgencyID,ModeOfTransportID,VesVoyID,CustomerID,BLNumber) " +
                                          " values (@ContainerID,@LocationID,@StatusCode,@DtMovement,@NextPortID,@DepotID,@AgencyID,@ModeOfTransportID,@VesVoyID,@CustomerID,@BLNumber) ";

                                Cmd.Parameters.Add(_dbFactory.GetParameter("@ContainerID", dtV.Rows[k]["CntrID"].ToString()));
                            
                                Cmd.Parameters.Add(_dbFactory.GetParameter("@LocationID", dtV.Rows[k]["LocationID"].ToString()));
                            Cmd.Parameters.Add(_dbFactory.GetParameter("@NextPortID", dtV.Rows[k]["LocationID"].ToString()));
                            
                            Cmd.Parameters.Add(_dbFactory.GetParameter("@StatusCode", dtV.Rows[k]["StatusCode"].ToString()));
                            string timet = DateTime.Parse(dtV.Rows[k]["DtMovement"].ToString()).ToString("MM/dd/yyyy");
                            if (dtV.Rows[k]["DtMovement"].ToString() != "")
                            {
                                Cmd.Parameters.Add(_dbFactory.GetParameter("@DtMovement", timet));
                            }
                            else
                            {
                                Cmd.Parameters.Add(_dbFactory.GetParameter("@DtMovement", DBNull.Value));
                            }

                                Cmd.Parameters.Add(_dbFactory.GetParameter("@AgencyID", dtV.Rows[k]["AgencyID"].ToString()));
                            
                                Cmd.Parameters.Add(_dbFactory.GetParameter("@DepotID", dtV.Rows[k]["DepotID"].ToString()));

                            Cmd.Parameters.Add(_dbFactory.GetParameter("@ModeOfTransportID", dtV.Rows[k]["TransitID"].ToString()));

                            Cmd.Parameters.Add(_dbFactory.GetParameter("@VesVoyID", dtV.Rows[k]["VesVoyID"].ToString()));

                            Cmd.Parameters.Add(_dbFactory.GetParameter("@CustomerID", dtV.Rows[k]["CustomerID"].ToString()));

                            Cmd.Parameters.Add(_dbFactory.GetParameter("@BLNumber", dtV.Rows[k]["BkgID"].ToString()));

                            LastCntrID = dtV.Rows[k]["CntrID"].ToString();
                            LastCntrportID = dtV.Rows[k]["LocationID"].ToString();
                            LastDepotID = dtV.Rows[k]["DepotID"].ToString();
                            LastAgencyID = dtV.Rows[k]["AgencyID"].ToString();
                            LastStatusCode = dtV.Rows[k]["StatusCode"].ToString();
                            LastTransitID = dtV.Rows[k]["TransitID"].ToString();
                            LastVesVoyID = dtV.Rows[k]["VesVoyID"].ToString();
                            LastCustomerID = dtV.Rows[k]["CustomerID"].ToString();
                            //string timet1 = DateTime.Parse(dtV.Rows[k]["DtMovement"].ToString()).ToString("MM/dd/yyyy");
                            LastMvDate = timet;

                            result = Cmd.ExecuteNonQuery();
                            Cmd.Parameters.Clear();
                           
                        }

                        Cmd.CommandText = "SELECT Ident_current('NVO_ContainerTxns')";
                        LastCntrTrnsID = Cmd.ExecuteScalar().ToString();

                        Cmd.CommandText = " Update NVO_Containers set StatusCode=@StatusCode,CurrentPortID=@CurrentPortID,LastMoveMentID=@LastMoveMentID,AgencyID=@AgencyID,DtModified=@DtModified,DepotID=@DepotID,ModeOfTransportID=@ModeOfTransportID,VesVoyID=@VesVoyID,CustomerID=@CustomerID where ID=@ID";

                        Cmd.Parameters.Add(_dbFactory.GetParameter("@ID", LastCntrID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@LastMoveMentID", LastCntrTrnsID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@CurrentPortID", LastCntrportID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@StatusCode", LastStatusCode));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@DepotID", LastDepotID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@AgencyID", LastAgencyID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@DtModified", LastMvDate));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@ModeOfTransportID", LastTransitID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@VesVoyID", LastVesVoyID));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@CustomerID", LastCustomerID));
                        result = Cmd.ExecuteNonQuery();
                        Cmd.Parameters.Clear();
                    }

                    trans.Commit();
                    result = 1;
                    return result;

                }
                catch (Exception ex)
                {
                    int CCC = rowdTCN;
                    int hh = rowdt;
                    string ss = ex.ToString();
                    trans.Rollback();
                    return 0;
                }

            }


            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (con != null && con.State != ConnectionState.Closed)
                    con.Close();

            }

        }
        public int InsertEmail(DataTable dt)
        {


            int indexrow = 0;
            int result = 0;
            int IDv = 0;

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
                try
                {
                    Cmd.CommandText = "truncate table NVO_TempCntrMovement";
                    Cmd.ExecuteNonQuery();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        IDv = 0;

                        Cmd.CommandText = " IF ((SELECT COUNT(*) FROM NVO_TempCntrMovement WHERE ID=@ID)<=0) " +
                                    " BEGIN " +
                                    " INSERT INTO NVO_TempCntrMovement (CntrNo,StatusCode,DtMovement,AgencyCode,LocationCode,Transit,Depot,Vessel_Voyage,BLNumber,Customer)" +
                                    " values (@CntrNo,@StatusCode,@DtMovement,@AgencyCode,@LocationCode,@Transit,@Depot,@Vessel_Voyage,@BLNumber,@Customer) " +
                                    " END " +
                                    " ELSE " +
                                    " UPDATE NVO_TempCntrMovement SET CntrNo=@CntrNo,StatusCode=@StatusCode,DtMovement=@DtMovement,AgencyCode=@AgencyCode,LocationCode=@LocationCode,Transit=@Transit,Depot=@Depot,Vessel_Voyage=@Vessel_Voyage,BLNumber=@BLNumber,Customer=@Customer WHERE ID=@ID ";

                        Cmd.Parameters.Add(_dbFactory.GetParameter("@ID", 0));

                        Cmd.Parameters.Add(_dbFactory.GetParameter("@CntrNo", dt.Rows[i]["CONTAINER NUMBER"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@StatusCode", dt.Rows[i]["MOVE CODE"].ToString()));

                        if (dt.Rows[i]["MOVE DATE"].ToString() != "")
                        {
                            string timet = DateTime.Parse(dt.Rows[i]["MOVE DATE"].ToString()).ToString("MM/dd/yyyy");

                            Cmd.Parameters.Add(_dbFactory.GetParameter("@DtMovement", timet));
                        }
                        else
                        {
                            Cmd.Parameters.Add(_dbFactory.GetParameter("@DtMovement", DBNull.Value));
                        }
                       
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@AgencyCode", dt.Rows[i]["AGENCY CODE"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@LocationCode", dt.Rows[i]["LOCATION"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@Transit", dt.Rows[i]["TRANSIT"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@Depot", dt.Rows[i]["DEPOT"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@Vessel_Voyage", dt.Rows[i]["VESSEL_VOYAGE"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@BLNumber", dt.Rows[i]["BL NUMBER"].ToString()));
                        Cmd.Parameters.Add(_dbFactory.GetParameter("@Customer", dt.Rows[i]["CUSTOMER"].ToString()));

                        Cmd.ExecuteNonQuery();
                        Cmd.Parameters.Clear();

                        indexrow = i;

                    }

                    trans.Commit();
                    result = 1;
                    return result;

                }
                catch (Exception ex)
                {
                    int hh = indexrow;
                    string ss = ex.ToString();
                    trans.Rollback();
                    return 0;
                }

            }


            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (con != null && con.State != ConnectionState.Closed)
                    con.Close();

            }

        }

        public DataTable GetPort(string PortName)
        {
            string _Query = "Select * from NVO_PortMaster where PortCode like '%" + PortName + "%'";
            return Mng.GetViewData(_Query, "");
        }
        public DataTable GetCntrs(string PortName)
        {
            string _Query = " Select * from NVO_Containers where CntrNo like '%" + PortName + "%'";
            return Mng.GetViewData(_Query, "");
        }
        public DataTable GetAgency(string PortName)
        {
            string _Query = " Select * from NVO_AgencyMaster where AgencyCode like '%" + PortName + "%'";
            return Mng.GetViewData(_Query, "");
        }
        public DataTable GetDEPOT(string PortName)
        {
            string _Query = " Select * from NVO_DepotMaster where DepName like '%" + PortName + "%'";
            return Mng.GetViewData(_Query, "");
        }

        public DataTable GetTempCntr()
        {
             string _Query = "Select DISTINCT NVO_Containers.ID,NVO_TempCntrMovement.CntrNo from NVO_TempCntrMovement inner Join NVO_Containers on  NVO_Containers.CntrNo = NVO_TempCntrMovement.CntrNo ORDER BY  NVO_Containers.ID";
            return Mng.GetViewData(_Query, "");
        }
        public DataTable GetCntrListDistinct(string CntrID)
        {
            string _Query = " Select DISTINCT T.CntrNo,C.ID AS CntrID,T.StatusCode,DtMovement,T.AgencyCode,LocationCode,Transit,Depot, isnull((Select top 1 ID from NVO_PortMaster where PortCode = T.LocationCode  ) ,0)AS LocationID , isnull((Select top 1 ID from NVO_AgencyMaster where AgencyCode = T.AgencyCode  ) ,0)AS AgencyID , isnull((Select top 1 ID from NVO_DepotMaster where DepName = T.Depot  ) ,0) AS DepotID, isnull((Select top 1 ID from NVO_GeneralMaster where GeneralName = T.Transit  ) ,0) AS TransitID, isnull((Select top 1 ID from NVO_View_VoyageDetails V where V.VesVoy = T.Vessel_Voyage  ) ,0) AS VesVoyID,isnull((Select top 1 ID from NVO_Booking where BookingNo = T.BLNumber  ) ,0) AS BkgID,isnull((select top 1 CID  from NVO_CusBranchLocation  inner join NVO_CustomerMaster on NVO_CustomerMaster.ID=NVO_CusBranchLocation.CustomerID where (CustomerName+'-'+Branch) = T.Customer  ) ,0) AS CustomerID from NVO_TempCntrMovement T inner Join NVO_Containers C on  C.CntrNo = T.CntrNo where C.ID = " + CntrID + " order by T.DtMovement asc, T.STATUSCODE DESC";
            return Mng.GetViewData(_Query, "");
        }
        
        public DataTable GetModeOfTransport(string ModeTransport)
        {
            string _Query = "Select * from NVO_GeneralMaster where GeneralName like '%" + ModeTransport + "%'";
            return Mng.GetViewData(_Query, "");
        }
    }
}