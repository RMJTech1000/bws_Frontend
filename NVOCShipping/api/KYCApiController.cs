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
using System.Net.Mail;

namespace NVOCShipping.api
{
    public class KYCApiController : ApiController
    {
        #region anand
        [ActionName("kycView")]
        public List<MyKYC> kycView(MyKYC Data)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.KYCCustomerMaster(Data);
            return st;
        }
        [ActionName("KYCDashBordCount")]
        public List<MyKYC> KYCDashBordCount(MyKYC Data)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.KYCDashBordCountMaster(Data);
            return st;
        }

        [ActionName("kycPartyDD")]
        public List<MyKYC> kycPartyDD(MyKYC Data)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.KYCPartyMaster(Data);
            return st;
        }

        [ActionName("ApprovalValues")]
        public List<MyKYC> ApprovalValues(MyKYC Data)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.KYCApprovalMaster(Data);
            return st;
        }

        [ActionName("ApprovalAgencyValues")]
        public List<MyKYC> ApprovalAgencyValues(MyKYC Data)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.KYC_Agency_ApprovalMaster(Data);
            return st;
        }

        [ActionName("KycBusinessTypes")]
        public List<MyCustomer> BusinessTypes(MyCustomer Data)
        {
            PartyManager cm = new PartyManager();
            List<MyCustomer> st = cm.GetBusinessTypeMaster(Data);
            return st;
        }
        [ActionName("InsertKYCApproval")]
        public List<MyKYC> InsertKYCApproval(MyKYC BussData)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.InsertKYCApproval(BussData);
            return st;
        }

        [ActionName("InsertKYC_Agency_Approval")]
        public List<MyKYC> InsertKYC_Agency_Approval(MyKYC BussData)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.InsertKYC_Agency_Approval(BussData);
            return st;
        }

        [ActionName("InsertKYC_Agency_Delete")]
        public List<MyKYC> InsertKYC_Agency_Delete(MyKYC BussData)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.InsertKYC_Agency_Delete(BussData);
            return st;
        }


        





        [ActionName("InsertKYCReject")]
        public List<MyKYC> InsertKYCReject(MyKYC BussData)
        {
            KYCManager cm = new KYCManager();
            List<MyKYC> st = cm.InsertKYCReject(BussData);
            return st;
        }

        [ActionName("EmailsendingKYC")]
        public List<MyKYC> EmailsendingEstimate(MyKYC Data)
        {

            SendEmailMNRManagerLocalValues cm = new SendEmailMNRManagerLocalValues();
            List<MyKYC> st = cm.EmailsendingKYCConfirm(Data);
            return st;
        }
        [ActionName("EmailsendingKYCReject")]
        public List<MyKYC> EmailsendingKYCRejectv(MyKYC Data)
        {

            SendEmailMNRManagerLocalValues cm = new SendEmailMNRManagerLocalValues();
            List<MyKYC> st = cm.EmailsendingKYCReject(Data);
            return st;
        }
        [ActionName("OnlinePaymentView")]
        public List<MyPayment> OnlinePaymentView(MyPayment Data)
        {
            KYCManager cm = new KYCManager();
            List<MyPayment> st = cm.OnlinePaymentView(Data);
            return st;
        }

        [ActionName("OnlinePaymentEdit")]
        public List<MyPayment> OnlinePaymentEdit(MyPayment Data)
        {
            KYCManager cm = new KYCManager();
            List<MyPayment> st = cm.OnlinePaymentEdit(Data);
            return st;
        }
        [ActionName("OnlinePaymentDtlsEdit")]
        public List<MyPayment> OnlinePaymentDtlsEdit(MyPayment Data)
        {
            KYCManager cm = new KYCManager();
            List<MyPayment> st = cm.OnlinePaymentDtlsEdit(Data);
            return st;
        }

        public class SendEmailMNRManagerLocalValues
        {
            MNRNewManager Manag = new MNRNewManager();
            public List<MyKYC> EmailsendingKYCConfirm(MyKYC Data)
            {
                List<MyKYC> ViewList = new List<MyKYC>();


                DateTime dtDate = Convert.ToDateTime(System.DateTime.Now.Date.ToShortDateString());
                var creation_date = String.Format("{0:dd-MMM-yyyy}", dtDate);
                string strSubHeader = "<td style='font-family:Arial;font-weight: bold;font-size:14px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;'>";
                string strSubHeader1 = "<td colspan='4' style='font-family:Arial;font-weight: bold;font-size:15px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;text-align:center;'>";
                string sHtml = "";
                sHtml = "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:17px; font-weight:bold; text-decoration:underline; color:#FF6600; '>System Generated Message:</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Your KYC has been approved,</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Your login credentials as follows </td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Party Name : "+Data.CompanyName +" </td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Login Id   : " + Data.MailId + " </td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Password   : " + Data.MailId + " </td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform on KYC Approved details  </td></tr>";


                DataTable _dtCom = GetCompanyDetails();
                DataTable dtE = GetKyc_AgencyEmailid(Data.ID.ToString());

                MailMessage EmailObject = new MailMessage();
                EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString());
                EmailObject.To.Add(new MailAddress(Data.MailId));

                for( int z=0; z< dtE.Rows.Count; z++)
                {
                    var EmailIDTo = dtE.Rows[z]["Emailid"].ToString().Split(',');
                    for (int y = 0; y < EmailIDTo.Length; y++)
                    {
                        if (EmailIDTo[y].ToString() != "")
                        {
                            EmailObject.CC.Add(new MailAddress(EmailIDTo[y].ToString()));
                        }
                    }
                }
                EmailObject.Body = sHtml;
                EmailObject.IsBodyHtml = true;
                EmailObject.Priority = MailPriority.Normal;
                EmailObject.Subject = "KYC Approved For " +Data.CompanyName+"";
                EmailObject.Priority = MailPriority.Normal;
                SmtpClient SMTPServer = new SmtpClient();
                SMTPServer.UseDefaultCredentials = true;
                SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());
                SMTPServer.Host = "smtp.office365.com";
                SMTPServer.ServicePoint.MaxIdleTime = 1;
                SMTPServer.Port = 587;
                SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                SMTPServer.EnableSsl = true;
                SMTPServer.Send(EmailObject);

                ViewList.Add(new MyKYC
                {
                    AlertMessage = "Email sent successfully"

                });
                return ViewList;


            }

            public List<MyKYC> EmailsendingKYCReject(MyKYC Data)
            {
                List<MyKYC> ViewList = new List<MyKYC>();


                DateTime dtDate = Convert.ToDateTime(System.DateTime.Now.Date.ToShortDateString());
                var creation_date = String.Format("{0:dd-MMM-yyyy}", dtDate);
                string strSubHeader = "<td style='font-family:Arial;font-weight: bold;font-size:14px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;'>";
                string strSubHeader1 = "<td colspan='4' style='font-family:Arial;font-weight: bold;font-size:15px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;text-align:center;'>";
                string sHtml = "";
                sHtml = "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:17px; font-weight:bold; text-decoration:underline; color:#FF6600; '>System Generated Message:</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Your KYC has been rejected,</td></tr>";
                //sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Your login credentials as follows </td></tr>";
                //sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Party Name : " + Data.CompanyName + " </td></tr>";
                //sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Login Id   : " + Data.MailId + " </td></tr>";
                //sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Password   : " + Data.MailId + " </td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform on KYC Rejected details  </td></tr>";


                DataTable _dtCom = GetCompanyDetails();

                MailMessage EmailObject = new MailMessage();
                EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString());
                EmailObject.To.Add(new MailAddress(Data.MailId));

                ////EmailObject.To.Add(new MailAddress(dtAuto.Rows[0]["UserEmailID"].ToString()));
                ////EmailObject.Bcc.Add(new MailAddress("ganesh@rmjtech.in"));
                EmailObject.Body = sHtml;
                EmailObject.IsBodyHtml = true;
                EmailObject.Priority = MailPriority.Normal;
                EmailObject.Subject = "KYC Rejected For " + Data.CompanyName + "";
                EmailObject.Priority = MailPriority.Normal;
                SmtpClient SMTPServer = new SmtpClient();
                SMTPServer.UseDefaultCredentials = true;
                SMTPServer.Credentials = new NetworkCredential(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailPwd"].ToString());
                SMTPServer.Host = "smtp.office365.com";
                SMTPServer.ServicePoint.MaxIdleTime = 1;
                SMTPServer.Port = 587;
                SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                SMTPServer.EnableSsl = true;
                SMTPServer.Send(EmailObject);

                ViewList.Add(new MyKYC
                {
                    AlertMessage = "Email sent successfully"

                });
                return ViewList;


            }

            public DataTable GetCompanyDetails()
            {
                string _Query = "select * from NVO_NewCompnayDetails";
                return Manag.GetViewData(_Query, "");
            }

            public DataTable GetKyc_AgencyEmailid(string id)
            {
                string _Query = " select EmailID from NVO_Online_Customer_AgencyDtls " +
                                " inner join NVO_AgencyEmailDtls on NVO_AgencyEmailDtls.AgencyID = NVO_Online_Customer_AgencyDtls.AgencyID " +
                                " where AlertTypeID = 349  and Online_id=" + id;
                return Manag.GetViewData(_Query, "");
            }



            #endregion
        }
    }
}
