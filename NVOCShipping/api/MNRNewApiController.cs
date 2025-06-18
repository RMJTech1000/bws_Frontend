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
using System;
using System.Net.Mail;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.IO;

namespace NVOCShipping.api
{
    public class MNRNewApiController : ApiController
    {
       
        [ActionName("MNRNewManagementView")]
        public List<MyMNRNew> MNRNewManagementView(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRNewManagementViewData(Data);
            return st;
        }

        [ActionName("MNRStatusCountView")]
        public List<MyMNRNew> MNRStatusCountView(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRStatusCountView(Data);
            return st;
        }
        [ActionName("BindMNRNewStatus")]
        public List<MyComponent> BindMNRNewStatus(MyComponent Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyComponent> st = cm.GetBindMNRNewStatus(Data);
            return st;
        }
        [ActionName("BindMNRSurveyStatus")]
        public List<MyComponent> BindMNRSurveyStatus(MyComponent Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyComponent> st = cm.GetBindMNRSurveyStatus(Data);
            return st;
        }
        
        [ActionName("MNRDetailsSurveyEdit")]
        public List<MyMNRNew> MNRDetailsSurveyEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRDetailsSurveyEdit(Data);
            return st;
        }
        [ActionName("MNRSurveyUpdate")]
        public List<MyMNRNew> MNRSurveyUpdate(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRSurveyUpdate(Data);
            return st;
        }

        [ActionName("MNRDetailsEstimateEdit")]
        public List<MyMNRNew> MNRDetailsEstimateEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRDetailsEstimateEdit(Data);
            return st;
        }
        [ActionName("MNREstimateUpdate")]
        public List<MyMNRNew> MNREstimateUpdate(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNREstimateUpdate(Data);
            return st;
        }

        [ActionName("MNRDetailsApproveEdit")]
        public List<MyMNRNew> MNRDetailsApproveEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRDetailsApproveEdit(Data);
            return st;
        }
        [ActionName("MNRApproveUpdate")]
        public List<MyMNRNew> MNRApproveUpdate(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRApproveUpdate(Data);
            return st;
        }

        [ActionName("MNRDetailCompleteEdit")]
        public List<MyMNRNew> MNRDetailCompleteEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRDetailCompleteEdit(Data);
            return st;
        }
        [ActionName("MNRCompleteUpdate")]
        public List<MyMNRNew> MNRCompleteUpdate(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRCompleteUpdate(Data);
            return st;
        }

        [ActionName("MNRNewUploadDocuments")]
        public List<MyMNRNew> MNRNewUploadDocuments(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRNewUploadDocuments(Data);
            return st;
        }
        [ActionName("MNRNewDocumentsEdit")]
        public List<MyMNRNew> MNRNewDocumentsEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRNewDocumentsEditView(Data);
            return st;
        }

        [ActionName("MNRHistoryView")]
        public List<MyMNRNew> MNRHistoryView(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRHistoryView(Data);
            return st;
        }

        [ActionName("MNRHistoryEdit")]
        public List<MyMNRNew> MNRHistoryEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRHistoryEdit(Data);
            return st;
        }


        [ActionName("BindMNRAttachTypeEstimate")]
        public List<MyMNRRepair> BindMNRAttachTypeEstimate(MyMNRRepair Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRRepair> st = cm.GetMNRAttachType(Data);
            return st;
        }

        [ActionName("BindMNRAttachTypeComplete")]
        public List<MyMNRRepair> BindMNRAttachTypeComplete(MyMNRRepair Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRRepair> st = cm.GetBindMNRAttachTypeComplete(Data);
            return st;
        }
        [ActionName("BindMNRAttachType")]
        public List<MyMNRRepair> BindMNRAttachType(MyMNRRepair Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRRepair> st = cm.GetBindMNRAttachType(Data);
            return st;
        }
        [ActionName("MNRNewDocumentsHistoryEdit")]
        public List<MyMNRNew> MNRNewDocumentsHistoryEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRNewDocumentsHistoryEditView(Data);
            return st;
        }
        [ActionName("MNRAttachmentsDelete")]
        public List<MyMNRNew> MNRAttachmentsDelete(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRAttachmentsDelete(Data);
            return st;
        }

        [ActionName("MNRBindCostToDropdown")]
        public List<MyMNRNew> MNRBindCostToDropdown(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.GetCostTo(Data);
            return st;
        }
        [ActionName("MNRLineItemDropdown")]
        public List<MyMNRNew> MNRLineItemDropdown(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.GetLineItemDropdown(Data);
            return st;
        }


        [ActionName("InsertMNRRecoveryDetails")]
        public List<MyMNRNew> InsertMNRRecoveryDetails(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.InsertMNRRecoveryDetailsValues(Data);
            return st;
        }
        [ActionName("MNRRecoveryDetailsEdit")]
        public List<MyMNRNew> MNRRecoveryDetailsEdit(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRRecoveryDetailsEditValues(Data);
            return st;
        }
        [ActionName("MNRRecoveryDelete")]
        public List<MyMNRNew> MNRRecoveryDelete(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRRecoveryDelete(Data);
            return st;
        }
        [ActionName("EmailsendingEstimate")]
        public List<MySendingMNREmailAlert> EmailsendingEstimate(MySendingMNREmailAlert Data)
        {

            SendEmailMNRManagerLocalValues AccMange = new SendEmailMNRManagerLocalValues();
            List<MySendingMNREmailAlert> st = AccMange.EmailsendingEstimateConfirm(Data);
            return st;
        }

        [ActionName("EmailsendingApproved")]
        public List<MySendingMNREmailAlert> EmailsendingApproved(MySendingMNREmailAlert Data)
        {

            SendEmailMNRManagerLocalValues AccMange = new SendEmailMNRManagerLocalValues();
            List<MySendingMNREmailAlert> st = AccMange.EmailsendingApprovedConfirm(Data);
            return st;
        }
        public class SendEmailMNRManagerLocalValues
        {
            MNRNewManager Manag = new MNRNewManager();
            public List<MySendingMNREmailAlert> EmailsendingEstimateConfirm(MySendingMNREmailAlert Data)
            {
                List<MySendingMNREmailAlert> ViewList = new List<MySendingMNREmailAlert>();
              
                DataTable _dtx = GetMNREstimateEmail(Data.ItemsMNRIDs);
                if (_dtx.Rows.Count > 0)
                {
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
                    sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Hi,</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings </td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform on EOR Pending details  </td></tr>";
                    sHtml += "</table> <br/>";
                    #region HTMl Display

                    string rowAlignleftLTHeader = "<td align='center' style='background-color:#e1e4e7; width:100px!important;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:12px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;white-space:nowrap!important'>";
                    string rowAlignRightAlter = "<td align='right' style='background-color:#FFFFFF;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:11px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                    string rowAlignleftAlter = "<td align='left' style='background-color:#FFFFFF;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:11px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                    string rowAlignleftHeader = "<td align='center' style='background-color:#e1e4e7;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:12px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;white-space:nowrap!important'>";
                    string rowAlignleftHeaderTotal = "<td align='center' style='background-color:#D6EAF8;  color:#B03A2E;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:13px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;' colspan='7'>";
                    string rowAlignleftHeaderAmt = "<td align='right' style='background-color:#D6EAF8;  color:#B03A2E;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:13px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                    sHtml += " <table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                    sHtml += "<tr>";
                    DataTable _dtCom = GetCompanyDetails();
                    sHtml += "<td style='font-family:Tw Cen MT Condensed; font-size:20px; font-weight:bold; text-align:center; background-color:#007999; color:#FFFFFF; border-right:0px solid #303297; border-left:0px solid #303297; border-top: 0px solid #303297; border-bottom: 0px solid #303297;' colspan='9'> "+ _dtCom.Rows[0]["CompanyName"].ToString() + "</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += "<td style='font-family:Arial; font-size:16px;  font-weight:bold; text-align:center; background-color:#716a6a; color:#FFFFFF; border-right:0px solid #336699; border-left:0px solid #336699; border-top: 0px solid #336699; border-bottom: 0px solid #336699;'  colspan='9'>EOR PENDING </td>";
                    sHtml += "</tr>";
                    sHtml += "<tr style='font-family:Tw Cen MT Condensed; font-size:12px; text-align:center;'>";
                    sHtml += rowAlignleftHeader + "Location</td>";
                    sHtml += rowAlignleftHeader + "Agency</td>";
                    sHtml += rowAlignleftLTHeader + "MNRNo</td>";
                    sHtml += rowAlignleftHeader + " Container No</td>";
                    sHtml += rowAlignleftLTHeader + "Container Type</td>";
                    sHtml += rowAlignleftLTHeader + "Empty In</td>";
                    sHtml += rowAlignleftHeader + "Depot</td>";
                    sHtml += rowAlignleftHeader + "Currency</td>";
                    sHtml += rowAlignleftHeader + "Amount</td>";
                  
                    sHtml += "</tr>";

                    decimal Amt = 0;
                   
                    for (int j = 0; j < _dtx.Rows.Count; j++)
                    {
                        sHtml += "<tr style='font-family:Tw Cen MT Condensed; font-size:12px; text-align:center;'>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["Location"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["Agency"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["MNRRefNo"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["CntrNo"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["CntrType"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["EmptyGateIN"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["Depot"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["MNRCurrency"].ToString() + "</td>";
                        sHtml += rowAlignRightAlter + _dtx.Rows[j]["Estimate"].ToString() + "</td>";

                        Amt += decimal.Parse(_dtx.Rows[j]["Estimate"].ToString());
                        sHtml += "</tr>";

                    }
                    sHtml += "<tr>";
                    sHtml += rowAlignleftHeaderTotal + "</td>";
                    sHtml += rowAlignleftHeaderAmt + "TOTAL" + "</td>";
                    sHtml += rowAlignleftHeaderAmt + Amt + " </td>";
                  

                    sHtml += "</tr>";
                    sHtml += "</table>";
                    #endregion

                    sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                    sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>BWS.</td></tr>";

                    sHtml += "</table>";

                   // DataTable _dtCom = GetCompanyDetails();
                    if (_dtCom.Rows.Count > 0)
                    {
                        MailMessage EmailObject = new MailMessage();
                        EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString());
                        DataTable dtAuto = GetMNREmailEorPendingsending(_dtx.Rows[0]["AgencyID"].ToString());
                        if (dtAuto.Rows.Count > 0)
                        {
                            var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                            for (int y = 0; y < EmailID.Length; y++)
                            {
                                if (EmailID[y].ToString() != "")
                                {
                                    EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                                }
                            }
                            //if (_dtx.Rows[0]["IsDepot"].ToString() =="1")
                            //{
                            //    DataTable dtAutoUser = GetMNREmailEorUserDepo(_dtx.Rows[0]["EstimatedBy"].ToString());
                            //    if (dtAutoUser.Rows.Count > 0)
                            //    {
                            //        var EmailIDUser = dtAutoUser.Rows[0]["EmailID"].ToString().Split(',');
                            //        for (int k = 0; k < EmailIDUser.Length; k++)
                            //        {
                            //            if (EmailIDUser[k].ToString() != "")
                            //            {
                            //                EmailObject.To.Add(new MailAddress(EmailIDUser[k].ToString()));
                            //            }
                            //        }

                            //    }
                            //}
                       
                            DataTable _dtattach = GetMNREstimateAtatchments(Data.ItemsMNRIDs);
  
                            if (_dtattach.Rows.Count > 0)
                            {
                                for (int k = 0; k < _dtattach.Rows.Count; k++)
                                {
                                    EmailObject.Attachments.Add(new Attachment(HttpContext.Current.Server.MapPath("~/MNRAttachments/" + _dtattach.Rows[k]["FileName"].ToString())));
                                }
                            }
                            //EmailObject.To.Add(new MailAddress(dtAuto.Rows[0]["UserEmailID"].ToString()));
                            EmailObject.To.Add(new MailAddress("ganeshchandran80@gmail.com"));
                            //EmailObject.Bcc.Add(new MailAddress("venkat@neridashipping.com"));
                            //EmailObject.Bcc.Add(new MailAddress("ganesh@rmjtech.in"));
                            EmailObject.Body = sHtml;
                            EmailObject.IsBodyHtml = true;
                            EmailObject.Priority = MailPriority.Normal;
                            EmailObject.Subject = "MNR ---- " + _dtx.Rows[0]["Location"].ToString() + "----" + _dtx.Rows[0]["Agency"].ToString() + "----" + "EOR PENDING";
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


                        }

                    }


                }


            
                ViewList.Add(new MySendingMNREmailAlert
                {
                    AlertMessage = "Email sent successfully"

                });
                return ViewList;


            }


            public List<MySendingMNREmailAlert> EmailsendingApprovedConfirm(MySendingMNREmailAlert Data)
            {
                List<MySendingMNREmailAlert> ViewList = new List<MySendingMNREmailAlert>();

              
              

                string strHTML = "";

                try
                {
                    DataTable _dtx = GetMNREstimateEmail(Data.ItemsMNRIDs);
                    if (_dtx.Rows.Count > 0)
                    {

                        Document doc = new Document();
                        Rectangle rec = new Rectangle(670, 900);
                        doc = new Document(rec);
                        Paragraph para = new Paragraph();
                        MemoryStream memoryStream = new MemoryStream();
                        PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                        doc.Open();
                        // PdfWriter pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
                        //// PdfWriter pdfWriter = PdfWriter.GetInstance(doc, Response.OutputStream);
                        //// pdfWriter = PdfWriter.GetInstance(doc, new FileStream(Server.MapPath("~/RRPDF\\" + dtv.Rows[0]["RatesheetNo"].ToString() + ".pdf"), FileMode.Create));
                        // doc.Open();

           
                        #region Header LOGO COMPANY NAME
                        //-------------HEADER-------------------//

                        iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(2);
                        tbllogo.Width = 100;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        //tbllogo.Cellpadding = 1;
                        tbllogo.BorderWidth = 0;
                        Cell cell = new Cell();
                        //cell.Width = 10;

                        var img = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/BWSLOGO.png"));
                        img.ScaleAbsolute(160f, 80f);
                        img.Alignment = Element.ALIGN_LEFT;
                        cell = new Cell(img);
                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        cell.Width = 20;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        tbllogo.AddCell(cell);

                        ///--SPACE--//
                        cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 16, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        //cell.Colspan = 3;
                        tbllogo.AddCell(cell);
                        ///----/////
                        DataTable dtc = GetCompanyDetails();
                        if (dtc.Rows.Count > 0)
                        {
                            cell = new Cell(new Phrase(dtc.Rows[0]["CompanyName"].ToString(), new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                        }


                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        //cell.Colspan = 2;
                        tbllogo.AddCell(cell);

                        cell = new Cell(new Phrase("EOR APPROVED ", new Font(Font.HELVETICA, 14, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        //cell.Colspan = 3;
                        tbllogo.AddCell(cell);

                        cell = new Cell(new Phrase(dtc.Rows[0]["CompanyAddress"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        tbllogo.Alignment = Element.ALIGN_LEFT;
                        cell.Colspan = 2;
                        tbllogo.AddCell(cell);

                        //cell = new Cell(new Phrase(DS.Tables[0].Rows[0]["Location"].ToString() + " - " + DS.Tables[0].Rows[0]["Pincode"].ToString() + " Tel # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Phone1"].ToString() + "   Fax # " + DS.Tables[0].Rows[0]["Areacode"].ToString() + "  " + DS.Tables[0].Rows[0]["Fax"].ToString(), new Font(Font.HELVETICA, 9, Font.BOLD)));
                        //cell.BorderWidth = 0;
                        //cell.Colspan = 6;
                        // tbllogo.AddCell(cell);

                        doc.Add(tbllogo);

                        para = new Paragraph("");
                        doc.Add(para);

                        para = new Paragraph("", new Font(Font.HELVETICA, 14.0F, Font.BOLD, Color.BLACK));
                        para.Alignment = Element.ALIGN_RIGHT;
                        doc.Add(para);

                        //----------SPACE----------------------------------
                        iTextSharp.text.Table Tblspace2 = new iTextSharp.text.Table(1);
                        doc.Add(Tblspace2);

                        //------------------------------------------------------------------------
                        #endregion

                        #region EOR BASIC details
                        //-------------------Bookingparty and Ratesheet details-----------
                        PdfContentByte content = pdfWriter.DirectContent;
                        PdfPTable mtable = new PdfPTable(2);
                        mtable.WidthPercentage = 100;
                        mtable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;


                        PdfPTable Tbl1 = new PdfPTable(1);
                        Tbl1.WidthPercentage = 50;
                        PdfPCell cell1 = new PdfPCell(new Phrase("Depot", new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK)));
                        cell1.Colspan = 6;
                        cell1.HorizontalAlignment = 1;
                        cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell1.BorderWidth = 0;
                        cell1.FixedHeight = 23f;
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.Colspan = 1;
                        Tbl1.AddCell(cell1);


                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["Depot"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.BorderWidth = 0;
                        Tbl1.AddCell(cell1);

                        var Addresss = Regex.Split(_dtx.Rows[0]["DepotAddress"].ToString(), "\r\n|\r|\n");
                        for (int a = 0; a < Addresss.Length; a++)
                        {
                            cell1 = new PdfPCell(new Phrase(Addresss[a].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                            cell1.BorderWidth = 0;
                            Tbl1.AddCell(cell1);
                        }

                        // cell1 = new PdfPCell(new Phrase("MUMBAI CITY, MAHARASHTRA, 400092", new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLUE)));

                        // cell1.BorderWidth = 0;
                        //Tbl1.AddCell(cell1);
                        mtable.AddCell(Tbl1);



                        Tbl1 = new PdfPTable(2);
                        Tbl1.WidthPercentage = 50;
                        Tbl1.HorizontalAlignment = Element.ALIGN_RIGHT;


                        cell1 = new PdfPCell(new Phrase("MNR NO", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);


                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["MNRRefNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("BL Number", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["BLNumber"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Container No", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["CntrNo"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Container Type", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["CntrType"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);


                        cell1 = new PdfPCell(new Phrase("Location", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["Location"].ToString(), new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BorderWidth = 1;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Tbl1.AddCell(cell1);

                        mtable.AddCell(Tbl1);
                        doc.Add(mtable);

                        //----------SPACE----------------------------------
                        iTextSharp.text.Table Tblspaces2 = new iTextSharp.text.Table(1);
                        doc.Add(Tblspaces2);

                        //------------------------------------------------------------------------
                        #endregion

                        #region Location POL POD AGENCY
                        // /----------------------- LocTable-----------------------///

                        PdfPTable TblLocs = new PdfPTable(new float[] { 2, 2, 2,  });
                        TblLocs.HorizontalAlignment = Element.ALIGN_LEFT;
                        TblLocs.WidthPercentage = 100;

                        cell1 = new PdfPCell(new Phrase("Agency", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblLocs.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Port Of Loading", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblLocs.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Port Of Discharge", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblLocs.AddCell(cell1);



                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["Agency"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblLocs.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["POL"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblLocs.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase(_dtx.Rows[0]["POD"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblLocs.AddCell(cell1);



                        doc.Add(TblLocs);



                        #endregion

                        #region EOR Details
                        //----------------- Booking Details--------------//

                        iTextSharp.text.Table Tbl3 = new iTextSharp.text.Table(1);
                        Tbl3.Width = 100;
                        Tbl3.Alignment = Element.ALIGN_LEFT;
                        Tbl3.Cellpadding = 0;
                        Tbl3.BorderWidth = 0;

                        //Sub Heading
                        cell = new Cell(new Phrase("EOR Details", new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK)));

                        cell.BorderWidth = 0;
                        cell.Colspan = 1;
                        Tbl3.AddCell(cell);
                        doc.Add(Tbl3);

                   

                        iTextSharp.text.Table Tbl5 = new iTextSharp.text.Table(12);
                        Tbl5.Width = 100;
                        Tbl5.Alignment = Element.ALIGN_LEFT;
                        Tbl5.Cellpadding = 1;
                        Tbl5.BorderWidth = 0;


                        //Caption
                        cell = new Cell(new Phrase("Empty In Date ", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 2;
                        Tbl5.AddCell(cell);
                        //Value
                        cell = new Cell(new Phrase(" : " + _dtx.Rows[0]["EmptyDateIN"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 4;
                        Tbl5.AddCell(cell);

                        //Caption
                        cell = new Cell(new Phrase("Survey Completion Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 3;
                        Tbl5.AddCell(cell);

                        //Value
                        cell = new Cell(new Phrase(" :  " + _dtx.Rows[0]["SurveyDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 3;
                        Tbl5.AddCell(cell);


                        //Caption
                        cell = new Cell(new Phrase("Estimation Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 2;
                        Tbl5.AddCell(cell);
                        //Value
                        cell = new Cell(new Phrase(" : " + _dtx.Rows[0]["EstimateDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 4;
                        Tbl5.AddCell(cell);


                        //Caption
                        cell = new Cell(new Phrase("Estimation Amount", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 3;
                        Tbl5.AddCell(cell);
                        //Value
                        cell = new Cell(new Phrase(" : " + _dtx.Rows[0]["Estimate"].ToString() +" / "+ _dtx.Rows[0]["Currency"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 3;
                        Tbl5.AddCell(cell);

                        //Caption
                        cell = new Cell(new Phrase("Approval Date", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 2;
                        Tbl5.AddCell(cell);
                        //Value
                        cell = new Cell(new Phrase(" : " + _dtx.Rows[0]["ApprovedDate"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 4;
                        Tbl5.AddCell(cell);

                        //Caption
                        cell = new Cell(new Phrase("Approval Amount", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 3;
                        Tbl5.AddCell(cell);
                        //Value
                        cell = new Cell(new Phrase(" : " + _dtx.Rows[0]["Approved"].ToString() + " / " + _dtx.Rows[0]["Currency"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 3;
                        Tbl5.AddCell(cell);

                        //Caption
                        cell = new Cell(new Phrase("Consignee", new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 2;
                        Tbl5.AddCell(cell);
                        //Value
                        cell = new Cell(new Phrase(" : " + _dtx.Rows[0]["Consignee"].ToString(), new Font(Font.HELVETICA, 10, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        cell.Colspan = 4;
                        Tbl5.AddCell(cell);

                   
                        doc.Add(Tbl5);


                        #endregion

                        //----------SPACE----------------------------------
                        iTextSharp.text.Table Tblspace4 = new iTextSharp.text.Table(1);
                        doc.Add(Tblspace4);

                        //------------------------------------------------------------------------

                        #region  Recovery Details


                        PdfPTable TblBreakUp = new PdfPTable(new float[] { 2, 2, 2, 2});
                        TblBreakUp.HorizontalAlignment = Element.ALIGN_LEFT;
                        TblBreakUp.WidthPercentage = 100;

                        cell1 = new PdfPCell(new Phrase("Accountability", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblBreakUp.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Item-No", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblBreakUp.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Amount", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblBreakUp.AddCell(cell1);

                        cell1 = new PdfPCell(new Phrase("Remarks", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell1.BackgroundColor = new Color(152, 178, 209);
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        TblBreakUp.AddCell(cell1);

                        DataTable _dtk = GetMNRRecoveryDetails(Data.ItemsMNRIDs);
                       
                        for (int k = 0; k < _dtk.Rows.Count; k++)
                        {

                            cell1 = new PdfPCell(new Phrase(_dtk.Rows[k]["Accountability"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            TblBreakUp.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtk.Rows[k]["ItemNo"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            TblBreakUp.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtk.Rows[k]["RecoveryAmount"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            TblBreakUp.AddCell(cell1);

                            cell1 = new PdfPCell(new Phrase(_dtk.Rows[k]["Remarks"].ToString(), new Font(Font.HELVETICA, 9, Font.NORMAL, Color.BLACK)));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            TblBreakUp.AddCell(cell1);

                        }

                        doc.Add(TblBreakUp);



                        #endregion

                        #region FOOTER

                        iTextSharp.text.Table Tbl7 = new iTextSharp.text.Table(1);
                        Tbl7.Width = 100;
                        Tbl7.Alignment = Element.ALIGN_LEFT;
                        Tbl7.DefaultCell.Border = 0;
                        Tbl7.DefaultCellBorder = Rectangle.NO_BORDER;
                        Tbl7.Border = Rectangle.NO_BORDER;

                        cell = new Cell(new Phrase(" \n \n \n \n \n \n \n \n \n \n \n \n", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        Tbl7.AddCell(cell);
                        doc.Add(Tbl7);

                        ///---------FOOTER----------------//
                        ///

                        iTextSharp.text.Table Tblline = new iTextSharp.text.Table(1);
                        Tblline.Width = 100;
                        Tblline.DefaultCellBorder = Rectangle.NO_BORDER;
                        Tblline.Border = Rectangle.NO_BORDER;
                        Tblline.Cellpadding = 1;

                        cell = new Cell(new Phrase("", new Font(Font.HELVETICA, 10, Font.NORMAL)));
                        cell.BorderWidthTop = 0;
                        cell.BorderWidthRight = 0;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthBottom = 0;
                        cell.BackgroundColor = new Color(98, 141, 214);
                        Tblline.AddCell(cell);
                        doc.Add(Tblline);

                        iTextSharp.text.Table Tbl12 = new iTextSharp.text.Table(4);
                        Tbl12.Width = 100;
                        Tbl12.Alignment = Element.ALIGN_LEFT;
                        Tbl12.Cellpadding = 1;
                        Tbl12.BorderWidth = 1;

                        cell = new Cell(new Phrase("Filed By : " + _dtx.Rows[0]["EORFiledBy"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        Tbl12.AddCell(cell);

                        cell = new Cell(new Phrase("Filed On :  " + _dtx.Rows[0]["EstimateDate"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        Tbl12.AddCell(cell);

                        cell = new Cell(new Phrase("Approved By :  " + _dtx.Rows[0]["EORApprovedBy"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        Tbl12.AddCell(cell);

                        cell = new Cell(new Phrase("Approved On : " + _dtx.Rows[0]["ApprovedDate"].ToString() + "", new Font(Font.HELVETICA, 10, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        cell.Colspan = 1;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        Tbl12.AddCell(cell);
                        doc.Add(Tbl12);
                        #endregion



                        pdfWriter.CloseStream = false;
                        doc.Close();

                        byte[] bytes = memoryStream.ToArray();
                        memoryStream.Close();
                       // DataTable _dtx = GetMNREstimateEmail(Data.ItemsMNRIDs);
                        if (_dtx.Rows.Count > 0)
                        {
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
                            sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Hi,</td></tr>";
                            sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings </td></tr>";
                            sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform on EOR Approved details  </td></tr>";
                            sHtml += "</table> <br/>";
                            #region HTMl Display

                            string rowAlignleftLTHeader = "<td align='center' style='background-color:#e1e4e7; width:100px!important;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:12px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;white-space:nowrap!important'>";
                            string rowAlignRightAlter = "<td align='right' style='background-color:#FFFFFF;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:11px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                            string rowAlignleftAlter = "<td align='left' style='background-color:#FFFFFF;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:11px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                            string rowAlignleftHeader = "<td align='center' style='background-color:#e1e4e7;  color: #000000;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:12px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;white-space:nowrap!important'>";
                            string rowAlignleftHeaderTotal = "<td align='center' style='background-color:#D6EAF8;  color:#B03A2E;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:13px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;' colspan='7'>";
                            string rowAlignleftHeaderAmt = "<td align='right' style='background-color:#D6EAF8;  color:#B03A2E;border-top: Black thin inset; border-bottom: Black thin inset; border-right: Black thin inset; border-Left: Black thin inset;  font-family: Arial, Helvetica, sans-serif; font-style:normal; font-size:13px;font-weight:bold;padding-Left:15px;padding-right:15px;vertical-align:middle;'>";
                            sHtml += " <table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                            sHtml += "<tr>";
                            DataTable _dtCom = GetCompanyDetails();
                            sHtml += "<td style='font-family:Tw Cen MT Condensed; font-size:20px; font-weight:bold; text-align:center; background-color:#007999; color:#FFFFFF; border-right:0px solid #303297; border-left:0px solid #303297; border-top: 0px solid #303297; border-bottom: 0px solid #303297;' colspan='13' > " + _dtCom.Rows[0]["CompanyName"].ToString() + "</td>";
                            sHtml += "</tr>";
                            sHtml += "<tr>";
                            sHtml += "<td style='font-family:Arial; font-size:16px;  font-weight:bold; text-align:center; background-color:#716a6a; color:#FFFFFF; border-right:0px solid #336699; border-left:0px solid #336699; border-top: 0px solid #336699; border-bottom: 0px solid #336699;border-right:2px solid #336699;'  colspan='9'>EOR Approved </td>";
                            sHtml += "<td style='font-family:Arial; font-size:16px;  font-weight:bold; text-align:center; background-color:#716a6a; color:#FFFFFF; border-right:0px solid #336699; border-left:0px solid #336699; border-top: 0px solid #336699; border-bottom: 0px solid #336699;'  colspan='4'> Approved Break Up </td>";
                            sHtml += "</tr>";
                            sHtml += "<tr style='font-family:Tw Cen MT Condensed; font-size:12px; text-align:center;'>";
                            sHtml += rowAlignleftHeader + "Location</td>";
                            sHtml += rowAlignleftHeader + "Agency</td>";
                            sHtml += rowAlignleftLTHeader + "MNRNo</td>";
                            sHtml += rowAlignleftHeader + " Container No</td>";
                            sHtml += rowAlignleftLTHeader + "Container Type</td>";
                            sHtml += rowAlignleftLTHeader + "Empty In</td>";
                            sHtml += rowAlignleftHeader + "Depot</td>";
                            sHtml += rowAlignleftHeader + "Currency</td>";
                            sHtml += rowAlignleftHeader + "Est.Amt</td>";
                            sHtml += rowAlignleftLTHeader + "Apr.Amt </td>";
                            sHtml += rowAlignleftHeader + "Principal</td>";
                            sHtml += rowAlignleftHeader + "Agency</td>";
                            sHtml += rowAlignleftHeader + "Customer</td>";
                            sHtml += "</tr>";

                            decimal Amt = 0, AprAmt = 0, PrAmt = 0, AgAmt = 0, CusAmt = 0;

                            for (int j = 0; j < _dtx.Rows.Count; j++)
                            {
                                sHtml += "<tr style='font-family:Tw Cen MT Condensed; font-size:12px; text-align:center;'>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["Location"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["Agency"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["MNRRefNo"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["CntrNo"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["CntrType"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["EmptyGateIN"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["Depot"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["MNRCurrency"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["Estimate"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["Approved"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["RecPrincipalAmt"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["RecAgentAmt"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtx.Rows[j]["RecCusAmt"].ToString() + "</td>";

                                Amt += decimal.Parse(_dtx.Rows[j]["Estimate"].ToString());
                                AprAmt += decimal.Parse(_dtx.Rows[j]["Approved"].ToString());
                                PrAmt += decimal.Parse(_dtx.Rows[j]["RecPrincipalAmt"].ToString());
                                AgAmt += decimal.Parse(_dtx.Rows[j]["RecAgentAmt"].ToString());
                                CusAmt += decimal.Parse(_dtx.Rows[j]["RecCusAmt"].ToString());
                                sHtml += "</tr>";

                            }

                            sHtml += "<tr>";
                            sHtml += rowAlignleftHeaderTotal + "</td>";
                            sHtml += rowAlignleftHeaderAmt + "TOTAL" + "</td>";
                            sHtml += rowAlignleftHeaderAmt + Amt + " </td>";
                            sHtml += rowAlignleftHeaderAmt + AprAmt + " </td>";
                            sHtml += rowAlignleftHeaderAmt + PrAmt + " </td>";
                            sHtml += rowAlignleftHeaderAmt + AgAmt + " </td>";
                            sHtml += rowAlignleftHeaderAmt + CusAmt + " </td>";
                            sHtml += "</tr>";
                            sHtml += "</table>";



                            sHtml += " <table border='0' cellpadding='0' cellspacing='0' width='50%' style='margin-top:30px!important;'>";
                            sHtml += "<td style='font-family:Arial; font-size:16px;  font-weight:bold; text-align:center; background-color:#716a6a; color:#FFFFFF; border-right:0px solid #336699; border-left:0px solid #336699; border-top: 0px solid #336699; border-bottom: 0px solid #336699;border-right:2px solid #336699;'  colspan='5'>Recovery Break-Up </td>";
                            sHtml += "<tr style='font-family:Tw Cen MT Condensed; font-size:12px; text-align:center;'>";
                            sHtml += rowAlignleftLTHeader + "S.No </td>";
                            sHtml += rowAlignleftHeader + "Accountability</td>";
                            sHtml += rowAlignleftHeader + "Item-No</td>";
                            sHtml += rowAlignleftHeader + "Amount</td>";
                            sHtml += rowAlignleftHeader + "Remarks</td>";
                            sHtml += "</tr>";
                            DataTable _dtj = GetMNRRecoveryDetails(Data.ItemsMNRIDs);
                            int sl = 1;
                            for (int k = 0; k < _dtj.Rows.Count; k++)
                            {
                                sHtml += "<tr style='font-family:Tw Cen MT Condensed; font-size:12px; text-align:center;'>";
                                sHtml += rowAlignRightAlter + sl + "</td>";
                                sHtml += rowAlignRightAlter + _dtj.Rows[k]["Accountability"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtj.Rows[k]["ItemNo"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtj.Rows[k]["RecoveryAmount"].ToString() + "</td>";
                                sHtml += rowAlignRightAlter + _dtj.Rows[k]["Remarks"].ToString() + "</td>";

                                sl++;
                                sHtml += "</tr>";
                            }
                            sHtml += "</table>";
                            #endregion

                            sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                            sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                            sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>BWS.</td></tr>";

                            sHtml += "</table>";

                            //DataTable _dtCom = GetCompanyDetails();
                            if (_dtCom.Rows.Count > 0)
                            {
                                MailMessage EmailObject = new MailMessage();
                                EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString());
                                DataTable dtAuto = GetMNREmailEorApprovalsending(_dtx.Rows[0]["AgencyID"].ToString());
                                if (dtAuto.Rows.Count > 0)
                                {
                                    var EmailID = dtAuto.Rows[0]["EmailID"].ToString().Split(',');
                                    for (int y = 0; y < EmailID.Length; y++)
                                    {
                                        if (EmailID[y].ToString() != "")
                                        {
                                            EmailObject.To.Add(new MailAddress(EmailID[y].ToString()));
                                        }
                                    }
                                    //if (_dtx.Rows[0]["IsDepot"].ToString() == "1")
                                    //{
                                    //    DataTable dtAutoUser = GetMNREmailEorUserDepo(_dtx.Rows[0]["EstimatedBy"].ToString());
                                    //    if (dtAutoUser.Rows.Count > 0)
                                    //    {
                                    //        var EmailIDUser = dtAutoUser.Rows[0]["EmailID"].ToString().Split(',');
                                    //        for (int k = 0; k < EmailIDUser.Length; k++)
                                    //        {
                                    //            if (EmailIDUser[k].ToString() != "")
                                    //            {
                                    //                EmailObject.To.Add(new MailAddress(EmailIDUser[k].ToString()));
                                    //            }
                                    //        }

                                    //    }
                                    //}
                                    DataTable _dtattach = GetMNREstimateAtatchments(Data.ItemsMNRIDs);

                                    //if (_dtattach.Rows.Count > 0)
                                    //{
                                    //    for (int k = 0; k < _dtattach.Rows.Count; k++)
                                    //    {
                                    //        if (_dtattach.Rows[k]["IsDepot"].ToString() == "1")
                                    //        {
                                    //             //EmailObject.Attachments.Add(new Attachment("https://avisa-containerlines.com/MNRAttachments/" + "983_10_CntrImges.zip"));
                                    //              EmailObject.Attachments.Add(new Attachment(HttpContext.Current.Request.MapPath("https://depo.oceanus-lines.com/MNRAttachments/" + "983_10_CntrImges.zip")));
                                    //        }
                                    //        else
                                    //        {
                                    //            EmailObject.Attachments.Add(new Attachment(HttpContext.Current.Server.MapPath("~/MNRAttachments/" + _dtattach.Rows[k]["FileName"].ToString())));
                                    //        }

                                          
                                    //    }
                                    //}
                                    EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "EOR_" + _dtx.Rows[0]["CntrNo"].ToString() + ".pdf"));
                                    //EmailObject.To.Add(new MailAddress(dtAuto.Rows[0]["UserEmailID"].ToString()));
                                    EmailObject.To.Add(new MailAddress("muthukrishnan1976k@gmail.com"));
                                    //EmailObject.Bcc.Add(new MailAddress("venkat@neridashipping.com"));
                                    //EmailObject.Bcc.Add(new MailAddress("ganesh@rmjtech.in"));
                                    EmailObject.Body = sHtml;
                                    EmailObject.IsBodyHtml = true;
                                    EmailObject.Priority = MailPriority.Normal;
                                    EmailObject.Subject = "MNR ---- " + _dtx.Rows[0]["Location"].ToString() + "----" + _dtx.Rows[0]["Agency"].ToString() + "----" + "EOR APPROVED";
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


                                }

                            }


                        }

                    }
                }
                catch (Exception ex)
                {
                    ViewList.Add(new MySendingMNREmailAlert
                    {
                        AlertMessage = ex.Message.ToString()

                    });
                }
                ViewList.Add(new MySendingMNREmailAlert
                {
                    AlertMessage = "Email sent successfully"

                });

                return ViewList;

            }


            public DataTable GetMNREstimateEmail(string MNRID)
            {
                string _Query = " select * from NVO_ViewMNRNewDetails where MNRID IN(" + MNRID + ") ";
                return Manag.GetViewData(_Query, "");
            }
            public DataTable GetMNREstimateAtatchments(string MNRID)
            {
                string _Query = " select *,isnull(IsDepot,0) as IsDepotv  from NVO_MNRCntrRepAttachments where RepairReqID IN(" + MNRID + ") ";
                return Manag.GetViewData(_Query, "");
            }
            
            public DataTable GetCompanyDetails()
            {
                string _Query = "select * from NVO_NewCompnayDetails";
                return Manag.GetViewData(_Query, "");
            }
            public DataTable GetMNREmailEorPendingsending(string AgentID)
            {
                string _Query = "select EmailID from NVO_AgencyEmailDtls  AE inner join NVO_GeneralMaster GM ON GM.ID = AE.AlertTypeID where GM.GeneralName = 'MNR EOR PENDING' " +
                    " and AE.AgencyID = " + AgentID;
                return Manag.GetViewData(_Query, "");
            }

            public DataTable GetMNREmailEorApprovalsending(string AgentID)
            {
                string _Query = "select EmailID from NVO_AgencyEmailDtls  AE inner join NVO_GeneralMaster GM ON GM.ID = AE.AlertTypeID where GM.GeneralName = 'MNR APPROVED' " +
                    " and AE.AgencyID = " + AgentID;
                return Manag.GetViewData(_Query, "");
            }
            public DataTable GetMNRRecoveryDetails(string MNRID)
            {

                string _Query = "select * from NVO_MNRNewRecoveryDtls WHERE MNRID=" + MNRID;

                return Manag.GetViewData(_Query, "");
            }
            public DataTable GetMNREmailEorUserDepo(string UserID)
            {
                string _Query = "select EmailID from NVO_UserDetails Where ID = " + UserID;
                return Manag.GetViewData(_Query, "");
            }
        }

        #region  MNR DEPOT

        [ActionName("MNRNewManagementDepot")]
        public List<MyMNRNew> MNRNewManagementDepot(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRNewManagementDepotData(Data);
            return st;
        }

        [ActionName("MNRStatusCountDepot")]
        public List<MyMNRNew> MNRStatusCountDepot(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRStatusCountDepot(Data);
            return st;
        }

        [ActionName("MNRDepotHistoryView")]
        public List<MyMNRNew> MNRDepotHistoryView(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRDepotHistoryView(Data);
            return st;
        }


        #endregion


        [ActionName("MNRRejectUpdate")]
        public List<MyMNRNew> MNRRejectUpdate(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.MNRRejectUpdate(Data);
            return st;
        }
        [ActionName("InsertMNRDetailsHistory")]
        public List<MyMNRNew> InsertMNRDetailsHistory(MyMNRNew Data)
        {
            MNRNewManager cm = new MNRNewManager();
            List<MyMNRNew> st = cm.InsertMNRDetailsHistoryValues(Data);
            return st;
        }
    }
}
