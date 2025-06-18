using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using DataManager;
using DataTier;
using System.Net.Mail;
using System.Text;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;


namespace NVOCShipping.api
{
    public class OnBard_shipmentController : ApiController
    {
        [ActionName("OnBordConfirmationTransaction")]
        public List<MyOnBoard> OnBordConfirmationTransaction(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.OnBoard_ConfirmTransation(Data);
            return st;
        }

        [ActionName("OnBordEmailSending")]
        public List<MyOnBoard> OnBordEmailSending(MyOnBoard Data)
        {

            SendEmailOnBarod Mange = new SendEmailOnBarod();
            List<MyOnBoard> st = Mange.OnBoardSendingEmail(Data);
            return st;
        }

        [ActionName("VesselCertificateVesselName")]
        public List<MyOnBoard> VesselCertificateVesselName(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.VeslCerficatificate_BLtablevessel(Data);
            return st;
        }

        [ActionName("VslVoyageCertificateInsert")]
        public List<MyOnBoard> VslVoyageCertificateInsert(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.VslcertificateInsert(Data);
            return st;
        }

        [ActionName("VslVoyageCertificateView")]
        public List<MyOnBoard> VslVoyageCertificateView(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.VesslCertificateView_dtls(Data);
            return st;
        }

        [ActionName("ExistingVslVoyageCertificate")]
        public List<MyOnBoard> ExistingVslVoyageCertificate(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.ExistingVesslCertificatevalue(Data);
            return st;
        }

        [ActionName("BLNumbervalues")]
        public List<MyOnBoard> BLNumbervalues(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.BLNumbervalue(Data);
            return st;
        }

        [ActionName("BkgNumbervalues")]
        public List<MyOnBoard> BkgNumbervalues(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.BkgNumbervalue(Data);
            return st;
        }

        [ActionName("BkgVesselCertificateVesselName")]
        public List<MyOnBoard> BKgVesselCertificateVesselName(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.VeslCerficatificate_Bkgtablevessel(Data);
            return st;
        }




        [ActionName("VesselCertificateEmailSending")]
        public List<MyOnBoard> VesselCertificateEmailSending(MyOnBoard Data)
        {

            SendEmailOnBarod Mange = new SendEmailOnBarod();
            List<MyOnBoard> st = Mange.VeselCertificationSendingEmail(Data);
            return st;
        }



        [ActionName("DelayCertificateEmailSending")]
        public List<MyOnBoard> DelayCertificateEmailSending(MyOnBoard Data)
        {

            SendEmailOnBarod Mange = new SendEmailOnBarod();
            List<MyOnBoard> st = Mange.delayCertificationSendingEmail(Data);
            return st;
        }


        [ActionName("ShippingCertificateEmailSending")]
        public List<MyOnBoard> ShippingCertificateEmailSending(MyOnBoard Data)
        {

            SendEmailOnBarod Mange = new SendEmailOnBarod();
            List<MyOnBoard> st = Mange.ShippingCertificationSendingEmail(Data);
            return st;
        }


        [ActionName("FreeTimeShippingCertificateEmailSending")]
        public List<MyOnBoard> FreeTimeShippingCertificateEmailSending(MyOnBoard Data)
        {

            SendEmailOnBarod Mange = new SendEmailOnBarod();
            List<MyOnBoard> st = Mange.FreeTimeShippingCertificationSendingEmail(Data);
            return st;
        }



        [ActionName("VesselEarlyArrivalCertificateEmailSending")]
        public List<MyOnBoard> VesselEarlyArrivalCertificateEmailSending(MyOnBoard Data)
        {

            SendEmailOnBarod Mange = new SendEmailOnBarod();
            List<MyOnBoard> st = Mange.VesselEarlyArrivalCertificationSendingEmail(Data);
            return st;
        }







        [ActionName("DelayCertificateView")]
        public List<MyOnBoard> DelayCertificateView(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.DeleyCertificateView_dtls(Data);
            return st;
        }


        [ActionName("DelayCertificateInsert")]
        public List<MyOnBoard> DelayCertificateInsert(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.DelaycertificateInsert(Data);
            return st;
        }

        [ActionName("ExistingDelayCertificate")]
        public List<MyOnBoard> ExistingDelayCertificate(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.ExistingDelayCertificatevalue(Data);
            return st;
        }


        [ActionName("DelayVessel")]
        public List<MyOnBoard> DelayVessel(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.VeslDelay(Data);
            return st;
        }

        [ActionName("DelayVesselwiseBL")]
        public List<MyOnBoard> DelayVesselwiseBL(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.VeslwiseBLDelay(Data);
            return st;
        }

        [ActionName("CertificateInsert")]
        public List<MyOnBoard> CertificateInsert(MyOnBoard Data)
        {
            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.certificateInsert(Data);
            return st;
        }

        [ActionName("ExistingCertificate")]
        public List<MyOnBoard> ExistingCertificate(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.ExistingCertificatevalue(Data);
            return st;
        }

        [ActionName("CertificateView")]
        public List<MyOnBoard> CertificateView(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.CertificateView_dtls(Data);
            return st;
        }

        [ActionName("Certificatedelete")]
        public List<MyOnBoard> Certificatedelete(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.deleteCertificate(Data);
            return st;
        }

        [ActionName("CertificateTitle")]
        public List<MyOnBoard> CertificateTitle(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.CertificateTitle(Data);
            return st;
        }

        [ActionName("CertificateSubjectBind")]
        public List<MyOnBoard> CertificateSubjectBind(MyOnBoard Data)
        {

            OnBoardConfirmation Mange = new OnBoardConfirmation();
            List<MyOnBoard> st = Mange.CertificateTitleBind(Data);
            return st;
        }

    }

    public class SendEmailOnBarod
    {
        DocumentManager Manag = new DocumentManager();
        public List<MyOnBoard> OnBoardSendingEmail(MyOnBoard Data)
        {
            List<MyOnBoard> ViewList = new List<MyOnBoard>();
            try
            {
                DataTable _dtCom = GetCompnayDetails();
                DataTable _dtx = BindOnBoardConfirm(Data);
                MailMessage EmailObject = new MailMessage();
                EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                var EmailIDTo = Data.Email.Split(',');
                for (int y = 0; y < EmailIDTo.Length; y++)
                {
                    if (EmailIDTo[y].ToString() != "")
                    {
                        EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                    }
                }


                // string strSubHeader = "<td style='font-family:Arial;font-weight: bold;font-size:14px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;'>";
                // string strSubHeader1 = "<td colspan='4' style='font-family:Arial;font-weight: bold;font-size:15px;padding-top:4px;padding-left:3px;padding-bottom:4px;padding-right:3px;background-color:#666;border:1px solid white;color:white;text-align:center;'>";
                string sHtml = "";
                sHtml = "<table border='0' cellpadding='0' cellspacing='0' width='100%'>";
                sHtml += "<tr><td style='font-family:Arial; font-size:17px; font-weight:bold;'>Dear Sir:</td></tr>";
                sHtml += "<br /><br /><tr><td style='font-family:Arial; font-size:14px;'>We are pleased  to inform you that below  mentioned  containers  are shipped On Board on vessel " + _dtx.Rows[0]["BLVesVoy"].ToString() + " sailed from " + _dtx.Rows[0]["POL"].ToString() + " Dated " + _dtx.Rows[0]["ETD"].ToString() + " </td></tr>";
                sHtml += "<br /><br /><tr><td style='font-family:Arial; font-size:14px;'>POD:  " + _dtx.Rows[0]["POD"].ToString() + " </td></tr>";
                sHtml += "<br /><br /><br /><br /><tr><td style='font-family:Arial; font-size:14px; font-weight:bold;'>Note :  This is an notification email & please do not respond/reply to this mail </td></tr>";
                sHtml += "<br /><tr><td style='font-family:Arial; font-size:14px; font-weight:bold;'>Further queries,you may contact/revert to custome service desk</td></tr>";
                sHtml += "<br /><tr><td style='font-family:Arial; font-size:14px; font-weight:bold;'></td></tr>";
                sHtml += "<br /><tr><td style='font-family:Arial; font-size:14px; font-weight:bold;'></td></tr>";

                sHtml += "<br /><tr><td style='font-family:Arial; font-size:18px; font-weight:bold;'>Regards</td></tr>";
                sHtml += "<br /><br /><tr><td style='font-family:Arial; font-size:18px; font-weight:bold;'>ADMIN</td></tr>";
                sHtml += "<br /><br /><tr><td style='font-family:Arial; font-size:18px; font-weight:bold;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
               
                EmailObject.Body = sHtml;
                EmailObject.IsBodyHtml = true;
                EmailObject.Priority = MailPriority.Normal;
                EmailObject.Subject = "Onboard confirmation : " + _dtx.Rows[0]["BLVesVoy"].ToString();
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
            catch (Exception ex)
            {
                ViewList.Add(new MyOnBoard
                {
                    AlertMessage = ex.Message

                });
            }
            ViewList.Add(new MyOnBoard
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;

        }

        public DataTable BindOnBoardConfirm(MyOnBoard Data)
        {
            string _Query = " select NVO_BOL.Id,BLNumber, BLVesVoy,(select top(1) PortName from NVO_PortMaster where NVO_PortMaster.ID = NVO_BOL.POLID) as POL, " +
                            " (select top(1) PortName from NVO_PortMaster where NVO_PortMaster.ID = NVO_BOL.PODID) as POD, " +
                            " (select top(1) PortName from NVO_PortMaster where NVO_PortMaster.ID = NVO_BOL.FPODID) as FPOD,BkgParty, " +
                            " (select count(CntrTypes) from NVO_BookingCntrTypes where BKgID = NVO_Booking.Id) as CntrCount, " +
                            " (select top(1)(select sum(TEUS) from NVO_tblCntrTypes where NVO_tblCntrTypes.ID = NVO_BookingCntrTypes.CntrTypes) " +
                            " from NVO_BookingCntrTypes where BKgID = NVO_Booking.Id) as CntrTeus,'' as Email, " +
                            " (select top(1)(select TerminalName from NVO_TerminalMaster where NVO_TerminalMaster.ID = NVO_VoyageRoute.TerminalID) from NVO_VoyageRoute  " +
                            " inner join NVO_PortMainMaster On NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID " +
                            " where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID ASC) as TerminalName, " +
                            " (select  top(1) Convert(varchar, ETD, 103) from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by NVO_VoyageRoute.RID ASC) as ETD " +
                            " from NVO_BOL " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_BOL.BkgID where NVO_BOL.Id=" + Data.BLID;
            return Manag.GetViewData(_Query, "");
        }


        public List<MyOnBoard> VeselCertificationSendingEmail(MyOnBoard Data)
        {
            List<MyOnBoard> ViewList = new List<MyOnBoard>();
            try
            {
                DataTable dtv = GetVesselData(Data);
                DataTable _dtx = getVesselCertificate(Data);

                Document doc = new Document();
                Rectangle rec = new Rectangle(670, 900);
                doc = new Document(rec);
                Paragraph para = new Paragraph();
                MemoryStream memoryStream = new MemoryStream();
                PdfWriter pdfWriter = PdfWriter.GetInstance(doc, memoryStream);
                doc.Open();

                PdfContentByte cb = pdfWriter.DirectContent;
                cb.SetColorStroke(Color.BLACK);

                BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader, 14);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


                iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/img/LOGO1.png"));
                png1.SetAbsolutePosition(40, 830);
                png1.ScalePercent(75f);
                doc.Add(png1);




                DataTable dtc = GetAgencyDetails(_dtx.Rows[0]["AgencyID"].ToString());
                BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bfheader3, 14);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 250, 855, 0);
                cb.SetFontAndSize(bfheader3, 9);
                int AddRow = 840;
                var LogoAddresss = Regex.Split(dtc.Rows[0]["Address"].ToString(), "\r\n|\r|\n");
                for (int a = 0; a < LogoAddresss.Length; a++)
                {
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, LogoAddresss[a].ToString(), 250, AddRow, 0);
                    AddRow -= 10;
                }


                BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.DARK_GRAY);

                cb.BeginText();
                //Border-Top//
                //cb.MoveTo(10, 935);
                //cb.LineTo(695, 935);

                //Top1//
                cb.MoveTo(10, 805);
                cb.LineTo(660, 805);
                //Top2//
                //cb.MoveTo(10, 770);
                //cb.LineTo(660, 770);

                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);

                //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                cb.EndText();
                cb.BeginText();
                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Date :" + dtv.Rows[0]["Date"].ToString(), 20, 783, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME / VOYAGE : " + dtv.Rows[0]["VSLVoyage"].ToString(), 20, 767, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BLNUMBER  : " + dtv.Rows[0]["BLNumber"].ToString(), 20, 753, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL  CERTIFICATE", 20, 730, 0);

                cb.SetFontAndSize(bfheader3, 8);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "THIS IS TO  CERTIFY THAT", 20, 718, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CERTIFICATE FROM  THE OWNER, CARRIER  OR CAPTAIN  OF THE CARRYING", 20, 680, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL  OT THEIR  AGENT SHOWING  ITS NAME FLAG AND NATIONALITY  AND ", 20, 670, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "CONFIRMING  THAT ", 20, 660, 0);



                string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                string[] Aaddsplit;
                int ColumnRows = 630;
                for (int x = 0; x < subjetv.Length; x++)
                {
                    Aaddsplit = subjetv[x].Split('\n');

                    for (int k = 0; k < Aaddsplit.Length; k++)
                    {

                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 20, ColumnRows, 0);
                        ColumnRows -= 13;
                    }
                }
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YOURS FAITHFULLY", 20, 595, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YOURS FAITHFULLY", 20, 400, 0);


                cb.SetFontAndSize(bfheader2, 9);
                cb.SetColorFill(Color.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "AS AGENT FOR THE CARRIER -", 20, 300, 0);

                cb.Stroke();

                cb.EndText();

                pdfWriter.CloseStream = false;
                doc.Close();



                byte[] bytes = memoryStream.ToArray();
                memoryStream.Close();


                DataTable _dtCom = GetCompnayDetails();
                MailMessage EmailObject = new MailMessage();

                EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                var EmailIDTo = Data.EmailTo.Split(',');
                var EmailIDCC = Data.EmailCC.Split(',');
                for (int y = 0; y < EmailIDTo.Length; y++)
                {
                    if (EmailIDTo[y].ToString() != "")
                    {
                        EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                    }
                }
                for (int y = 0; y < EmailIDCC.Length; y++)
                {
                    if (EmailIDCC[y].ToString() != "")
                    {
                        EmailObject.CC.Add(new MailAddress(EmailIDCC[y].ToString()));
                    }
                }


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
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Vessel Certificate</td></tr>";
                sHtml += "</table>";
                sHtml += "<br />";
                sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                sHtml += "<tr>";
                sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >Vessel Certificate</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + " </td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader + " Booking Party</td>";
                sHtml += strSubHeader + " Booking No</td>";
                sHtml += strSubHeader + " Vessel/Voyage</td>";
                sHtml += strSubHeader + " Port of Loading</td>";
                sHtml += "</tr>";
                sHtml += "<tr>";
                sHtml += strSubHeader + _dtx.Rows[0]["BkgParty"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["BLNumber"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["VSLVoyage"].ToString() + "</td>";
                sHtml += strSubHeader + _dtx.Rows[0]["POL"].ToString() + "</td>";
                sHtml += "</tr>";
                sHtml += "</table>";

                sHtml += "</table>";
                sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";

                sHtml += "</table>";
                EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "VesselCerifcation" + ".pdf"));
                EmailObject.Body = sHtml;
                EmailObject.IsBodyHtml = true;
                EmailObject.Priority = MailPriority.Normal;
                EmailObject.Subject = "Vessel Certificate : " + _dtx.Rows[0]["VSLVoyage"].ToString();
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
            catch (Exception ex)
            {
                ViewList.Add(new MyOnBoard
                {
                    AlertMessage = ex.Message

                });
            }
            ViewList.Add(new MyOnBoard
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;

        }


        public DataTable getVesselCertificate(MyOnBoard Data)
        {
            string _Query = " Select BLID,BLNumber as BLNumber,AgencyID, CertificateType,CertificateTitle,VSLVoyage,Subject, Date,BkgParty,POL " +
                            " from NVO_VesselCertificate "+
                            " inner join NVO_BOL on NVO_BOL.ID = NVO_VesselCertificate.BLID "+
                            " inner join NVO_Booking on NVO_Booking.Id = NVO_BOL.BkgID where NVO_BOL.Id=" + Data.BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetAgencyDetails(string AgencyID)
        {

            string _Query = "Select * from nvo_agencymaster where ID=" + AgencyID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetVesselData(MyOnBoard Data)
        {
            string _Query = " Select BLID,(SELECT TOP(1) BLNumber FROM nvo_BOL where Id =NVO_VesselCertificate.BLID) as BLNumber," +
                            " CertificateType, CertificateTitle, VSLVoyage, Subject, convert(varchar,Date,106) as Date  from NVO_VesselCertificate where BLID=" + Data.BLID;
            return Manag.GetViewData(_Query, "");
        }

        public List<MyOnBoard> delayCertificationSendingEmail(MyOnBoard Data)
        {

            List<MyOnBoard> ViewList = new List<MyOnBoard>();
            try
            {
                string[] Array = Data.Items.Split(new[] { "Insert:" }, StringSplitOptions.None);
                for (int i = 1; i < Array.Length; i++)
                {
                    var CharSplit = Array[i].ToString().TrimEnd(',').Split(',');
                    Data.BLID = CharSplit[0].ToString();
                    var EmailIDs = CharSplit[1].ToString();

                    #region ff
                    DataTable _dtCom = GetCompnayDetails();
                    DataTable dtv = getDelayCertificate(Data.BLID.ToString(), Data.CertificateType);

                    MemoryStream memoryStream = new MemoryStream();
                    //doc.Open();

                    Document doc = new Document();
                    Rectangle rec = new Rectangle(670, 900);
                    doc = new Document(rec);
                    Paragraph para = new Paragraph();


                    PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);
                    doc.Open();


                    PdfContentByte cb = writer.DirectContent;
                    cb.SetColorStroke(Color.BLACK);
                    int _Xp = 10, _Yp = 785, YDiff = 10;

                    BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader, 14);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                    //iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                    png2.SetAbsolutePosition(45, 750);
                    png2.ScalePercent(17f);
                    doc.Add(png2);


                    DataTable dtc = GetAgencyDetails(dtv.Rows[0]["AgencyID"].ToString());
                    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader3, 14);
                    cb.SetColorFill(Color.BLACK);
                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtc.Rows[0]["AgencyName"].ToString(), 300, 855, 0);
                    cb.SetFontAndSize(bfheader3, 9);
                    int AddRow = 840;


                    BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                    cb.SetFontAndSize(bfheader2, 9);
                    cb.SetColorFill(Color.DARK_GRAY);

                    cb.BeginText();



                    cb.SetFontAndSize(bfheader2, 10);
                    cb.SetColorFill(Color.BLACK);

                    //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PAYMENT VOUCHER", 275, 888, 0);
                    cb.EndText();
                    cb.BeginText();

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL DELAY NOTICE ", 250, 730, 0);

                    cb.MoveTo(250, 728);
                    cb.LineTo(370, 728);

                    //--------------------
                    cb.SetFontAndSize(bfheader2, 9);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE", 65, 700, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 95, 700, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO", 65, 680, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 95, 680, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 100, 700, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO VALUABLE CUSTOMER", 100, 680, 0);

                    BaseFont BHHeaderSubject = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(BHHeaderSubject, 11);
                    Font underlineFont = new Font(BHHeaderSubject, 11, Font.UNDERLINE);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUBJECT: VESSEL ARRIVAL DELAYED – " +
                        dtv.Rows[0]["VesVoy"].ToString() + " ETA:  " + dtv.Rows[0]["ETA"].ToString(), 65, 650, 0);


                    cb.MoveTo(65, 648);
                    cb.LineTo(500, 648);
                    BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);




                    int ColumnRows = 620;

                    cb.SetFontAndSize(bfheader7, 9);
                    string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                    string[] Aaddsplit;

                    for (int x = 0; x < subjetv.Length; x++)
                    {
                        Aaddsplit = subjetv[x].Split('\n');

                        for (int k = 0; k < Aaddsplit.Length; k++)
                        {

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 65, ColumnRows, 0);
                            ColumnRows -= 14;
                        }
                    }
                    //box
                    BaseFont bfheader11 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetColorFill(Color.BLACK);
                    cb.SetFontAndSize(bfheader11, 9);

                    ColumnRows -= 20;
                    int LineVLeft = ColumnRows;
                    cb.MoveTo(65, ColumnRows);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME/VOYAGE NO", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, ColumnRows, 0);

                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "LOAD PORT ETA", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 260, ColumnRows, 0);
                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YARD CLOSING", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["YardClosing"].ToString(), 260, ColumnRows, 0);

                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VoyageNotes"].ToString(), 260, ColumnRows, 0);


                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesselID"].ToString(), 260, ColumnRows, 0);

                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "NEXT PORT ETA", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["NextPortETA"].ToString(), 260, ColumnRows, 0);

                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REVISED ETA", 70, ColumnRows -= 20, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 260, ColumnRows, 0);

                    cb.MoveTo(65, ColumnRows -= 5);
                    cb.LineTo(500, ColumnRows);


                    //vertical
                    cb.MoveTo(65, LineVLeft);
                    cb.LineTo(65, ColumnRows);

                    cb.MoveTo(250, LineVLeft);
                    cb.LineTo(250, ColumnRows);

                    cb.MoveTo(500, LineVLeft);
                    cb.LineTo(500, ColumnRows);

                    cb.Stroke();
                    cb.EndText();


                    writer.CloseStream = false;
                    doc.Close();


                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();



                    MailMessage EmailObject = new MailMessage();

                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    var EmailIDTo = EmailIDs.Split(';');

                    for (int y = 0; y < EmailIDTo.Length; y++)
                    {
                        if (EmailIDTo[y].ToString() != "")
                        {
                            EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                        }
                    }


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
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Delay Certificate</td></tr>";
                    sHtml += "</table>";
                    sHtml += "<br />";
                    sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                    sHtml += "<tr>";
                    sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >Delay Certificate</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + " </ td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + " Booking Party</td>";
                    sHtml += strSubHeader + " Booking No</td>";
                    sHtml += strSubHeader + " Vessel/Voyage</td>";
                    sHtml += strSubHeader + " Port of Loading</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + dtv.Rows[0]["BkgParty"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["BLNumber"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["VesVoy"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["POL"].ToString() + "</td>";
                    sHtml += "</tr>";
                    sHtml += "</table>";

                    sHtml += "</table>";
                    sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                    sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + ".</td></tr>";

                    sHtml += "</table>";
                    EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "DelayCerifcation" + ".pdf"));
                    EmailObject.Body = sHtml;
                    EmailObject.IsBodyHtml = true;
                    EmailObject.Priority = MailPriority.Normal;
                    EmailObject.Subject = "Delay Certificate : " + dtv.Rows[0]["VesVoy"].ToString();
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

                    #endregion
                }
            }
            catch (Exception ex)
            {
                ViewList.Add(new MyOnBoard
                {
                    AlertMessage = ex.Message

                });
            }
            ViewList.Add(new MyOnBoard
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;

        }

        public List<MyOnBoard> ShippingCertificationSendingEmail(MyOnBoard Data)
        {

            List<MyOnBoard> ViewList = new List<MyOnBoard>();
            try
            {
                string[] Array = Data.Items.Split(new[] { "Insert:" }, StringSplitOptions.None);
                for (int i = 1; i < Array.Length; i++)
                {
                    var CharSplit = Array[i].ToString().TrimEnd(',').Split(',');
                    Data.BLID = CharSplit[0].ToString();
                    var EmailIDs = CharSplit[1].ToString();

                    #region ff
                    DataTable _dtCom = GetCompnayDetails();
                    DataTable dtv = getDelayCertificate(Data.BLID.ToString(),Data.CertificateType);

                    Document doc = new Document();
                    Rectangle rec = new Rectangle(700, 900);
                    doc = new Document(rec);
                    Paragraph para = new Paragraph();
                    MemoryStream memoryStream = new MemoryStream();
                    PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);

                    doc.Open();

                    PdfContentByte cb = writer.DirectContent;
                    cb.SetColorStroke(Color.BLACK);
                    int _Xp = 10, _Yp = 785, YDiff = 10;

                    BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader, 14);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);


                    iTextSharp.text.Image png2 = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));

                    png2.SetAbsolutePosition(25, 780);     //logo fixed location
                    png2.ScalePercent(15f);
                    doc.Add(png2);

                    BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader2, 15);
                    cb.SetColorFill(Color.BLACK);
                    cb.BeginText();

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SHIPPING CERTIFICATE", 260, 700, 0);//right

                    cb.MoveTo(260, 698);
                    cb.LineTo(440, 698);  //BORDER LINE CODE




                    BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader5, 14);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO WHOM IT MAY CONCERN", 260, 780, 0);//right


                    BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader7, 12);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE  :", 70, 730, 0);//right


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 660, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 640, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 620, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 600, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ": ", 250, 580, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING NUMBER  ", 70, 660, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE OF SHIPMENT", 70, 640, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ", 70, 620, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING  ", 70, 600, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF DISCHARGE  ", 70, 580, 0);//right

                    cb.MoveTo(60, 560);
                    cb.LineTo(650, 560);


                    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader3, 12);
                    cb.SetColorFill(Color.BLACK);
                   

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 130, 730, 0);



                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 260, 660, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Bldate1"].ToString(), 260, 640, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, 620, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POL"].ToString(), 260, 600, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POD"].ToString(), 260, 580, 0);




                    int ColumnRows = 530;
                    cb.SetFontAndSize(bfheader3, 9);
                    string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                    string[] Aaddsplit;

                    for (int x = 0; x < subjetv.Length; x++)
                    {
                        Aaddsplit = subjetv[x].Split('\n');

                        for (int k = 0; k < Aaddsplit.Length; k++)
                        {

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 70, ColumnRows, 0);
                            ColumnRows -= 14;
                        }
                    }





                    cb.Stroke();
                    cb.EndText();

                   
                    writer.CloseStream = false;
                    doc.Close();


                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();



                    MailMessage EmailObject = new MailMessage();

                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    var EmailIDTo = EmailIDs.Split(';');

                    for (int y = 0; y < EmailIDTo.Length; y++)
                    {
                        if (EmailIDTo[y].ToString() != "")
                        {
                            EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                        }
                    }


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
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Shipping Certificate</td></tr>";
                    sHtml += "</table>";
                    sHtml += "<br />";
                    sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                    sHtml += "<tr>";
                    sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >SHIPPING CERTIFICATE</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + " </ td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + " Booking Party</td>";
                    sHtml += strSubHeader + " Booking No</td>";
                    sHtml += strSubHeader + " Vessel/Voyage</td>";
                    sHtml += strSubHeader + " Port of Loading</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + dtv.Rows[0]["BkgParty"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["BLNumber"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["VesVoy"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["POL"].ToString() + "</td>";
                    sHtml += "</tr>";
                    sHtml += "</table>";

                    sHtml += "</table>";
                    sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                    sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + ".</td></tr>";

                    sHtml += "</table>";
                    EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "ShippingCerifcation" + ".pdf"));
                    EmailObject.Body = sHtml;
                    EmailObject.IsBodyHtml = true;
                    EmailObject.Priority = MailPriority.Normal;
                    EmailObject.Subject = "Shipping Certificate : " + dtv.Rows[0]["VesVoy"].ToString();
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

                    #endregion
                }
            }
            catch (Exception ex)
            {
                ViewList.Add(new MyOnBoard
                {
                    AlertMessage = ex.Message

                });
            }
            ViewList.Add(new MyOnBoard
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;

        }


        public List<MyOnBoard> FreeTimeShippingCertificationSendingEmail(MyOnBoard Data)
        {

            List<MyOnBoard> ViewList = new List<MyOnBoard>();
            try
            {
                string[] Array = Data.Items.Split(new[] { "Insert:" }, StringSplitOptions.None);
                for (int i = 1; i < Array.Length; i++)
                {
                    var CharSplit = Array[i].ToString().TrimEnd(',').Split(',');
                    Data.BLID = CharSplit[0].ToString();
                    var EmailIDs = CharSplit[1].ToString();

                    #region ff
                    DataTable _dtCom = GetCompnayDetails();
                    DataTable dtv = getDelayCertificate(Data.BLID.ToString(), Data.CertificateType);

                    Document doc = new Document();
                    Rectangle rec = new Rectangle(700, 900);
                    doc = new Document(rec);
                    Paragraph para = new Paragraph();
                    MemoryStream memoryStream = new MemoryStream();
                    PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);

                    doc.Open();

                    PdfContentByte cb = writer.DirectContent;
                    cb.SetColorStroke(Color.BLACK);
                    int _Xp = 10, _Yp = 785, YDiff = 10;

                    BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader, 14);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                    png.SetAbsolutePosition(25, 780);     //logo fixed location
                    png.ScalePercent(15f);
                    doc.Add(png);

                    BaseFont bfheader2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader2, 15);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "FREE TIME SHIPPING CERTIFICATE", 220, 780, 0);//right

                    cb.MoveTo(220, 778);
                    cb.LineTo(490, 778);  //BORDER LINE CODE


                    BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader5, 14);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO WHOM IT MAY CONCERN", 260, 740, 0);//right


                    BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader7, 12);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 680, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 660, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 640, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 620, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 250, 600, 0);//right


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE  ", 70, 680, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "BILL OF LADING NUMBER  ", 70, 660, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL & VOYAGE  ", 70, 640, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF LOADING  ", 70, 620, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "PORT OF DISCHARGE  ", 70, 600, 0);//right



                    cb.MoveTo(60, 590);
                    cb.LineTo(650, 590);


                    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader3, 12);
                    cb.SetColorFill(Color.BLACK);
                    cb.BeginText();


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Bldate1"].ToString(), 260, 680, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["BLNumber"].ToString(), 260, 660, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 260, 640, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POL"].ToString(), 260, 620, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["POD"].ToString(), 260, 600, 0);


                    int ColumnRows = 570;
                    cb.SetFontAndSize(bfheader3, 9);
                    string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                    string[] Aaddsplit;

                    for (int x = 0; x < subjetv.Length; x++)
                    {
                        Aaddsplit = subjetv[x].Split('\n');

                        for (int k = 0; k < Aaddsplit.Length; k++)
                        {

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 70, ColumnRows, 0);
                            ColumnRows -= 14;
                        }
                    }





                    cb.Stroke();
                    cb.EndText();


                    writer.CloseStream = false;
                    doc.Close();


                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();



                    MailMessage EmailObject = new MailMessage();

                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    var EmailIDTo = EmailIDs.Split(';');

                    for (int y = 0; y < EmailIDTo.Length; y++)
                    {
                        if (EmailIDTo[y].ToString() != "")
                        {
                            EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                        }
                    }


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
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Free Time Shipping Certificate</td></tr>";
                    sHtml += "</table>";
                    sHtml += "<br />";
                    sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                    sHtml += "<tr>";
                    sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >FREE TIME SHIPPING CERTIFICATE</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + " </ td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + " Booking Party</td>";
                    sHtml += strSubHeader + " Booking No</td>";
                    sHtml += strSubHeader + " Vessel/Voyage</td>";
                    sHtml += strSubHeader + " Port of Loading</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + dtv.Rows[0]["BkgParty"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["BLNumber"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["VesVoy"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["POL"].ToString() + "</td>";
                    sHtml += "</tr>";
                    sHtml += "</table>";

                    sHtml += "</table>";
                    sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                    sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + ".</td></tr>";

                    sHtml += "</table>";
                    EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "FreeTimeShippingCerifcation" + ".pdf"));
                    EmailObject.Body = sHtml;
                    EmailObject.IsBodyHtml = true;
                    EmailObject.Priority = MailPriority.Normal;
                    EmailObject.Subject = "FreeTimeShipping Certificate: " + dtv.Rows[0]["VesVoy"].ToString();
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

                    #endregion
                }
            }
            catch (Exception ex)
            {
                ViewList.Add(new MyOnBoard
                {
                    AlertMessage = ex.Message

                });
            }
            ViewList.Add(new MyOnBoard
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;

        }



        public List<MyOnBoard> VesselEarlyArrivalCertificationSendingEmail(MyOnBoard Data)
        {

            List<MyOnBoard> ViewList = new List<MyOnBoard>();
            try
            {
                string[] Array = Data.Items.Split(new[] { "Insert:" }, StringSplitOptions.None);
                for (int i = 1; i < Array.Length; i++)
                {
                    var CharSplit = Array[i].ToString().TrimEnd(',').Split(',');
                    Data.BLID = CharSplit[0].ToString();
                    var EmailIDs = CharSplit[1].ToString();

                    #region ff
                    DataTable _dtCom = GetCompnayDetails();
                    DataTable dtv = getDelayCertificate(Data.BLID.ToString(), Data.CertificateType);

                    Document doc = new Document();
                    Rectangle rec = new Rectangle(700, 900);
                    doc = new Document(rec);
                    Paragraph para = new Paragraph();
                    MemoryStream memoryStream = new MemoryStream();
                    PdfWriter writer = PdfWriter.GetInstance(doc, memoryStream);

                    doc.Open();

                    PdfContentByte cb = writer.DirectContent;
                    cb.SetColorStroke(Color.BLACK);
                    int _Xp = 10, _Yp = 785, YDiff = 10;

                    BaseFont bfheader = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader, 14);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "", 300, 820, 0);

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(System.Web.Hosting.HostingEnvironment.MapPath("~/assets/agentlogo/BWSLOGO.jpg"));
                    png.SetAbsolutePosition(25, 780);     //logo fixed location
                    png.ScalePercent(15f);
                    doc.Add(png);

                    cb.BeginText();


                    BaseFont bfheader3 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader3, 20);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL EARLY ARRIVAL NOTICE", 220, 750, 0);//right

                    cb.MoveTo(220, 748);
                    cb.LineTo(560, 748);  //BORDER LINE     CODE


                    BaseFont bfheader7 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader7, 13);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "DATE  ", 70, 680, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO    ", 70, 660, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 680, 0);//right

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, ":", 120, 660, 0);//right


                    BaseFont bfheader4 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader4, 12);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL NAME/VOYAGE NO", 110, 235, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "REVISED ETA", 110, 215, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "YARD CLOSING", 110, 195, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SCN", 110, 175, 0);//right
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "VESSEL ID", 110, 155, 0);//right
                                                                                            //cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA INNSA", 120, 85, 0);//right


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 310, 235, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["RevisedETA"].ToString(), 310, 215, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["YardClosing"].ToString(), 310, 195, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VoyageNotes"].ToString(), 310, 175, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesselID"].ToString(), 310, 155, 0);





                    BaseFont bfheader5 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    cb.SetFontAndSize(bfheader5, 14);
                    cb.SetColorFill(Color.BLACK);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "SUBJECT : VESSEL ARRIVE EARLY  -  ", 70, 600, 0);//right

                    cb.MoveTo(70, 598);
                    cb.LineTo(560, 598);  //BORDER LINE     CODE

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "ETA MYPKG", 70, 580, 0);//right


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["Date"].ToString(), 130, 680, 0);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "TO VALUABLE CUSTOMER", 130, 660, 0);


                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["VesVoy"].ToString(), 330, 600, 0);

                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, dtv.Rows[0]["ETA"].ToString(), 175, 580, 0);



                    cb.MoveTo(70, 578);
                    cb.LineTo(300, 578);  //BORDER LINE     CODE

                    int ColumnRows = 550;
                    cb.SetFontAndSize(bfheader3, 9);
                    string[] subjetv = Regex.Split(dtv.Rows[0]["Subject"].ToString().Trim().ToUpper() + "\r", char.ConvertFromUtf32(13));
                    string[] Aaddsplit;

                    for (int x = 0; x < subjetv.Length; x++)
                    {
                        Aaddsplit = subjetv[x].Split('\n');

                        for (int k = 0; k < Aaddsplit.Length; k++)
                        {

                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Aaddsplit[k].ToString(), 70, ColumnRows, 0);
                            ColumnRows -= 14;
                        }
                    }


                    //Line

                    cb.MoveTo(100, 250);
                    cb.LineTo(500, 250);  //

                    cb.MoveTo(100, 230);
                    cb.LineTo(500, 230);  //

                    cb.MoveTo(100, 210);
                    cb.LineTo(500, 210);

                    cb.MoveTo(100, 190);
                    cb.LineTo(500, 190);

                    cb.MoveTo(100, 170);
                    cb.LineTo(500, 170);

                    cb.MoveTo(100, 150);
                    cb.LineTo(500, 150);

                    //cb.MoveTo(100, 80);
                    //cb.LineTo(500, 80);  

                    //cb.MoveTo(100, 60);
                    //cb.LineTo(500, 60);

                    //vertical

                    cb.MoveTo(500, 250);
                    cb.LineTo(500, 150);

                    cb.MoveTo(300, 250);
                    cb.LineTo(300, 150);


                    cb.MoveTo(100, 250);
                    cb.LineTo(100, 150);

                    //cb.MoveTo(500, 660);
                    //cb.LineTo(500, 540);





                    cb.Stroke();
                    cb.EndText();

                    writer.CloseStream = false;
                    doc.Close();


                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();



                    MailMessage EmailObject = new MailMessage();

                    EmailObject.From = new MailAddress(_dtCom.Rows[0]["EmailID"].ToString(), _dtCom.Rows[0]["EmailHeader"].ToString());
                    var EmailIDTo = EmailIDs.Split(';');

                    for (int y = 0; y < EmailIDTo.Length; y++)
                    {
                        if (EmailIDTo[y].ToString() != "")
                        {
                            EmailObject.To.Add(new MailAddress(EmailIDTo[y].ToString()));
                        }
                    }


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
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Greetings from " + _dtCom.Rows[0]["EmailHeader"].ToString() + "</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>This is system generated email to inform Vessel Early Arrival Notice Certificate</td></tr>";
                    sHtml += "</table>";
                    sHtml += "<br />";
                    sHtml += "<table border='1' cellpadding='0' cellspacing='0' width='75%'>";
                    sHtml += "<tr>";
                    sHtml += "<td colspan='4' style='font-family:Tw Cen MT Condensed; font-size:26px; font-weight:bold; text-align:center; background-color:#007999; color:#fff; border-right:0px solid #007999; border-left:0px solid #007999; border-top: 0px solid #007999; border-bottom: 0px solid #007999;'  colspan='8' >VESSEL EARLY ARRIVAL NOTICE</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader1 + _dtCom.Rows[0]["EmailHeader"].ToString() + " </ td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + " Booking Party</td>";
                    sHtml += strSubHeader + " Booking No</td>";
                    sHtml += strSubHeader + " Vessel/Voyage</td>";
                    sHtml += strSubHeader + " Port of Loading</td>";
                    sHtml += "</tr>";
                    sHtml += "<tr>";
                    sHtml += strSubHeader + dtv.Rows[0]["BkgParty"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["BLNumber"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["VesVoy"].ToString() + "</td>";
                    sHtml += strSubHeader + dtv.Rows[0]["POL"].ToString() + "</td>";
                    sHtml += "</tr>";
                    sHtml += "</table>";

                    sHtml += "</table>";
                    sHtml += "<br/><br/><tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>Do not reply on the auto mail.</td></tr>";
                    sHtml += "<br/><tr><td style='font-family:Arial; font-size:15px; font-style:italic'>Thank you and Regards,</td></tr>";
                    sHtml += "<tr><td style='font-family:Arial; font-size:15px; font-weight:bold; font-style:italic;'>" + _dtCom.Rows[0]["EmailHeader"].ToString() + ".</td></tr>";

                    sHtml += "</table>";
                    EmailObject.Attachments.Add(new Attachment(new MemoryStream(bytes), "VesselEarlyArrivalNoticeCerifcation" + ".pdf"));
                    EmailObject.Body = sHtml;
                    EmailObject.IsBodyHtml = true;
                    EmailObject.Priority = MailPriority.Normal;
                    EmailObject.Subject = "VesselEarlyArrivalNoticeCertifcate : " + dtv.Rows[0]["VesVoy"].ToString();
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

                    #endregion
                }
            }
            catch (Exception ex)
            {
                ViewList.Add(new MyOnBoard
                {
                    AlertMessage = ex.Message

                });
            }
            ViewList.Add(new MyOnBoard
            {
                AlertMessage = "Email sent successfully"

            });
            return ViewList;

        }











        public DataTable GetdelayData(MyOnBoard Data)
        {
            string _Query = " Select BLID,(SELECT TOP(1) BLNumber FROM nvo_BOL where Id =NVO_DelayCertificate.BLID) as BLNumber," +
                            " (select (select top(1) PortName from NVO_PortMaster where NVO_PortMaster.Id =nvo_BOL.POLID ) from  nvo_BOL where nvo_BOL.ID=BLID) as POL," +
                            " CertificateType, CertificateTitle, VSLVoyage, Subject, convert(varchar,Date,106) as Date  from NVO_DelayCertificate where BLID=" + Data.BLID;
            return Manag.GetViewData(_Query, "");
        }

      
        public DataTable getDelayCertificate(string ID, string CertificateType)
        {
            string _Query = " Select NVO_Booking.ID as BkgID,NVO_Booking.VesVoy,BookingNo as BLNumber,AgentID as AgencyID, CertificateType,CertificateTitle,VSLVoyage, Subject, convert(varchar,Date,106) as Date ,BkgParty,POL,POD, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 282 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as VoyageNotes, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 333 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as YardClosing, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 335 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as NextPortETA, " +
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 340 and NVO_VoyageNotesDtls.VoyageID=NVO_Booking.VesVoyID) as RevisedETA, " +
                            " (select  top(1) convert(varchar, BLDate, 103) as BLDate from NVO_BLRelease where NVO_BLRelease.BkgID = NVO_Booking.ID) as Bldate1, " +
                            " (select (select top(1) VesselID from NVO_VesselMaster where NVO_VesselMaster.ID=NVO_Voyage.VesselID ) from NVO_Voyage where NVO_Voyage.Id=NVO_Booking.VesVoyID) as VesselID, " +
                            " (select  top(1) convert(varchar,ETA, 103) as ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID) as ETA, " +
                            " (select  top(1) convert(varchar,ETA, 103) as ETA from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID and NVO_VoyageRoute.PortID=NVO_Booking.PODID) as ETA1, " +
                            " (select top(1) (select top(1) PortName from NVO_PortMainMaster where NVO_PortMainMaster.ID =NVO_VoyageRoute.PortID)   from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by RID asc) as Port1, " +
                            " (select top(1)(select top(1) PortName from NVO_PortMainMaster where NVO_PortMainMaster.ID = NVO_VoyageRoute.PortID)  from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_Booking.VesVoyID order by RID desc) as Port2 " +
                            " from NVO_DelayCertificate " +
                            " inner join NVO_Booking on NVO_Booking.VesVoyID = NVO_DelayCertificate.VSLVoyage where NVO_Booking.Id=" + ID;
            if (CertificateType == "5")
                _Query += " and CertificateType= 5";  //SHIPPING CERTIFICATE
            if (CertificateType == "9")
                _Query += " and CertificateType= 9";  //DELAY CERTIFICATE
            if (CertificateType == "10")
                _Query += " and CertificateType= 10";  //FREE TIME SHIPPINGCERTIFICATE
            if (CertificateType == "8")
                _Query += " and CertificateType= 8";  //VESSEL EARLY ARRIVAL NOTICE CERTIFICATE


            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetCompnayDetails()
        {
            string _Query = "select * from NVO_NewCompnayDetails";
            return Manag.GetViewData(_Query, "");
        }


      

    }

   
}