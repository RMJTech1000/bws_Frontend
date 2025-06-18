using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DataManager;
using System.ComponentModel;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Web.UI;

namespace NVOCShipping.Controllers
{
    public class MNRPDFController : Controller
    {
        MNRManager Manag = new MNRManager();
        // GET: MNRPDF
        public ActionResult Index()
        {
            return View();

        }
        public ActionResult MNRPDF(string ReqID, string IsforAppvd)
        {
            //if (Session["GeoLocID"] != null)
            //    geolocid = Session["GeoLocID"].ToString();

            //if (Session["vGeoLocID"] != null)
            //    geolocid = Session["vGeoLocID"].ToString();

            DataSet DS = new DataSet();
            string[,] qr = new string[2, 2];
            qr[0, 0] = "SELECT Distinct (SELECT TOP 1 CONVERT(VARCHAR, DtMovement, 103) + ' ' + CONVERT(VARCHAR, DtMovement, 108) FROM NVO_Containertxns where  statusCode IN('MA')  and " +
          " Containerid = CRR.CntrID ORDER BY DtMovement DESC) Gate_In_Dt,CONVERT(VARCHAR, C.DtManufacture, 103) DtMFD, (SELECT TOP 1 CntrNo FROM NVO_Containers where  ID = CRR.CntrID ) AS CntrNo," +
         " NVO_tblCntrTypes.type,NVO_tblCntrTypes.Size,Upper(NVO_tblCntrTypes.ISOCode) ISOCode, CRR.ReqRefNo,CONVERT(Varchar, CRR.DtRequest, 103) DtReq, isnull(tVendor.CustomerName ,'') as Vendor, " +
        " isnull ((SELECT TOP 1 Address FROM NVO_CusBranchLocation where  CustomerID = tVendor.ID  ) ,'') AS Address," +
        "  isnull((SELECT TOP 1 City FROM NVO_CusBranchLocation where  CustomerID = tVendor.ID  ) ,'') AS City, " +
        " isnull((SELECT TOP 1 TelNo FROM NVO_CusBranchLocation where  CustomerID = tVendor.ID  ),'') AS TelNo, " +
         " isnull ((SELECT TOP 1 Fax FROM NVO_CusBranchLocation where  CustomerID = tVendor.ID  ),'') AS Fax, " +
         " isnull( (SELECT TOP 1 CustomerName FROM NVO_CustomerMaster where ID = CRR.SurveyorID   ),'')  AS SurveyorName, " +
          " isnull( (SELECT TOP 1 CountryName FROM NVO_CountryMaster where ID = tVendor.CountryID ),''  ) AS Country, " +
              " CRR.ApprovalNo,CRR.CostPerManHrx100,CRR.DtRepaired,CRR.DtApproved,CRR.SurveyRef,CRR.DtSvrReq,CRR.DtSvrComplete, " +
   " NVO_CurrencyMaster.CurrencyCode,isnull((SELECT top 1 UserName from NVO_UserDetails WHERE ID = CRR.ApprovedByID),'')  as ApprovedBy FROM NVO_MNRCntrRepairReq CRR " +
      "  INNER JOIN NVO_Containers C ON C.ID = CRR.CntrID INNER JOIN NVO_tblCntrTypes ON NVO_tblCntrTypes.ID = C.TypeID " +
     "   left outer JOIN NVO_CustomerMaster tVendor ON tVendor.ID = Crr.VendorID AND tVendor.CustomerType = 44 " +
     " INNER JOIN NVO_CurrencyMaster on NVO_CurrencyMaster.Id = CRR.CurrID  " +
                   " WHERE CRR.ID=" + ReqID;

            qr[0, 1] = "RepairDtls";
            if (IsforAppvd != null)
            {
                qr[1, 0] = "  SELECT Distinct (SELECT TOP 1 RepairCode FROM NVO_MNRRepairMaster where  ID = DT.RepairTypeID) AS RepairCode, " +
                 "  (SELECT TOP 1 LocationCode FROM NVO_MNRLocationMaster where  ID = DT.LocCodeID ) AS LocationCode, " +
                 "  (SELECT TOP 1 DamageCode FROM NVO_MNRDamageMaster where  ID = DT.DamageTypeID ) AS DamageCode, " +
                  "  Measurement, MeasureUnit, (SELECT TOP 1 ComponentCode FROM NVO_MNRComponentMaster where  ID = DT.ComponentID ) AS ComponentCode, " +
                 " cast(round(MaterialCostx100 / 100.00 + TotalLabourCostx100 / 100.00, 2, 0) as decimal(18, 2)) " +
                 " as EstTotalCost,LabourHrs,AppDescription, (SELECT TOP 1 GeneralName FROM NVO_GeneralMaster where  ID = DT.CostToID AND SeqNo = 38  ) AS CostTO, cast(round(LabourCostx100 / 100.00, 2, 0) as decimal(18, 2)) as TotalHr, " +
                  " cast(round(TotalLabourCostx100 / 100.00, 2, 0) as decimal(18, 2)) as LabourRate,cast(round(MaterialCostx100 / 100.00, 2, 0) as decimal(18, 2)) as MatCost, " +
                 " DT.Qty,DT.Description  FROM NVO_MNRCntrRepairReqDtls DT " +
                  "  WHERE DT.RepairReqID = " + ReqID;
            }
            else
            {
                //Approved Report
                qr[1, 0] = "        SELECT Distinct (SELECT TOP 1 RepairCode FROM NVO_MNRRepairMaster where  ID = DT.RepairTypeID) AS RepairCode, " +
                   "  (SELECT TOP 1 LocationCode FROM NVO_MNRLocationMaster where  ID = DT.LocCodeID ) AS LocationCode, " +
                   " (SELECT TOP 1 DamageCode FROM NVO_MNRDamageMaster where  ID = DT.DamageTypeID ) AS DamageCode," +
                 " Measurement, MeasureUnit,  (SELECT TOP 1 ComponentCode FROM NVO_MNRComponentMaster where  ID = DT.ComponentID ) AS ComponentCode, " +
                   " cast(round(AppvdMaterialCostx100 / 100.00 + AppvdTotalLabCostx100 / 100.00, 2, 0) as decimal(18, 2)) " +
                  " as EstTotalCost,LabourHrs,AppDescription, (SELECT TOP 1 GeneralName FROM NVO_GeneralMaster where  ID = DT.CostToID AND SeqNo = 38  ) AS CostTO, cast(round(AppvdLabCostx100 / 100.00, 2, 0) as decimal(18, 2)) as TotalHr, " +
                   " cast(round(AppvdTotalLabCostx100 / 100.00, 2, 0) as decimal(18, 2)) as LabourRate,cast(round(AppvdMaterialCostx100 / 100.00, 2, 0) as decimal(18, 2)) as MatCost, " +
                "  DT.Qty,DT.Description FROM NVO_MNRCntrRepairReqDtls DT  " +
                 "  WHERE DT.RepairReqID = " + ReqID;
            }
            qr[1, 1] = "CostDtls";

            DS = Manag.GetData(qr, 2);

            //if (DS.Tables["RepairDtls"].Rows.Count == 0 || DS.Tables["CostDtls"].Rows.Count == 0)
            //{
            //    AlertBox("No Record exist.");
            //    return;
            //}

            MNREORPdf(DS, IsforAppvd);
            return View();

        }
        protected void MNREORPdf(DataSet DS, string IsforAppvd)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {

                    Document pdfDoc = new Document(PageSize.A4, 25, 10, 25, 10);

                    Paragraph para = new Paragraph();
                    PdfPCell cell = new PdfPCell();

                    PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
                    pdfDoc.Open();
                    pdfDoc.NewPage();

                    PdfContentByte content = pdfWriter.DirectContent;
                    Font fnt = new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK);

                    iTextSharp.text.Table tbllogo = new iTextSharp.text.Table(1);
                    tbllogo.Width = 60;
                    tbllogo.Alignment = Element.ALIGN_LEFT;
                    //tbllogo.Cellpadding = 1;
                    tbllogo.BorderWidth = 0;
                    Cell cell1 = new Cell();
                    cell1.Width = 10;

                    var img = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/BWSLOGO.png"));
                    img.Alignment = Element.ALIGN_LEFT;
                    cell1 = new Cell(img);
                    cell1.BorderWidth = 0;
                    cell1.Colspan = 1;
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell1.VerticalAlignment = Element.ALIGN_LEFT;
                    tbllogo.AddCell(cell1);


                    pdfDoc.Add(tbllogo);

                    //iTextSharp.text.Table tbllogo1 = new iTextSharp.text.Table(1);
                    //tbllogo1.Width = 60;
                    //tbllogo1.Alignment = Element.ALIGN_LEFT;
                    ////tbllogo.Cellpadding = 1;
                    //tbllogo1.BorderWidth = 0;
                    //Cell cell2 = new Cell();
                    //cell1.Width = 10;

                    //var img1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~/assets/img/neridaheader.png"));
                    //img.Alignment = Element.ALIGN_LEFT;
                    //cell2 = new Cell(img1);
                    //cell2.BorderWidth = 0;
                    //cell2.Colspan = 1;
                    //cell2.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell2.VerticalAlignment = Element.ALIGN_LEFT;
                    //tbllogo1.AddCell(cell2);


                    //pdfDoc.Add(tbllogo1);

                    Paragraph Text;

                    //Text = new Paragraph(DS.Tables["RepairDtls"].Rows[0]["Vendor"].ToString(), fnt);
                    //Text.Alignment = Element.ALIGN_CENTER;
                    //pdfDoc.Add(Text);

                    //Text = new Paragraph(DS.Tables["RepairDtls"].Rows[0]["Address"].ToString(), fnt);
                    //Text.Alignment = Element.ALIGN_CENTER;
                    //pdfDoc.Add(Text);

                    //Text = new Paragraph(DS.Tables["RepairDtls"].Rows[0]["CtryCode"].ToString(), fnt);
                    //Text.Alignment = Element.ALIGN_CENTER;
                    //pdfDoc.Add(Text);

                    fnt = new Font(Font.HELVETICA, 12, Font.BOLD, Color.BLACK);

                    if (IsforAppvd != null)
                    {
                        Text = new Paragraph("ESTIMATE OF REPAIR", fnt);
                    }
                    else
                    {
                        Text = new Paragraph("APPROVED REPAIR", fnt);
                    }

                    Text.Alignment = Element.ALIGN_RIGHT;
                    pdfDoc.Add(Text);

                    fnt = new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK);

                    Text = new Paragraph(" ", fnt);
                    Text.Alignment = Element.ALIGN_CENTER;
                    Text.SpacingAfter = 20;
                    pdfDoc.Add(Text);

                    Paragraph Text2;
                    fnt = new Font(Font.HELVETICA, 10, Font.BOLD | Font.UNDERLINE, Color.BLACK);
                    Text2 = new Paragraph("EOR Details", fnt);
                    Text2.Alignment = Element.ALIGN_CENTER;
                    Text2.SpacingAfter = 20;
                    pdfDoc.Add(Text2);

                    //DataTable dtRepairreq = GetRepairReq(Request.QueryString["CRRID"].ToString());

                    #region 6 Columns per table

                    PdfPTable TblHead = new PdfPTable(6);
                    TblHead.WidthPercentage = 100;
                    TblHead.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    TblHead.DefaultCell.BorderWidth = 0;

                    //  DataTable dtLoc = GetLocDesc(geolocid);

                    cell = new PdfPCell(new Phrase("GeoLocation :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead.AddCell(cell);


                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Depot :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["Vendor"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead.AddCell(cell);


                    cell = new PdfPCell(new Phrase("Req Ref# :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;

                    TblHead.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["ReqRefNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead.AddCell(cell);
                    TblHead.SpacingAfter = 10;
                    pdfDoc.Add(TblHead);

                    #endregion


                    #region 6 Columns per table

                    PdfPTable TblHead1 = new PdfPTable(6);
                    TblHead1.WidthPercentage = 100;
                    TblHead1.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    TblHead1.DefaultCell.BorderWidth = 0;


                    cell = new PdfPCell(new Phrase("Status:", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead1.AddCell(cell);

                    if (DS.Tables["RepairDtls"].Rows[0]["ApprovalNo"].ToString() != "")
                    {

                        cell = new PdfPCell(new Phrase("APPROVED ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    }
                    if (DS.Tables["RepairDtls"].Rows[0]["ApprovalNo"].ToString() == "")
                    {
                        cell = new PdfPCell(new Phrase("PENDING ", new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    }
                    cell.BorderWidth = 0;
                    TblHead1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Cntr No :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead1.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["CntrNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Type-Size :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead1.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["type"].ToString() + " - " + DS.Tables["RepairDtls"].Rows[0]["Size"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead1.AddCell(cell);
                    TblHead1.SpacingAfter = 20;
                    pdfDoc.Add(TblHead1);

                    #endregion

                    #region 6 Columns per table

                    PdfPTable TblHead2 = new PdfPTable(6);
                    TblHead2.WidthPercentage = 100;
                    TblHead2.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    TblHead2.DefaultCell.BorderWidth = 0;


                    cell = new PdfPCell(new Phrase("Mty in Date:", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead2.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["Gate_In_Dt"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Request Date :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead2.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["DtReq"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Cost/ManHr.* :", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead2.AddCell(cell);



                    float outF = 0.00f;
                    float.TryParse(DS.Tables["RepairDtls"].Rows[0]["CostPerManHrx100"].ToString(), out outF);

                    cell = new PdfPCell(new Phrase((outF / 100.00).ToString("##0.00"), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    TblHead2.AddCell(cell);
                    TblHead2.SpacingAfter = 20;
                    pdfDoc.Add(TblHead2);


                    #endregion

                    #region 2 Columns per table

                    PdfPTable TblHead3 = new PdfPTable(2);
                    TblHead3.WidthPercentage = 35;
                    TblHead3.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    TblHead3.DefaultCell.BorderWidth = 0;


                    cell = new PdfPCell(new Phrase("Currency:", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    TblHead3.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["CurrencyCode"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 0;
                    cell.Colspan = 1;
                    TblHead3.AddCell(cell);

                    TblHead3.SpacingAfter = 20;
                    pdfDoc.Add(TblHead3);

                    #endregion

                    #region cost details
                    // COST DETAILS HEADER
                    Paragraph Text1;
                    fnt = new Font(Font.HELVETICA, 10, Font.BOLD | Font.UNDERLINE, Color.BLACK);
                    Text1 = new Paragraph("Cost Details", fnt);
                    Text1.Alignment = Element.ALIGN_CENTER;
                    Text1.SpacingAfter = 20;
                    pdfDoc.Add(Text1);

                    #region 15 Columns TABLE HEADER
                    PdfPTable Tbl2 = new PdfPTable(13);
                    Tbl2.WidthPercentage = 100;
                    Tbl2.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    Tbl2.DefaultCell.BorderWidth = 0;
                    Tbl2.SetWidths(new int[] { 10, 8, 8, 10, 15, 10, 8, 17, 17, 17, 17, 6, 15, });



                    cell = new PdfPCell(new Phrase("Comp Cd", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Loc Cd", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Dm Cd", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Rep Type", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);


                    cell = new PdfPCell(new Phrase("Measurement W*H", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Unit", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Lab Hr", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Lab Cost", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Total Lab Cost", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Mat cost", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);

                    if (IsforAppvd != null)
                    {
                        cell = new PdfPCell(new Phrase("Est Total Cost", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl2.AddCell(cell);
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase("Approved Total Cost", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl2.AddCell(cell);
                    }

                    cell = new PdfPCell(new Phrase("Qty", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl2.AddCell(cell);


                    if (IsforAppvd != null)
                    {
                        cell = new PdfPCell(new Phrase("Description", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl2.AddCell(cell);
                    }
                    else
                    {

                        cell = new PdfPCell(new Phrase("Approver Description	", new Font(Font.HELVETICA, 7, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl2.AddCell(cell);
                    }
                    pdfDoc.Add(Tbl2);
                    #endregion



                    //#region   ROW VALUES



                    PdfPTable Tbl3 = new PdfPTable(13);
                    Tbl3.WidthPercentage = 100;
                    Tbl3.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    Tbl3.DefaultCell.BorderWidth = 0;
                    Tbl3.SetWidths(new int[] { 10, 8, 8, 10, 15, 10, 8, 17, 17, 17, 17, 6, 15, });

                    int s;
                    for (s = 0; s < DS.Tables["CostDtls"].Rows.Count; s++)
                    {


                        //Column1 component code
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["ComponentCode"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column2 loc code
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["LocationCode"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column3 damagecode
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["DamageCode"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column4 RepairCode
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["RepairCode"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);


                        //Column6 Measurement
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["Measurement"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column7 unit
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["MeasureUnit"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column8 lab hrs
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["LabourHrs"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column9 lab cost

                        string num = decimal.Parse(DS.Tables["CostDtls"].Rows[s]["TotalHr"].ToString()).ToString("0.00");
                        cell = new PdfPCell(new Phrase(num, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column10 total lab cost

                        string num1 = decimal.Parse(DS.Tables["CostDtls"].Rows[s]["LabourRate"].ToString()).ToString("0.00");
                        cell = new PdfPCell(new Phrase(num1, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        ////Column11 mat cost

                        string num2 = decimal.Parse(DS.Tables["CostDtls"].Rows[s]["MatCost"].ToString()).ToString("0.00");
                        cell = new PdfPCell(new Phrase(num2, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        ////Column12  est  total cost

                        string num3 = decimal.Parse(DS.Tables["CostDtls"].Rows[s]["EstTotalCost"].ToString()).ToString("0.00");
                        cell = new PdfPCell(new Phrase(num3, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column13 qty
                        cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["qty"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 1;
                        Tbl3.AddCell(cell);

                        //Column14 Appover //Description
                        if (IsforAppvd != null)
                        {

                            cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["Description"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                            cell.BorderWidth = 1;
                            Tbl3.AddCell(cell);
                        }
                        else
                        {
                            cell = new PdfPCell(new Phrase(DS.Tables["CostDtls"].Rows[s]["AppDescription"].ToString(), new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                            cell.BorderWidth = 1;
                            Tbl3.AddCell(cell);
                        }
                    }
                    pdfDoc.Add(Tbl3);

                    #endregion

                    #region TOTAL

                    PdfPTable Tbl4 = new PdfPTable(13);
                    Tbl4.WidthPercentage = 100;
                    Tbl4.HorizontalAlignment = Element.ALIGN_LEFT;
                    //Tbl1.DefaultCell.Padding = 1;
                    Tbl4.DefaultCell.BorderWidth = 0;
                    Tbl4.SetWidths(new int[] { 10, 8, 8, 10, 15, 10, 8, 17, 17, 17, 17, 6, 15, });


                    //Column1 component code
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column2 loc code
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column3 damagecode
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column4 RepairCode
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column6 Measurement
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column7 unit
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column8 lab hrs
                    string num4 = decimal.Parse(DS.Tables["CostDtls"].Compute("Sum(LabourHrs)", "").ToString()).ToString("0.00");
                    cell = new PdfPCell(new Phrase(num4, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column9 lab cost
                    string num5 = decimal.Parse(DS.Tables["CostDtls"].Compute("Sum(TotalHr)", "").ToString()).ToString("0.00");
                    cell = new PdfPCell(new Phrase(num5, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column10 total lab cost
                    string num6 = decimal.Parse(DS.Tables["CostDtls"].Compute("Sum(LabourRate)", "").ToString()).ToString("0.00");
                    cell = new PdfPCell(new Phrase(num6, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column11 mat cost
                    string num7 = decimal.Parse(DS.Tables["CostDtls"].Compute("Sum(MatCost)", "").ToString()).ToString("0.00");
                    cell = new PdfPCell(new Phrase(num7, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column12  est mat cost
                    string num8 = decimal.Parse(DS.Tables["CostDtls"].Compute("Sum(EstTotalCost)", "").ToString()).ToString("0.00");
                    cell = new PdfPCell(new Phrase(num8, new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);


                    //Column13 qty
                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    //Column14 Approver Description

                    cell = new PdfPCell(new Phrase("", new Font(Font.HELVETICA, 7, Font.NORMAL, Color.BLACK)));
                    cell.BorderWidth = 1;
                    Tbl4.AddCell(cell);

                    Tbl4.SpacingAfter = 20;
                    pdfDoc.Add(Tbl4);

                    #endregion


                    //  Approval / Rejection  Details

                    if (IsforAppvd == null)
                    {
                        Paragraph Text3;
                        fnt = new Font(Font.HELVETICA, 12, Font.BOLD | Font.UNDERLINE, Color.BLACK);
                        Text3 = new Paragraph("Approval / Rejection Details", fnt);
                        Text3.Alignment = Element.ALIGN_CENTER;
                        Text3.SpacingAfter = 20;
                        pdfDoc.Add(Text3);

                        #region 4 Columns per table

                        PdfPTable TblApprove = new PdfPTable(4);
                        TblApprove.WidthPercentage = 100;
                        TblApprove.HorizontalAlignment = Element.ALIGN_LEFT;
                        //Tbl1.DefaultCell.Padding = 1;
                        TblApprove.DefaultCell.BorderWidth = 0;


                        cell = new PdfPCell(new Phrase("Applied By Name", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["Vendor"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove.AddCell(cell);

                        cell = new PdfPCell(new Phrase("Approval By", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;

                        TblApprove.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["ApprovedBy"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove.AddCell(cell);

                        TblApprove.SpacingAfter = 10;
                        pdfDoc.Add(TblApprove);

                        #endregion

                        #region 4 Columns per table

                        PdfPTable TblApprove1 = new PdfPTable(4);
                        TblApprove1.WidthPercentage = 100;
                        TblApprove1.HorizontalAlignment = Element.ALIGN_LEFT;
                        //Tbl1.DefaultCell.Padding = 1;
                        TblApprove1.DefaultCell.BorderWidth = 0;


                        cell = new PdfPCell(new Phrase("Approval #", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove1.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["ApprovalNo"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove1.AddCell(cell);

                        cell = new PdfPCell(new Phrase("App.Dt", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;

                        TblApprove1.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["DtApproved"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove1.AddCell(cell);

                        TblApprove1.SpacingAfter = 10;
                        pdfDoc.Add(TblApprove1);

                        #endregion

                        #region 4 Columns per table

                        PdfPTable TblApprove2 = new PdfPTable(4);
                        TblApprove2.WidthPercentage = 100;
                        TblApprove2.HorizontalAlignment = Element.ALIGN_LEFT;
                        //Tbl1.DefaultCell.Padding = 1;
                        TblApprove2.DefaultCell.BorderWidth = 0;


                        cell = new PdfPCell(new Phrase("Repair Dt", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove2.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["DtRepaired"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove2.AddCell(cell);

                        cell = new PdfPCell(new Phrase("Surveyor", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;

                        TblApprove2.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["SurveyorName"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove2.AddCell(cell);

                        TblApprove2.SpacingAfter = 10;
                        pdfDoc.Add(TblApprove2);

                        #endregion

                        #region 4 Columns per table

                        PdfPTable TblApprove3 = new PdfPTable(4);
                        TblApprove3.WidthPercentage = 100;
                        TblApprove3.HorizontalAlignment = Element.ALIGN_LEFT;
                        //Tbl1.DefaultCell.Padding = 1;
                        TblApprove3.DefaultCell.BorderWidth = 0;


                        cell = new PdfPCell(new Phrase("Survey Complete Date", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove3.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["DtSvrComplete"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove3.AddCell(cell);

                        cell = new PdfPCell(new Phrase("Survey Request Date", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        cell.BorderWidth = 0;

                        TblApprove3.AddCell(cell);

                        cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["DtSvrReq"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        cell.BorderWidth = 0;
                        TblApprove3.AddCell(cell);

                        TblApprove3.SpacingAfter = 10;
                        pdfDoc.Add(TblApprove3);

                        #endregion
                        //---Surveyor ref//
                        //#region 2 Columns per table

                        //PdfPTable TblApprove4 = new PdfPTable(2);
                        //TblApprove4.WidthPercentage = 100;
                        //TblApprove4.HorizontalAlignment = Element.ALIGN_LEFT;
                        ////Tbl1.DefaultCell.Padding = 1;
                        //TblApprove4.DefaultCell.BorderWidth = 0;

                        //cell = new PdfPCell(new Phrase("Survey Ref", new Font(Font.HELVETICA, 8, Font.BOLD, Color.BLACK)));
                        //cell.BorderWidth = 0;
                        //TblApprove4.AddCell(cell);

                        //cell = new PdfPCell(new Phrase(DS.Tables["RepairDtls"].Rows[0]["SurveyRef"].ToString(), new Font(Font.HELVETICA, 8, Font.NORMAL, Color.BLACK)));
                        //cell.BorderWidth = 0;
                        //TblApprove4.AddCell(cell);


                        //TblApprove4.SpacingAfter = 10;
                        //pdfDoc.Add(TblApprove4);

                        //#endregion
                    }

                    pdfWriter.CloseStream = false;
                    pdfDoc.Close();
                    Response.Buffer = true;
                    Response.ContentType = "application/pdf";
                    //Response.AddHeader("content-disposition", "attachment;filename=EstimateOfRepair.pdf");
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    //Response.Write(pdfDoc);
                    Response.End();

                }

            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
            }
            return;

        }
    }
}