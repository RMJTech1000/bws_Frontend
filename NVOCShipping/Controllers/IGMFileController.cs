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
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using DataTier;
using System.IO.Compression;

namespace NVOCShipping.Controllers
{
    public class IGMFileController : Controller
    {
        // GET: IGMFile
        DocumentManager Manag = new DocumentManager();


        public ActionResult IGMCreatefile(string vesid, string ManifestID, string TerminalID)
        {
           
            try
            {
                DataTable _dt = GetImgCargoValues(vesid);
                if (_dt.Rows.Count > 0)
                {
                    // string directoryPath = Server.MapPath("~/XMLFile");
                    string folderName = "xml_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string directoryPath = Server.MapPath("~/XMLFile/" + folderName);
                    Directory.CreateDirectory(directoryPath);

                    for (int x = 0; x < _dt.Rows.Count; x++)
                    {
                        string filePath = Path.Combine(directoryPath, "IGMfile_" + x + ".xml");
                        // Create XML document
                        XmlDocument xmlDoc = new XmlDocument();


                        XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
                        xmlDoc.AppendChild(xmlDeclaration);



                        // Create root element
                        XmlElement root = xmlDoc.CreateElement("RootElement");
                        xmlDoc.AppendChild(root);

                        //XmlElement root = xmlDoc.CreateElement("RootElement");
                        //xmlDoc.InsertBefore(Headervalues, root);

                        // Create arrays elements
                        XmlElement array1Header = xmlDoc.CreateElement("Header");
                        XmlElement array2Body = xmlDoc.CreateElement("Body");

                        // Append arrays elements to root
                        root.AppendChild(array1Header);
                        root.AppendChild(array2Body);

                        XmlElement itemElement = xmlDoc.CreateElement("DocumentFormat");

                        XmlElement itemElementHDoc = xmlDoc.CreateElement("DocumentFormatIdentifier");
                        itemElementHDoc.InnerText = "TRADEDOC";
                        itemElement.AppendChild(itemElementHDoc);

                        XmlElement itemElementHCon = xmlDoc.CreateElement("ControllingAgency");
                        itemElementHCon.InnerText = "DNT";
                        itemElement.AppendChild(itemElementHCon);

                        array1Header.AppendChild(itemElement);

                        XmlElement DocumentType = xmlDoc.CreateElement("DocumentType");
                        XmlElement DocumentTypeCode = xmlDoc.CreateElement("DocumentTypeCode");
                        DocumentTypeCode.InnerText = "GIG";
                        DocumentType.AppendChild(DocumentTypeCode);
                        array1Header.AppendChild(DocumentType);

                        XmlElement DocumentIdentification = xmlDoc.CreateElement("DocumentIdentification");
                        XmlElement UniqueReferenceNo = xmlDoc.CreateElement("UniqueReferenceNo");
                        UniqueReferenceNo.InnerText = "SNRF123";
                        DocumentIdentification.AppendChild(UniqueReferenceNo);

                        XmlElement MessageFunction = xmlDoc.CreateElement("MessageFunction");
                        MessageFunction.InnerText = "9";
                        DocumentIdentification.AppendChild(MessageFunction);
                        array1Header.AppendChild(DocumentIdentification);

                        XmlElement InterchangeSender = xmlDoc.CreateElement("InterchangeSender");
                        XmlElement SenderID = xmlDoc.CreateElement("SenderID");
                        SenderID.InnerText = "info@bluewave.com.my";
                        InterchangeSender.AppendChild(SenderID);
                        array1Header.AppendChild(InterchangeSender);

                        XmlElement InterchangeRecipient = xmlDoc.CreateElement("InterchangeRecipient");
                        XmlElement RecipientID = xmlDoc.CreateElement("RecipientID");
                        RecipientID.InnerText = "info@bluewave.com.my";
                        InterchangeRecipient.AppendChild(RecipientID);
                        array1Header.AppendChild(InterchangeRecipient);


                        XmlElement DateorTimeofPreparation = xmlDoc.CreateElement("DateorTimeofPreparation");
                        XmlElement DateofPreparation = xmlDoc.CreateElement("DateofPreparation");
                        DateofPreparation.InnerText = _dt.Rows[0]["CurrDate"].ToString();
                        DateorTimeofPreparation.AppendChild(DateofPreparation);

                        XmlElement TimeofPreparation = xmlDoc.CreateElement("TimeofPreparation");
                        TimeofPreparation.InnerText =  _dt.Rows[0]["CurrTime"].ToString();
                        DateorTimeofPreparation.AppendChild(TimeofPreparation);
                        array1Header.AppendChild(DateorTimeofPreparation);



                        XmlElement BeginningofMessage = xmlDoc.CreateElement("BeginningofMessage");



                        XmlElement DocumentNo = xmlDoc.CreateElement("DocumentNo");
                        DocumentNo.InnerText = "1";
                        BeginningofMessage.AppendChild(DocumentNo);

                        XmlElement MessageCode = xmlDoc.CreateElement("MessageCode");
                        MessageCode.InnerText = "785";
                        BeginningofMessage.AppendChild(MessageCode);

                        XmlElement DocMessageJobNo = xmlDoc.CreateElement("DocMessageJobNo");
                        DocMessageJobNo.InnerText = _dt.Rows[x]["BLNumberv"].ToString();
                        BeginningofMessage.AppendChild(DocMessageJobNo);

                        XmlElement IssueDate = xmlDoc.CreateElement("IssueDate");
                        IssueDate.InnerText = _dt.Rows[x]["IssuesDate"].ToString();
                        BeginningofMessage.AppendChild(IssueDate);
                        array2Body.AppendChild(BeginningofMessage);

                        XmlElement ProcessingIndicator = xmlDoc.CreateElement("ProcessingIndicator");

                        XmlElement ProcessingIndicatorCode = xmlDoc.CreateElement("ProcessingIndicatorCode");
                        ProcessingIndicatorCode.InnerText = "23";
                        ProcessingIndicator.AppendChild(ProcessingIndicatorCode);

                        XmlElement ProcessingIndicatorValue = xmlDoc.CreateElement("ProcessingIndicatorValue");
                        ProcessingIndicatorValue.InnerText = "MNT";
                        ProcessingIndicator.AppendChild(ProcessingIndicatorValue);
                        array2Body.AppendChild(ProcessingIndicator);

                        DataTable dtc = GetBLCustomerValues(_dt.Rows[x]["BLID"].ToString());

                        for (int y = 0; y < dtc.Rows.Count; y++)
                        {

                            XmlElement Parties = xmlDoc.CreateElement("Parties");

                            XmlElement PartyType = xmlDoc.CreateElement("PartyType");
                            PartyType.InnerText = dtc.Rows[y]["DESP"].ToString();
                            Parties.AppendChild(PartyType);

                            XmlElement OrganizationIdentification = xmlDoc.CreateElement("OrganizationIdentification");
                            OrganizationIdentification.InnerText = dtc.Rows[y]["PartyCode"].ToString();
                            Parties.AppendChild(OrganizationIdentification);

                            XmlElement OrganizationName = xmlDoc.CreateElement("OrganizationName");
                            OrganizationName.InnerText = dtc.Rows[y]["PartyName"].ToString();
                            Parties.AppendChild(OrganizationName);

                            XmlElement AddressInformation = xmlDoc.CreateElement("AddressInformation");
                            //var ArrayAddress = SplitByLenght(dtc.Rows[y]["PartyAddress"].ToString(), 35);

                            XmlElement Text = xmlDoc.CreateElement("Text");
                            Text.InnerText = dtc.Rows[y]["PartyAddress"].ToString();
                            //AddressInformation.AppendChild(Text);
                            //var ArrayAddress = fn_WrapDescriptionWords(dtc.Rows[y]["PartyAddress"].ToString(), 55);
                             var ArrayAddress = dtc.Rows[y]["PartyAddress"].ToString().Split('\n');
                            for (int i = 0; i < ArrayAddress.Length; i++)
                            {
                                if (i == 3)
                                {
                                    break;
                                }
                                else
                                {
                                    XmlElement Text1 = xmlDoc.CreateElement("Text");
                                    Text1.InnerText = ArrayAddress[i].ToString().Trim();
                                    AddressInformation.AppendChild(Text1);
                                }

                                 
                            }

                            Parties.AppendChild(AddressInformation);
                            array2Body.AppendChild(Parties);
                        }


                        // for loop

                        DataTable dtp = GetPortCodeValues(_dt.Rows[x]["BLID"].ToString());
                        for (int k = 0; k < dtp.Rows.Count; k++)
                        {

                            XmlElement PlaceOrLocation = xmlDoc.CreateElement("PlaceOrLocation");
                            XmlElement PlaceOrLocCodeQualifier = xmlDoc.CreateElement("PlaceOrLocCodeQualifier");
                            PlaceOrLocCodeQualifier.InnerText = dtp.Rows[k]["HNo"].ToString();
                            PlaceOrLocation.AppendChild(PlaceOrLocCodeQualifier);
                            XmlElement PlaceOrLocCode = xmlDoc.CreateElement("PlaceOrLocCode");
                            PlaceOrLocCode.InnerText = dtp.Rows[k]["PortCode"].ToString();
                            PlaceOrLocation.AppendChild(PlaceOrLocCode);
                            array2Body.AppendChild(PlaceOrLocation);
                        }
                        //end for loop


                        // for loop


                        XmlElement DateTimeDetail = xmlDoc.CreateElement("DateTimeDetail");
                        XmlElement DateTimeFuncCode = xmlDoc.CreateElement("DateTimeFuncCode");
                        DateTimeFuncCode.InnerText = "182";
                        DateTimeDetail.AppendChild(DateTimeFuncCode);
                        XmlElement DateTimeValue = xmlDoc.CreateElement("DateTimeValue");
                        DateTimeValue.InnerText = _dt.Rows[x]["CurrDate"].ToString();
                        DateTimeDetail.AppendChild(DateTimeValue);
                        array2Body.AppendChild(DateTimeDetail);


                        XmlElement DateTimeDetail_ETA = xmlDoc.CreateElement("DateTimeDetail");
                        XmlElement DateTimeFuncCode_ETA = xmlDoc.CreateElement("DateTimeFuncCode");
                        DateTimeFuncCode_ETA.InnerText = "132";
                        DateTimeDetail_ETA.AppendChild(DateTimeFuncCode_ETA);
                        XmlElement DateTimeValue_ETA = xmlDoc.CreateElement("DateTimeValue");
                        DateTimeValue_ETA.InnerText = _dt.Rows[x]["ETA"].ToString();
                        DateTimeDetail_ETA.AppendChild(DateTimeValue_ETA);
                        array2Body.AppendChild(DateTimeDetail_ETA);


                     

                        //end for loop
                        DataTable dtz = GetPortCodeDocValues(_dt.Rows[x]["BLID"].ToString(), "BKMTO1");
                        for (int z = 0; z < dtz.Rows.Count; z++)
                        {
                            XmlElement DocumentReferenceInfo = xmlDoc.CreateElement("DocumentReferenceInfo");
                            XmlElement DocRefDetail = xmlDoc.CreateElement("DocRefDetail");
                            DocumentReferenceInfo.AppendChild(DocRefDetail);
                           
                            
                            XmlElement DocRefCode = xmlDoc.CreateElement("DocRefCode");
                            DocRefCode.InnerText = dtz.Rows[z]["desvalue"].ToString();
                            DocRefDetail.AppendChild(DocRefCode);


                            XmlElement DocRefNo = xmlDoc.CreateElement("DocRefNo");
                            DocRefNo.InnerText = dtz.Rows[z]["valuesv"].ToString();
                            DocRefDetail.AppendChild(DocRefNo);


                            array2Body.AppendChild(DocumentReferenceInfo);
                        }



                        XmlElement TransportDetails = xmlDoc.CreateElement("TransportDetails");

                        XmlElement TransportStageCodeQualifier = xmlDoc.CreateElement("TransportStageCodeQualifier");
                        TransportStageCodeQualifier.InnerText = "11";
                        TransportDetails.AppendChild(TransportStageCodeQualifier);

                        XmlElement ModeofTransportCode = xmlDoc.CreateElement("ModeofTransportCode");
                        ModeofTransportCode.InnerText = "1";
                        TransportDetails.AppendChild(ModeofTransportCode);

                        //voyage No
                        XmlElement ConveyanceReferenceNo = xmlDoc.CreateElement("ConveyanceReferenceNo");
                        ConveyanceReferenceNo.InnerText = _dt.Rows[x]["ImportVoyageCd"].ToString();
                        TransportDetails.AppendChild(ConveyanceReferenceNo);


                        XmlElement TransportMeans = xmlDoc.CreateElement("TransportMeans");
                        TransportDetails.AppendChild(TransportMeans);

                        XmlElement TransportTypeID = xmlDoc.CreateElement("TransportTypeID");
                        //innerText
                        TransportMeans.AppendChild(TransportTypeID);

                        XmlElement TransportTypeName = xmlDoc.CreateElement("TransportTypeName");
                        //innerText
                        TransportMeans.AppendChild(TransportTypeName);


                        XmlElement TransportIdentification = xmlDoc.CreateElement("TransportIdentification");
                        TransportDetails.AppendChild(TransportIdentification);

                        XmlElement TransportTypeID_1 = xmlDoc.CreateElement("TransportID");
                        TransportTypeID_1.InnerText = _dt.Rows[x]["VesselCallSign"].ToString();
                        TransportIdentification.AppendChild(TransportTypeID_1);

                        XmlElement TransportMeansName = xmlDoc.CreateElement("TransportMeansName");
                        TransportMeansName.InnerText = _dt.Rows[x]["VesselName"].ToString();
                        TransportIdentification.AppendChild(TransportMeansName);

                        XmlElement Nationality = xmlDoc.CreateElement("Nationality");
                        Nationality.InnerText = _dt.Rows[x]["Nationality"].ToString();
                        TransportIdentification.AppendChild(Nationality);


                        array2Body.AppendChild(TransportDetails);


                        DataTable dtx = getCntrDetails(_dt.Rows[x]["BLID"].ToString());
                        for (int j = 0; j < dtx.Rows.Count; j++)
                        {

                            //forloop

                            XmlElement LineItem = xmlDoc.CreateElement("LineItem");
                            XmlElement LineItemNo = xmlDoc.CreateElement("LineItemNo");
                            LineItemNo.InnerText = "1";
                            LineItem.AppendChild(LineItemNo);

                            XmlElement ItemIdentifiers = xmlDoc.CreateElement("ItemIdentifiers");
                            LineItem.AppendChild(ItemIdentifiers);

                            XmlElement TariffCode = xmlDoc.CreateElement("TariffCode");
                            TariffCode.InnerText = "850220";
                            ItemIdentifiers.AppendChild(TariffCode);

                            XmlElement ProductDesc = xmlDoc.CreateElement("ProductDesc");
                            //ProductDesc.InnerText = dtx.Rows[j]["CargoDes"].ToString();
                            ProductDesc.InnerText = dtx.Rows[j]["ImpDesc"].ToString();
                            ItemIdentifiers.AppendChild(ProductDesc);

                            //loop

                            if (_dt.Rows[0]["BLCommodityTypeID"].ToString() == "4")
                            {
                                DataTable dtPC = GetProductCharValues();
                                for (int a = 0; a < dtPC.Rows.Count; a++)
                                {
                                    XmlElement ProductCharacteristics = xmlDoc.CreateElement("ProductCharacteristics");
                                    ItemIdentifiers.AppendChild(ProductCharacteristics);

                                    XmlElement ProductCharType = xmlDoc.CreateElement("ProductCharType");
                                    ProductCharType.InnerText = dtPC.Rows[a]["desvalue"].ToString();
                                    ProductCharacteristics.AppendChild(ProductCharType);

                                    XmlElement ProductCharValue = xmlDoc.CreateElement("ProductCharValue");
                                    ProductCharValue.InnerText = dtPC.Rows[a]["valuesv"].ToString();
                                    ProductCharacteristics.AppendChild(ProductCharValue);
                                }
                            }

                            //end

                            XmlElement PackageandWeightSummary = xmlDoc.CreateElement("PackageandWeightSummary");
                            LineItem.AppendChild(PackageandWeightSummary);

                            XmlElement TotalPackage = xmlDoc.CreateElement("TotalPackage");
                            PackageandWeightSummary.AppendChild(TotalPackage);

                            XmlElement NumberofPackages = xmlDoc.CreateElement("NumberofPackages");
                            NumberofPackages.InnerText = dtx.Rows[j]["NoOfPkg"].ToString();
                            TotalPackage.AppendChild(NumberofPackages);

                            XmlElement PackageTypeCode = xmlDoc.CreateElement("PackageTypeCode");
                            PackageTypeCode.InnerText = dtx.Rows[j]["PakgeName"].ToString();
                            TotalPackage.AppendChild(PackageTypeCode);


                            XmlElement Measurement_pac = xmlDoc.CreateElement("Measurement");
                            PackageandWeightSummary.AppendChild(Measurement_pac);

                            XmlElement MeasurementQualifier_pac = xmlDoc.CreateElement("MeasurementQualifier");
                            MeasurementQualifier_pac.InnerText = "WT";
                            Measurement_pac.AppendChild(MeasurementQualifier_pac);

                            XmlElement MeasurementDimensionCode = xmlDoc.CreateElement("MeasurementDimensionCode");
                            MeasurementDimensionCode.InnerText = "AAD";
                            Measurement_pac.AppendChild(MeasurementDimensionCode);

                            XmlElement MeasurementUnitCode_pac = xmlDoc.CreateElement("MeasurementUnitCode");
                            MeasurementUnitCode_pac.InnerText = "KGM";
                            Measurement_pac.AppendChild(MeasurementUnitCode_pac);

                            XmlElement Value_pac = xmlDoc.CreateElement("Value");
                            Value_pac.InnerText = dtx.Rows[j]["GrsWt"].ToString() + 0;
                            Measurement_pac.AppendChild(Value_pac);




                            //Measurment2


                            XmlElement Measurement_pac_2 = xmlDoc.CreateElement("Measurement");
                            PackageandWeightSummary.AppendChild(Measurement_pac_2);

                            XmlElement MeasurementQualifier_pac_2 = xmlDoc.CreateElement("MeasurementQualifier");
                            MeasurementQualifier_pac_2.InnerText = "WT";
                            Measurement_pac_2.AppendChild(MeasurementQualifier_pac_2);

                            XmlElement MeasurementDimensionCode_2 = xmlDoc.CreateElement("MeasurementDimensionCode");
                            MeasurementDimensionCode_2.InnerText = "MFT";
                            Measurement_pac_2.AppendChild(MeasurementDimensionCode_2);

                            XmlElement MeasurementUnitCode_pac_2 = xmlDoc.CreateElement("MeasurementUnitCode");
                            MeasurementUnitCode_pac_2.InnerText = "KGM";
                            Measurement_pac_2.AppendChild(MeasurementUnitCode_pac_2);

                            XmlElement Value_pac_2 = xmlDoc.CreateElement("Value");
                            Value_pac_2.InnerText = dtx.Rows[j]["GrsWt"].ToString() + 0;
                            Measurement_pac_2.AppendChild(Value_pac_2);

                            //end


                            //Measurment3


                            XmlElement Measurement_pac_3 = xmlDoc.CreateElement("Measurement");
                            PackageandWeightSummary.AppendChild(Measurement_pac_3);

                            XmlElement MeasurementQualifier_pac_3 = xmlDoc.CreateElement("MeasurementQualifier");
                            MeasurementQualifier_pac_3.InnerText = "VOL";
                            Measurement_pac_3.AppendChild(MeasurementQualifier_pac_3);

                            XmlElement MeasurementDimensionCode_3 = xmlDoc.CreateElement("MeasurementDimensionCode");
                            MeasurementDimensionCode_3.InnerText = "BRH";
                            Measurement_pac_3.AppendChild(MeasurementDimensionCode_3);

                            XmlElement MeasurementUnitCode_pac_3 = xmlDoc.CreateElement("MeasurementUnitCode");
                            MeasurementUnitCode_pac_3.InnerText = "MTQ";
                            Measurement_pac_3.AppendChild(MeasurementUnitCode_pac_3);

                            XmlElement Value_pac_3 = xmlDoc.CreateElement("Value");
                            Value_pac_3.InnerText = dtx.Rows[j]["GrsWt"].ToString() + 0;
                            Measurement_pac_3.AppendChild(Value_pac_3);

                            //end


                            XmlElement AdditionalInformation = xmlDoc.CreateElement("AdditionalInformation");
                            LineItem.AppendChild(AdditionalInformation);


                            XmlElement InformationType = xmlDoc.CreateElement("InformationType");
                            InformationType.InnerText = "PAC";
                            AdditionalInformation.AppendChild(InformationType);

                            XmlElement Text_add_info = xmlDoc.CreateElement("Text");
                            Text_add_info.InnerText =  _dt.Rows[0]["MarkNo"].ToString();
                            AdditionalInformation.AppendChild(Text_add_info);

                            XmlElement Container = xmlDoc.CreateElement("Container");
                            LineItem.AppendChild(Container);

                            XmlElement ContainerIdentifier = xmlDoc.CreateElement("ContainerIdentifier");
                            ContainerIdentifier.InnerText = dtx.Rows[j]["CntrNo"].ToString();
                            Container.AppendChild(ContainerIdentifier);

                            array2Body.AppendChild(LineItem);


                        }

                        for (int y = 0; y < dtx.Rows.Count; y++)
                        {

                            XmlElement Container_list = xmlDoc.CreateElement("Container");

                            XmlElement ContainerItemNo = xmlDoc.CreateElement("ContainerItemNo");
                            ContainerItemNo.InnerText = "1";
                            Container_list.AppendChild(ContainerItemNo);

                            XmlElement ContainerIdentifier_cntr = xmlDoc.CreateElement("ContainerIdentifier");
                            ContainerIdentifier_cntr.InnerText = dtx.Rows[y]["CntrNo"].ToString();
                            Container_list.AppendChild(ContainerIdentifier_cntr);

                            XmlElement ContainerSizeandCodes = xmlDoc.CreateElement("ContainerSizeandCodes");
                            Container_list.AppendChild(ContainerSizeandCodes);

                            XmlElement ContainerSizeandDetailedTypeCode = xmlDoc.CreateElement("ContainerSizeandDetailedTypeCode");
                            ContainerSizeandDetailedTypeCode.InnerText = dtx.Rows[y]["Size"].ToString();
                            ContainerSizeandCodes.AppendChild(ContainerSizeandDetailedTypeCode);

                            XmlElement EquipmentStatus = xmlDoc.CreateElement("EquipmentStatus");
                            EquipmentStatus.InnerText = "3";
                            Container_list.AppendChild(EquipmentStatus);

                            XmlElement FullEmptyIndicator = xmlDoc.CreateElement("FullEmptyIndicator");
                            FullEmptyIndicator.InnerText = "1";
                            Container_list.AppendChild(FullEmptyIndicator);

                            XmlElement Seal = xmlDoc.CreateElement("Seal");
                            Seal.InnerText = dtx.Rows[y]["SealNo"].ToString();
                            Container_list.AppendChild(Seal);

                            XmlElement SealingParty = xmlDoc.CreateElement("SealingParty");
                            SealingParty.InnerText = "CA";
                            Container_list.AppendChild(SealingParty);

                            // Mar-09-2025

                            XmlElement MeasurementQualifier = xmlDoc.CreateElement("Measurement");
                            Container_list.AppendChild(MeasurementQualifier);

                            XmlElement Measurement = xmlDoc.CreateElement("MeasurementQualifier");
                            Measurement.InnerText = "TE";
                            MeasurementQualifier.AppendChild(Measurement);

                            XmlElement MeasurementUnitcode = xmlDoc.CreateElement("MeasurementUnitCode");
                            MeasurementUnitcode.InnerText = "FAH";
                            MeasurementQualifier.AppendChild(MeasurementUnitcode);

                            XmlElement MeasurementUnitgrowswt = xmlDoc.CreateElement("Value");
                            MeasurementUnitgrowswt.InnerText = dtx.Rows[y]["Grswt"].ToString() + 0;
                            MeasurementQualifier.AppendChild(MeasurementUnitgrowswt);



                            //end





                            array2Body.AppendChild(Container_list);
                        }

                        XmlElement GrandTotal = xmlDoc.CreateElement("GrandTotal");


                        DataTable dtT = getCntrDetailsTotal(_dt.Rows[x]["BLID"].ToString());



                        XmlElement TotalMeasurement = xmlDoc.CreateElement("TotalMeasurement");
                        GrandTotal.AppendChild(TotalMeasurement);

                        XmlElement MeasurementQualifier_gt = xmlDoc.CreateElement("MeasurementQualifier");
                        MeasurementQualifier_gt.InnerText = "AAE";
                        TotalMeasurement.AppendChild(MeasurementQualifier_gt);

                        XmlElement MeasurementDimensionCode_gt = xmlDoc.CreateElement("MeasurementDimensionCode");
                        MeasurementDimensionCode_gt.InnerText = "AAD";
                        TotalMeasurement.AppendChild(MeasurementDimensionCode_gt);

                        XmlElement MeasurementUnitCode_gt = xmlDoc.CreateElement("MeasurementUnitCode");
                        MeasurementUnitCode_gt.InnerText = "KGM";
                        TotalMeasurement.AppendChild(MeasurementUnitCode_gt);

                        XmlElement Value_gt = xmlDoc.CreateElement("Value");
                        Value_gt.InnerText = dtT.Rows[0]["GRsWtTotal"].ToString() + 0;
                        TotalMeasurement.AppendChild(Value_gt);


                        XmlElement TotalMeasurement1 = xmlDoc.CreateElement("TotalMeasurement");
                        GrandTotal.AppendChild(TotalMeasurement1);

                        XmlElement MeasurementQualifier_gt1 = xmlDoc.CreateElement("MeasurementQualifier");
                        MeasurementQualifier_gt1.InnerText = "AAE";
                        TotalMeasurement1.AppendChild(MeasurementQualifier_gt1);

                        XmlElement MeasurementUnitCode_gt1 = xmlDoc.CreateElement("MeasurementUnitCode");
                        MeasurementUnitCode_gt1.InnerText = "MTQ";
                        TotalMeasurement1.AppendChild(MeasurementUnitCode_gt1);

                        XmlElement Value_gt1 = xmlDoc.CreateElement("Value");
                        Value_gt1.InnerText = dtT.Rows[0]["GRsWtTotal"].ToString() + 0;
                        TotalMeasurement1.AppendChild(Value_gt1);



                        XmlElement TotalMeasurement2 = xmlDoc.CreateElement("TotalMeasurement");
                        GrandTotal.AppendChild(TotalMeasurement2);

                        XmlElement MeasurementQualifier_gt2 = xmlDoc.CreateElement("MeasurementQualifier");
                        MeasurementQualifier_gt2.InnerText = "11";
                        TotalMeasurement2.AppendChild(MeasurementQualifier_gt2);

                        XmlElement MeasurementUnitCode_gt2 = xmlDoc.CreateElement("MeasurementUnitCode");
                        MeasurementUnitCode_gt2.InnerText = "CT";
                        TotalMeasurement2.AppendChild(MeasurementUnitCode_gt2);

                        XmlElement Value_gt2 = xmlDoc.CreateElement("Value");
                        Value_gt2.InnerText = dtT.Rows[0]["TotalPkg"].ToString();
                        
                        TotalMeasurement2.AppendChild(Value_gt2);



                        XmlElement TotalMeasurement3 = xmlDoc.CreateElement("TotalMeasurement");
                        GrandTotal.AppendChild(TotalMeasurement3);

                        XmlElement MeasurementQualifier_gt3 = xmlDoc.CreateElement("MeasurementQualifier");

                        MeasurementQualifier_gt3.InnerText = dtT.Rows[0]["NoofCntr"].ToString(); ; //no of cntr
                        TotalMeasurement3.AppendChild(MeasurementQualifier_gt3);

                        XmlElement MeasurementUnitCode_gt3 = xmlDoc.CreateElement("MeasurementUnitCode");
                        //MeasurementUnitCode_gt3.InnerText = "CT";
                        TotalMeasurement3.AppendChild(MeasurementUnitCode_gt3);

                        XmlElement Value_gt21 = xmlDoc.CreateElement("Value");
                        Value_gt21.InnerText = dtT.Rows[0]["NoofCntr"].ToString();
                        TotalMeasurement3.AppendChild(Value_gt21);




                        array2Body.AppendChild(GrandTotal);

                        XmlElement AdditionalInformation_last = xmlDoc.CreateElement("AdditionalInformation");

                        XmlElement InformationType_ad = xmlDoc.CreateElement("InformationType");
                        InformationType_ad.InnerText = "AAA";
                        AdditionalInformation_last.AppendChild(InformationType_ad);


                        XmlElement Text_ad = xmlDoc.CreateElement("Text");
                        Text_ad.InnerText = _dt.Rows[x]["ImpDesc"].ToString();
                        AdditionalInformation_last.AppendChild(Text_ad);


                        array2Body.AppendChild(AdditionalInformation_last);
                       
                        xmlDoc.Save(filePath);
                        string xmlString = xmlDoc.OuterXml;


                    }
                    //return Content("XML document saved successfully.");
                    string zipFilePath = Server.MapPath("~/XMLFile/" + folderName + ".zip");
                    ZipFile.CreateFromDirectory(directoryPath, zipFilePath);

                    byte[] fileBytes = System.IO.File.ReadAllBytes(zipFilePath);
                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Zip, folderName + ".zip");
                }
                else
                {
                    string errorScript = $"<script>alert('There is no record');</script>";
                    return Content(errorScript);
                }
            }

            catch (Exception ex)
            {
                string errorScript = $"<script>alert('Error saving XML document: {ex.Message}');</script>";
                return Content(errorScript);
               
            }
        }

       

        public void CreateCSIGMFile(string vesId, string ManifestID, string TerminalID)
        {
            StringWriter stringWriter = new StringWriter();
            StringBuilder sb = new StringBuilder();


            DataTable dtVoy = GetImgVesVoyValues(vesId, TerminalID);
            if (dtVoy.Rows.Count > 0)
            {
            }
           
        }


        public void CreateIGMFile(string vesId, string ManifestID, string TerminalID)
        {
            StringWriter stringWriter = new StringWriter();
            StringBuilder sb = new StringBuilder();
            //char quote = '"';
            int c1 = 29;
            char c = (char)c1;
            string DtIGM = "";

            DataTable dtVoy = GetImgVesVoyValues(vesId, TerminalID);
            if (dtVoy.Rows.Count> 0)
            {

                sb.Append("HREC" + c + "ZZ" + c + dtVoy.Rows[0]["LineCode"].ToString() + c + "ZZ" + c + dtVoy.Rows[0]["POL"].ToString() + c + "ICES1_5" + c + "T" + c + c + "SACHI01" + c + "123" + c + System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString("00") + System.DateTime.Now.Day.ToString() + c + System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString());
                sb.Append("\r\n");
                sb.Append("<manifest>");
                sb.Append("\r\n");
                sb.Append("<vesinfo>");
                sb.Append("\r\n");
                sb.Append("F");
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["POL"].ToString());
                sb.Append(c);
                sb.Append(c);
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["IMONumber"].ToString());
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["VesselCallSign"].ToString());
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["ImportVoyageCd"].ToString());
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["LineCode"].ToString());// panNo
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["PanNo"].ToString());
                sb.Append(c);
                sb.Append("CAPT" + dtVoy.Rows[0]["VesselName"].ToString());
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["FloppyCode"].ToString());
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["POL"].ToString());
                sb.Append(c);
                sb.Append(c);
                sb.Append(c);
                sb.Append("C");
                sb.Append(c);
                sb.Append("2");//ShipSDecl
                sb.Append(c);

                sb.Append("Y");//CrewLDecl
                sb.Append(c);
                sb.Append("CONTAINERISED");  //PList
                sb.Append(c);

                sb.Append(System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString("00") + System.DateTime.Now.Day.ToString() + c + System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString()); //CrewEDecl
                sb.Append(c);
                sb.Append(c);
                sb.Append("Y");
                sb.Append(c);
                sb.Append("Y");
                sb.Append(c);
                sb.Append("Y");
                sb.Append(c);
                sb.Append("N");
                sb.Append(c);
                sb.Append("Y");
                sb.Append(c);
                sb.Append("Y");
                sb.Append(c);
                sb.Append(dtVoy.Rows[0]["TerminalCode"].ToString());

                sb.Append("\r\n");
                sb.Append("<END-vesinfo>");
                sb.Append("\r\n");
                sb.Append("<cargo>");
                sb.Append("\r\n");
                int RowIndex = 1;
                DataTable _dtc = GetImgCargoValues("1");
                for (int i = 0; i < _dtc.Rows.Count; i++)
                {
                    sb.Append("F");
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["POL"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["IMONumber"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["VesselCallSign"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["ImportVoyageCd"].ToString());
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append(RowIndex);
                    sb.Append(c);
                    sb.Append("0");
                    sb.Append(_dtc.Rows[i]["BookingNo"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["BkgDate"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["POLCode"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["FPODCode"].ToString());
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["Shipper"].ToString());
                    sb.Append(c);
                    sb.Append(_dtc.Rows[i]["ShipperAddress"].ToString());

                    RowIndex++;
                    sb.Append("\r\n");
                }

                sb.Append("<END-cargo>");
                sb.Append("\r\n");

                //#region Container
                sb.Append("<contain>");
                sb.Append("\r\n");

                DataTable _dtCntr = GetImgCntrNoValues("1");
                for (int y = 0; y < _dtCntr.Rows.Count; y++)
                {
                    sb.Append("F");
                    sb.Append(c);
                    sb.Append(dtVoy.Rows[0]["POL"].ToString());
                    sb.Append(c);
                    sb.Append(dtVoy.Rows[0]["IMONumber"].ToString());
                    sb.Append(c);
                    sb.Append(dtVoy.Rows[0]["VesselCallSign"].ToString());
                    sb.Append(c);
                    sb.Append(dtVoy.Rows[0]["ImportVoyageCd"].ToString());
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append(c);
                    sb.Append("1");
                    sb.Append(c);
                    sb.Append("0");
                    sb.Append(_dtCntr.Rows[y]["CntrNo"].ToString());
                    sb.Append(c);
                    sb.Append(_dtCntr.Rows[y]["SealNo"].ToString());
                    sb.Append(c);
                    sb.Append(dtVoy.Rows[0]["PanNo"].ToString());
                    sb.Append(c);
                    sb.Append(dtVoy.Rows[0]["PanNo"].ToString());
                }
                sb.Append("\r\n");
                sb.Append("<END-contain>");

                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment;filename=" + 1 + "ImpEDI.txt");
                Response.Charset = "";
                Response.ContentType = "text/plain";
                Response.Output.Write(sb.ToString());
                Response.ContentEncoding = Encoding.ASCII;
                Response.Flush();
                Response.End();
            }
        }


        private string[] SplitByLenght(string Values, int split)
        {

            List<string> list = new List<string>();
            int SplitTheLoop = Values.Length / split;
            for (int i = 0; i < SplitTheLoop; i++)
                list.Add(Values.Substring(i * split, split));
            if (SplitTheLoop * split != Values.Length)
                list.Add(Values.Substring(SplitTheLoop * split));

            return list.ToArray();
        }

        public static string fn_WrapDescriptionWords(string description, int lineWidth)
        {
            if (string.IsNullOrWhiteSpace(description))
                return string.Empty;

            var words = description.Split(' ');
            var result = new StringBuilder();
            var line = new StringBuilder();

            foreach (var word in words)
            {
                if ((line.Length + word.Length + 1) > lineWidth)
                {
                    result.AppendLine(line.ToString().TrimEnd());
                    line.Clear();
                }

                line.Append(word + " ");
            }

            if (line.Length > 0)
            {
                result.AppendLine(line.ToString().TrimEnd());
            }

            return result.ToString();
        }

        public DataTable GetImgVesVoyValues(string vesVoyID, string TerminalID)
        {
            //string _Query = " Select VesselName,VesselOwner,ExportVoyageCd,ImportVoyageCd,IMONumber,NVO_VoyageRoute.PortID,TerminalID,LineCode,AgentCode,PanNo,FloppyCode,IMONumber, " +
            //                " VesselCallSign,(select top(1) PortCode from NVO_PortMaster inner join NVO_Booking on NVO_Booking.POLID = NVO_PortMaster.ID where VesVoyID = NVO_VoyageRoute.RID) as POL," +
            //                " (select top(1) FPOD from  NVO_Booking where VesVoyID = NVO_VoyageRoute.RID) as FPOD, " +
            //                " (select top(1) FloppyCode from NVO_Customs_PortCode inner join NVO_Booking on NVO_Booking.FPODID = NVO_Customs_PortCode.PortID where VesVoyID = NVO_VoyageRoute.RID) as FPOD_FloppyCode " +
            //                " from NVO_Voyage " +
            //                " inner join NVO_VoyageRoute on NVO_VoyageRoute.VoyageID = NVO_Voyage.ID " +
            //                " inner join NVO_VesselMaster on NVO_VesselMaster.ID = NVO_Voyage.VesselID " +
            //                " inner join NVO_Customs_AgentCode on NVO_Customs_AgentCode.PortID = NVO_VoyageRoute.PortID " +
            //                " inner join NVO_Customs_PortCode on NVO_Customs_PortCode.PortID =NVO_VoyageRoute.PortID " +
            //                " where NVO_Voyage.ID = 1 and TerminalID = 99";
            //return Manag.GetViewData(_Query, "");


            string _Query = " select VesselName, VesselOwner, ExportVoyageCd, ImportVoyageCd, IMONumber, NVO_VoyageRoute.PortID, TerminalID, " +
                            " (select top(1) TerminalCode from NVO_TerminalMaster where Id = TerminalID) as TerminalCode," +
                            " (select top(1) PanNo from NVO_Customs_AgentCode where NVO_Customs_AgentCode.PortID = NVO_VoyageRoute.PortID) as PanNo, " +
                            " (select top(1) AgentCode from NVO_Customs_AgentCode where NVO_Customs_AgentCode.PortID = NVO_VoyageRoute.PortID) as AgentCode, " +
                            " (select top(1) LineCode from NVO_Customs_AgentCode where NVO_Customs_AgentCode.PortID = NVO_VoyageRoute.PortID) as LineCode, " +
                            " (select top(1) FloppyCode from NVO_Customs_PortCode where NVO_Customs_PortCode.PortID = NVO_VoyageRoute.PortID) as FloppyCode,IMONumber,  VesselCallSign, " +
                            " (select top(1) IGMNo from NVO_VoyageManifestDtls where VoyageID= NVO_Voyage.ID) as IGMNO," +
                            " (select top(1) IGMDate from NVO_VoyageManifestDtls where VoyageID = NVO_Voyage.ID) as IGMDate," +
                            " (select top(1) PortCode from NVO_PortMaster inner join NVO_Booking on NVO_Booking.POLID = NVO_PortMaster.ID where VesVoyID = NVO_VoyageRoute.RID) as POL, " +
                            " (select top(1) FPOD from  NVO_Booking where VesVoyID = NVO_VoyageRoute.RID) as FPOD, " +
                            " (select top(1) FloppyCode from NVO_Customs_PortCode inner join NVO_Booking on NVO_Booking.FPODID = NVO_Customs_PortCode.PortID where VesVoyID = NVO_VoyageRoute.RID) as FPOD_FloppyCode " +
                            " from NVO_Voyage inner join NVO_VoyageRoute on NVO_VoyageRoute.VoyageID = NVO_Voyage.ID" +
                            " inner join NVO_VesselMaster on NVO_VesselMaster.ID = NVO_Voyage.VesselID " +
                            " where NVO_Voyage.ID = 1 and TerminalID = 142";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetImgCargoValues(string vesVoyID)
        {
            string _Query = " select distinct NVO_BOL.ID as BLID,format(getdate(),'yyyyMMdd') as CurrDate, format(getdate(),'HHmm') as CurrTime, " +
                            " BLNumber,SUBSTRING(BLNumber, 1, 12) as BLNumberv , format(getdate(), 'yyyyMMddHHmmss') as BLDate, format(NVO_BOLImpVoyageDetails.ETA,'yyyyMMddHHmm') as ETA," +
                            " (select top(1) VesVoy from NVO_View_VoyageDetails where NVO_View_VoyageDetails.ID = NVO_BOLImpVoyageDetails.VesVoyID) as Voy, " +
                            " (select(select top(1) VesselName from NVO_VesselMaster where NVO_VesselMaster.ID = v.VesselID) " +
                            " from NVO_Voyage v where V.ID = NVO_BOLImpVoyageDetails.VesVoyID)  as VesselName,  " +
                            " (select(select top(1) VesselCallSign from NVO_VesselMaster where NVO_VesselMaster.ID = v.VesselID) from NVO_Voyage v where V.ID = NVO_BOLImpVoyageDetails.VesVoyID)  as VesselCallSign,  "+
                            " (select top(1) Notes from NVO_VoyageNotesDtls where NotesTypeID= 336  and NVO_VoyageNotesDtls.VoyageID=NVO_BOLImpVoyageDetails.VesVoyID) as Nationality, "+
                            " (select top(1) ExportVoyageCd from NVO_VoyageRoute where VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) as ImportVoyageCd, " +
                            " (select top(1) PortCode from  NVO_PortMaster where ID = NVO_BOL.PODID) as PortCode, " +
                            " replace(convert(NVARCHAR, ETA, 106), ' ', '-') as ArriveDatev, " +
                            " (select top(1) RIN from NVO_Voyage where ID = NVO_BOLImpVoyageDetails.VesVoyID) as RinNo, " +
                            " 'MFI' as MessageTypes,'1' as NoofInstalment,'Mv000003' as AMSNo, " +
                            " format(getdate(), 'yyyyMMddHHmmss') as IssuesDate, " +
                            " BkgParty,ShipmentType,ServiceType,CommodityType,Shipper,PickUpDepot,POD,POL,POO,FPOD,   " +
                            " (select(select top(1) PartyAddress from NVO_BOLCustomerDetails " +
                            " where NVO_BOLCustomerDetails.BLID = NVO_BOL.ID and PartyTypeID = 1) from NVO_BOL where BLTypes = 40 " +
                            " and BkgID = NVO_Booking.ID) as ShipperAddress, " +
                            " case when isnull(ImpMarks,'')= '' then MarkNo else ImpMarks end MarkNo,case when isnull(ImpDesc,'')= '' then CagoDescription else ImpDesc end ImpDesc,BLCommodityTypeID " +
                            " from NVO_BOLImpVoyageDetails " +
                            " inner join NVO_BOL on NVO_BOL.ID = NVO_BOLImpVoyageDetails.BLID " +
                            " inner join NVO_Booking on NVO_Booking.ID = NVO_BOL.BkgID " +
                            " where LegInformation = 74 and NVO_BOLImpVoyageDetails.VesVoyID =" + vesVoyID;


            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetImgCntrNoValues(string BLID)
        {
            string _Query = "  select distinct  CntrID,(select top(1) TypeID from NVO_Containers where ID = CntrID)  as TypeID, " +
                            " (select top(1) CntrNo from NVO_Containers where ID = CntrID) as CntrNo, " +
                            " (select(select  top(1) Size from NVO_tblCntrTypes where ID = TypeID) from NVO_Containers where Id = CntrID) as Size, " +
                            " ISOCode,CntrID,SealNo,NoOfPkg,PakgType,PakgTypeName,GrsWt,NtWt,VGM,CBM from NVO_BOLCntrDetails where  BkgId = " + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetBLCustomerValues(string BLID)
        {
            string _Query = " select 1 as sno, 0 as PartyTypeID,  'SA' as Desp, 'BLUE WAVE SHIPPING (M) SDN BHD' AS PartyName, 'NO.31-1,JALAN RAMIN 1,\nBANDAR AMBANG BOTANIC, 41200 KLANG,\nSELANGOR DARUL EHSAN,MALAYSIA ' as PartyAddress, " +
                            " 'CM' + right('0000' + convert(varchar(4),7 ), 4) as PartyCode " +
                            " union " +
                            " select 2 as sno, 0 as PartyTypeID, 'PSA' as Desp, 'BLUE WAVE SHIPPING (M) SDN BHD' AS PartyName, 'NO.31-1,JALAN RAMIN 1,\nBANDAR AMBANG BOTANIC, 41200 KLANG,\nSELANGOR DARUL EHSAN,MALAYSIA' as PartyAddress, " +
                            " 'CM' + right('0000' + convert(varchar(4),10 ), 4) as PartyCode " +
                            " union " +
                            " select 3 as sno, PartyTypeID, 'CN' as Desp, SUBSTRING(PartyName, 1,34) as PartyName, (select top(1) dbo.fn_WrapDescriptionWords(SUBSTRING(Address,1,105), 50) from NVO_CusBranchLocation where NVO_CusBranchLocation.CID =NVO_BOLCustomerDetails.PartID) as PartyAddress,'CN' + right('0000' + convert(varchar(4), PartID), 4) as PartyCode  from NVO_BOLCustomerDetails where PartyTypeID in (2) and BLID= " + BLID +
                            " union " +
                            " select 4 as sno, PartyTypeID, 'CZ' as Desp, SUBSTRING(PartyName, 1,34) as PartyName, (select top(1) dbo.fn_WrapDescriptionWords(SUBSTRING(Address,1,105), 50) from NVO_CusBranchLocation where NVO_CusBranchLocation.CID =NVO_BOLCustomerDetails.PartID) as PartyAddress,'CZ' + right('0000' + convert(varchar(4), PartID), 4) as PartyCode  from NVO_BOLCustomerDetails where PartyTypeID in (1) and BLID =" + BLID +
                            " UNION " +
                            " select 5 as sno, PartyTypeID, 'NI' as Desp, SUBSTRING(PartyName, 1,34) as PartyName, (select top(1) dbo.fn_WrapDescriptionWords(SUBSTRING(Address,1,105), 50) from NVO_CusBranchLocation where NVO_CusBranchLocation.CID =NVO_BOLCustomerDetails.PartID) as PartyAddress,'NI' + right('0000' + convert(varchar(4), PartID), 4) as PartyCode  from NVO_BOLCustomerDetails where PartyTypeID in (3) and BLID = " + BLID;
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetPortCodeValues(string BLID)
        {
            string _Query = "select* from v_PortCodeBkgwise where Id=" + BLID;
            return Manag.GetViewData(_Query, "");
        }

        //public DataTable GetPortCodeDocValues(string BLID)
        //{
        //    string _Query = " select 1 sno, 'CSC' as desvalue, 'H10' as valuesv " +
        //                    " union " +
        //                    " select 2 sno, 'CS1' as desvalue, 'B10' as valuesv " +
        //                    " union " +
        //                    " select 3 sno, 'SCN' as desvalue, Notes as valuesv from NVO_VoyageNotesDtls where NotesTypeID = 282  and VoyageID = 76 " +
        //                    " union " +
        //                    " select 4 sno,  'SC1' as desvalue, '11C4' as valuesv " +
        //                    " union " +
        //                    " select 5 sno, 'CAC' as desvalue, '2' as valuesv " +
        //                    " union " +
        //                    " select 6 sno, 'BM' as desvalue,  (select top(1) BLNumber from NVO_BOL where ID = " + BLID + ") as valuesv " +
        //                    " Union " +
        //                    " select 7 sno, 'TOP' as desvalue,  " +
        //                    " (select(select(select top(1) TerminalCode from NVO_TerminalMaster where NVO_TerminalMaster.ID = " +
        //                    " NVO_VoyageRoute.TerminalID )  from NVO_VoyageRoute where NVO_VoyageRoute.VoyageID = NVO_BOLImpVoyageDetails.VesVoyID) " +
        //                    " from NVO_BOLImpVoyageDetails where NVO_BOLImpVoyageDetails.BLID = " + BLID + ") as valuesv " +
        //                    " union " +
        //                    " select 8 sno, 'AA' as desvalue, 'PCA' as valuesv";
        //    return Manag.GetViewData(_Query, "");
        //}


        public DataTable GetPortCodeDocValues(string BLID, string Topvalues)
        {
            string _Query = " select 1 sno, 'CSC' as desvalue, 'H10' as valuesv " +
                            " union " +

                            " select 2 sno, 'SCN' as desvalue, Notes as valuesv from NVO_VoyageNotesDtls where NotesTypeID = 282  and VoyageID = 76 " +
                            " union " +

                            " select 3 sno, 'CAC' as desvalue, '2' as valuesv " +
                            " union " +
                            " select 4 sno, 'BM' as desvalue,  (select top(1) substring(BLNumber,1,15) from NVO_BOL where ID = " + BLID + ") as valuesv " +
                            " Union " +
                            " select 5 sno, 'TOP' as desvalue, '" + Topvalues + "' as valuesv" +
                            " Union " +
                            " select 6 sno, 'AA' as desvalue, 'PCA' as valuesv";
                            
            return Manag.GetViewData(_Query, "");
        }


        public DataTable GetProductCharValues()
        {
            string _Query = " select  'IMD' as desvalue, '9' as valuesv " +
                            " union " +
                            " select  'UN' as desvalue, '3082' as valuesv " +
                            " union " +
                            " select  'LPK' as desvalue, '1' as valuesv";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable GetProductTotalMesCharValues()
        {
            string _Query = " select  'AAE' as desvalue, '9' as valuesv " +
                            " union " +
                            " select  'UN' as desvalue, '3082' as valuesv " +
                            " union " +
                            " select  'LPK' as desvalue, '1' as valuesv";
            return Manag.GetViewData(_Query, "");
        }

        public DataTable getCntrDetails(string BLID)
        {
            string _Query = " select(select top(1) PkgCode from NVO_CargoPkgMaster where ID = PakgType) as PakgeName,NoOfPkg, " +
                            " (select top(1) CntrNo  from NVO_Containers where NVO_Containers.ID =NVO_BOLCntrDetails.CntrID) as CntrNo," +
                            " * from NVO_BOL " +
                            " inner join NVO_BOLCntrDetails on NVO_BOLCntrDetails.BLID = NVO_BOL.Id " +
                            " where ID =" + BLID;
            return Manag.GetViewData(_Query, "");
        }


        public DataTable getCntrDetailsTotal(string BLID)
        {
            string _Query = " select sum(GRSwt) as GRsWtTotal, sum(NtWt) as NTWtTotal,sum(CBM) as CBMTotal, sum(NoofPkg) as TotalPkg, count(CntrID) as NoofCntr" +
                " from NVO_BOLCntrDetails where  BLID= " + BLID;


            return Manag.GetViewData(_Query, "");
        }

       
    }
}
