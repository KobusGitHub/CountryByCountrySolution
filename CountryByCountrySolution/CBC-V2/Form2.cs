using CBC_V2.models;
using CountryByCountryReportV2;
using CountryByCountryReportV2.models;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Text.Json;
using CBC_V2_ALL_COUNTRIES;
using System.Xml.Serialization;
using CBC_V2.Service.AllCountries;
using CBC_V2.Service.SARS;

namespace CBC_V2
{
    public partial class Form2 : Form
    {
        // DocSpec_Type gloabalDocSpec = null;
        FileInfo xlsxFile = null;
        string xmlFilePath = null;
        string destLogFilePath = null;
        bool canWriteToLogFile = true;
        Guid myGuid;

        List<ReceivingCountryClass> receivingCountryClass = new List<ReceivingCountryClass>();
        List<ConstituentEntitiesSummary> ConstituentEntitiesSummaries = new List<ConstituentEntitiesSummary>();

        public Form2()
        {
            InitializeComponent();
            lstLog.Items.Clear();

        }

        private void brnBrowseSource_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = ".xlsx";
            openFileDialog1.Filter = "Excel Worksheets|*.xlsx";
            openFileDialog1.ShowDialog();
            txtSource.Text = openFileDialog1.FileName;
        }

        private void btnBrowseDest_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtDestFolder.Text = folderBrowserDialog1.SelectedPath;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            myGuid = Guid.NewGuid();
            lstLog.Items.Clear();


            var sourceFilePath = txtSource.Text;
            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show("Invalid Source File", "Alert");
                return;
            }

            var destFileName = txtDestFileName.Text;
            if (!destFileName.Contains(".xml"))
            {
                destFileName = destFileName + ".xml";
            }

            var destFilePath = txtDestFolder.Text + "\\" + destFileName;
            if (File.Exists(destFilePath))
            {
                File.Delete(destFilePath);
            }

            destLogFilePath = txtDestFolder.Text + "\\" + destFileName.Replace(".xml", "_" + myGuid + "_Log.log");
            if (File.Exists(destLogFilePath))
            {
                File.Delete(destLogFilePath);
            }

            xlsxFile = new FileInfo(sourceFilePath);
            xmlFilePath = destFilePath;

            var newExcelFilePath = txtDestFolder.Text + "\\" + destFileName.Replace(".xml", "_" + myGuid + xlsxFile.Extension);
            var newxlsxFile = new FileInfo(newExcelFilePath);

            try
            {
                this.StartWork(newxlsxFile);
            }
            catch (Exception ex)
            {
                AddMessageToListBox("Generating XML Failed");

                MessageBox.Show(ex.Message, "Failed");
                logMessage("File Completed");

                if (File.Exists(destLogFilePath))
                {
                    MessageBox.Show("Please check log file for more information:", "Failed");
                    Process.Start("explorer.exe", txtDestFolder.Text);
                }

                return;
            }

            AddMessageToListBox("Generating XML Success");

            MessageBox.Show("File created successfully.", "Success");


            try
            {
                logMessage("");

                logMessage("XML Validation Start");
                XmlValidatorService.ValidateXml(xmlFilePath);
                logMessage("XML Validation Success");

                MessageBox.Show("XML validation successful.", "Success");
            }
            catch (Exception ex)
            {
                logMessage(ex.Message);
                logMessage("XML Validation Failed");
                MessageBox.Show(ex.Message);
                return;
            }




            if (File.Exists(xmlFilePath))
            {
                Process.Start("explorer.exe", txtDestFolder.Text);
            }

        }



        public DocSpec_Type GetDocSpec(ExcelPackage package, string docTypeIndic, string docRefId, string corrDocRefId, string corrMessageRefId)
        {
            DocSpec_Type docSpec = new DocSpec_Type
            {
                DocTypeIndic = AllCountriesEnumLookup.GetOECDDocTypeIndicEnumType(docTypeIndic),
                DocRefId = docRefId,
            };


            if (!string.IsNullOrEmpty(corrDocRefId))
            {
                docSpec.CorrDocRefId = corrDocRefId;
            }

            if (!string.IsNullOrEmpty(corrMessageRefId))
            {
                docSpec.CorrMessageRefId = corrMessageRefId;
            }


            return docSpec;
        }


       
        private void StartWork(FileInfo newExcelFile)
        {
            // xsd.exe CbcXML_v1.0.1.xsd /Classes oecdtypes_v4.1.xsd /Classes isocbctypes_v1.0.1.xsd


            // xsd.exe CbcXML_v2.0.xsd /Classes isocbctypes_v1.1.xsd /Classes oecdcbctypes_v5.0.xsd


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            if (ckSars.Checked)
            {
                SarsGenerator generator = new SarsGenerator(this.myGuid);
                generator.LogMessageEvent += LogEventHandlerFired;
                generator.GenerateFile(xlsxFile, xmlFilePath, rad8.Checked);
            }
            else
            {
                AllCountriesGenerator generator = new AllCountriesGenerator(this.myGuid);
                generator.LogMessageEvent += LogEventHandlerFired;
                generator.GenerateFile(xlsxFile, xmlFilePath, rad8.Checked);
            }


            return;






            CBC_V2_SARS.CBC_SARS_Structure CBC_SARS = null;


            var cbcfd = new CBC_V2_ALL_COUNTRIES.CBC_OECD();
            cbcfd.version = "2.0";


            using (var package = new ExcelPackage(xlsxFile))
            {
                this.SetupReferences(package);

                cbcfd.MessageSpec = GetMessageSpec(package);
                cbcfd.CbcBody = GetCbcBodies(package).ToArray();

                if (ckSars.Checked)
                {
                    CBC_SARS = GetCBC_SARS_Structure(package);
                }

                package.SaveAs(newExcelFile);
            }

            var xml = "";
            if (ckSars.Checked)
            {
                // SARS

                //var sarsCbcds = new CBC_V2_SARS.CountryByCountryDeclarationStructure();
                //sarsCbcds.CBC_OECD = new CBC_V2_SARS.CBC_OECD();
                //sarsCbcds.CBC_SARS = new CBC_V2_SARS.CBC_SARS_Structure();

                //sarsCbcds.CBC_OECD.version = "2.0";
                //sarsCbcds.CBC_OECD.MessageSpec = ConvertMessageSpecToSars(cbcfd.MessageSpec);
                //sarsCbcds.CBC_OECD.CbcBody = ConvertCbcBodyToSars(cbcfd.CbcBody);
                //sarsCbcds.CBC_SARS = CBC_SARS;


                //logMessage("Data Completed");

                //if (rad8.Checked)
                //{
                //    xml = XmlConverter.GetUtf8XmlStringSars(sarsCbcds);
                //}
                //else
                //{
                //    xml = XmlConverter.GetUtf16XmlStringSars(sarsCbcds);
                //}

            }
            else
            {
                // ALL Countries

                logMessage("Data Completed");

                if (rad8.Checked)
                {
                    xml = XmlConverter.GetUtf8XmlStringAllCountries(cbcfd);
                }
                else
                {
                    xml = XmlConverter.GetUtf16XmlStringAllCountries(cbcfd);
                }

            }

            using (StreamWriter file = new StreamWriter(xmlFilePath))
            {
                file.Write(xml);
                logMessage("File Completed");
            }

        }

        private void LogEventHandlerFired(object sender, string e)
        {
            this.logMessage(e);
        }

        private CBC_V2_SARS.MessageSpec_Type ConvertMessageSpecToSars(MessageSpec_Type messageSpec)
        {
            var serialisedObject = JsonSerializer.Serialize(messageSpec);
            var returnObject = JsonSerializer.Deserialize<CBC_V2_SARS.MessageSpec_Type>(serialisedObject);

            return returnObject;
        }

        private CBC_V2_SARS.CbcBody_Type[] ConvertCbcBodyToSars(CbcBody_Type[] cbcBody)
        {
            var serialisedObject = JsonSerializer.Serialize(cbcBody);
            var returnObject = JsonSerializer.Deserialize<CBC_V2_SARS.CbcBody_Type[]>(serialisedObject);

            return returnObject;
        }

        //public string GetUtf8XmlString(CBC_OECD cbcfd)
        //{
        //    var serializer = new XmlSerializer(typeof(CBC_OECD));
        //    var xml = "";
        //    using (StringWriter writer = new Utf8StringWriter())
        //    {
        //        serializer.Serialize(writer, cbcfd);
        //        xml = writer.ToString();
        //    }
        //    return xml;
        //}

        //public string GetUtf16XmlString(CBC_OECD cbcfd)
        //{
        //    var xml = "";
        //    XmlSerializer xsSubmit = new XmlSerializer(typeof(CBC_OECD));
        //    using (var sww = new StringWriter())
        //    {
        //        XmlWriterSettings settings = new XmlWriterSettings();
        //        settings.Indent = true;
        //        settings.NewLineOnAttributes = true;



        //        using (XmlWriter writer = XmlWriter.Create(sww, settings))
        //        {
        //            xsSubmit.Serialize(writer, cbcfd);
        //            xml = sww.ToString(); // Your XML
        //        }
        //    }

        //    return xml;

        //}

        private CBC_V2_SARS.CBC_SARS_Structure GetCBC_SARS_Structure(ExcelPackage package)
        {
            var sarsStructure = new CBC_V2_SARS.CBC_SARS_Structure();

            sarsStructure.ContactDetails = new CBC_V2_SARS.CBC_SARS_StructureContactDetails
            {
                Surname = this.GetExcelStringValue(package, "CoverPage", "B22"),
                FirstNames = this.GetExcelStringValue(package, "CoverPage", "B23"),
                BusTelNo1 = this.GetExcelStringValue(package, "CoverPage", "B24"),
                //BusTelNo2 = this.GetExcelStringValue(package, "CoverPage", "B25"),
                CellNo = this.GetExcelStringValue(package, "CoverPage", "B26"),
                EmailAddress = this.GetExcelStringValue(package, "CoverPage", "B27")
            };

            var busTelCont2Value = this.GetExcelStringValue(package, "CoverPage", "B25");
            if (!string.IsNullOrEmpty(busTelCont2Value))
            {
                sarsStructure.ContactDetails.BusTelNo2 = busTelCont2Value;
            }

            sarsStructure.DeclarationDate = DateTime.Parse(this.GetExcelStringValue(package, "CoverPage", "B21"));
            sarsStructure.DeclarationDateSpecified = true;

            sarsStructure.TotalConsolidatedMNEGroupRevenue = new CBC_V2_SARS.FinancialAmtWithCurrencyStructure()
            {
                CurrencyCode = "ZAR",
                Amount = Convert.ToInt64(this.GetExcelDoubleValue(package, "CoverPage", "B28").Value)
            };
            sarsStructure.NoOfTaxJurisdictions = Convert.ToInt16(this.receivingCountryClass.Count());


            List<CBC_V2_SARS.CBC_SARS_StructureTaxJurisdiction> stjList = new List<CBC_V2_SARS.CBC_SARS_StructureTaxJurisdiction>();

            foreach (var summary in this.ConstituentEntitiesSummaries)
            {
                stjList.Add(new CBC_V2_SARS.CBC_SARS_StructureTaxJurisdiction
                {
                    NoOfConstituentEntities = Convert.ToInt16(summary.ConstituentEntityCount)
                });
            }


            sarsStructure.TaxJurisdictions = stjList.ToArray();

            return sarsStructure;

        }

        private MessageSpec_Type GetMessageSpec(ExcelPackage package)
        {
            var messageSpec = new MessageSpec_Type();
            this.receivingCountryClass = new List<ReceivingCountryClass>();

            // cbc_oecd.MessageSpec.SendingEntityIN
            messageSpec.TransmittingCountry = CountryCode_Type.MU;// TODO 20240305 get FROM file Somewhere SHould be cover page B6


            // ReceivingCountry
            List<CountryCode_Type> ReceivingCountries = new List<CountryCode_Type>();

            var rowNumber = 2;
            while (true)
            {
                var cellValue = GetExcelStringValue(package, "SUMMARY", "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }


                ReceivingCountries.Add(AllCountriesEnumLookup.GetCountryCodeEnumType(cellValue)); // (S:SUMMARY - Cells:A)

                this.receivingCountryClass.Add(new ReceivingCountryClass
                {
                    RowNumber = rowNumber,
                    CountryCode = cellValue
                });
                rowNumber++;
            }
            //messageSpec.ReceivingCountry = new CountryCode_Type[] { CountryCode_Type.MU,  };  // TODO 20240305 get FROM file Somewhere SHould be cover page B6

            messageSpec.ReceivingCountry = new CountryCode_Type[] {
                CountryCode_Type.AE,
                CountryCode_Type.BE,
                CountryCode_Type.BI,
                CountryCode_Type.BW,
                CountryCode_Type.CD,
                CountryCode_Type.CV,
                CountryCode_Type.GB,
                CountryCode_Type.GH,
                CountryCode_Type.JE,
                CountryCode_Type.KE,
                CountryCode_Type.LS,
                CountryCode_Type.MT,
                CountryCode_Type.MU,
                CountryCode_Type.MW,
                CountryCode_Type.MZ,
                CountryCode_Type.NG,
                CountryCode_Type.RW,
                CountryCode_Type.SA,
                CountryCode_Type.SD,
                CountryCode_Type.TG,
                CountryCode_Type.TZ,
                CountryCode_Type.UG,
                CountryCode_Type.US,
                CountryCode_Type.ZA,
                CountryCode_Type.ZM,
                CountryCode_Type.ZW,
                CountryCode_Type.VG,
                CountryCode_Type.IL,
                CountryCode_Type.CA,
                CountryCode_Type.CL,
                CountryCode_Type.MX,
                CountryCode_Type.PE,
                CountryCode_Type.AR,
                CountryCode_Type.PY,
                CountryCode_Type.CO,
                CountryCode_Type.PL
            };
            // TODO 20240305 get FROM file Somewhere SHould be cover page B6





            // MessageType
            messageSpec.MessageType = MessageType_EnumType.CBC;

            // Language
            messageSpec.Language = LanguageCode_Type.EN;
            messageSpec.LanguageSpecified = true;

            // Warning
            messageSpec.Warning = GetExcelStringValue(package, "CoverPage", "B20");

            // Contact
            // messageSpec.Contact = GetExcelStringValue(package, "CoverPage", "B1"); // (S: CoverPage; Cells:B1)
            var cont = GetExcelStringValue(package, "CoverPage", "B1");
            if (!string.IsNullOrEmpty(cont))
            {
                messageSpec.Contact = cont;
            }

            // MessageRefId
            messageSpec.MessageRefId = GetExcelStringValue(package, "CoverPage", "B2");// "ZA2019BAW"; // (S: CoverPage; Cells:B2)

            if (!string.IsNullOrEmpty(GetExcelStringValue(package, "CoverPage", "B3")))
            {
                messageSpec.CorrMessageRefId = new string[] { GetExcelStringValue(package, "CoverPage", "B3") };
            }

            // MessageTypeIndic
            messageSpec.MessageTypeIndic = AllCountriesEnumLookup.GetCbcMessageTypeIndicEnumType(GetExcelStringValue(package, "CoverPage", "B4")); // CbcMessageTypeIndic_EnumType.CBC401; //  (S: CoverPage; Cells:B3)

            // Removed in V2.0
            // messageSpec.MessageTypeIndicSpecified = true;

            // ReportingPeriod
            messageSpec.ReportingPeriod = DateTime.Parse(GetExcelStringValue(package, "CoverPage", "B5"));

            messageSpec.Timestamp = DateTime.Now;


            messageSpec.SendingEntityIN = GetExcelStringValue(package, "CoverPage", "B8"); ; // TODO Check with SARS B10

            // messageSpec.CorrMessageRefId = null;  // This data element is not used for CbC reporting



            return messageSpec;

        }


        private List<CbcBody_Type> GetCbcBodies(ExcelPackage package)
        {
            var cbcBodies = new List<CbcBody_Type>();

            var cbcBody = new CbcBody_Type();

            cbcBody.ReportingEntity = GetReportingEntity(package);
            cbcBody.AdditionalInfo = GetCorrectableAdditionalInfo(package).ToArray();

            cbcBody.CbcReports = GetCbcReports(package).ToArray();

            cbcBodies.Add(cbcBody);


            return cbcBodies;
        }

        private CorrectableReportingEntity_Type GetReportingEntity(ExcelPackage package)
        {
            var repEnt = new CorrectableReportingEntity_Type();


            var resCountryCode = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B6"));
            var tinIssueBy = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B7"));
            var tinValue = GetExcelStringValue(package, "CoverPage", "B8");
            var OrgInTypeIssueBy = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B9"));
            var orgInTypeValue = GetExcelStringValue(package, "CoverPage", "B10");
            var nameOrg = GetExcelStringValue(package, "CoverPage", "B11");
            var addCountryCode = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B12"));
            var adds = GetExcelStringValue(package, "CoverPage", "B13").Split(';');
            var legAddType = AllCountriesEnumLookup.GetOECDLegalAddressTypeEnumType(GetExcelStringValue(package, "CoverPage", "B14"));

            repEnt.Entity = GetOrganisationPartyType(resCountryCode, tinIssueBy, tinValue, OrgInTypeIssueBy, orgInTypeValue, nameOrg, addCountryCode, adds, legAddType);



            repEnt.ReportingRole = AllCountriesEnumLookup.GetCbcReportingRoleEnumType(GetExcelStringValue(package, "CoverPage", "B16")); // CbcReportingRole_EnumType.CBC701; // (S: CoverPage; Cells: B17)


            var docTypeIndic = GetExcelStringValue(package, "CoverPage", "B17");
            var docRefId = GetExcelStringValue(package, "CoverPage", "B18");
            var corrDocRefId = GetExcelStringValue(package, "CoverPage", "B19");
            var corrMessageRefId = GetExcelStringValue(package, "CoverPage", "B3");
            repEnt.DocSpec = GetDocSpec(package, docTypeIndic, docRefId, corrDocRefId, corrMessageRefId);

            repEnt.NameMNEGroup = GetExcelStringValue(package, "CoverPage", "B15");

            var reportEndString = GetExcelStringValue(package, "CoverPage", "B5");
            var reportEnd = DateTime.Parse(reportEndString);
            var reportStart = reportEnd.AddYears(-1);
            reportStart = reportStart.AddDays(1);

            repEnt.ReportingPeriod = new ReportingEntity_TypeReportingPeriod()
            {
                StartDate = reportStart,
                EndDate = reportEnd
            };

            return repEnt;
        }

        private OrganisationParty_Type GetOrganisationPartyType(
           CountryCode_Type resCountryCode,
           CountryCode_Type tinIssuedBy,
           string tinValue,
           CountryCode_Type organisationINTypeIssuedBy,
           string organisationINTypeValue,
           string nameOrganisation,
           CountryCode_Type addressCountryCode,
           string[] address,
           OECDLegalAddressType_EnumType legalAddressType
           )
        {
            // Entity
            var entity = new OrganisationParty_Type();

            entity.ResCountryCode = new CountryCode_Type[] { resCountryCode };

            // Entity.TIN
            entity.TIN = new TIN_Type();
            entity.TIN.issuedBy = tinIssuedBy;
            entity.TIN.issuedBySpecified = true;
            entity.TIN.Value = tinValue;


            OrganisationIN_Type organisationIN_Type = new OrganisationIN_Type
            {
                INType = "Company Registration Number",
                issuedBy = organisationINTypeIssuedBy,
                issuedBySpecified = true,
                Value = organisationINTypeValue
            };
            entity.IN = new OrganisationIN_Type[] { organisationIN_Type };


            NameOrganisation_Type nameOrganisation_Type = new NameOrganisation_Type();
            nameOrganisation_Type.Value = nameOrganisation;
            entity.Name = new NameOrganisation_Type[] { nameOrganisation_Type };


            var adr = string.Join(",", address);
            var addr = new Address_Type()
            {
                CountryCode = addressCountryCode,
                Items = new string[] { adr },
                legalAddressType = legalAddressType,
                legalAddressTypeSpecified = true
            };
            entity.Address = new Address_Type[] { addr };

            return entity;

        }


        private List<CorrectableAdditionalInfo_Type> GetCorrectableAdditionalInfo(ExcelPackage package)
        {
            var list = new List<CorrectableAdditionalInfo_Type>();

            int rowNumber = 2;
            while (true)
            {
                var cellValue = GetExcelStringValue(package, "Additional Information", "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }

                var info = new CorrectableAdditionalInfo_Type();

                var docTypeIndic = GetExcelStringValue(package, "Additional Information", "D" + rowNumber);
                var docRefId = GetExcelStringValue(package, "Additional Information", "E" + rowNumber);
                var corrDocRefId = GetExcelStringValue(package, "Additional Information", "F" + rowNumber);
                var corrMessageRefId = GetExcelStringValue(package, "CoverPage", "B3");
                info.DocSpec = GetDocSpec(package, docTypeIndic, docRefId, corrDocRefId, corrMessageRefId);

                if (cellValue.Length >= 4000)
                {
                    throw new ArgumentException("Other Information not allowed more than 4000 characters in one cell");
                }


                var otherInfoList = new List<StringMin1Max4000WithLang_Type>();
                otherInfoList.Add(new StringMin1Max4000WithLang_Type()
                {
                    language = LanguageCode_Type.EN,
                    languageSpecified = true,
                    Value = cellValue
                });
                info.OtherInfo = otherInfoList.ToArray();

                var resCountryCodeExcelValue = GetExcelStringValue(package, "Additional Information", "B" + rowNumber);
                info.ResCountryCode = GetAdditionalInfoResCountryCodes(resCountryCodeExcelValue);


                var resSummaryRegExcelValue = GetExcelStringValue(package, "Additional Information", "C" + rowNumber);
                info.SummaryRef = GetAdditionalInfoSummaryRefs(resSummaryRegExcelValue);

                list.Add(info);

                rowNumber++;
            }

            return list;
        }

        private CountryCode_Type[] GetAdditionalInfoResCountryCodes(string commaDelimetedCountryCode)
        {
            List<CountryCode_Type> codeList = new List<CountryCode_Type>();

            var array = commaDelimetedCountryCode.Split(',');
            foreach (var arr in array)
            {
                var val = arr.Trim();
                if (string.IsNullOrEmpty(val))
                {
                    continue;
                }
                codeList.Add(AllCountriesEnumLookup.GetCountryCodeEnumType(val));
            }


            return codeList.ToArray();
        }

        private CbcSummaryListElementsType_EnumType[] GetAdditionalInfoSummaryRefs(string commaDelimitedSummaryRef)
        {
            List<CbcSummaryListElementsType_EnumType> refList = new List<CbcSummaryListElementsType_EnumType>();

            var array = commaDelimitedSummaryRef.Split(',');
            foreach (var arr in array)
            {
                var val = arr.Trim();
                if (string.IsNullOrEmpty(val))
                {
                    continue;
                }
                refList.Add(AllCountriesEnumLookup.GetCbcSummaryListElementsType_EnumType(val));
            }

            return refList.ToArray();
        }


        private List<CorrectableCbcReport_Type> GetCbcReports(ExcelPackage package)
        {
            var cbcReports = new List<CorrectableCbcReport_Type>();

            // Some Loop
            foreach (var recCountryCls in receivingCountryClass)
            {
                var cbcRep = new CorrectableCbcReport_Type();

                cbcRep.ResCountryCode = AllCountriesEnumLookup.GetCountryCodeEnumType(recCountryCls.CountryCode);

                var docTypeIndic = GetExcelStringValue(package, "SUMMARY", "M" + recCountryCls.RowNumber);
                var docRefId = GetExcelStringValue(package, "SUMMARY", "N" + recCountryCls.RowNumber);
                var corrDocRefId = GetExcelStringValue(package, "SUMMARY", "O" + recCountryCls.RowNumber);
                var corrMessageRefId = GetExcelStringValue(package, "CoverPage", "B3");
                cbcRep.DocSpec = GetDocSpec(package, docTypeIndic, docRefId, corrDocRefId, corrMessageRefId);


                // Summary
                cbcRep.Summary = GetSummary(currCode_Type.ZAR,
                    GetExcelStringValue(package, "SUMMARY", "K" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "C" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "D" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "L" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "E" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "G" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "F" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "H" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "J" + recCountryCls.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "I" + recCountryCls.RowNumber));

                // Const Entities
                cbcRep.ConstEntities = GetConstituentEntities(package, recCountryCls).ToArray();

                cbcReports.Add(cbcRep);
            }

            return cbcReports;
        }


        private string RemoveDecimalFromString(string value)
        {
            var dblValue = double.Parse(value);
            var intValue = Convert.ToInt64(dblValue);
            var returnValue = intValue.ToString();
            return returnValue;
        }

        private CorrectableCbcReport_TypeSummary GetSummary(currCode_Type currCode,
           string assets,
           string capital,
           string earnings,
           string nbEmployees,
           string profitOrLoss,
           string revenues_Related,
           string revenues_Unrelated,
           string revenues_Total,
           string taxAccrued,
           string taxPaid)
        {

            assets = this.RemoveDecimalFromString(assets);
            capital = this.RemoveDecimalFromString(capital);
            earnings = this.RemoveDecimalFromString(earnings);
            nbEmployees = this.RemoveDecimalFromString(nbEmployees);
            profitOrLoss = this.RemoveDecimalFromString(profitOrLoss);
            revenues_Related = this.RemoveDecimalFromString(revenues_Related);
            revenues_Unrelated = this.RemoveDecimalFromString(revenues_Unrelated);
            revenues_Total = this.RemoveDecimalFromString(revenues_Total);
            taxAccrued = this.RemoveDecimalFromString(taxAccrued);
            taxPaid = this.RemoveDecimalFromString(taxPaid);

            var summary = new CorrectableCbcReport_TypeSummary();
            summary.Assets = new MonAmnt_Type()
            {
                currCode = currCode,
                Value = assets
            };

            summary.Capital = new MonAmnt_Type()
            {
                currCode = currCode,
                Value = capital
            };
            summary.Earnings = new MonAmnt_Type
            {
                currCode = currCode,
                Value = earnings
            };

            summary.NbEmployees = nbEmployees;
            summary.ProfitOrLoss = new MonAmnt_Type
            {
                currCode = currCode,
                Value = profitOrLoss
            };

            summary.Revenues = new CorrectableCbcReport_TypeSummaryRevenues()
            {
                Related = new MonAmnt_Type
                {
                    currCode = currCode,
                    Value = revenues_Related
                },
                Unrelated = new MonAmnt_Type
                {
                    currCode = currCode,
                    Value = revenues_Unrelated
                },
                Total = new MonAmnt_Type
                {
                    currCode = currCode,
                    Value = revenues_Total
                }
            };

            summary.TaxAccrued = new MonAmnt_Type
            {
                currCode = currCode,
                Value = taxAccrued
            };

            summary.TaxPaid = new MonAmnt_Type
            {
                currCode = currCode,
                Value = taxPaid
            };

            return summary;
        }


        private List<ConstituentEntity_Type> GetConstituentEntities(ExcelPackage package, ReceivingCountryClass receivingCountryClass)
        {

            var workbookName = "CE_" + receivingCountryClass.CountryCode;

            var constEntities = new List<ConstituentEntity_Type>();

            int rowNumber = 2;
            while (true)
            {
                var cellValue = GetExcelStringValue(package, workbookName, "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }


                var constEntity = new ConstituentEntity_Type();
                var bizActivities = new List<CbcBizActivityType_EnumType>();

                var excelActValue = GetExcelStringValue(package, workbookName, "H" + rowNumber);
                var actCodes = excelActValue.Split(';');
                foreach (var actCode in actCodes)
                {
                    if (!string.IsNullOrEmpty(actCode))
                    {
                        bizActivities.Add(AllCountriesEnumLookup.GetCbcBizActivityTypeEnumType(actCode));
                    }
                }
                constEntity.BizActivities = bizActivities.ToArray();



                var resCountryCode = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "G" + rowNumber));
                var tinIssueBy = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "E" + rowNumber));
                var tinValue = GetExcelStringValue(package, workbookName, "D" + rowNumber);
                var OrgInTypeIssueBy = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "C" + rowNumber));
                var orgInTypeValue = GetExcelStringValue(package, workbookName, "B" + rowNumber);
                var nameOrg = GetExcelStringValue(package, workbookName, "A" + rowNumber);
                var addCountryCode = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "K" + rowNumber));
                var adds = GetExcelStringValue(package, workbookName, "J" + rowNumber).Split(';');
                var legAddType = AllCountriesEnumLookup.GetOECDLegalAddressTypeEnumType(GetExcelStringValue(package, workbookName, "L" + rowNumber));

                constEntity.ConstEntity = GetOrganisationPartyType(resCountryCode, tinIssueBy, tinValue, OrgInTypeIssueBy, orgInTypeValue, nameOrg, addCountryCode, adds, legAddType);


                constEntity.IncorpCountryCode = AllCountriesEnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "F" + rowNumber));
                constEntity.IncorpCountryCodeSpecified = true;

                constEntity.OtherEntityInfo = GetExcelStringValue(package, workbookName, "I" + rowNumber);


                constEntity.Role = UltimateParentEntityRole_EnumType.CBC801;
                constEntity.RoleSpecified = true;

                constEntities.Add(constEntity);

                rowNumber++;
            }

            this.ConstituentEntitiesSummaries.Add(new ConstituentEntitiesSummary { CountryCode = receivingCountryClass.CountryCode, ConstituentEntityCount = rowNumber - 2 });

            return constEntities;
        }


        private void SetExcelStringValue(ExcelPackage package, string workbook, string cell, string value)
        {

            logMessage(string.Format("WRITE: Worksheet: '{0}', Cell '{1}', Value '{2}'", workbook, cell, value));


            package.Workbook.Worksheets[workbook].Cells[cell].Value = value;
        }

        private string GetExcelStringValue(ExcelPackage package, string workbook, string cell)
        {

            logMessage(string.Format("READ: Worksheet: '{0}', Cell '{1}', Converting to a String", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString().Trim();
        }

        private int? GetExcelIntValue(ExcelPackage package, string workbook, string cell)
        {
            logMessage(string.Format("READ: Worksheet: '{0}', Cell '{1}', Converting to a INTEGER", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return int.Parse(package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString().Trim());
        }

        private double? GetExcelDoubleValue(ExcelPackage package, string workbook, string cell)
        {
            logMessage(string.Format("READ: Worksheet: '{0}', Cell '{1}', Converting to a Double", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return Convert.ToDouble(package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString().Trim());
        }


        private void logMessage(string message)
        {
            if (!canWriteToLogFile)
            {
                return;
            }

            try
            {
                AddMessageToListBox(message);

                using (StreamWriter w = File.AppendText(destLogFilePath))
                {
                    w.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to write to log file!", "Log File Failed");
                var exep = ex;
                canWriteToLogFile = false;
                throw;
            }
        }

        private void AddMessageToListBox(string message)
        {
            lstLog.Items.Add(message);
            // scoll to bottom
            lstLog.TopIndex = lstLog.Items.Count - 1;
        }

        private void SetupReferences(ExcelPackage p)
        {

            var docRefPrefix = GetExcelStringValue(p, "CoverPage", "B29");

            // OECD0 -> Resent Data
            // OECD1 -> New Data
            // OECD2 -> Corrected Data
            // OECD3 -> Deletion of Data
            // OECD10 -> Resent Test Data
            // OECD11 -> New Test Data
            // OECD12 -> Corrected Test Data
            // OECD13 -> Deletion of Test Data


            // MessageRefID

            var messageTypeIndic = GetExcelStringValue(p, "CoverPage", "B4");
            if (messageTypeIndic == "CBC402") // Corrected Data
            {
                var docRefValue = GetExcelStringValue(p, "CoverPage", "B2");
                if (string.IsNullOrEmpty(docRefValue))
                {
                    logMessage("Correction on empty reference INVALID!");
                    throw new ArgumentException("Correction on empty reference INVALID!");
                }
                this.SetExcelStringValue(p, "CoverPage", "B3", docRefValue);

            }
            else if (messageTypeIndic == "CBC401") // New Data
            {
                this.SetExcelStringValue(p, "CoverPage", "B3", "");
                this.SetExcelStringValue(p, "CoverPage", "B4", "CBC401");

            }
            this.SetExcelStringValue(p, "CoverPage", "B2", docRefPrefix + myGuid);




            // ReportingEnt-DocSpec-DocRefID
            var repEntDocTypeIndic = GetExcelStringValue(p, "CoverPage", "B17");
            if (repEntDocTypeIndic == "OECD2") // Corrected Data
            {
                var docRefValue = GetExcelStringValue(p, "CoverPage", "B18");
                if (string.IsNullOrEmpty(docRefValue))
                {
                    logMessage("Correction on empty reference INVALID!");
                    throw new ArgumentException("Correction on empty reference INVALID!");
                }
                this.SetExcelStringValue(p, "CoverPage", "B19", docRefValue);
            }
            else if (repEntDocTypeIndic == "OECD1") // New Data
            {
                this.SetExcelStringValue(p, "CoverPage", "B19", "");
            }
            this.SetExcelStringValue(p, "CoverPage", "B18", docRefPrefix + Guid.NewGuid());


            // SUMMARY
            var rowNumber = 2;
            while (true)
            {
                var cellValue = GetExcelStringValue(p, "SUMMARY", "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }

                var docTypeIndec = GetExcelStringValue(p, "SUMMARY", "M" + rowNumber);
                if (docTypeIndec == "OECD2") // Corrected Data
                {
                    var docRefValue = GetExcelStringValue(p, "SUMMARY", "N" + rowNumber);
                    if (string.IsNullOrEmpty(docRefValue))
                    {
                        logMessage("Correction on empty reference INVALID!");
                        throw new ArgumentException("Correction on empty reference INVALID!");
                    }
                    this.SetExcelStringValue(p, "SUMMARY", "O" + rowNumber, docRefValue);
                }
                else if (docTypeIndec == "OECD1") // New Data
                {
                    this.SetExcelStringValue(p, "SUMMARY", "O" + rowNumber, "");
                }
                this.SetExcelStringValue(p, "SUMMARY", "N" + rowNumber, docRefPrefix + Guid.NewGuid());

                rowNumber++;
            }


            // Additional Information
            rowNumber = 2;
            while (true)
            {
                var cellValue = GetExcelStringValue(p, "Additional Information", "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }

                var docTypeIndec = GetExcelStringValue(p, "Additional Information", "D" + rowNumber);
                if (docTypeIndec == "OECD2") // Corrected Data
                {
                    var docRefValue = GetExcelStringValue(p, "Additional Information", "E" + rowNumber);
                    if (string.IsNullOrEmpty(docRefValue))
                    {
                        logMessage("Correction on empty reference INVALID!");
                        throw new ArgumentException("Correction on empty reference INVALID!");
                    }
                    this.SetExcelStringValue(p, "Additional Information", "F" + rowNumber, docRefValue);
                }
                else if (docTypeIndec == "OECD1") // New Data
                {
                    this.SetExcelStringValue(p, "Additional Information", "F" + rowNumber, "");
                }
                this.SetExcelStringValue(p, "Additional Information", "E" + rowNumber, docRefPrefix + Guid.NewGuid());

                rowNumber++;
            }


        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }



    public class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }
}
