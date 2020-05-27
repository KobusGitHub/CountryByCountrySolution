using CBC_V1.models;
using CountryByCountryReportV1;
using CountryByCountryReportV1.models;
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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace CBC_V1
{
    public partial class Form1 : Form
    {
        // DocSpec_Type gloabalDocSpec = null;
        FileInfo xlsxFile = null;
        string xmlFilePath = null;
        string destLogFilePath = null;
        bool canWriteToLogFile = true;
        List<ReceivingCountryClass> receivingCountryClass = new List<ReceivingCountryClass>();
        List<ConstituentEntitiesSummary> ConstituentEntitiesSummaries = new List<ConstituentEntitiesSummary>();

        public Form1()
        {
            InitializeComponent();
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
            var sourceFilePath = txtSource.Text;
            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show("Invalid Source File", "Alert");
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

            destLogFilePath = txtDestFolder.Text + "\\" + destFileName.Replace(".xml", "_Log.log");
            if (File.Exists(destLogFilePath))
            {
                File.Delete(destLogFilePath);
            }

            xlsxFile = new FileInfo(sourceFilePath);
            xmlFilePath = destFilePath;

            try
            {
                this.StartWork();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed");
                logMessage("File Completed");

                if (File.Exists(destLogFilePath))
                {
                    MessageBox.Show("Please check log file for more information:", "Failed");
                    Process.Start("explorer.exe", txtDestFolder.Text);
                }

                return;
            }


            MessageBox.Show("File created successfully.", "Success");

            if (File.Exists(xmlFilePath))
            {
                Process.Start("explorer.exe", txtDestFolder.Text);
            }

        }


        public DocSpec_Type GetDocSpec(ExcelPackage package, string docTypeIndic, string docRefId, string corrDocRefId)
        {
            DocSpec_Type docSpec;

            if (string.IsNullOrEmpty(corrDocRefId))
            {
                docSpec = new DocSpec_Type
                {
                    DocTypeIndic = EnumLookup.GetOECDDocTypeIndicEnumType(docTypeIndic),
                    DocRefId = docRefId,
                };
            }
            else
            {
                docSpec = new DocSpec_Type
                {
                    DocTypeIndic = EnumLookup.GetOECDDocTypeIndicEnumType(docTypeIndic),
                    DocRefId = docRefId,
                    CorrDocRefId = "",
                    CorrMessageRefId = ""
                };
            }

            return docSpec;
        }
        private void StartWork()
        {
            // xsd.exe CbcXML_v1.0.1.xsd /Classes oecdtypes_v4.1.xsd /Classes isocbctypes_v1.0.1.xsd


            var cbcfd = new CountryByCountryDeclarationStructure();
            cbcfd.CBC_OECD = new CBC_OECD();
            cbcfd.CBC_SARS = new CBC_SARS_Structure();



            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(xlsxFile))
            {

                //gloabalDocSpec = new DocSpec_Type
                //{
                //    DocTypeIndic = EnumLookup.GetOECDDocTypeIndicEnumType(GetExcelStringValue(package, "CoverPage", "B16")), //   (S: CoverPage; Cells: B24)
                //    DocRefId = GetExcelStringValue(package, "CoverPage", "B17"), // "ZA2018DOCBAW", //   (S: CoverPage; Cells: B25)
                //    CorrDocRefId = "",
                //    CorrMessageRefId = ""
                //};

                cbcfd.CBC_OECD.version = "1.0";
                cbcfd.CBC_OECD.MessageSpec = GetMessageSpec(package);
                cbcfd.CBC_OECD.CbcBody = GetCbcBodies(package).ToArray();

                cbcfd.CBC_SARS = GetCBC_SARS_Structure(package);
            }

            logMessage("Data Completed");

            var xml = "";
            XmlSerializer xsSubmit = new XmlSerializer(typeof(CountryByCountryDeclarationStructure));
            using (var sww = new StringWriter())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                settings.NewLineOnAttributes = true;

                using (XmlWriter writer = XmlWriter.Create(sww, settings))
                {
                    xsSubmit.Serialize(writer, cbcfd);
                    xml = sww.ToString(); // Your XML
                }
            }


            using (StreamWriter file = new StreamWriter(xmlFilePath))
            {
                file.Write(xml);
                logMessage("File Completed");
            }

        }


        private CBC_SARS_Structure GetCBC_SARS_Structure(ExcelPackage package)
        {
            var sarsStructure = new CBC_SARS_Structure();

            sarsStructure.ContactDetails = new CBC_SARS_StructureContactDetails
            {
                Surname = this.GetExcelStringValue(package, "CoverPage", "B21"),
                FirstNames = this.GetExcelStringValue(package, "CoverPage", "B22"),
                BusTelNo1 = this.GetExcelStringValue(package, "CoverPage", "B23"),
                BusTelNo2 = this.GetExcelStringValue(package, "CoverPage", "B24"),
                CellNo = this.GetExcelStringValue(package, "CoverPage", "B25"),
                EmailAddress = this.GetExcelStringValue(package, "CoverPage", "B26")
            };

            sarsStructure.DeclarationDate = DateTime.Parse(this.GetExcelStringValue(package, "CoverPage", "B20"));
            sarsStructure.DeclarationDateSpecified = true;

            sarsStructure.TotalConsolidatedMNEGroupRevenue = new FinancialAmtWithCurrencyStructure()
            {
                CurrencyCode = "ZAR",
                Amount = Convert.ToInt64(this.GetExcelDoubleValue(package, "CoverPage", "B27").Value)
            };
            sarsStructure.NoOfTaxJurisdictions = Convert.ToInt16(this.receivingCountryClass.Count());


            List<CBC_SARS_StructureTaxJurisdiction> stjList = new List<CBC_SARS_StructureTaxJurisdiction>();

            foreach (var summary in this.ConstituentEntitiesSummaries)
            {
                stjList.Add(new CBC_SARS_StructureTaxJurisdiction
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
            messageSpec.TransmittingCountry = CountryCode_Type.ZA;


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


                ReceivingCountries.Add(EnumLookup.GetCountryCodeEnumType(cellValue)); // (S:SUMMARY - Cells:A)

                this.receivingCountryClass.Add(new ReceivingCountryClass
                {
                    RowNumber = rowNumber,
                    CountryCode = cellValue
                });
                rowNumber++;
            }
            messageSpec.ReceivingCountry = new CountryCode_Type[] { CountryCode_Type.ZA };

            // MessageType
            messageSpec.MessageType = MessageType_EnumType.CBC;

            // Language
            messageSpec.Language = LanguageCode_Type.EN;
            messageSpec.LanguageSpecified = true;

            // Warning
            messageSpec.Warning = GetExcelStringValue(package, "CoverPage", "B19");

            // Contact
            messageSpec.Contact = GetExcelStringValue(package, "CoverPage", "B1"); // (S: CoverPage; Cells:B1)

            // MessageRefId
            messageSpec.MessageRefId = GetExcelStringValue(package, "CoverPage", "B2");// "ZA2019BAW"; // (S: CoverPage; Cells:B2)

            // MessageTypeIndic
            messageSpec.MessageTypeIndic = EnumLookup.GetCbcMessageTypeIndicEnumType(GetExcelStringValue(package, "CoverPage", "B3")); // CbcMessageTypeIndic_EnumType.CBC401; //  (S: CoverPage; Cells:B3)
            messageSpec.MessageTypeIndicSpecified = true;

            // ReportingPeriod
            // messageSpec.ReportingPeriod = new DateTime(GetExcelIntValue(package, "CoverPage", "B4").Value, GetExcelIntValue(package, "CoverPage", "B5").Value, GetExcelIntValue(package, "CoverPage", "B6").Value); // new DateTime(2019, 9, 30);  //  (S: CoverPage; Cells:B4, B5, B6)
            messageSpec.ReportingPeriod = DateTime.Parse(GetExcelStringValue(package, "CoverPage", "B4"));

            messageSpec.Timestamp = DateTime.Now;


            // messageSpec.SendingEntityIN = null; // TODO Check with SARS

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


            repEnt.Entity = GetOrganisationPartyType(EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B5")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B7)
                EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B6")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B8)
                GetExcelStringValue(package, "CoverPage", "B7"), //"9000051715", // (S: CoverPage; Cells: B9)
                EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B8")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B10)
                GetExcelStringValue(package, "CoverPage", "B9"), //"1918/000095/06", // (S: CoverPage; Cells: B11)
                GetExcelStringValue(package, "CoverPage", "B10"), //"Barloworld Limited", // (S: CoverPage; Cells: B12)
                EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B11")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B13)
                GetExcelStringValue(package, "CoverPage", "B12").Split(';'), // new object[] { "61 Katherine Street", "Sandton", "2196" }, // (S: CoverPage; Cells: B14) "61 Katherine Street;Sandton;2196" (Split on ;)
                EnumLookup.GetOECDLegalAddressTypeEnumType(GetExcelStringValue(package, "CoverPage", "B13"))); //OECDLegalAddressType_EnumType.OECD304);// (S: CoverPage; Cells: B15)


            repEnt.ReportingRole = EnumLookup.GetCbcReportingRoleEnumType(GetExcelStringValue(package, "CoverPage", "B15")); // CbcReportingRole_EnumType.CBC701; // (S: CoverPage; Cells: B17)


            var docTypeIndic = GetExcelStringValue(package, "CoverPage", "B16");
            var docRefId = GetExcelStringValue(package, "CoverPage", "B17");
            var corrDocRefId = GetExcelStringValue(package, "CoverPage", "B18");
            repEnt.DocSpec = GetDocSpec(package, docTypeIndic, docRefId, corrDocRefId);

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


            //if (address.Count() == 3)
            //{
            //    var addrItem = new AddressFix_Type()
            //    {
            //        Street = address[0].Trim(),
            //        City = address[1].Trim(),
            //        PostCode = address[2].Trim()
            //    };

            //    var addr = new Address_Type()
            //    {
            //        CountryCode = addressCountryCode,
            //        Items = new AddressFix_Type[] { addrItem },
            //        legalAddressType = legalAddressType,
            //        legalAddressTypeSpecified = true
            //    };
            //    entity.Address = new Address_Type[] { addr };
            //}
            //else
            //{
                var adr = string.Join(",", address);
                var addr = new Address_Type()
                {
                    CountryCode = addressCountryCode,
                    Items = new string[] { adr },
                    legalAddressType = legalAddressType,
                    legalAddressTypeSpecified = true
                };
                entity.Address = new Address_Type[] { addr };
            //}


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

                var docTypeIndic = GetExcelStringValue(package, "Additional Information", "B" + rowNumber);
                var docRefId = GetExcelStringValue(package, "Additional Information", "C" + rowNumber);
                var corrDocRefId = GetExcelStringValue(package, "Additional Information", "D" + rowNumber);
                info.DocSpec = GetDocSpec(package, docTypeIndic, docRefId, corrDocRefId);

                if (cellValue.Length >= 4000)
                {
                    throw new ArgumentException("Other Information not allowed more than 4000 characters in one cell");
                }

                info.OtherInfo = cellValue;

                // info.ResCountryCode = null; - Optional
                // info.SummaryRef = null; - Optional

                list.Add(info);

                rowNumber++;
            }

            return list;
        }



        private List<CorrectableCbcReport_Type> GetCbcReports(ExcelPackage package)
        {
            var cbcReports = new List<CorrectableCbcReport_Type>();

            // Some Loop
            foreach (var recCountryCls in receivingCountryClass)
            {
                var cbcRep = new CorrectableCbcReport_Type();

                cbcRep.ResCountryCode = EnumLookup.GetCountryCodeEnumType(recCountryCls.CountryCode);

                var docTypeIndic = GetExcelStringValue(package, "SUMMARY", "M" + recCountryCls.RowNumber);
                var docRefId = GetExcelStringValue(package, "SUMMARY", "N" + recCountryCls.RowNumber);
                var corrDocRefId = GetExcelStringValue(package, "SUMMARY", "O" + recCountryCls.RowNumber);
                cbcRep.DocSpec = GetDocSpec(package, docTypeIndic, docRefId, corrDocRefId);


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
                        bizActivities.Add(EnumLookup.GetCbcBizActivityTypeEnumType(actCode));
                    }
                }
                constEntity.BizActivities = bizActivities.ToArray();


                constEntity.ConstEntity = GetOrganisationPartyType(EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "G" + rowNumber)),
                    EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "E" + rowNumber)),
                    GetExcelStringValue(package, workbookName, "D" + rowNumber),
                    EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "C" + rowNumber)),
                    GetExcelStringValue(package, workbookName, "B" + rowNumber),
                    GetExcelStringValue(package, workbookName, "A" + rowNumber),
                    EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "K" + rowNumber)),
                    GetExcelStringValue(package, workbookName, "J" + rowNumber).Split(';'),
                    EnumLookup.GetOECDLegalAddressTypeEnumType(GetExcelStringValue(package, workbookName, "L" + rowNumber)));


                // TODO - It doesnt want to serialize this
                constEntity.IncorpCountryCode = EnumLookup.GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "F" + rowNumber));
                constEntity.IncorpCountryCodeSpecified = true;

                constEntity.OtherEntityInfo = GetExcelStringValue(package, workbookName, "I" + rowNumber);


                constEntities.Add(constEntity);

                rowNumber++;
            }

            this.ConstituentEntitiesSummaries.Add(new ConstituentEntitiesSummary { CountryCode = receivingCountryClass.CountryCode, ConstituentEntityCount = rowNumber - 2 });

            return constEntities;
        }




        private string GetExcelStringValue(ExcelPackage package, string workbook, string cell)
        {

            logMessage(string.Format("Worksheet: '{0}', Cell '{1}', Converting to a String", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString().Trim();
        }

        private int? GetExcelIntValue(ExcelPackage package, string workbook, string cell)
        {
            logMessage(string.Format("Worksheet: '{0}', Cell '{1}', Converting to a INTEGER", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return int.Parse(package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString().Trim());
        }

        private double? GetExcelDoubleValue(ExcelPackage package, string workbook, string cell)
        {
            logMessage(string.Format("Worksheet: '{0}', Cell '{1}', Converting to a Double", workbook, cell));

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
                using (StreamWriter w = File.AppendText(destLogFilePath))
                {
                    w.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to write to log file!", "Log File Failed");

                canWriteToLogFile = false;
                throw;
            }
        }


    }
}
