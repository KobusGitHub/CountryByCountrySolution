using OfficeOpenXml;
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

namespace CBC_V2
{
    public partial class Form1 : Form
    {
        DocSpec_Type gloabalDocSpec = null;
        FileInfo xlsxFile = null;
        string xmlFilePath = null;
        string destLogFilePath = null;
        bool canWriteToLogFile = true;

        List<ReceivingCountryClass> receivingCountryClass = new List<ReceivingCountryClass>();

        public Form1()
        {
            InitializeComponent();
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

        private string GetExcelStringValue(ExcelPackage package, string workbook, string cell)
        {

            logMessage(string.Format("Worksheet: '{0}', Cell '{1}', Converting to a String", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString();
        }

        private int? GetExcelIntValue(ExcelPackage package, string workbook, string cell)
        {
            logMessage(string.Format("Worksheet: '{0}', Cell '{1}', Converting to a INTEGER", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return int.Parse(package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString());
        }

        private double? GetExcelDoubleValue(ExcelPackage package, string workbook, string cell)
        {
            logMessage(string.Format("Worksheet: '{0}', Cell '{1}', Converting to a Double", workbook, cell));

            var cellObject = package.Workbook.Worksheets[workbook].Cells[cell].Value;
            if (cellObject == null)
            {
                return null;
            }
            return Convert.ToDouble(package.Workbook.Worksheets[workbook].Cells[cell].Value.ToString());
        }

        private void StartWork()
        {



            // xsd.exe CbcXML_v2.0.xsd /Classes oecdcbctypes_v5.0.xsd /Classes isocbctypes_v1.1.xsd


            var cbc_oecd = new CBC_OECD();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(xlsxFile))
            {


                gloabalDocSpec = new DocSpec_Type
                {
                    DocTypeIndic = GetOECDDocTypeIndicEnumType(GetExcelStringValue(package, "CoverPage", "B24")), //   (S: CoverPage; Cells: B24)
                    DocRefId = GetExcelStringValue(package, "CoverPage", "B25"), // "ZA2018DOCBAW", //   (S: CoverPage; Cells: B25)
                    CorrDocRefId = "",
                    CorrMessageRefId = ""
                };

                cbc_oecd.version = "2.0";
                cbc_oecd.MessageSpec = GetMessageSpec(package);
                cbc_oecd.CbcBody = GetCbcBodies(package).ToArray();

            }


            var xml = "";
            XmlSerializer xsSubmit = new XmlSerializer(typeof(CBC_OECD));
            using (var sww = new StringWriter())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                settings.NewLineOnAttributes = true;


                using (XmlWriter writer = XmlWriter.Create(sww, settings))
                {
                    xsSubmit.Serialize(writer, cbc_oecd);
                    xml = sww.ToString(); // Your XML
                }
            }


            using (StreamWriter file = new StreamWriter(xmlFilePath))
            {
                file.Write(xml);

            }
        }


        private OECDDocTypeIndic_EnumType GetOECDDocTypeIndicEnumType(string oECDDocTypeIndic)
        {
            return (OECDDocTypeIndic_EnumType)Enum.Parse(typeof(OECDDocTypeIndic_EnumType), oECDDocTypeIndic);
        }

        private CountryCode_Type GetCountryCodeEnumType(string countryCodeType)
        {
            return (CountryCode_Type)Enum.Parse(typeof(CountryCode_Type), countryCodeType);
        }

        private CbcMessageTypeIndic_EnumType GetCbcMessageTypeIndicEnumType(string CbcMessageTypeIndicType)
        {
            return (CbcMessageTypeIndic_EnumType)Enum.Parse(typeof(CbcMessageTypeIndic_EnumType), CbcMessageTypeIndicType);
        }


        private OECDLegalAddressType_EnumType GetOECDLegalAddressTypeEnumType(string OECDLegalAddressTypeType)
        {
            return (OECDLegalAddressType_EnumType)Enum.Parse(typeof(OECDLegalAddressType_EnumType), OECDLegalAddressTypeType);
        }

        private CbcReportingRole_EnumType GetCbcReportingRoleEnumType(string CbcReportingRoleType)
        {
            return (CbcReportingRole_EnumType)Enum.Parse(typeof(CbcReportingRole_EnumType), CbcReportingRoleType);
        }



        private CbcBizActivityType_EnumType GetCbcBizActivityTypeEnumType(string CbcBizActivityTypeEnumType)
        {
            return (CbcBizActivityType_EnumType)Enum.Parse(typeof(CbcBizActivityType_EnumType), CbcBizActivityTypeEnumType);
        }

        private UltimateParentEntityRole_EnumType GetUltimateParentEntityRoleEnumType(string UltimateParentEntityRoleEnumType)
        {
            return (UltimateParentEntityRole_EnumType)Enum.Parse(typeof(UltimateParentEntityRole_EnumType), UltimateParentEntityRoleEnumType);
        }




        private MessageSpec_Type GetMessageSpec(ExcelPackage package)
        {
            var messageSpec = new MessageSpec_Type();

            //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //using (var package = new ExcelPackage(xlsxFile))
            //{

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


                ReceivingCountries.Add(GetCountryCodeEnumType(cellValue)); // (S:SUMMARY - Cells:A)

                this.receivingCountryClass.Add(new ReceivingCountryClass
                {
                    RowNumber = rowNumber,
                    CountryCode = cellValue
                });
                rowNumber++;
            }
            messageSpec.ReceivingCountry = ReceivingCountries.ToArray();

            // MessageType
            messageSpec.MessageType = MessageType_EnumType.CBC;

            // Language
            messageSpec.Language = LanguageCode_Type.EN;

            // Warning
            messageSpec.Warning = "";

            // Contact
            messageSpec.Contact = GetExcelStringValue(package, "CoverPage", "B1"); // (S: CoverPage; Cells:B1)

            // MessageRefId
            messageSpec.MessageRefId = GetExcelStringValue(package, "CoverPage", "B2");// "ZA2019BAW"; // (S: CoverPage; Cells:B2)

            // MessageTypeIndic
            messageSpec.MessageTypeIndic = GetCbcMessageTypeIndicEnumType(GetExcelStringValue(package, "CoverPage", "B3")); // CbcMessageTypeIndic_EnumType.CBC401; //  (S: CoverPage; Cells:B3)

            //// CorrMessageRefId
            //messageSpec.CorrMessageRefId = new string[];

            // ReportingPeriod
            messageSpec.ReportingPeriod = new DateTime(GetExcelIntValue(package, "CoverPage", "B4").Value, GetExcelIntValue(package, "CoverPage", "B5").Value, GetExcelIntValue(package, "CoverPage", "B6").Value); // new DateTime(2019, 9, 30);  //  (S: CoverPage; Cells:B4, B5, B6)

            messageSpec.Timestamp = DateTime.Now;

            //}
            return messageSpec;

        }


        private List<CbcBody_Type> GetCbcBodies(ExcelPackage package)
        {
            var cbcBodies = new List<CbcBody_Type>();


            // We need to loop through something
            // for (int i = 0; i < 10; i++)  // (S: SUMMARY; Cells:B Count)
            //foreach (var receivingCountry in receivingCountries)
            //{
            var cbcBody = new CbcBody_Type();

            cbcBody.ReportingEntity = GetReportingEntity(package);
            cbcBody.AdditionalInfo = GetCorrectableAdditionalInfo(package).ToArray();

            cbcBody.CbcReports = GetCbcReports(package).ToArray();

            cbcBodies.Add(cbcBody);
            //}



            return cbcBodies;
        }


        private List<CorrectableCbcReport_Type> GetCbcReports(ExcelPackage package)
        {
            var cbcReports = new List<CorrectableCbcReport_Type>();

            // Some Loop
            // for (int i = 0; i < 10; i++) // (S: SUMMARY; Cells: B Count)
            foreach (var receivingCountryClass in receivingCountryClass)
            {
                var cbcRep = new CorrectableCbcReport_Type();

                cbcRep.ResCountryCode = GetCountryCodeEnumType(receivingCountryClass.CountryCode); // CountryCode_Type.AO; // (S: SUMMARY; Cells: A)

                cbcRep.DocSpec = gloabalDocSpec;

                // Summary
                // (S: SUMMARY; Cells:                         ,K              ,C              ,D                ,L     ,E              ,G              ,F              ,H                 ,J            ,I)
                // cbcRep.Summary = GetSummary(currCode_Type.ZAR, "766604888.40", "-69810833.20", "-1217694218.00", "363", "-217527466.40", "509150539.20", "-771290869.60", "-1280441408.80", "55195379.39", "53236237.67");
                cbcRep.Summary = GetSummary(currCode_Type.ZAR,
                    GetExcelStringValue(package, "SUMMARY", "K" + receivingCountryClass.RowNumber),
                    GetExcelStringValue(package, "SUMMARY", "C" + receivingCountryClass.RowNumber),  // "-69810833.20",
                    GetExcelStringValue(package, "SUMMARY", "D" + receivingCountryClass.RowNumber),  // "-1217694218.00",
                    GetExcelStringValue(package, "SUMMARY", "L" + receivingCountryClass.RowNumber),  // "363",
                    GetExcelStringValue(package, "SUMMARY", "E" + receivingCountryClass.RowNumber),  // "-217527466.40",
                    GetExcelStringValue(package, "SUMMARY", "G" + receivingCountryClass.RowNumber),  // "509150539.20",
                    GetExcelStringValue(package, "SUMMARY", "F" + receivingCountryClass.RowNumber),  // "-771290869.60",
                    GetExcelStringValue(package, "SUMMARY", "H" + receivingCountryClass.RowNumber),  // "-1280441408.80",
                    GetExcelStringValue(package, "SUMMARY", "J" + receivingCountryClass.RowNumber),  // "55195379.39",
                    GetExcelStringValue(package, "SUMMARY", "I" + receivingCountryClass.RowNumber));  // "53236237.67");

                // Const Entities
                cbcRep.ConstEntities = GetConstituentEntities(package, receivingCountryClass).ToArray();

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
                currCode = currCode_Type.ZAR,
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
                // (S: CE_XX like CE_AO; Cells: X Count)
                var cellValue = GetExcelStringValue(package, workbookName, "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }

                //   // Loop
                //for (int i = 0; i < 10; i++) // (S: CE_XX like CE_AO; Cells: X Count)
                //{
                var constEntity = new ConstituentEntity_Type();

                var bizActivities = new List<CbcBizActivityType_EnumType>();


                //for (int bizA = 0; bizA < 10; bizA++) // (S: CE_XX; Cells: X Count)
                //{
                //    bizActivities.Add(CbcBizActivityType_EnumType.CBC505); // (S: CE_XX; Cells: H)
                //}
                var excelActValue = GetExcelStringValue(package, workbookName, "H" + rowNumber);
                var actCodes = excelActValue.Split(';');
                foreach (var actCode in actCodes)
                {
                    if (!string.IsNullOrEmpty(actCode))
                    {
                        bizActivities.Add(GetCbcBizActivityTypeEnumType(actCode));
                    }
                }
                constEntity.BizActivities = bizActivities.ToArray();


                constEntity.ConstEntity = GetOrganisationPartyType(GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "G" + rowNumber)), // CountryCode_Type.AO, // (S: CE_XX; Cells: G)
                    GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "E" + rowNumber)), // CountryCode_Type.AO, // (S: CE_XX; Cells: E)
                    GetExcelStringValue(package, workbookName, "D" + rowNumber), //"5410000595", // (S: CE_XX; Cells: D)
                    GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "C" + rowNumber)), // CountryCode_Type.AO, // (S: CE_XX; Cells: C)
                    GetExcelStringValue(package, workbookName, "B" + rowNumber), //"1996.1", // (S: CE_XX; Cells: B)
                    GetExcelStringValue(package, workbookName, "A" + rowNumber), //"Barloworld Equipamentos Angola Limitada", // (S: CE_XX; Cells: A)
                    GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "K" + rowNumber)), // CountryCode_Type.AO, // (S: CE_XX; Cells: K)
                    GetExcelStringValue(package, workbookName, "J" + rowNumber).Split(';'), //new object[] { "", "" }, // (S: CE_XX; Cells: J)
                    GetOECDLegalAddressTypeEnumType(GetExcelStringValue(package, workbookName, "L" + rowNumber))); // OECDLegalAddressType_EnumType.OECD304); // (S: CE_XX; Cells: L)

                constEntity.IncorpCountryCode = GetCountryCodeEnumType(GetExcelStringValue(package, workbookName, "F" + rowNumber)); // CountryCode_Type.AO; // (S: CE_XX; Cells: F)
                constEntity.IncorpCountryCodeSpecified = true;
                constEntity.OtherEntityInfo = GetExcelStringValue(package, workbookName, "I" + rowNumber); // ""; // (S: CE_XX; Cells: I)
                constEntity.Role = GetUltimateParentEntityRoleEnumType(GetExcelStringValue(package, workbookName, "M" + rowNumber)); // UltimateParentEntityRole_EnumType.CBC803; // (S: CE_XX; Cells: M)
                constEntity.RoleSpecified = true;

                constEntities.Add(constEntity);

                rowNumber++;
            }
            return constEntities;
        }



        private List<CorrectableAdditionalInfo_Type> GetCorrectableAdditionalInfo(ExcelPackage package)
        {
            var list = new List<CorrectableAdditionalInfo_Type>();

            //for (int ai = 0; ai < 10; ai++)
            //{
            var info = new CorrectableAdditionalInfo_Type();
            info.DocSpec = gloabalDocSpec;



            var otherInformationList = new List<StringMin1Max4000WithLang_Type>();


            int rowNumber = 1;
            while (true)
            {
                var cellValue = GetExcelStringValue(package, "Additional Information", "A" + rowNumber);
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }


                if (cellValue.Length >= 4000)
                {
                    throw new ArgumentException("Other Information not allowed more than 4000 characters in one cell");
                }

                otherInformationList.Add(new StringMin1Max4000WithLang_Type() //  (S: Additional Information; Cells:A)
                {
                    language = LanguageCode_Type.EN,
                    languageSpecified = true,
                    Value = cellValue
                });

                rowNumber++;
            }


            info.OtherInfo = otherInformationList.ToArray();

            ////// TODO Add all note from spreadsheet
            ////var countryCodes = new List<CountryCode_Type>();
            ////for (int cc = 0; cc < 10; cc++)
            ////{
            ////    countryCodes.Add(CountryCode_Type.AO);
            ////    //countryCodes.Add(CountryCode_Type.BW);
            ////}
            ////info.ResCountryCode = countryCodes.ToArray();


            //// info.SummaryRef = new CbcSummaryListElementsType_EnumType[] { CbcSummaryListElementsType_EnumType.CBC601 };

            list.Add(info);
            //}
            return list;
        }



        private CorrectableReportingEntity_Type GetReportingEntity(ExcelPackage package)
        {
            var repEnt = new CorrectableReportingEntity_Type();


            repEnt.Entity = GetOrganisationPartyType(GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B7")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B7)
                GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B8")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B8)
                GetExcelStringValue(package, "CoverPage", "B9"), //"9000051715", // (S: CoverPage; Cells: B9)
                GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B10")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B10)
                GetExcelStringValue(package, "CoverPage", "B11"), //"1918/000095/06", // (S: CoverPage; Cells: B11)
                GetExcelStringValue(package, "CoverPage", "B12"), //"Barloworld Limited", // (S: CoverPage; Cells: B12)
                GetCountryCodeEnumType(GetExcelStringValue(package, "CoverPage", "B13")), // CountryCode_Type.ZA, // (S: CoverPage; Cells: B13)
                GetExcelStringValue(package, "CoverPage", "B14").Split(';'), // new object[] { "61 Katherine Street", "Sandton", "2196" }, // (S: CoverPage; Cells: B14) "61 Katherine Street;Sandton;2196" (Split on ;)
                GetOECDLegalAddressTypeEnumType(GetExcelStringValue(package, "CoverPage", "B15"))); //OECDLegalAddressType_EnumType.OECD304);// (S: CoverPage; Cells: B15)


            repEnt.NameMNEGroup = GetExcelStringValue(package, "CoverPage", "B16"); // "Barloworld Limited"; // (S: CoverPage; Cells: B16)
            repEnt.ReportingRole = GetCbcReportingRoleEnumType(GetExcelStringValue(package, "CoverPage", "B17")); // CbcReportingRole_EnumType.CBC701; // (S: CoverPage; Cells: B17)
            repEnt.ReportingPeriod = new ReportingEntity_TypeReportingPeriod()
            {
                StartDate = new DateTime(GetExcelIntValue(package, "CoverPage", "B18").Value, GetExcelIntValue(package, "CoverPage", "B19").Value, GetExcelIntValue(package, "CoverPage", "B20").Value), // new DateTime(2018, 10, 1),  // (S: CoverPage; Cells: B18, B19, B20)
                EndDate = new DateTime(GetExcelIntValue(package, "CoverPage", "B21").Value, GetExcelIntValue(package, "CoverPage", "B22").Value, GetExcelIntValue(package, "CoverPage", "B23").Value), // new DateTime(2019, 09, 30)  // (S: CoverPage; Cells: B21, B22, B23)
            };

            repEnt.DocSpec = gloabalDocSpec;

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
            object[] address,
            OECDLegalAddressType_EnumType legalAddressType
            )
        {
            var entity = new OrganisationParty_Type();

            // Entity
            entity = new OrganisationParty_Type();

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


            var addr = new Address_Type()
            {
                CountryCode = addressCountryCode,
                Items = address,
                legalAddressType = legalAddressType,
                legalAddressTypeSpecified = true
            };
            entity.Address = new Address_Type[] { addr };


            return entity;

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
                var excep = ex;
                canWriteToLogFile = false;
                throw;
            }
        }

    }

    public class ReceivingCountryClass
    {
        public int RowNumber { get; set; }
        public string CountryCode { get; set; }
    }
}
