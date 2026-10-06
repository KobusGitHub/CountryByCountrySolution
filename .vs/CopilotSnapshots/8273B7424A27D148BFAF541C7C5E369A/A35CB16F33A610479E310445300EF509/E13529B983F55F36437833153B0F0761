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
using CBC_V2_OTHER_COUNTRIES;
using System.Xml.Serialization;
using CBC_V2.Service.OtherCountries;
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

            if (!string.IsNullOrEmpty(txtSource.Text) && File.Exists(txtSource.Text))
            {
                txtDestFolder.Text = Path.GetDirectoryName(txtSource.Text);
            }
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
       
        private void StartWork(FileInfo newExcelFile)
        {
            // xsd.exe CbcXML_v1.0.1.xsd /Classes oecdtypes_v4.1.xsd /Classes isocbctypes_v1.0.1.xsd
            // xsd.exe CbcXML_v2.0.xsd /Classes isocbctypes_v1.1.xsd /Classes oecdcbctypes_v5.0.xsd

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            if (ckSars.Checked)
            {
                SarsGenerator generator = new SarsGenerator(this.myGuid);
                generator.LogMessageEvent += LogEventHandlerFired;
                generator.GenerateFile(xlsxFile, newExcelFile, xmlFilePath, rad8.Checked);
            }
            else
            {
                OtherCountriesGenerator generator = new OtherCountriesGenerator(this.myGuid);
                generator.LogMessageEvent += LogEventHandlerFired;
                generator.GenerateFile(xlsxFile, newExcelFile, xmlFilePath, rad8.Checked);
            }
        }

        private void LogEventHandlerFired(object sender, string e)
        {
            this.logMessage(e);
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

    }



    public class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }
}
