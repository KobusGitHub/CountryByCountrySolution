using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Schema;

namespace CBC_V2
{
    public class XmlValidatorService
    {
        public static void ValidateXml(string xmlFilePath)
        {
            var path = new Uri(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)).LocalPath;
            XmlSchemaSet schema = new XmlSchemaSet();
            schema.Add("urn:oecd:ties:cbc:v2", path + "\\xsdFiles\\CbcXML_v2.0.xsd");
            schema.Add("urn:oecd:ties:isocbctypes:v1", path + "\\xsdFiles\\isocbctypes_v1.1.xsd");
            schema.Add("urn:oecd:ties:cbcstf:v5", path + "\\xsdFiles\\oecdcbctypes_v5.0.xsd");
            // XmlReader rd = XmlReader.Create(xmlFilePath);
            //XDocument doc = XDocument.Load(rd);


            using (StreamReader s = new StreamReader(xmlFilePath, true))
            {
                XDocument xdoc = XDocument.Load(s);
                xdoc.Validate(schema, ValidationEventHandler);
            }


            //doc.Validate(schema, ValidationEventHandler);
        }

        private static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            XmlSeverityType type = XmlSeverityType.Warning;
            if (Enum.TryParse<XmlSeverityType>("Error", out type))
            {
                if (type == XmlSeverityType.Error)
                {
                    // MessageBox.Show(e.Message);
                    throw new Exception(e.Message);
                }
            }
        }
    }
}
