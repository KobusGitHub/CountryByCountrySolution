using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Xml;
using System.IO;

namespace CBC_V2
{
    public static class XmlConverter
    {

        public static string GetUtf8XmlStringAllCountries(CBC_V2_ALL_COUNTRIES.CBC_OECD cbcfd)
        {
            var serializer = new XmlSerializer(typeof(CBC_V2_ALL_COUNTRIES.CBC_OECD));
            var xml = "";
            using (StringWriter writer = new Utf8StringWriter())
            {
                serializer.Serialize(writer, cbcfd);
                xml = writer.ToString();
            }
            return xml;
        }

        public static string GetUtf16XmlStringAllCountries(CBC_V2_ALL_COUNTRIES.CBC_OECD cbcfd)
        {
            var xml = "";
            XmlSerializer xsSubmit = new XmlSerializer(typeof(CBC_V2_ALL_COUNTRIES.CBC_OECD));
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

            return xml;

        }


       



    }
}
