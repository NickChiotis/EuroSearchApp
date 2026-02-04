using System;
using System.IO;
using System.Net;
using System.Text;
using System.Xml.Linq;

namespace EuroSearchApp.Services
{
    public class AadeResult
    {
        public bool Success { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string ErrorMessage { get; set; }
    }

    public static class AadeService
    {
        private static string AadeUsername => Properties.Settings.Default.AadeUser;
        private static string AadePassword => Properties.Settings.Default.AadePass;
        private static string MyAfm => Properties.Settings.Default.MyAfm; // Το ΑΦΜ αυτού που καλεί

        public static AadeResult GetDetails(string afmToSearch)
        {
            var result = new AadeResult();

            try
            {
                // 1. Το XML που ζητάει η ΑΑΔΕ (SOAP Request)
                string soapRequest = $@"<env:Envelope xmlns:env=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns1=""http://gr/gsis/rgwspublic/RgWsPublic.wsdl"" xmlns:ns2=""http://xml.apache.org/xml-soap"" xmlns:ns3=""http://gr/gsis/rgwspublic/RgWsPublicDefinitions.xsd"">
                <env:Header/>
                    <env:Body>
                    <ns1:rgWsPublicAfmMethod>
                        <ns1:auth>
                        <ns3:username>{AadeUsername}</ns3:username>
                        <ns3:password>{AadePassword}</ns3:password>
                        </ns1:auth>
                        <ns1:afm_called_by>999977386</ns1:afm_called_by> <ns1:afm_called_for>{afmToSearch}</ns1:afm_called_for>
                    </ns1:rgWsPublicAfmMethod>
                    </env:Body>
                </env:Envelope>";

                // 2. Ρυθμίζουμε την κλήση στο Server
                var request = (HttpWebRequest)WebRequest.Create("https://www1.gsis.gr/webtax2/wsgsis/RgWsPublic/RgWsPublicPort");
                request.Headers.Add("SOAPAction", "");
                request.ContentType = "text/xml; charset=\"utf-8\"";
                request.Method = "POST";

                // Στέλνουμε το XML
                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    streamWriter.Write(soapRequest);
                }

                // 3. Παίρνουμε την απάντηση
                using (var response = request.GetResponse())
                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    string xmlResponse = streamReader.ReadToEnd();

                    // 4. Διαβάζουμε το XML της απάντησης
                    XDocument doc = XDocument.Parse(xmlResponse);

                    // Ψάχνουμε τα πεδία μέσα στο XML (Namespaces are tricky in SOAP)
                    // Κάνουμε μια απλή αναζήτηση με LocalName για να μην μπλέκουμε με namespaces
                    var onomasia = GetValue(doc, "onomasia");
                    var address = GetValue(doc, "postal_address");
                    var city = GetValue(doc, "postal_area_description");

                    // Αν υπάρχει error από την ΑΑΔΕ
                    var errorDescr = GetValue(doc, "pErrorRec_out");

                    if (!string.IsNullOrEmpty(onomasia))
                    {
                        result.Success = true;
                        result.Name = onomasia;
                        result.Address = $"{address}, {city}".Trim(',', ' ');
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = "Δεν βρέθηκαν στοιχεία ή το ΑΦΜ είναι λάθος.";
                        if (!string.IsNullOrEmpty(errorDescr)) result.ErrorMessage += $" ({errorDescr})";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = "Σφάλμα σύνδεσης: " + ex.Message;
            }

            return result;
        }

        private static string GetValue(XDocument doc, string tag)
        {
            foreach (var element in doc.Descendants())
            {
                if (element.Name.LocalName == tag)
                    return element.Value;
            }
            return "";
        }
    }
}