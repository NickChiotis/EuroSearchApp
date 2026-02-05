using System;
using System.IO;
using System.Net;
using System.Text;
using System.Xml.Linq;
using System.Windows;
using System.Linq;

namespace EuroSearchApp.Services
{
    public class AadeResult
    {
        public bool Success { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Profession { get; set; } // Πρόσθεσα και το επάγγελμα
        public string ErrorMessage { get; set; }
    }

    public static class AadeService
    {
        private static string AadeUsername => Properties.Settings.Default.AadeUser;
        private static string AadePassword => Properties.Settings.Default.AadePass;

        public static AadeResult GetDetails(string afmToSearch)
        {
            var result = new AadeResult();

            try
            {
                // 1. TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string safeUser = AadeUsername?.Trim();
                string safePass = AadePassword?.Trim();
                string safeSearchAfm = afmToSearch?.Trim();

                // 2. Το XML ακριβώς όπως στο Euromedica (String Interpolation)
                string soapRequest = $@"<?xml version=""1.0"" encoding=""UTF-8""?><env:Envelope xmlns:env=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns1=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:ns2=""http://rgwspublic2/RgWsPublic2Service"" xmlns:ns3=""http://rgwspublic2/RgWsPublic2"">" +
                    $"<env:Header>" +
                    $"<ns1:Security>" +
                    $"<ns1:UsernameToken>" +
                    $"<ns1:Username>{safeUser}</ns1:Username>" +
                    $"<ns1:Password>{safePass}</ns1:Password>" +
                    $"</ns1:UsernameToken>" +
                    $"</ns1:Security>" +
                    $"</env:Header>" +
                    $"<env:Body>" +
                    $"<ns2:rgWsPublic2AfmMethod><ns2:INPUT_REC>" +
                    $"<ns3:afm_called_by xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\" />" +
                    $"<ns3:afm_called_for>{safeSearchAfm}</ns3:afm_called_for>" +
                    $"</ns2:INPUT_REC></ns2:rgWsPublic2AfmMethod>" +
                    $"</env:Body>" +
                    $"</env:Envelope>";

                // 3. WebRequest (Βασισμένο στη μέθοδο CreateWebRequest του Euromedica)
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"https://www1.gsis.gr/wsaade/RgWsPublic2/RgWsPublic2");

                // ΠΡΟΣΟΧΗ: Αυτές οι 3 γραμμές είναι το μυστικό που έλειπε
                request.Headers.Add(@"SOAP:Action");
                request.ContentType = "application/soap+xml;charset=UTF-8;action=\"http://rgwspublic2/RgWsPublic2Service:rgWsPublic2AfmMethod\"";
                request.Method = "POST";

                // Αποστολή
                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    streamWriter.Write(soapRequest);
                }

                // Λήψη
                using (var response = request.GetResponse())
                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    string xmlResponse = streamReader.ReadToEnd();

                    // XDocument για parsing (πιο γρήγορο από το XmlDocument + File Save)
                    XDocument doc = XDocument.Parse(xmlResponse);

                    // 4. Έλεγχος για λάθη (όπως στο Euromedica)
                    var errorCode = GetValue(doc, "error_code");
                    var errorDescr = GetValue(doc, "error_descr");

                    if (!string.IsNullOrEmpty(errorCode))
                    {
                        result.Success = false;
                        result.ErrorMessage = errorDescr; // Επιστροφή του μηνύματος της ΑΑΔΕ
                        return result;
                    }

                    // 5. Ανάγνωση στοιχείων (basic_rec)
                    // Ψάχνουμε τα tags ανεξάρτητα από namespace
                    var onomasia = GetValue(doc, "onomasia");
                    var commer_title = GetValue(doc, "commer_title");

                    var address = GetValue(doc, "postal_address");
                    var addressNo = GetValue(doc, "postal_address_no");
                    var zip = GetValue(doc, "postal_zip_code");
                    var city = GetValue(doc, "postal_area_description");
                    var firm_descr = GetValue(doc, "firm_act_descr"); // Επάγγελμα

                    // Λογική ονόματος (Αν έχει Εμπορικό Τίτλο, πάρε αυτόν)
                    if (!string.IsNullOrWhiteSpace(commer_title))
                    {
                        result.Name = commer_title;
                    }
                    else
                    {
                        result.Name = onomasia;
                    }

                    // Σύνθεση Διεύθυνσης
                    string fullAddress = $"{address} {addressNo}".Trim();
                    result.Address = $"{fullAddress}, {zip} {city}".Trim(',', ' ');
                    result.Profession = firm_descr;

                    if (!string.IsNullOrEmpty(result.Name))
                    {
                        result.Success = true;
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = "Δεν βρέθηκαν στοιχεία (Κενό Όνομα).";
                    }
                }
            }
            catch (WebException wex)
            {
                result.Success = false;
                if (wex.Response != null)
                {
                    using (var errorReader = new StreamReader(wex.Response.GetResponseStream()))
                    {
                        string rawError = errorReader.ReadToEnd();
                        // MessageBox.Show("SERVER ERROR:\n" + rawError); // Debug αν χρειαστεί
                        result.ErrorMessage = "Σφάλμα Server ΑΑΔΕ.";
                    }
                }
                else
                {
                    result.ErrorMessage = "Σφάλμα Δικτύου: " + wex.Message;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = "Γενικό Σφάλμα: " + ex.Message;
            }

            return result;
        }

        // Βοηθητική μέθοδος για εύρεση τιμής XML tag
        private static string GetValue(XDocument doc, string tag)
        {
            foreach (var element in doc.Descendants())
            {
                if (element.Name.LocalName.Equals(tag, StringComparison.OrdinalIgnoreCase))
                    return element.Value;
            }
            return "";
        }
    }
}