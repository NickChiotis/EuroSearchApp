using iText.Html2pdf;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace EuroSearchApp
{
    class PDFGenerator
    {
        public void CreatePdfWithSignature(string dest, string pharmacyName, string promoter, string city, string phone, string email, string program, string client, string presentation, string sales, string notes, InkCanvas canvas)
        {
            try
            {
                // 1. Μετατροπή της υπογραφής σε Base64
                string signatureBase64 = "";
                if (canvas != null && canvas.Strokes.Count > 0)
                {
                    byte[] sigBytes = GetCanvasImage(canvas);
                    if (sigBytes != null)
                    {
                        signatureBase64 = Convert.ToBase64String(sigBytes);
                    }
                }

                // 2. Μετατροπή του Λογότυπου σε Base64 (Η ΑΠΟΛΥΤΗ ΛΥΣΗ ΓΙΑ ΝΑ ΕΜΦΑΝΙΣΤΕΙ)
                string logoBase64 = "";
                string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Eurologo.png");
                if (File.Exists(logoPath))
                {
                    byte[] logoBytes = File.ReadAllBytes(logoPath);
                    logoBase64 = Convert.ToBase64String(logoBytes);
                }

                // Ελέγχουμε αν βρέθηκε το λογότυπο, αλλιώς βάζουμε απλό κείμενο για να μην "σπάσει" η εμφάνιση
                string logoHtml = !string.IsNullOrEmpty(logoBase64)
                    ? $"<img src='data:image/png;base64,{logoBase64}' class='logo' />"
                    : "<div class='title'>EUROPHARMACY IKE</div>";

                // 3. ΤΟ PREMIUM DESIGN ΣΕ HTML/CSS
                string htmlContent = $@"
                <html>
                <head>
                    <style>
                        /* Στήσιμο Σελίδας */
                        @page {{ size: A4; margin: 40px; }}
                        body {{ 
                            font-family: 'Helvetica', 'Arial', sans-serif; 
                            color: #333; 
                            background-color: #fff;
                            margin: 0;
                            padding: 0;
                        }}

                        /* Header με το Λογότυπο */
                        .header {{ 
                            border-bottom: 3px solid #0A3D70; /* Βαθύ Μπλε */
                            padding-bottom: 15px; 
                            margin-bottom: 30px; 
                            width: 100%; 
                        }}
                        .header table {{ width: 100%; }}
                        .logo {{ max-height: 65px; width: auto; }}
                        
                        .title-box {{ text-align: right; }}
                        .title {{ color: #0A3D70; font-size: 20px; font-weight: 900; letter-spacing: 1px; text-transform: uppercase; }}
                        .subtitle {{ color: #666; font-size: 11px; margin-top: 4px; letter-spacing: 0.5px; }}

                        /* Κεντρικός Πίνακας (Στυλ Κάρτας) */
                        .content-card {{ 
                            background: #fff; 
                            border-radius: 8px; 
                            padding: 10px; 
                            border: 1px solid #E2E8F0; 
                            margin-bottom: 25px; 
                        }}

                        table.form-table {{ width: 100%; border-collapse: collapse; }}
                        table.form-table td, table.form-table th {{ 
                            padding: 12px 15px; 
                            border-bottom: 1px solid #EDF2F7; 
                            font-size: 12px; 
                            vertical-align: middle;
                        }}
                        table.form-table tr:last-child td, table.form-table tr:last-child th {{ border-bottom: none; }}
                        
                        /* Χρώματα στις Επικεφαλίδες του πίνακα */
                        table.form-table th {{ 
                            background-color: #F0F4F8; /* Πολύ απαλό γαλάζιο/γκρι */
                            color: #0A3D70; 
                            font-weight: bold; 
                            text-align: left; 
                            width: 25%; 
                            border-right: 3px solid #fff; /* Κενό ανάμεσα στα κελιά */
                        }}

                        /* Ενότητα Συναίνεσης */
                        .consent-box {{ 
                            background-color: #F8FAFC; 
                            border-left: 5px solid #007BFF; /* Έντονο μπλε αριστερά */
                            padding: 15px 20px; 
                            border-radius: 4px; 
                            display: table; 
                            width: 100%; 
                            box-sizing: border-box; 
                            margin-bottom: 30px;
                        }}
                        .consent-check {{ 
                            display: table-cell; 
                            vertical-align: middle; 
                            width: 45px; 
                            font-size: 28px; 
                            color: #007BFF; 
                        }}
                        .consent-text {{ 
                            display: table-cell; 
                            vertical-align: middle; 
                            font-size: 11px; 
                            color: #4A5568; 
                            line-height: 1.5; 
                        }}

                        /* Περιοχή Υπογραφής */
                        .signature-area {{ width: 100%; margin-top: 30px; }}
                        .signature-box {{ float: right; width: 250px; text-align: center; }}
                        .sig-date {{ text-align: right; font-size: 11px; margin-bottom: 15px; color: #718096; }}
                        .sig-title {{ font-weight: bold; color: #0A3D70; margin-bottom: 5px; font-size: 13px; }}
                        .sig-img {{ max-width: 160px; max-height: 70px; border-bottom: 1px dashed #CBD5E0; padding-bottom: 5px; margin-bottom: 5px; }}
                        .sig-label {{ font-size: 10px; color: #A0AEC0; }}

                        /* Footer */
                        .footer {{ 
                            position: fixed; 
                            bottom: 10px; 
                            left: 0; 
                            right: 0; 
                            text-align: center; 
                            font-size: 9px; 
                            color: #A0AEC0; 
                            border-top: 1px solid #EDF2F7; 
                            padding-top: 15px; 
                        }}
                    </style>
                </head>
                <body>
                    <div class='header'>
                        <table>
                            <tr>
                                <td style='width: 50%; vertical-align: bottom;'>
                                    {logoHtml}
                               </td>
                                <td style='width: 50%; vertical-align: bottom;' class='title-box'>
                                    <div class='title'>ΔΕΛΤΙΟ ΕΠΙΚΟΙΝΩΝΙΑΣ</div>
                                    <div class='subtitle'>ΥΠΕΥΘΥΝΗ ΔΗΛΩΣΗ ΠΕΛΑΤΗ / PROMOTER</div>
                                </td>
                            </tr>
                        </table>
                    </div>

                    <div class='content-card'>
                        <table class='form-table'>
                            <tr><th>ΕΠΩΝΥΜΙΑ ΦΑΡΜΑΚΕΙΟΥ</th><td colspan='3'><strong>{pharmacyName}</strong></td></tr>
                            <tr>
                                <th>PROMOTER</th><td>{promoter}</td>
                                <th style='width:15%'>ΠΟΛΗ</th><td>{city}</td>
                            </tr>
                            <tr>
                                <th>ΤΗΛΕΦΩΝΟ</th><td>{phone}</td>
                                <th style='width:15%'>E-MAIL</th><td>{email}</td>
                            </tr>
                            <tr><th>ΠΡΟΓΡΑΜΜΑ</th><td colspan='3'>{program}</td></tr>
                            <tr><th>ΠΕΛΑΤΗΣ</th><td colspan='3'>{client}</td></tr>
                            <tr>
                                <th>PRESENTATION</th><td>{presentation}</td>
                                <th style='width:15%'>SALES</th><td>{sales}</td>
                            </tr>
                            <tr><th>ΠΑΡΑΤΗΡΗΣΕΙΣ</th><td colspan='3'>{notes}</td></tr>
                        </table>
                    </div>

                    <div class='consent-box'>
                        <div class='consent-check'>☑</div>
                        <div class='consent-text'>
                            <strong>ΣΥΝΑΙΝΩ</strong> για την τήρηση των προσωπικών μου στοιχείων από τη <strong>Europharmacy ΙΚΕ</strong> για μελλοντική επικοινωνία με σκοπό την ενημέρωσή μου.
                        </div>
                    </div>

                    <div class='signature-area'>
                        <div class='signature-box'>
                            <div class='sig-date'>Ημερομηνία: {DateTime.Now:dd/MM/yyyy}</div>
                            <div class='sig-title'>Ο - Η Δηλ.</div>
                            <img src='data:image/png;base64,{signatureBase64}' class='sig-img' />
                            <div class='sig-label'>(Υπογραφή)</div>
                        </div>
                    </div>

                    <div class='footer'>
                        Στην περίπτωση που επιθυμείτε να διαγραφούν τα προσωπικά σας στοιχεία από τη βάση δεδομένων της Europharmacy, μπορείτε να ανακαλέσετε εγγράφως τη συγκατάθεσή σας ανά πάσα στιγμή.
                    </div>
                </body>
                </html>";

                // 4. Δημιουργία PDF
                using (FileStream pdfDest = new FileStream(dest, FileMode.Create))
                {
                    ConverterProperties converterProperties = new ConverterProperties();
                    HtmlConverter.ConvertToPdf(htmlContent, pdfDest, converterProperties);
                    var driveService = new GoogleDriveVault();
                    Task.Run(() => driveService.UploadFileToDrive(dest));
                }
            }
            catch (Exception ex) { throw new Exception("Σφάλμα Premium Design PDF: " + ex.Message); }
        }

        private byte[] GetCanvasImage(InkCanvas canvas)
        {
            System.Windows.Rect rect = new System.Windows.Rect(canvas.RenderSize);
            if (rect.Width <= 0 || rect.Height <= 0) return null;

            RenderTargetBitmap rtb = new RenderTargetBitmap((int)rect.Right, (int)rect.Bottom, 96d, 96d, System.Windows.Media.PixelFormats.Default);
            rtb.Render(canvas);
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));
            using (MemoryStream ms = new MemoryStream()) { encoder.Save(ms); return ms.ToArray(); }
        }
    }
}