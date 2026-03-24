using iText.Html2pdf;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
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
                    <meta charset='UTF-8'>
                    <style>
                        body {{ font-family: 'Arial', sans-serif; color: #333; margin: 0; padding: 0; }}
                        .wrapper {{ width: 95%; max-width: 700px; margin: 20px auto; }}
        
                        .header-table {{ width: 100%; border-bottom: 3px solid #0A3D70; margin-bottom: 20px; padding-bottom: 10px; }}
                        .brand-title {{ color: #0A3D70; font-size: 20px; font-weight: bold; text-align: left; }}
                        .brand-subtitle {{ color: #666; font-size: 10px; text-align: right; }}

                        .info-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
                        .info-table th {{ 
                            background-color: #F1F5F9; 
                            color: #475569; 
                            text-align: left; 
                            padding: 10px; 
                            font-size: 11px; 
                            border: 1px solid #CBD5E0;
                            width: 30%;
                        }}
                        .info-table td {{ 
                            padding: 10px; 
                            font-size: 13px; 
                            border: 1px solid #CBD5E0; 
                        }}

                        .consent-box {{ 
                            width: 100%; 
                            background-color: #F0F9FF; 
                            border: 1px solid #BEE3F8;
                            padding: 15px;
                            margin-bottom: 25px;
                        }}

                        .sig-table {{ width: 100%; margin-top: 30px; }}
                        .sig-line {{ border-bottom: 2px solid #0A3D70; width: 200px; margin: 10px auto; }}

                        .footer {{ 
                            text-align: center; 
                            font-size: 9px; 
                            color: #94A3B8; 
                            margin-top: 30px; 
                            border-top: 1px solid #E2E8F0; 
                            padding-top: 10px; 
                        }}
                        .sig-img {{
                            max - width: 180px; /* Ή όποιο πλάτος θέλεις */
                            height: auto;
        
                            /* ΑΥΤΕΣ ΟΙ ΓΡΑΜΜΕΣ ΚΡΑΤΑΝΕ ΤΗΝ ΥΠΟΓΡΑΦΗ ΚΑΘΑΡΗ */
                            image-rendering: crisp-edges; /* Για Firefox */
                            image-rendering: pixelated;   /* Για Chrome/Safari */
                            -ms-interpolation-mode: nearest-neighbor; /* Για IE */
                        }}
                    </style>
                </head>
                <body>
                    <div class='wrapper'>
                        <table class='header-table'>
                            <tr>
                                <td class='brand-title'>EUROPHARMACY IKE</td>
                                <td>
                                    <div class='brand-subtitle'>ΔΕΛΤΙΟ ΕΠΙΚΟΙΝΩΝΙΑΣ</div>
                                </td>
                            </tr>
                        </table>

                        <table class='info-table'>
                            <tr><th>ΦΑΡΜΑΚΕΙΟ</th><td><strong>{pharmacyName}</strong></td></tr>
                            <tr><th>PROMOTER</th><td>{promoter}</td></tr>
                            <tr><th>ΠΟΛΗ</th><td>{city}</td></tr>
                            <tr><th>ΤΗΛΕΦΩΝΟ</th><td>{phone}</td></tr>
                            <tr><th>E-MAIL</th><td>{email}</td></tr>
                            <tr><th>ΠΕΛΑΤΗΣ</th><td>{client}</td></tr>
                            <tr><th>ΠΡΟΓΡΑΜΜΑ</th><td>{program}</td></tr>
                            <tr><th>PRESENTATION</th><td>{presentation}</td></tr>
                            <tr><th>SALES</th><td>{sales}</td></tr>
                            <tr><th>ΠΑΡΑΤΗΡΗΣΕΙΣ</th><td>{notes}</td></tr>
                        </table>

                        <div class='consent-box'>
                            <table width='100%' border='0' cellspacing='0' cellpadding='0'>
                                <tr>
                                    <td width='30' valign='top' style='font-size: 18px; color: #3182CE;'>☑</td>
                                    <td style='font-size: 11px; color: #2C5282;'>
                                        <strong>ΣΥΝΑΙΝΩ</strong> για την τήρηση των προσωπικών μου στοιχείων από τη <strong>Europharmacy ΙΚΕ</strong> για μελλοντική επικοινωνία με σκοπό την ενημέρωσή μου.
                                    </td>
                                </tr>
                            </table>
                        </div>

                        <table class='sig-table'>
                            <tr>
                                <td width='50%' valign='bottom' style='font-size: 11px; color: #64748B;'>
                                    Ημερομηνία: {DateTime.Now:dd/MM/yyyy}
                                </td>
                                <td width='50%' align='right'>
                                    <div style='text-align: center; width: 220px;'>
                                        <div style='font-size: 12px; font-weight: bold; color: #0A3D70; text-align: right;'>Ο - Η Δηλών/ούσα</div>
                                        <div class='sig-line'>
                                            <img src='data:image/png;base64,{signatureBase64}' style='max-width: 180px; height: auto; 
                                                image-rendering: crisp-edges; 
                                                image-rendering: pixelated; 
                                                -ms-interpolation-mode: nearest-neighbor; 
                                                display: block; 
                                                margin-left: auto; 
                                                margin-right: auto;' />
                                        </div>
                                        <div style='font-size: 9px; color: #94A3B8; text-align: right;'>ΨΗΦΙΑΚΗ ΥΠΟΓΡΑΦΗ</div>
                                    </div>
                                </td>
                            </tr>
                        </table>

                        <div class='footer'>
                            Στην περίπτωση που επιθυμείτε να διαγραφούν τα προσωπικά σας στοιχεία από τη βάση δεδομένων της Europharmacy, μπορείτε να ανακαλέσετε
εγγράφως τη συγκατάθεσή σας ανά πάσα στιγμή.
                        </div>
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
            // 1. Έλεγχος αν υπάρχουν διαστάσεις
            if (canvas.ActualWidth <= 0 || canvas.ActualHeight <= 0) return null;

            // 2. Ορισμός High DPI (300 DPI για κρυστάλλινη ποιότητα)
            double dpi = 300;
            double scale = dpi / 96.0;

            // 3. Υπολογισμός νέων διαστάσεων
            int width = (int)(canvas.ActualWidth * scale);
            int height = (int)(canvas.ActualHeight * scale);

            // 4. Δημιουργία του RenderTargetBitmap με υψηλή ανάλυση
            RenderTargetBitmap rtb = new RenderTargetBitmap(
                width,
                height,
                dpi,
                dpi,
                System.Windows.Media.PixelFormats.Pbgra32);

            // 5. Χρήση VisualBrush για τέλειο rendering των strokes
            DrawingVisual dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                VisualBrush vb = new VisualBrush(canvas);
                dc.DrawRectangle(vb, null, new System.Windows.Rect(new System.Windows.Point(0, 0), new System.Windows.Point(canvas.ActualWidth, canvas.ActualHeight)));
            }

            rtb.Render(dv);

            // 6. Κωδικοποίηση σε PNG
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));

            using (MemoryStream ms = new MemoryStream())
            {
                encoder.Save(ms);
                return ms.ToArray();
            }
        }
    }
}