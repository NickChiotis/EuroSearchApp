using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using System.Threading.Tasks;

namespace EuroSearchApp
{
    // Το μοντέλο για τα αρχεία PDF
    public class StatementFile
    {
        public string FileName { get; set; }
        public string FullPath { get; set; }
        public string DateCreated { get; set; }
    }

    public partial class StatementDashboard : Window
    {
        // Λίστα που κρατάει όλα τα αρχεία για να μπορούμε να φιλτράρουμε
        private List<StatementFile> allFiles = new List<StatementFile>();

        public StatementDashboard()
        {
            InitializeComponent();
            LoadPdfFiles();
        }

        private void LoadPdfFiles()
        {
            try
            {
                // Φάκελος EuroStatementPdf στο directory της εφαρμογής
                string folderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EuroStatementPdf");

                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                // Διάβασμα των PDF και αποθήκευση στη λίστα allFiles
                allFiles = Directory.GetFiles(folderPath, "*.pdf")
                    .Select(f => new StatementFile
                    {
                        FileName = System.IO.Path.GetFileName(f),
                        FullPath = f,
                        DateCreated = File.GetCreationTime(f).ToString("dd/MM/yyyy HH:mm")
                    })
                    .OrderByDescending(x => x.DateCreated)
                    .ToList();

                // Εμφάνιση στο DataGrid
                DtgFiles.ItemsSource = allFiles;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Σφάλμα φόρτωσης: {ex.Message}");
            }
        }

        // --- Η ΜΕΘΟΔΟΣ ΓΙΑ ΤΗΝ ΑΝΑΖΗΤΗΣΗ (Αυτό έλειπε) ---
        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (allFiles == null) return;

            string searchText = TxtSearch.Text.ToLower();

            // Φιλτράρουμε τη λίστα με βάση το όνομα του αρχείου
            var filteredList = allFiles
                .Where(f => f.FileName.ToLower().Contains(searchText))
                .ToList();

            DtgFiles.ItemsSource = filteredList;
        }

        // Άνοιγμα με διπλό κλικ
        private void DtgFiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DtgFiles.SelectedItem is StatementFile selected)
            {
                OpenFile(selected.FullPath);
            }
        }

        // Άνοιγμα με το κουμπί στη στήλη "Ενέργειες"
        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button?.DataContext is StatementFile selected)
            {
                OpenFile(selected.FullPath);
            }
        }


        private void OpenFile(string path)
        {
            if (System.IO.File.Exists(path))
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("Το αρχείο δεν βρέθηκε πλέον στον φάκελο.");
            }
        }
        private async void AsyncBtnAddNewStatement_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Άνοιγμα φόρμας στοιχείων
                var infoWin = new LeadInfoWindow();
                infoWin.Owner = this;

                if (infoWin.ShowDialog() == true)
                {
                    // 2. Άνοιγμα παραθύρου υπογραφής
                    var sigWin = new SignatureWindow();
                    sigWin.Owner = this;

                    if (sigWin.ShowDialog() == true)
                    {
                        // 3. Προετοιμασία φακέλου και ονόματος αρχείου
                        string folderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EuroStatementPdf");
                        if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

                        string finalProgram = (infoWin.ComboProgram.SelectedItem == infoWin.ItemOther)
                          ? infoWin.TxtOtherProgram.Text.Trim()
                          : infoWin.Program;

                        if (string.IsNullOrEmpty(finalProgram)) finalProgram = "Μη καθορισμένο";

                        // Όνομα αρχείου: Φαρμακείο_Ημερομηνία.pdf
                        string fileName = $"{infoWin.TxtPharmacy.Text}_{DateTime.Now:yyyyMMdd_HHmm}.pdf";
                        string fullDestPath = System.IO.Path.Combine(folderPath, fileName);

                        // 4. Δημιουργία PDF με όλα τα πεδία (στυλ Υπεύθυνης Δήλωσης)
                        PDFGenerator generator = new PDFGenerator();
                        generator.CreatePdfWithSignature(
                            fullDestPath,
                            infoWin.Pharmacy,      // Αντί για infoWin.TxtPharmacy.Text
                            infoWin.Promoter,      // Χρησιμοποιεί το SelectionBoxItem αυτόματα
                            infoWin.City,          // Αντί για infoWin.TxtCity.Text
                            infoWin.Phone,         // Αντί για infoWin.TxtPhone.Text
                            infoWin.Email,         // Αντί για infoWin.TxtEmail.Text
                            finalProgram,       // Αντί για infoWin.TxtProgram.Text
                            infoWin.Client,        // Χρησιμοποιεί το SelectionBoxItem (ΝΑΙ/ΟΧΙ)
                            infoWin.Presentation,  // Χρησιμοποιεί το SelectionBoxItem (ΝΑΙ/ΟΧΙ)
                            infoWin.Sales,         // Χρησιμοποιεί το SelectionBoxItem (ΝΑΙ/ΟΧΙ)
                            infoWin.Notes,         // Αντί για infoWin.TxtNotes.Text
                            sigWin.IsConsentChecked,
                            sigWin.SignCanvas
                        );

                        // 4.5 ΑΠΟΣΤΟΛΗ ΣΤΟ GOOGLE DRIVE
                        try
                        {
                            // Δημιουργούμε το αντικείμενο του Drive Vault
                            GoogleDriveVault driveVault = new GoogleDriveVault();

                            // Ξεκινάει το ανέβασμα
                            await driveVault.UploadFileToDrive(fullDestPath);

                            // Αν φτάσει εδώ, πέτυχε
                            MessageBox.Show("Η αναφορά ανέβηκε επιτυχώς στο Google Drive!", "Drive Sync", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        catch (Exception ex)
                        {
                            // Αν κοπεί το ίντερνετ ή γίνει λάθος, το PDF υπάρχει ήδη τοπικά, οπότε απλά προειδοποιούμε
                            MessageBox.Show("Η αναφορά σώθηκε τοπικά, αλλά απέτυχε η αποστολή στο Drive.\nΣφάλμα: " + ex.Message,
                                            "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }

                        // 5. Ανανέωση λίστας (αυτό το έχεις ήδη)
                        LoadPdfFiles();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Σφάλμα κατά τη διαδικασία: {ex.Message}");
            }
        }
    }
}