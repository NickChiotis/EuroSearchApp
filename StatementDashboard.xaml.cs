using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;

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
        private void BtnAddNewStatement_Click(object sender, RoutedEventArgs e)
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

                        // Όνομα αρχείου: Φαρμακείο_Ημερομηνία.pdf
                        string fileName = $"{infoWin.TxtPharmacy.Text}_{DateTime.Now:yyyyMMdd_HHmm}.pdf";
                        string fullDestPath = System.IO.Path.Combine(folderPath, fileName);

                        // 4. Δημιουργία PDF με όλα τα πεδία (στυλ Υπεύθυνης Δήλωσης)
                        PDFGenerator generator = new PDFGenerator();
                        generator.CreatePdfWithSignature(
                            fullDestPath,
                            infoWin.TxtPharmacy.Text,    // ΕΠΩΝΥΜΙΑ ΦΑΡΜΑΚΕΙΟΥ [cite: 25]
                            infoWin.TxtSales.Text,       // PROMOTER [cite: 24]
                            infoWin.TxtCity.Text,        // ΠΟΛΗ 
                            infoWin.TxtPhone.Text,       // ΤΗΛΕΦΩΝΟ 
                            infoWin.TxtEmail.Text,       // e-mail 
                            infoWin.TxtProgram.Text,     // Πρόγραμμα 
                            infoWin.TxtClient.Text,      // Πελάτης 
                            infoWin.TxtPresentation.Text,// PRESENTATION 
                            infoWin.TxtSales.Text,       // SALES 
                            infoWin.TxtNotes.Text,       // ΠΑΡΑΤΗΡΗΣΕΙΣ [cite: 28]
                            sigWin.SignCanvas            // Υπογραφή 
                        );

                        // 5. Ανανέωση λίστας
                        LoadPdfFiles();

                        MessageBox.Show($"Το έγγραφο δημιουργήθηκε επιτυχώς!", "Ολοκλήρωση", MessageBoxButton.OK, MessageBoxImage.Information);
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