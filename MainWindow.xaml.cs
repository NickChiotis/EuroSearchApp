using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using EuroSearchApp.Models;
using EuroSearchApp.Services;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Microsoft.Win32;
using System.Windows.Interop;
using WinForms = System.Windows.Forms;
using System.Windows.Threading;
using System.Data;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EuroSearchApp
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private ICollectionView _recordsView;
        public ICollectionView RecordsView
        {
            get { return _recordsView; }
            set
            {
                _recordsView = value;
                OnPropertyChanged(nameof(RecordsView));
            }
        }

        // --- Πεδία Αναζήτησης ---
        private string _nameQuery = "";
        public string NameQuery
        {
            get { return _nameQuery; }
            set
            {
                _nameQuery = value;
                OnPropertyChanged(nameof(NameQuery));
                RefreshFilter();
            }
        }

        private string _phoneQuery = "";
        public string PhoneQuery
        {
            get { return _phoneQuery; }
            set
            {
                _phoneQuery = value;
                OnPropertyChanged(nameof(PhoneQuery));
                RefreshFilter();
            }
        }

        private string _afmQuery = "";
        public string AfmQuery
        {
            get { return _afmQuery; }
            set
            {
                _afmQuery = value;
                OnPropertyChanged(nameof(AfmQuery));
                RefreshFilter();
            }
        }

        private int _statusFilterIndex = 0; // 0=Όλα, 1=Checked, 2=Unchecked
        public int StatusFilterIndex
        {
            get { return _statusFilterIndex; }
            set
            {
                if (_statusFilterIndex != value)
                {
                    _statusFilterIndex = value;
                    OnPropertyChanged(nameof(StatusFilterIndex));
                    RefreshFilter();
                }
            }
        }

        private void ResetFilters_Click(object sender, RoutedEventArgs e)
        {
            // Καθαρίζουμε όλα τα πεδία
            NameQuery = "";
            PhoneQuery = "";
            AfmQuery = "";
            StatusFilterIndex = 0; // Επιστροφή στο "Όλα"

            // Επειδή έχουμε κάνει Bindings, το UI θα ενημερωθεί αυτόματα
            // και θα τρέξει και το RefreshFilter μόνο του!
        }

        private bool _forcingKiosk;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "Template.xlsx");

            if (File.Exists(defaultPath))
            {
                LoadData(defaultPath);
            }
            else
            {
                MessageBox.Show($"Δεν βρέθηκε το αρχείο αυτόματης φόρτωσης.\n\n" +
                                $"Έψαξα σε αυτή τη διαδρομή:\n{defaultPath}\n\n" +
                                $"Σιγουρέψου ότι στα Properties του αρχείου στο Visual Studio " +
                                $"το 'Copy to Output Directory' είναι 'Copy always'.",
                                "Το αρχείο λείπει", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            TxtName.Focus();
        }

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            // Ελέγχουμε αν πατήθηκε το ENTER
            if (e.Key == Key.Enter)
            {
                // Βλέπουμε ποιο TextBox το κάλεσε
                var source = sender as TextBox;

                if (source == TxtName)
                {
                    // Από Όνομα -> Τηλέφωνο
                    TxtPhone.Focus();
                    TxtPhone.SelectAll(); // Επιλέγει το κείμενο για γρήγορη αντικατάσταση
                }
                else if (source == TxtPhone)
                {
                    // Από Τηλέφωνο -> ΑΦΜ
                    TxtAfm.Focus();
                    TxtAfm.SelectAll();
                }
                else if (source == TxtAfm)
                {
                    // Από ΑΦΜ -> Πίνακας Αποτελεσμάτων (ή κρύψιμο πληκτρολογίου)
                    RecordsGrid.Focus();
                }
            }
        }

        // ΠΡΟΣΟΧΗ: Πρόσθεσε το 'async' πριν το void γιατί θα κάνουμε κλήση στο δίκτυο
        private async void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            // 1. Δημιουργία του παραθύρου ρυθμίσεων
            AadeSettingsWindow settingsWin = new AadeSettingsWindow();
            settingsWin.Owner = this;
            settingsWin.Topmost = true;
            settingsWin.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            // 2. Εμφάνιση παραθύρου (ShowDialog)
            // Αν ο χρήστης πατήσει "Αποθήκευση", το αποτέλεσμα είναι true
            if (settingsWin.ShowDialog() == true)
            {
                // 3. ΛΟΓΙΚΗ VALIDATION (Όπως στο παράδειγμά σου)
                // Μόλις κλείσει το παράθυρο με επιτυχία, κάνουμε αυτόματη δοκιμή σύνδεσης.

                // Χρησιμοποιούμε το ΑΦΜ της ΓΓΠΣ (999977386) για το τεστ, όπως ακριβώς στον κώδικά σου.
                string testAfm = "999977386";

                // Δείχνουμε έναν κέρσορα αναμονής γιατί μπορεί να πάρει 1-2 δευτερόλεπτα
                Mouse.OverrideCursor = Cursors.Wait;

                try
                {
                    // Κάνουμε την κλήση στην ΑΑΔΕ (χρησιμοποιώντας τα νέα settings που μόλις σώθηκαν)
                    var result = await System.Threading.Tasks.Task.Run(() => AadeService.GetDetails(testAfm));

                    Mouse.OverrideCursor = null; // Επαναφορά κέρσορα

                    if (result.Success)
                    {
                        // Αντιστοιχεί στο: activeService = valid;
                        MessageBox.Show("Οι ρυθμίσεις αποθηκεύτηκαν και η υπηρεσία ΑΑΔΕ είναι ΕΝΕΡΓΗ!",
                                        "Επιτυχής Σύνδεση", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Αν οι κωδικοί είναι λάθος
                        MessageBox.Show($"Οι ρυθμίσεις αποθηκεύτηκαν, αλλά ο έλεγχος σύνδεσης απέτυχε.\n\nΑιτία: {result.ErrorMessage}",
                                        "Πρόβλημα Σύνδεσης", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    Mouse.OverrideCursor = null;
                    MessageBox.Show($"Σφάλμα κατά τον έλεγχο: {ex.Message}");
                }
            }
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            // Κρύβει το παράθυρο στη γραμμή εργασιών
            WindowState = WindowState.Minimized;
        }

        private void Maximize_Click(object sender, RoutedEventArgs e)
        {
            // Ελέγχουμε: Αν είναι ήδη μεγάλο -> το κάνουμε κανονικό.
            // Αν είναι κανονικό -> το κάνουμε μεγάλο.
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
            }
            else
            {
                WindowState = WindowState.Maximized;
            }
        }   

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void LoadData(string filePath)
        {
            try
            {
                var rawList = ExcelLoader.Load(filePath);

                if (rawList == null || rawList.Count == 0)
                {
                    MessageBox.Show("Το αρχείο είναι κενό ή δεν φορτώθηκε σωστά.");
                    return;
                }

                RecordsView = CollectionViewSource.GetDefaultView(rawList);
                RecordsView.Filter = FilterRecords;

                RefreshFilter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Σφάλμα κατά τη φόρτωση του αρχείου: {ex.Message}");
            }
        }

        private bool FilterRecords(object item)
        {
            var person = item as PersonRecord;
            if (person == null) return false;
            // -------------------------

            // 1. Φίλτρο Ονόματος
            if (!string.IsNullOrWhiteSpace(NameQuery))
            {
                if (string.IsNullOrEmpty(person.Επωνυμία) ||
                    person.Επωνυμία.IndexOf(NameQuery, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            // 2. Φίλτρο ΑΦΜ
            if (!string.IsNullOrWhiteSpace(AfmQuery))
            {
                // Χρησιμοποιούμε StartsWith που είναι πιο γρήγορο για ΑΦΜ
                // Αν δεν υπάρχει ΑΦΜ ή δεν ξεκινάει με τα ψηφία που έγραψες -> false
                if (string.IsNullOrEmpty(person.ΑΦΜ) ||
                    !person.ΑΦΜ.StartsWith(AfmQuery, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            // 3. Φίλτρο Τηλεφώνου
            if (!string.IsNullOrWhiteSpace(PhoneQuery))
            {
                string cleanQuery = NormalizeDigits(PhoneQuery);

                if (!string.IsNullOrEmpty(cleanQuery))
                {
                    // Ελέγχουμε αν είναι null πριν καλέσουμε το Normalize
                    string p1 = person.Τηλέφωνο != null ? NormalizeDigits(person.Τηλέφωνο) : "";
                    string p2 = person.Τηλέφωνο2 != null ? NormalizeDigits(person.Τηλέφωνο2) : "";

                    bool match1 = p1.Contains(cleanQuery);
                    bool match2 = p2.Contains(cleanQuery);

                    if (!match1 && !match2) return false;
                }
            }

            // 4. Φίλτρο Κατάστασης
            if (StatusFilterIndex == 1 && !person.Selected) return false;
            if (StatusFilterIndex == 2 && person.Selected) return false;

            return true;
        }

        private void RefreshFilter()
        {
            if (RecordsView != null)
            {
                RecordsView.Refresh();

                // Υπολογισμός πλήθους ορατών εγγραφών
                int count = 0;
                foreach (var item in RecordsView) count++;

                ResultCount = count.ToString("N0"); // Το "N0" βάζει τελείες στις χιλιάδες (π.χ. 1.500)
            }
        }

        private string _resultCount = "0";
        public string ResultCount
        {
            get { return _resultCount; }
            set
            {
                _resultCount = value;
                OnPropertyChanged(nameof(ResultCount));
            }
        }

        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return " - ";

            // Χρήση unsafe pointer-like logic με char array για ταχύτητα
            char[] buffer = new char[s.Length];
            int idx = 0;

            foreach (char c in s)
            {
                if (char.IsDigit(c))
                {
                    buffer[idx++] = c;
                }
            }

            return new string(buffer, 0, idx);
        }

        private void RecordsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
        }

        protected override void OnStateChanged(EventArgs e)
        {
            base.OnStateChanged(e);
        }



        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void Aade_Click(object sender, RoutedEventArgs e)
        {
            // 1. Παίρνουμε το ΑΦΜ από το πεδίο αναζήτησης
            string afmFromFilter = TxtAfm.Text.Trim();

            if (string.IsNullOrEmpty(afmFromFilter))
            {
                MessageBox.Show("Παρακαλώ πληκτρολογήστε ένα ΑΦΜ στο πεδίο αναζήτησης.", "Λείπει ΑΦΜ", MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtAfm.Focus();
                return;
            }

            // 2. Ανοίγουμε το παράθυρο της ΑΑΔΕ (χωρίς να μας νοιάζει αν υπάρχει ο πελάτης)
            AadeWindow aadeWin = new AadeWindow(afmFromFilter);

            // --- ΠΡΟΣΘΕΣΕ ΑΥΤΕΣ ΤΙΣ 3 ΓΡΑΜΜΕΣ ---
            aadeWin.Owner = this; // 1. Συνδέει τα παράθυρα ώστε να μην χάνεται από πίσω
            aadeWin.WindowStartupLocation = WindowStartupLocation.CenterOwner; // 2. Κεντράρισμα
            // ------------------------------------

            if (aadeWin.ShowDialog() == true)
            {
                // Ο χρήστης πάτησε "Αποθήκευση" στο παράθυρο της ΑΑΔΕ.
                // Τώρα πρέπει να βρούμε αν αυτό το ΑΦΜ υπάρχει ήδη στη λίστα μας.

                // Παίρνουμε την αρχική λίστα δεδομένων
                var sourceList = RecordsView.SourceCollection as List<PersonRecord>;

                if (sourceList == null) return;

                // Ψάχνουμε αν υπάρχει ήδη εγγραφή με αυτό το ΑΦΜ
                var existingPerson = sourceList.FirstOrDefault(p => p.ΑΦΜ == afmFromFilter);

                if (existingPerson != null)
                {
                    // ΠΕΡΙΠΤΩΣΗ 1: Υπάρχει ήδη -> Τον ενημερώνουμε
                    existingPerson.Επωνυμία = aadeWin.FetchedName;
                    // existingPerson.Διεύθυνση = aadeWin.FetchedAddress; // Αν έχεις διεύθυνση

                    MessageBox.Show($"Ο πελάτης βρέθηκε στη λίστα και ενημερώθηκε!\n\nΕπωνυμία: {existingPerson.Επωνυμία}", "Ενημέρωση Υπάρχοντος", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    // ΠΕΡΙΠΤΩΣΗ 2: Δεν υπάρχει -> Φτιάχνουμε ΝΕΟ
                    var newPerson = new PersonRecord
                    {
                        ΑΦΜ = afmFromFilter,
                        Επωνυμία = aadeWin.FetchedName,
                        Τηλέφωνο = "", // Κενό ή μπορείς να βάλεις "N/A"
                        Comments = "Προστέθηκε από ΑΑΔΕ",
                        Selected = true // Τον επιλέγουμε για να φαίνεται έντονα
                    };

                    // Τον προσθέτουμε στη λίστα
                    sourceList.Add(newPerson);

                    MessageBox.Show($"Δημιουργήθηκε νέα εγγραφή!\n\nΕπωνυμία: {newPerson.Επωνυμία}", "Νέος Πελάτης", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                // 3. Σημαντικό: Ανανεώνουμε τον πίνακα για να φανεί η αλλαγή (ή η νέα εγγραφή)
                RefreshFilter();
            }
        }

        // --- ΒΟΗΘΗΤΙΚΗ: Παίρνει μόνο τα ορατά ---
        private List<PersonRecord> GetVisibleRecords()
        {
            var list = new List<PersonRecord>();
            if (RecordsView != null)
            {
                foreach (var item in RecordsView)
                {
                    if (item is PersonRecord record)
                    {
                        list.Add(record);
                    }
                }
            }
            return list;
        }

        // --- ΤΟ ΝΕΟ ΚΟΥΜΠΙ: Ένα κλικ = Εξαγωγή όσων βλέπεις ---
        private void ExportVisible_Click(object sender, RoutedEventArgs e)
        {
            RecordsGrid.CommitEdit(); // Αποθήκευση σχολίων
            RecordsGrid.CommitEdit();

            var visibleRecords = GetVisibleRecords();
            ExportToExcel(visibleRecords, "Export");
        }

        private void ExportToExcel(List<PersonRecord> records, string suffix)
        {
            if (records.Count == 0)
            {
                MessageBox.Show("Δεν υπάρχουν εγγραφές για εξαγωγή (με βάση τα φίλτρα σας).", "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"ΛΙΣΤΑ_ΕΠΩΝΥΜΙΩΝ_{DateTime.Now:dd_MM_yyyy_HH-mm}.xlsx",
                Title = "Εξαγωγή για Pylon (Pro)"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    if (File.Exists(saveFileDialog.FileName)) File.Delete(saveFileDialog.FileName);

                    using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                    {
                        var ws = package.Workbook.Worksheets.Add("ΠΕΛΑΤΟΛΟΓΙΟ");

                        // --- 1. ΕΠΙΚΕΦΑΛΙΔΕΣ ---
                        string[] headers = {
                            "Επιλογή", "Όνομα Συμμετέχοντος", "Επωνυμία", "Επαφές - Α.Φ.Μ", "Τηλέφωνο 1", "Διακριτικός Τίτλος",
                            "Έγινε Επίδειξη", "Πόλεις (Μεγέθυνση) - Όνομα", "Επαφές - Ημερομηνία 1",
                            "E-mail 1", "E-mail 2", "Τηλέφωνο 2", "GDPR", "Παρουσίαση TWO", 
                            "Παλιός Πελάτης", "ΠΡΟΓΡΑΜΜΑ"
                        };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            ws.Cells[1, i + 1].Value = headers[i];
                        }

                        // --- 2. ΓΕΜΙΣΜΑ ΔΕΔΟΜΕΝΩΝ ---
                        int row = 2;
                        foreach (var item in records)
                        {
                            ws.Cells[row, 1].Value = item.Selected ? "ΝΑΙ" : "ΟΧΙ";
                            ws.Cells[row, 2].Value = item.Comments;
                            ws.Cells[row, 3].Value = item.Επωνυμία;
                            ws.Cells[row, 4].Value = item.ΑΦΜ;
                            ws.Cells[row, 11].Value = item.Τηλέφωνο;
                            ws.Cells[row, 15].Value = item.Τηλέφωνο2;

                            row++;
                        }

                        // --- 3. PRO ΜΟΡΦΟΠΟΙΗΣΗ ---
                        var dataRange = ws.Cells[1, 1, row - 1, headers.Length];

                        var table = ws.Tables.Add(dataRange, "PylonData");
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                        table.ShowFilter = true;

                        ws.Cells.Style.Font.Name = "Segoe UI";
                        ws.Cells.Style.Font.Size = 10;

                        ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Column(11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ws.Cells.AutoFitColumns();
                        for (int i = 1; i <= headers.Length; i++)
                        {
                            ws.Column(i).Width = ws.Column(i).Width + 2;
                        }

                        ws.View.FreezePanes(2, 1);

                        package.Save();
                    }

                    MessageBox.Show($"Εγινε εξαγωγή Excel με!\n{saveFileDialog.FileName}", "Επιτυχία", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Σφάλμα: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}