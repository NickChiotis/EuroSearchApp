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
        // 1. Ορισμός της λίστας δώρων (Μόνο μία φορά)
        private ObservableCollection<GiftRecord> _giftsList = new ObservableCollection<GiftRecord>();
        public ObservableCollection<GiftRecord> GiftsList
        {
            get => _giftsList;
            set { _giftsList = value; OnPropertyChanged(nameof(GiftsList)); }
        }

        // Το path για το αρχείο των δώρων (βεβαιώσου ότι το όνομα είναι σωστό, π.χ. EuroGifts.xlsx)
        private string giftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");

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

        private int _statusFilterIndex = 0;
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

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "Template.xlsx");
            string tempPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "temp_data.xlsx");

            // 1. Έλεγχος αν υπάρχει το Temp (Προτεραιότητα)
            if (File.Exists(tempPath))
            {
                LoadData(tempPath);
            }
            // 2. Αν δεν υπάρχει, φόρτωσε το κανονικό Template
            else if (File.Exists(defaultPath))
            {
                LoadData(defaultPath);
            }
            else
            {
                MessageBox.Show("Δεν βρέθηκε κανένα αρχείο δεδομένων (Template ή Temp).");
            }

            // Φόρτωση των δώρων στο dropdown και focus στο όνομα
            LoadGifts();
            TxtName.Focus();
        }

        private void LoadGifts()
        {
            try
            {
                if (!File.Exists(giftPath)) return;

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(giftPath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return;

                    int rowCount = ws.Dimension.End.Row;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        GiftsList.Clear();
                        // ΠΡΟΣΘΗΚΗ PLACEHOLDER: Βάζουμε χειροκίνητα την πρώτη επιλογή
                        GiftsList.Add(new GiftRecord { Eidos = "Επιλογή Δώρου" });
                    });

                    for (int row = 2; row <= rowCount; row++) // Ξεκινάμε από 2 για να πηδήξουμε το "ΕΙΔΟΣ"
                    {
                        var val = ws.Cells[row, 1].Value?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(val))
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                GiftsList.Add(new GiftRecord { Eidos = val });
                            });
                        }
                    }
                }
            }
            catch (Exception ex) { /* Handle error */ }
        }

        private void GiftInfo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Έλεγχος αν υπάρχει το αρχείο
                if (!File.Exists(giftPath)) return;

                // 2. Διάβασμα του Excel και γέμισμα του DataTable (όπως το κάναμε πριν)
                DataTable dt = new DataTable();
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(giftPath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return;

                    for (int i = 1; i <= ws.Dimension.End.Column; i++)
                    {
                        string header = ws.Cells[1, i].Value?.ToString()?.Trim() ?? $"Column {i}";
                        dt.Columns.Add(header);
                    }

                    for (int rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int colNum = 1; colNum <= ws.Dimension.End.Column; colNum++)
                        {
                            dr[colNum - 1] = ws.Cells[rowNum, colNum].Value?.ToString()?.Trim() ?? "";
                        }
                        dt.Rows.Add(dr);
                    }
                }

                // 3. ΑΝΟΙΓΜΑ ΤΟΥ ΠΑΡΑΘΥΡΟΥ
                var viewer = new GiftViewerWindow();
                viewer.Owner = this;
                viewer.GiftsGrid.ItemsSource = dt.DefaultView;

                viewer.GiftsGrid.ColumnWidth = new DataGridLength(1, DataGridLengthUnitType.Star);
                // Εδώ είναι το "μαγικό" σημείο 3:
                // Περιμένουμε να κλείσει το παράθυρο. Αν ο χρήστης έκανε διπλό κλικ (DialogResult = true)
                if (viewer.ShowDialog() == true)
                {
                    // Α) Βρίσκουμε ποια γραμμή είναι επιλεγμένη στον ΚΕΝΤΡΙΚΟ πίνακα (MainWindow)
                    var selectedPerson = RecordsGrid.SelectedItem as PersonRecord;

                    if (selectedPerson != null)
                    {
                        // Β) Παίρνουμε το όνομα του δώρου που αποθηκεύτηκε στον Viewer
                        // (Πρέπει να έχεις φτιάξει την ιδιότητα SelectedGiftName στον Viewer - δες παρακάτω)
                        selectedPerson.SelectedGift = viewer.SelectedGiftName;

                        // Γ) Ενημερώνουμε το UI
                        RecordsGrid.Items.Refresh();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // --- Υπόλοιπες Μέθοδοι (Αμετάβλητες) ---

        private void ResetFilters_Click(object sender, RoutedEventArgs e)
        {
            NameQuery = "";
            PhoneQuery = "";
            AfmQuery = "";
            StatusFilterIndex = 0;
        }

        private void LoadData(string filePath)
        {
            try
            {
                var rawList = ExcelLoader.Load(filePath);
                if (rawList == null || rawList.Count == 0) return;

                RecordsView = CollectionViewSource.GetDefaultView(rawList);
                RecordsView.Filter = FilterRecords;
                RefreshFilter();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private bool FilterRecords(object item)
        {
            var person = item as PersonRecord;
            if (person == null) return false;

            // 1. Φίλτρο ΟΝΟΜΑΤΟΣ (Starts With)
            if (!string.IsNullOrWhiteSpace(NameQuery))
            {
                if (string.IsNullOrEmpty(person.Επωνυμία) ||
                    !person.Επωνυμία.StartsWith(NameQuery, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            // 2. Φίλτρο ΑΦΜ (Starts With)
            if (!string.IsNullOrWhiteSpace(AfmQuery))
            {
                if (string.IsNullOrEmpty(person.ΑΦΜ) ||
                    !person.ΑΦΜ.StartsWith(AfmQuery, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            // 3. Φίλτρο Τηλεφώνου (Διορθωμένο StartsWith)
            if (!string.IsNullOrWhiteSpace(PhoneQuery))
            {
                // Καθαρίζουμε αυτό που πληκτρολογεί ο χρήστης
                string cleanQuery = NormalizeDigits(PhoneQuery).Trim();

                if (!string.IsNullOrEmpty(cleanQuery))
                {
                    // Καθαρίζουμε τα τηλέφωνα της εγγραφής από κενά, παύλες κλπ
                    string p1 = !string.IsNullOrEmpty(person.Τηλέφωνο) ? NormalizeDigits(person.Τηλέφωνο) : "";
                    string p2 = !string.IsNullOrEmpty(person.Τηλέφωνο2) ? NormalizeDigits(person.Τηλέφωνο2) : "";

                    // Έλεγχος αν ΞΕΚΙΝΑΕΙ το καθαρό τηλέφωνο με το καθαρό query
                    bool match1 = p1.StartsWith(cleanQuery);
                    bool match2 = p2.StartsWith(cleanQuery);

                    if (!match1 && !match2) return false;
                }
            }

            // 4. Φίλτρο ΚΑΤΑΣΤΑΣΗΣ (Selected/Unselected)
            if (StatusFilterIndex == 1 && !person.Selected) return false;
            if (StatusFilterIndex == 2 && person.Selected) return false;

            return true;
        }

        private void RefreshFilter()
        {
            if (RecordsView != null)
            {
                RecordsView.Refresh();
                int count = 0;
                foreach (var item in RecordsView) count++;
                ResultCount = count.ToString("N0");
            }
        }

        private string _resultCount = "0";
        public string ResultCount
        {
            get { return _resultCount; }
            set { _resultCount = value; OnPropertyChanged(nameof(ResultCount)); }
        }

        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return new string(s.Where(char.IsDigit).ToArray());
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var source = sender as TextBox;
                if (source == TxtName) TxtPhone.Focus();
                else if (source == TxtPhone) TxtAfm.Focus();
                else if (source == TxtAfm) RecordsGrid.Focus();
            }
        }

        private async void Aade_Click(object sender, RoutedEventArgs e)
        {
            string afmFromFilter = TxtAfm.Text.Trim();
            if (string.IsNullOrEmpty(afmFromFilter)) return;

            AadeWindow aadeWin = new AadeWindow(afmFromFilter);
            aadeWin.Owner = this;
            aadeWin.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            if (aadeWin.ShowDialog() == true)
            {
                var sourceList = RecordsView.SourceCollection as List<PersonRecord>;
                if (sourceList == null) return;

                var existingPerson = sourceList.FirstOrDefault(p => p.ΑΦΜ == afmFromFilter);
                if (existingPerson != null) existingPerson.Επωνυμία = aadeWin.FetchedName;
                else sourceList.Add(new PersonRecord { ΑΦΜ = afmFromFilter, Επωνυμία = aadeWin.FetchedName, Selected = true });

                RefreshFilter();
            }
        }

        private void ExportVisible_Click(object sender, RoutedEventArgs e)
        {
            RecordsGrid.CommitEdit();
            var visibleRecords = new List<PersonRecord>();
            foreach (var item in RecordsView) visibleRecords.Add(item as PersonRecord);
            ExportToExcel(visibleRecords, "Export");
        }

        private void ExportToExcel(List<PersonRecord> records, string suffix)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            SaveFileDialog saveFileDialog = new SaveFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx", FileName = $"ΛΙΣΤΑ_{DateTime.Now:dd_MM_yyyy}.xlsx" };

            if (saveFileDialog.ShowDialog() == true)
            {
                using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets.Add("ΠΕΛΑΤΟΛΟΓΙΟ");
                    string[] headers = { "Επιλογή", "Συμμετέχων", "Επωνυμία", "ΑΦΜ", "Τηλέφωνο 1", "ΔΩΡΟ" };
                    for (int i = 0; i < headers.Length; i++) ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var item in records)
                    {
                        ws.Cells[row, 1].Value = item.Selected ? "ΝΑΙ" : "ΟΧΙ";
                        ws.Cells[row, 2].Value = item.Comments;
                        ws.Cells[row, 3].Value = item.Επωνυμία;
                        ws.Cells[row, 4].Value = item.ΑΦΜ;
                        ws.Cells[row, 5].Value = item.Τηλέφωνο;
                        ws.Cells[row, 6].Value = item.SelectedGift; // Εξαγωγή του επιλεγμένου δώρου
                        row++;
                    }
                    package.Save();
                }
                MessageBox.Show("Η εξαγωγή ολοκληρώθηκε!");
            }
        }

        private void SaveToTemp()
        {
            try
            {
                string tempPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Templates", "temp_data.xlsx");
                var allRecords = RecordsView.SourceCollection as IEnumerable<PersonRecord>;
                if (allRecords == null) return;

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("TempData");

                    // ΑΥΤΑ ΤΑ ΟΝΟΜΑΤΑ ΠΡΕΠΕΙ ΝΑ ΕΙΝΑΙ ΙΔΙΑ ΜΕ ΤΟ TEMPLATE.XLSX
                    string[] headers = {
                "Επιλογή",
                "Συμμετέχων",
                "Επωνυμία",
                "Επαφές - Α.Φ.Μ",
                "Τηλέφωνο 1",
                "Τηλέφωνο 2",
                "SelectedGift"
            };

                    for (int i = 0; i < headers.Length; i++) ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var item in allRecords)
                    {
                        ws.Cells[row, 1].Value = item.Selected ? "ΝΑΙ" : "ΟΧΙ";
                        ws.Cells[row, 2].Value = item.Comments;
                        ws.Cells[row, 3].Value = item.Επωνυμία;
                        ws.Cells[row, 4].Value = item.ΑΦΜ;
                        ws.Cells[row, 5].Value = item.Τηλέφωνο;
                        ws.Cells[row, 6].Value = item.Τηλέφωνο2;
                        ws.Cells[row, 7].Value = item.SelectedGift;
                        row++;
                    }
                    package.SaveAs(new FileInfo(tempPath));
                }
            }
            catch { }
        }
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            SaveToTemp(); // Σώζει τα πάντα πριν κλείσει
            base.OnClosing(e);
        }


        private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void Maximize_Click(object sender, RoutedEventArgs e) => WindowState = (WindowState == WindowState.Maximized) ? WindowState.Normal : WindowState.Maximized;
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
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
    }

    public class GiftRecord
    {
        public string Eidos { get; set; }
    }
}