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
        // --- Λίστες και Views ---
        private ObservableCollection<GiftRecord> _giftsList = new ObservableCollection<GiftRecord>();
        public ObservableCollection<GiftRecord> GiftsList
        {
            get => _giftsList;
            set { _giftsList = value; OnPropertyChanged(nameof(GiftsList)); }
        }

        private string giftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");

        private ICollectionView _selectedRecordsView; // Το View που βλέπει ο χρήστης στο DataGrid
        public ICollectionView SelectedRecordsView
        {
            get => _selectedRecordsView;
            set { _selectedRecordsView = value; OnPropertyChanged(nameof(SelectedRecordsView)); }
        }
        private string originalGiftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");
        private string tempGiftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "temp_gift.xlsx");

        private ICollectionView _recordsView; // Το "κρυφό" View για τα Suggestions και το Enter
        public ICollectionView RecordsView
        {
            get { return _recordsView; }
            set { _recordsView = value; OnPropertyChanged(nameof(RecordsView)); }
        }

        // --- Properties για τα Suggestion Popups ---
        private bool _isNamePopupOpen;
        public bool IsNamePopupOpen { get => _isNamePopupOpen; set { _isNamePopupOpen = value; OnPropertyChanged(nameof(IsNamePopupOpen)); } }
        private ObservableCollection<string> _nameSuggestions = new ObservableCollection<string>();
        public ObservableCollection<string> NameSuggestions { get => _nameSuggestions; set { _nameSuggestions = value; OnPropertyChanged(nameof(NameSuggestions)); } }

        private bool _isPhonePopupOpen;
        public bool IsPhonePopupOpen { get => _isPhonePopupOpen; set { _isPhonePopupOpen = value; OnPropertyChanged(nameof(IsPhonePopupOpen)); } }
        private ObservableCollection<string> _phoneSuggestions = new ObservableCollection<string>();
        public ObservableCollection<string> PhoneSuggestions { get => _phoneSuggestions; set { _phoneSuggestions = value; OnPropertyChanged(nameof(PhoneSuggestions)); } }

        private bool _isAfmPopupOpen;
        public bool IsAfmPopupOpen { get => _isAfmPopupOpen; set { _isAfmPopupOpen = value; OnPropertyChanged(nameof(IsAfmPopupOpen)); } }
        private ObservableCollection<string> _afmSuggestions = new ObservableCollection<string>();
        public ObservableCollection<string> AfmSuggestions { get => _afmSuggestions; set { _afmSuggestions = value; OnPropertyChanged(nameof(AfmSuggestions)); } }

        // --- Πεδία Αναζήτησης (Queries) ---
        private string _nameQuery = "";
        public string NameQuery
        {
            get { return _nameQuery; }
            set
            {
                _nameQuery = value;
                OnPropertyChanged(nameof(NameQuery));
                UpdateSuggestions("Name", value);
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
                UpdateSuggestions("Phone", value);
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
                UpdateSuggestions("Afm", value);
                RefreshFilter();
            }
        }

        private int _statusFilterIndex = 0;
        public int StatusFilterIndex
        {
            get { return _statusFilterIndex; }
            set { if (_statusFilterIndex != value) { _statusFilterIndex = value; OnPropertyChanged(nameof(StatusFilterIndex)); RefreshFilter(); } }
        }

        private string _resultCount = "0";
        public string ResultCount
        {
            get { return _resultCount; }
            set { _resultCount = value; OnPropertyChanged(nameof(ResultCount)); }
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

            if (File.Exists(tempPath)) LoadData(tempPath);
            else if (File.Exists(defaultPath)) LoadData(defaultPath);
            else MessageBox.Show("Δεν βρέθηκε κανένα αρχείο δεδομένων (Template ή Temp).");

            LoadGifts();
            TxtName.Focus();
        }

        // --- Logic για Suggestions ---
        private void UpdateSuggestions(string type, string value)
        {
            if (RecordsView == null || RecordsView.SourceCollection == null || string.IsNullOrWhiteSpace(value) || value.Length < 2)
            {
                ClosePopups();
                return;
            }

            var source = RecordsView.SourceCollection as IEnumerable<PersonRecord>;
            if (source == null) return;

            if (type == "Name")
            {
                var matches = source.Where(p => p.Επωνυμία != null && p.Επωνυμία.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
                                    .Select(p => p.Επωνυμία).Distinct().Take(10).ToList();
                NameSuggestions = new ObservableCollection<string>(matches);
                IsNamePopupOpen = NameSuggestions.Any();
            }
            else if (type == "Phone")
            {
                string cleanVal = NormalizeDigits(value);
                var matches = source.Where(p => (p.Τηλέφωνο != null && NormalizeDigits(p.Τηλέφωνο).Contains(cleanVal)) || (p.Τηλέφωνο2 != null && NormalizeDigits(p.Τηλέφωνο2).Contains(cleanVal)))
                                    .Select(p => p.Τηλέφωνο).Distinct().Take(10).ToList();
                PhoneSuggestions = new ObservableCollection<string>(matches);
                IsPhonePopupOpen = PhoneSuggestions.Any();
            }
            else if (type == "Afm")
            {
                var matches = source.Where(p => p.ΑΦΜ != null && p.ΑΦΜ.Contains(value))
                                    .Select(p => p.ΑΦΜ).Distinct().Take(10).ToList();
                AfmSuggestions = new ObservableCollection<string>(matches);
                IsAfmPopupOpen = AfmSuggestions.Any();
            }
        }

        private void ClosePopups()
        {
            IsNamePopupOpen = IsPhonePopupOpen = IsAfmPopupOpen = false;
        }

        private void Suggestion_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var lb = sender as ListBox;
            if (lb?.SelectedItem != null)
            {
                string selectedVal = lb.SelectedItem.ToString();
                if (lb == LstNameSuggestions) { _nameQuery = selectedVal; OnPropertyChanged(nameof(NameQuery)); IsNamePopupOpen = false; }
                else if (lb.Name == "LstPhoneSuggestions") { _phoneQuery = selectedVal; OnPropertyChanged(nameof(PhoneQuery)); IsPhonePopupOpen = false; }
                else if (lb.Name == "LstAfmSuggestions") { _afmQuery = selectedVal; OnPropertyChanged(nameof(AfmQuery)); IsAfmPopupOpen = false; }

                RefreshFilter();
                lb.SelectedItem = null;
            }
        }

        // --- Φιλτράρισμα και Φόρτωση ---
        private void LoadData(string filePath)
        {
            try
            {
                var rawList = ExcelLoader.Load(filePath);
                if (rawList == null || rawList.Count == 0) return;

                // View για την εσωτερική αναζήτηση
                RecordsView = CollectionViewSource.GetDefaultView(rawList);
                RecordsView.Filter = FilterRecords;

                // View για το DataGrid (δείχνει μόνο τα Selected)
                SelectedRecordsView = CollectionViewSource.GetDefaultView(rawList);
                SelectedRecordsView.Filter = (item) => ((PersonRecord)item).Selected;

                RefreshFilter();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private bool FilterRecords(object item)
        {
            var person = item as PersonRecord;
            if (person == null) return false;

            // Αν η εγγραφή είναι ήδη επιλεγμένη, τη δείχνουμε πάντα
            if (person.Selected) return true;

            // Κρύβουμε τα πάντα αν δεν υπάρχει φίλτρο
            bool hasActiveFilter = !string.IsNullOrWhiteSpace(NameQuery) || !string.IsNullOrWhiteSpace(PhoneQuery) || !string.IsNullOrWhiteSpace(AfmQuery);
            if (!hasActiveFilter) return false;

            if (!string.IsNullOrWhiteSpace(NameQuery) && (person.Επωνυμία == null || person.Επωνυμία.IndexOf(NameQuery, StringComparison.OrdinalIgnoreCase) < 0)) return false;
            if (!string.IsNullOrWhiteSpace(AfmQuery) && (person.ΑΦΜ == null || !person.ΑΦΜ.StartsWith(AfmQuery))) return false;
            if (!string.IsNullOrWhiteSpace(PhoneQuery))
            {
                string q = NormalizeDigits(PhoneQuery);
                string p1 = NormalizeDigits(person.Τηλέφωνο);
                string p2 = NormalizeDigits(person.Τηλέφωνο2);
                if (!p1.StartsWith(q) && !p2.StartsWith(q)) return false;
            }

            return true;
        }

        private void RefreshFilter()
        {
            RecordsView?.Refresh();
            SelectedRecordsView?.Refresh();
            if (SelectedRecordsView != null)
            {
                int count = 0;
                foreach (var item in SelectedRecordsView) count++;
                ResultCount = count.ToString("N0");
            }
        }

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (RecordsView == null || RecordsView.SourceCollection == null) return;
                var sourceList = RecordsView.SourceCollection as IEnumerable<PersonRecord>;
                if (sourceList == null) return;

                // Βρίσκουμε την πρώτη εγγραφή που ταιριάζει στα φίλτρα
                var match = sourceList.FirstOrDefault(p => FilterRecords(p));

                if (match != null)
                {
                    match.Selected = true;
                    _nameQuery = _phoneQuery = _afmQuery = "";
                    OnPropertyChanged(nameof(NameQuery)); OnPropertyChanged(nameof(PhoneQuery)); OnPropertyChanged(nameof(AfmQuery));
                    ClosePopups();
                    RefreshFilter();
                    TxtName.Focus();
                }
            }
        }

        // --- Υπόλοιπες Λειτουργίες (AADE, Excel, Save) ---
        private void LoadGifts()
        {
            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                // ΛΟΓΙΚΗ: Αν υπάρχει το Temp διάβασε το, αλλιώς διάβασε το Original
                string fileToRead = File.Exists(tempGiftPath) ? tempGiftPath : originalGiftPath;

                if (!File.Exists(fileToRead)) return;

                using (var package = new ExcelPackage(new FileInfo(fileToRead)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return;

                    int rowCount = ws.Dimension.End.Row;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        GiftsList.Clear();
                        GiftsList.Add(new GiftRecord { Eidos = "Επιλογή Δώρου" });
                    });

                    for (int row = 2; row <= rowCount; row++)
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
            catch { }
        }

        private void SaveToTemp()
        {
            try
            {
                if (RecordsView == null || RecordsView.SourceCollection == null) return;
                var allRecords = RecordsView.SourceCollection as IEnumerable<PersonRecord>;
                if (allRecords == null) return;

                string tempPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Templates", "temp_data.xlsx");
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("TempData");
                    string[] headers = { "Επιλογή", "Συμμετέχων", "Επωνυμία", "Επαφές - Α.Φ.Μ", "Τηλέφωνο 1", "Τηλέφωνο 2", "SelectedGift" };
                    for (int i = 0; i < headers.Length; i++) ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var item in allRecords)
                    {
                        if (item == null) continue;
                        ws.Cells[row, 1].Value = item.Selected ? "ΝΑΙ" : "ΟΧΙ";
                        ws.Cells[row, 2].Value = item.Comments;
                        ws.Cells[row, 3].Value = item.Επωνυμία;
                        ws.Cells[row, 4].Value = item.ΑΦΜ;
                        ws.Cells[row, 5].Value = item.Τηλέφωνο;
                        ws.Cells[row, 6].Value = item.Τηλέφωνο2;
                        ws.Cells[row, 7].Value = item.SelectedGift;
                        row++;
                    }
                    var fileInfo = new FileInfo(tempPath);
                    if (!fileInfo.Directory.Exists) fileInfo.Directory.Create();
                    package.SaveAs(fileInfo);
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("SaveToTemp Error: " + ex.Message); }
        }

        private void ExportVisible_Click(object sender, RoutedEventArgs e)
        {
            RecordsGrid.CommitEdit();
            var list = new List<PersonRecord>();
            if (SelectedRecordsView != null) foreach (var item in SelectedRecordsView) list.Add(item as PersonRecord);
            ExportToExcel(list, "Export");
        }

        private void ExportToExcel(List<PersonRecord> records, string suffix)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx", FileName = $"ΛΙΣΤΑ_{DateTime.Now:dd_MM_yyyy}.xlsx" };
            if (sfd.ShowDialog() == true)
            {
                using (var package = new ExcelPackage(new FileInfo(sfd.FileName)))
                {
                    var ws = package.Workbook.Worksheets.Add("ΠΕΛΑΤΟΛΟΓΙΟ");
                    string[] h = { "Επιλογή", "Συμμετέχων", "Επωνυμία", "ΑΦΜ", "Τηλέφωνο 1", "ΔΩΡΟ" };
                    for (int i = 0; i < h.Length; i++) ws.Cells[1, i + 1].Value = h[i];
                    int r = 2;
                    foreach (var item in records)
                    {
                        ws.Cells[r, 1].Value = item.Selected ? "ΝΑΙ" : "ΟΧΙ";
                        ws.Cells[r, 2].Value = item.Comments;
                        ws.Cells[r, 3].Value = item.Επωνυμία;
                        ws.Cells[r, 4].Value = item.ΑΦΜ;
                        ws.Cells[r, 5].Value = item.Τηλέφωνο;
                        ws.Cells[r, 6].Value = item.SelectedGift;
                        r++;
                    }
                    package.Save();
                }
                MessageBox.Show("Η εξαγωγή ολοκληρώθηκε!");
            }
        }

        private async void Aade_Click(object sender, RoutedEventArgs e)
        {
            string afm = TxtAfm.Text.Trim();
            if (string.IsNullOrEmpty(afm)) return;
            AadeWindow win = new AadeWindow(afm) { Owner = this };
            if (win.ShowDialog() == true)
            {
                var source = RecordsView.SourceCollection as List<PersonRecord>;
                var exist = source?.FirstOrDefault(p => p.ΑΦΜ == afm);
                if (exist != null) exist.Επωνυμία = win.FetchedName;
                else source?.Add(new PersonRecord { ΑΦΜ = afm, Επωνυμία = win.FetchedName, Selected = true });
                RefreshFilter();
            }
        }

        protected override void OnClosing(CancelEventArgs e) { SaveToTemp(); base.OnClosing(e); }
        private void GiftInfo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ΛΟΓΙΚΗ: Αν υπάρχει το Temp διάβασε το, αλλιώς διάβασε το Original
                string fileToRead = File.Exists(tempGiftPath) ? tempGiftPath : originalGiftPath;

                if (!File.Exists(fileToRead)) return;

                DataTable dt = new DataTable();
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(fileToRead)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return;

                    for (int i = 1; i <= ws.Dimension.End.Column; i++) dt.Columns.Add(ws.Cells[1, i].Value?.ToString()?.Trim() ?? $"Column {i}");

                    for (int rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        DataRow dr = dt.NewRow();
                        bool hasValue = false;
                        for (int colNum = 1; colNum <= ws.Dimension.End.Column; colNum++)
                        {
                            var val = ws.Cells[rowNum, colNum].Value?.ToString()?.Trim() ?? "";
                            dr[colNum - 1] = val;
                            if (!string.IsNullOrEmpty(val)) hasValue = true;
                        }
                        if (hasValue) dt.Rows.Add(dr);
                    }
                }

                var viewer = new GiftViewerWindow();
                viewer.Owner = this;
                viewer.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                viewer.GiftsGrid.ItemsSource = dt.DefaultView;
                viewer.GiftsGrid.ColumnWidth = new DataGridLength(1, DataGridLengthUnitType.Star);

                // Ανοίγουμε το παράθυρο
                bool? result = viewer.ShowDialog();

                // ΞΑΝΑΦΟΡΤΩΝΟΥΜΕ ΤΑ ΔΩΡΑ ΣΤΟ ΚΥΡΙΩΣ ΠΑΡΑΘΥΡΟ (μήπως έγινε προσθήκη)
                LoadGifts();

                if (result == true)
                {
                    var button = sender as Button;
                    var selectedPerson = button.DataContext as PersonRecord;
                    if (selectedPerson != null)
                    {
                        selectedPerson.SelectedGift = viewer.SelectedGiftName;
                        RecordsGrid.Items.Refresh();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void Maximize_Click(object sender, RoutedEventArgs e) => WindowState = (WindowState == WindowState.Maximized) ? WindowState.Normal : WindowState.Maximized;
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
        private static string NormalizeDigits(string s) => string.IsNullOrEmpty(s) ? "" : new string(s.Where(char.IsDigit).ToArray());

        private async void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            AadeSettingsWindow settingsWin = new AadeSettingsWindow { Owner = this };
            if (settingsWin.ShowDialog() == true)
            {
                Mouse.OverrideCursor = Cursors.Wait;
                try
                {
                    var result = await System.Threading.Tasks.Task.Run(() => AadeService.GetDetails("999977386"));
                    Mouse.OverrideCursor = null;
                    if (result.Success) MessageBox.Show("Η υπηρεσία ΑΑΔΕ είναι ΕΝΕΡΓΗ!", "Επιτυχία", MessageBoxButton.OK, MessageBoxImage.Information);
                    else MessageBox.Show($"Πρόβλημα Σύνδεσης: {result.ErrorMessage}", "Σφάλμα", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (Exception ex) { Mouse.OverrideCursor = null; MessageBox.Show(ex.Message); }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void ResetFilters_Click(object sender, RoutedEventArgs e)
        {
            // Καθαρίζουμε τα πεδία κειμένου
            _nameQuery = "";
            _phoneQuery = "";
            _afmQuery = "";
            StatusFilterIndex = 0;

            // Ενημερώνουμε το UI για τις αλλαγές
            OnPropertyChanged(nameof(NameQuery));
            OnPropertyChanged(nameof(PhoneQuery));
            OnPropertyChanged(nameof(AfmQuery));

            // Κλείνουμε τυχόν ανοιχτά πλαίσια (Popups)
            IsNamePopupOpen = false;
            IsPhonePopupOpen = false;
            IsAfmPopupOpen = false;

            // Ανανεώνουμε τα φίλτρα στη λίστα
            RefreshFilter();
        }
    }

    public class GiftRecord { public string Eidos { get; set; } }
}