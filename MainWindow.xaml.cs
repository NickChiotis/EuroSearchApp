using System;
using System.Collections.Generic;
using System.Linq;
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
        // 1. Λίστα Δώρων
        private ObservableCollection<GiftRecord> _giftsList = new ObservableCollection<GiftRecord>();
        public ObservableCollection<GiftRecord> GiftsList
        {
            get => _giftsList;
            set { _giftsList = value; OnPropertyChanged(nameof(GiftsList)); }
        }

        // 2. Λίστα Επιλεγμένων Συμμετεχόντων (για τον κάτω πίνακα)
        private ObservableCollection<PersonRecord> _selectedParticipants = new ObservableCollection<PersonRecord>();
        public ObservableCollection<PersonRecord> SelectedParticipants
        {
            get => _selectedParticipants;
            set { _selectedParticipants = value; OnPropertyChanged(nameof(SelectedParticipants)); }
        }

        // --- Μέθοδος Προσθήκης (ΔΙΟΡΘΩΜΕΝΗ) ---
        public void AddParticipant(PersonRecord person)
        {
            bool exists = false;

            // 1. Αν έχει ΑΦΜ, ελέγχουμε βάσει ΑΦΜ
            if (!string.IsNullOrEmpty(person.ΑΦΜ))
            {
                exists = SelectedParticipants.Any(p => p.ΑΦΜ == person.ΑΦΜ);
            }
            // 2. Αν ΔΕΝ έχει ΑΦΜ, ελέγχουμε να μην υπάρχει ακριβώς το ίδιο Όνομα ΚΑΙ Τηλέφωνο
            else
            {
                exists = SelectedParticipants.Any(p => p.Επωνυμία == person.Επωνυμία && p.Τηλέφωνο == person.Τηλέφωνο);
            }

            if (!exists)
            {
                var newEntry = new PersonRecord
                {
                    Επωνυμία = person.Επωνυμία,
                    ΑΦΜ = person.ΑΦΜ,
                    Τηλέφωνο = person.Τηλέφωνο,
                    Τηλέφωνο2 = person.Τηλέφωνο2,
                    Comments = person.Comments,
                    Selected = true,
                    SelectedGift = person.SelectedGift
                };
                SelectedParticipants.Add(newEntry);
            }
        }

        // Paths
        private string giftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");

        // 3. View για τα αποτελέσματα αναζήτησης
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

            // Παρακολουθούμε τις αλλαγές στον κάτω πίνακα για να ενημερώνουμε το σύνολο
            SelectedParticipants.CollectionChanged += (s, e) => UpdateTotalCount();
        }

        private void UpdateTotalCount()
        {
            ResultCount = SelectedParticipants.Count.ToString("N0");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "Template.xlsx");
            string tempPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "temp_data.xlsx");

            // 1. Φόρτωση Κύριας Βάσης (Αναζήτηση) -> Template.xlsx
            if (File.Exists(defaultPath))
            {
                LoadDataForSearch(defaultPath);
            }
            else
            {
                MessageBox.Show("Δεν βρέθηκε το αρχείο δεδομένων Template.xlsx");
            }

            // 2. Φόρτωση Αποθηκευμένων (Κάτω Πίνακας) -> temp_data.xlsx
            if (File.Exists(tempPath))
            {
                LoadSavedParticipants(tempPath);
            }

            LoadGifts();
            TxtName.Focus();
            UpdateTotalCount();
        }

        private void LoadDataForSearch(string filePath)
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

        private void LoadSavedParticipants(string filePath)
        {
            try
            {
                var savedList = ExcelLoader.Load(filePath);
                if (savedList != null)
                {
                    SelectedParticipants.Clear();
                    foreach (var item in savedList)
                    {
                        item.Selected = true;
                        SelectedParticipants.Add(item);
                    }
                    UpdateTotalCount();
                }
            }
            catch { }
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

        private void GiftInfo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Έλεγχος αν υπάρχει το αρχείο δώρων
                if (!File.Exists(giftPath)) return;

                // 2. Διάβασμα του Excel και γέμισμα του DataTable
                DataTable dt = new DataTable();
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(giftPath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return;

                    // Δημιουργία στηλών
                    for (int i = 1; i <= ws.Dimension.End.Column; i++)
                    {
                        string header = ws.Cells[1, i].Value?.ToString()?.Trim() ?? $"Column {i}";
                        dt.Columns.Add(header);
                    }

                    // Γέμισμα γραμμών
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

                // 3. Άνοιγμα του παραθύρου προβολής
                // Βεβαιώσου ότι έχεις την κλάση GiftViewerWindow στο project σου
                var viewer = new GiftViewerWindow();
                viewer.Owner = this;
                viewer.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // Πέρνα τα δεδομένα στο Grid του viewer (υποθέτουμε ότι το DataGrid εκεί λέγεται GiftsGrid)
                viewer.GiftsGrid.ItemsSource = dt.DefaultView;
                viewer.GiftsGrid.ColumnWidth = new DataGridLength(1, DataGridLengthUnitType.Star);

                // 4. Λογική Επιλογής (Αν ο χρήστης διαλέξει δώρο από το παράθυρο)
                if (viewer.ShowDialog() == true)
                {
                    // Βρίσκουμε ποια γραμμή πατήθηκε στον ΚΕΝΤΡΙΚΟ πίνακα (MainWindow)
                    // Χρησιμοποιούμε το κουμπί που πατήθηκε (sender) για να βρούμε την εγγραφή
                    var button = sender as Button;
                    var selectedPerson = button.DataContext as PersonRecord;

                    if (selectedPerson != null && !string.IsNullOrEmpty(viewer.SelectedGiftName))
                    {
                        // Ενημερώνουμε το δώρο του συγκεκριμένου ατόμου
                        selectedPerson.SelectedGift = viewer.SelectedGiftName;

                        // Ανανεώνουμε τον πίνακα για να φανεί η αλλαγή
                        ParticipantsGrid.Items.Refresh();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά το άνοιγμα των δώρων: " + ex.Message);
            }
        }

        private bool FilterRecords(object item)
        {
            var person = item as PersonRecord;
            if (person == null) return false;

            if (!string.IsNullOrWhiteSpace(NameQuery))
            {
                if (string.IsNullOrEmpty(person.Επωνυμία) || !person.Επωνυμία.StartsWith(NameQuery, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            if (!string.IsNullOrWhiteSpace(AfmQuery))
            {
                if (string.IsNullOrEmpty(person.ΑΦΜ) || !person.ΑΦΜ.StartsWith(AfmQuery, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            if (!string.IsNullOrWhiteSpace(PhoneQuery))
            {
                string cleanQuery = NormalizeDigits(PhoneQuery).Trim();
                if (!string.IsNullOrEmpty(cleanQuery))
                {
                    string p1 = !string.IsNullOrEmpty(person.Τηλέφωνο) ? NormalizeDigits(person.Τηλέφωνο) : "";
                    string p2 = !string.IsNullOrEmpty(person.Τηλέφωνο2) ? NormalizeDigits(person.Τηλέφωνο2) : "";
                    bool match1 = p1.StartsWith(cleanQuery);
                    bool match2 = p2.StartsWith(cleanQuery);
                    if (!match1 && !match2) return false;
                }
            }

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

                // ΔΕΝ ενημερώνουμε το ResultCount εδώ πλέον (το κάνουμε από το SelectedParticipants)

                bool hasText = !string.IsNullOrEmpty(NameQuery) || !string.IsNullOrEmpty(PhoneQuery) || !string.IsNullOrEmpty(AfmQuery);

                if (hasText && count > 0)
                {
                    SearchPopup.IsOpen = true;
                }
                else
                {
                    SearchPopup.IsOpen = false;
                }
            }
        }

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                SearchPopup.IsOpen = false;
                return;
            }

            if (e.Key == Key.Enter && !SearchPopup.IsOpen)
            {
                var source = sender as TextBox;
                if (source == TxtName) TxtPhone.Focus();
                else if (source == TxtPhone) TxtAfm.Focus();
                else if (source == TxtAfm && SuggestionList.Items.Count > 0)
                {
                    SearchPopup.IsOpen = true;
                    SuggestionList.SelectedIndex = 0;
                    var item = SuggestionList.ItemContainerGenerator.ContainerFromIndex(0) as ListBoxItem;
                    item?.Focus();
                    SuggestionList.Focus();
                }
                return;
            }

            if (e.Key == Key.Down && SearchPopup.IsOpen)
            {
                if (SuggestionList.Items.Count > 0)
                {
                    SuggestionList.SelectedIndex = 0;
                    var item = SuggestionList.ItemContainerGenerator.ContainerFromIndex(0) as ListBoxItem;
                    item?.Focus();
                    SuggestionList.Focus();
                    e.Handled = true;
                }
            }
        }

        private void SuggestionList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var person = SuggestionList.SelectedItem as PersonRecord;
                if (person != null)
                {
                    AddParticipant(person);
                    ResetFilters_Click(null, null);
                }
                e.Handled = true;
            }
            if (e.Key == Key.Escape)
            {
                SearchPopup.IsOpen = false;
                TxtName.Focus();
            }
        }

        private void RemoveParticipant_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var person = button.DataContext as PersonRecord;
            if (person != null)
            {
                SelectedParticipants.Remove(person);
            }
        }

        private void ResetFilters_Click(object sender, RoutedEventArgs e)
        {
            _nameQuery = ""; OnPropertyChanged(nameof(NameQuery));
            _phoneQuery = ""; OnPropertyChanged(nameof(PhoneQuery));
            _afmQuery = ""; OnPropertyChanged(nameof(AfmQuery));

            StatusFilterIndex = 0;
            SearchPopup.IsOpen = false;
            TxtName.Focus();

            RefreshFilter();
        }

        private void ResultsList_KeyDown(object sender, KeyEventArgs e) { }

        private async void Aade_Click(object sender, RoutedEventArgs e)
        {
            string afm = TxtAfm.Text.Trim();
            if (string.IsNullOrEmpty(afm)) return;

            AadeWindow aadeWin = new AadeWindow(afm);
            aadeWin.Owner = this;
            aadeWin.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            if (aadeWin.ShowDialog() == true)
            {
                var sourceList = RecordsView.SourceCollection as List<PersonRecord>;
                if (sourceList != null)
                {
                    var existingPerson = sourceList.FirstOrDefault(p => p.ΑΦΜ == afm);
                    if (existingPerson != null) existingPerson.Επωνυμία = aadeWin.FetchedName;
                    else
                    {
                        var newPerson = new PersonRecord { ΑΦΜ = afm, Επωνυμία = aadeWin.FetchedName, Selected = true };
                        sourceList.Add(newPerson);
                        AddParticipant(newPerson);
                    }
                    RefreshFilter();
                }
            }
        }

        private void SaveToTemp()
        {
            try
            {
                string tempPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Templates", "temp_data.xlsx");

                // ΣΩΖΟΥΜΕ ΜΟΝΟ ΤΟΥΣ SelectedParticipants
                var itemsToSave = SelectedParticipants;

                if (itemsToSave == null) return;

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("TempData");
                    // Τα headers πρέπει να είναι ίδια με το Template για να δουλεύει ο Loader
                    string[] headers = { "Επιλογή", "Συμμετέχων", "Επωνυμία", "Επαφές - Α.Φ.Μ", "Τηλέφωνο 1", "Τηλέφωνο 2", "SelectedGift" };

                    for (int i = 0; i < headers.Length; i++) ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var item in itemsToSave)
                    {
                        ws.Cells[row, 1].Value = "ΝΑΙ";
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

        private void ExportToExcel(List<PersonRecord> records, string suffix) { }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            SaveToTemp();
            base.OnClosing(e);
        }

        private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void Maximize_Click(object sender, RoutedEventArgs e) => WindowState = (WindowState == WindowState.Maximized) ? WindowState.Normal : WindowState.Maximized;
        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private async void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            AadeSettingsWindow settingsWin = new AadeSettingsWindow();
            settingsWin.Owner = this;
            settingsWin.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            if (settingsWin.ShowDialog() == true)
            {
                string testAfm = "999977386";
                Mouse.OverrideCursor = Cursors.Wait;
                try
                {
                    var result = await System.Threading.Tasks.Task.Run(() => AadeService.GetDetails(testAfm));
                    Mouse.OverrideCursor = null;
                    if (result.Success)
                        MessageBox.Show("Επιτυχής Σύνδεση με ΑΑΔΕ!", "Επιτυχία", MessageBoxButton.OK, MessageBoxImage.Information);
                    else
                        MessageBox.Show("Πρόβλημα Σύνδεσης: " + result.ErrorMessage, "Σφάλμα", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    Mouse.OverrideCursor = null;
                    MessageBox.Show("Σφάλμα: " + ex.Message);
                }
            }
        }

        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return new string(s.Where(char.IsDigit).ToArray());
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class GiftRecord
    {
        public string Eidos { get; set; }
    }
}