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
        // --- 1. ΣΤΑΤΙΚΗ ΑΝΑΦΟΡΑ ΓΙΑ TON CONVERTER ---
        public static MainWindow AppInstance;

        // Λίστα Επιλεγμένων Συμμετεχόντων
        private ObservableCollection<PersonRecord> _selectedParticipants = new ObservableCollection<PersonRecord>();
        public ObservableCollection<PersonRecord> SelectedParticipants
        {
            get => _selectedParticipants;
            set { _selectedParticipants = value; OnPropertyChanged(nameof(SelectedParticipants)); }
        }

        // --- Λίστες και Views ---
        private ObservableCollection<GiftRecord> _giftsList = new ObservableCollection<GiftRecord>();
        public ObservableCollection<GiftRecord> GiftsList
        {
            get => _giftsList;
            set { _giftsList = value; OnPropertyChanged(nameof(GiftsList)); }
        }

        private void BtnHelperMenu_Click(object sender, RoutedEventArgs e)
        {
            // Βρίσκουμε το κουμπί που πατήθηκε
            Button btn = sender as Button;

            // Αν το κουμπί έχει ContextMenu, το ανοίγουμε
            if (btn != null && btn.ContextMenu != null)
            {
                // Ορίζουμε το Target στο κουμπί για να ανοίξει στη σωστή θέση
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.IsOpen = true;
            }
        }

        private string giftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");

        private ICollectionView _selectedRecordsView;
        public ICollectionView SelectedRecordsView
        {
            get => _selectedRecordsView;
            set { _selectedRecordsView = value; OnPropertyChanged(nameof(SelectedRecordsView)); }
        }
        private string originalGiftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");
        private string tempGiftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "temp_gift.xlsx");

        private ICollectionView _recordsView;
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

        private string _resultCount = "0";
        public string ResultCount
        {
            get { return _resultCount; }
            set { _resultCount = value; OnPropertyChanged(nameof(ResultCount)); }
        }

        public MainWindow()
        {
            InitializeComponent();

            // --- 2. ΑΡΧΙΚΟΠΟΙΗΣΗ ΤΟΥ AppInstance ---
            AppInstance = this;

            DataContext = this;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "Template.xlsx");
            string tempPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "temp_data.xlsx");

            if (File.Exists(defaultPath))
            {
                var mainData = ExcelLoader.Load(defaultPath);
                RecordsView = CollectionViewSource.GetDefaultView(mainData);
                RecordsView.Filter = FilterRecords;
            }
            else
            {
                MessageBox.Show("Δεν βρέθηκε το αρχείο Template.xlsx για την αναζήτηση.");
            }

            // 2. Φορτώνουμε το temp_data.xlsx ΜΟΝΟ για τη λίστα συμμετεχόντων (κάτω πλαίσιο)
            if (File.Exists(tempPath))
            {
                var savedParticipants = ExcelLoader.Load(tempPath);
                SelectedParticipants.Clear();
                foreach (var p in savedParticipants)
                {
                    // Προαιρετικά: Αν θες να ταυτίζονται τα αντικείμενα με τη βάση, 
                    // αλλά για απλή εμφάνιση αρκεί να τα προσθέσεις:
                    SelectedParticipants.Add(p);
                }
            }

            // Αφού φορτώσαμε τα πάντα, ενημερώνουμε το νούμερο!
            if (SelectedParticipants != null)
            {
                ResultCount = SelectedParticipants.Count.ToString();
            }
            else
            {
                ResultCount = "0";
            }

            LoadGifts();
            TxtName.Focus();
        }

        // --- ΒΟΗΘΗΤΙΚΗ ΜΕΘΟΔΟΣ ΓΙΑ ΤΟ FOCUS ΣΤΗ ΛΙΣΤΑ ---
        private void ForceFocusToList(ListBox list)
        {
            if (list == null) return;
            Application.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
            {
                if (list.Items.Count > 0)
                {
                    try
                    {
                        list.SelectedIndex = 0;
                        list.UpdateLayout();
                        var item = list.ItemContainerGenerator.ContainerFromIndex(0) as ListBoxItem;
                        if (item != null)
                        {
                            item.Focus();
                            Keyboard.Focus(item);
                        }
                        else list.Focus();
                    }
                    catch { }
                }
            }));
        }

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
                // 1. Καθαρίζουμε το κείμενο που πληκτρολόγησε ο χρήστης από τόνους και το κάνουμε κεφαλαία
                string cleanQuery = RemoveAccents(value);

                // 2. Αναζητούμε στη λίστα μετατρέποντας προσωρινά και την Επωνυμία σε "καθαρή" μορφή
                var matches = source.Where(p => p.Επωνυμία != null &&
                                               RemoveAccents(p.Επωνυμία).Contains(cleanQuery))
                                    .Select(p => p.Επωνυμία)
                                    .Distinct()
                                    .Take(10)
                                    .ToList();

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

                RecordsView = CollectionViewSource.GetDefaultView(rawList);
                RecordsView.Filter = FilterRecords;

                SelectedParticipants.Clear();
                foreach (var person in rawList)
                {
                    // Αν το ExcelLoader είδε δώρο, έχει κάνει το person.Selected = true.
                    // Εμείς απλά τον βάζουμε στη σωστή λίστα.
                    if (person.Selected)
                    {
                        SelectedParticipants.Add(person);
                    }
                }

                RefreshFilter();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private bool FilterRecords(object item)
        {
            var person = item as PersonRecord;
            if (person == null) return false;

            bool hasActiveFilter = !string.IsNullOrWhiteSpace(NameQuery) || !string.IsNullOrWhiteSpace(PhoneQuery) || !string.IsNullOrWhiteSpace(AfmQuery);
            if (!hasActiveFilter) return false;

            if (!string.IsNullOrWhiteSpace(NameQuery))
            {
                string cleanName = RemoveAccents(person.Επωνυμία);
                string cleanQuery = RemoveAccents(NameQuery);
                if (!cleanName.Contains(cleanQuery)) return false;
            }

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
            if (RecordsGrid != null)
            {
                RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            }

            RecordsView?.Refresh();
            SelectedRecordsView?.Refresh();

            if (SelectedRecordsView != null)
            {
                int count = 0;
                foreach (var item in SelectedRecordsView) count++;
                ResultCount = count.ToString("N0");
            }
        }

        public void AddParticipant(PersonRecord person)
        {
            if (person == null) return;

            // ΕΛΕΓΧΟΣ: Είναι ήδη ο πελάτης στον κάτω πίνακα;
            // Ελέγχουμε αν υπάρχει ήδη στην ObservableCollection των επιλεγμένων
            if (SelectedParticipants.Any(p =>
                (p.ΑΦΜ == person.ΑΦΜ && !string.IsNullOrEmpty(p.ΑΦΜ)) ||
                (p.Τηλέφωνο == person.Τηλέφωνο && !string.IsNullOrEmpty(p.Τηλέφωνο)) ||
                (RemoveAccents(p.Επωνυμία) == RemoveAccents(person.Επωνυμία)))) // Αλλαγή εδώ
            {
                MessageBox.Show("Αυτός ο Πελάτης έχει ήδη προστεθεί στις σημερινές εγγραφές!",
                                "Διπλότυπη Εγγραφή",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                return; // Σταματάει τη διαδικασία εδώ
            }

            // Αν δεν υπάρχει, συνεχίζουμε κανονικά:
            RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true);

            person.Selected = true;
            if (!SelectedParticipants.Contains(person))
            {
                SelectedParticipants.Add(person);
            }

            RefreshFilter();
            SaveToTemp();

            // Αυτόματο άνοιγμα παραθύρου δώρων
            Application.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
            {
                OpenGiftSelectionForPerson(person);
            }));
        }

        private void OpenGiftSelectionForPerson(PersonRecord person)
        {
            try
            {
                if (person == null) return;

                // Ενημερώνουμε τα αποθέματα πριν ανοίξει το παράθυρο
                LoadGifts();

                var allRecords = RecordsView.SourceCollection as IEnumerable<PersonRecord>;
                var usedGifts = SelectedParticipants?
                    .Where(p => !string.IsNullOrEmpty(p.SelectedGift))
                    .GroupBy(p => p.SelectedGift)
                    .ToDictionary(g => g.Key, g => g.Count()) ?? new Dictionary<string, int>();

                List<GiftStockItem> stockItems = new List<GiftStockItem>();
                string fileToRead = File.Exists(tempGiftPath) ? tempGiftPath : originalGiftPath;

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(fileToRead)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension != null)
                    {
                        for (int row = 2; row <= ws.Dimension.End.Row; row++)
                        {
                            var name = ws.Cells[row, 1].Value?.ToString()?.Trim();
                            int totalQty = 0;
                            if (ws.Cells[row, 2].Value != null) int.TryParse(ws.Cells[row, 2].Value.ToString(), out totalQty);

                            if (!string.IsNullOrEmpty(name))
                            {
                                int used = usedGifts.ContainsKey(name) ? usedGifts[name] : 0;
                                stockItems.Add(new GiftStockItem { Name = name, TotalQty = totalQty, UsedQty = used });
                            }
                        }
                    }
                }

                var viewer = new GiftViewerWindow();
                viewer.Owner = this;
                viewer.GiftItems = stockItems;
                viewer.GiftsGrid.ItemsSource = stockItems;

                if (viewer.ShowDialog() == true)
                {
                    // Ανέθεσε το δώρο στον πελάτη
                    person.SelectedGift = viewer.SelectedGiftName;

                    // Αποθήκευση και ανανέωση
                    SaveToTemp();
                    LoadGifts();
                    RecordsGrid.Items.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά την αυτόματη επιλογή δώρου: " + ex.Message);
            }
        }
        public void RefreshStockDisplay()
        {
            // Αναγκάζει το DataGrid να ξανασχεδιάσει τα κελιά, 
            // οπότε ο Converter θα ξαναδιαβάσει τα νέα νούμερα Remaining
            SelectedRecordsView?.Refresh();
        }

        // --- 1. Διαχείριση πλήκτρων στα TextBox ---
        private void Input_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox == null) return;

            System.Windows.Controls.Primitives.Popup targetPopup = null;
            ListBox targetList = null;

            if (textBox == TxtName) { targetPopup = PopName; targetList = LstNameSuggestions; }
            else if (textBox == TxtPhone) { targetPopup = PopPhone; targetList = LstPhoneSuggestions; }
            else if (textBox == TxtAfm) { targetPopup = PopAfm; targetList = LstAfmSuggestions; }

            if (targetPopup == null || targetList == null) return;

            if (e.Key == Key.Down)
            {
                if (targetPopup.IsOpen && targetList.Items.Count > 0)
                {
                    e.Handled = true;
                    ForceFocusToList(targetList);
                }
            }
            else if (e.Key == Key.Enter)
            {
                if (RecordsView == null || RecordsView.SourceCollection == null) return;
                var sourceList = RecordsView.SourceCollection as IEnumerable<PersonRecord>;

                // Αν είμαστε στο ΑΦΜ και το popup είναι κλειστό (σημαίνει ότι ήδη συμπληρώθηκε)
                if (textBox == TxtAfm && !PopAfm.IsOpen)
                {
                    var match = sourceList?.FirstOrDefault(p => p.ΑΦΜ == TxtAfm.Text.Trim());
                    if (match != null)
                    {
                        // Αν βρέθηκε κάποιος, τον προσθέτουμε κανονικά
                        AddParticipant(match);
                        ResetFilters_Click(null, null);
                        e.Handled = true;
                        return;
                    }
                    else
                    {
                        // ΑΝ ΔΕΝ ΒΡΕΘΗΚΕ ΤΑΙΡΙΑΣΜΑ:
                        // Ελέγχουμε αν ο κέρσορας είναι στο TextBox του ονόματος και αν έχει γραφτεί κάτι
                        if (textBox == TxtName && !string.IsNullOrWhiteSpace(TxtName.Text))
                        {
                            // Καλούμε τη μέθοδο του κουμπιού "+" που φτιάξαμε
                            BtnQuickAdd_Click(null, null);
                            e.Handled = true;
                            return;
                        }
                    }
                }

                // Η κανονική λογική για τα υπόλοιπα πεδία
                var anyMatch = sourceList?.FirstOrDefault(p => FilterRecords(p));
                if (anyMatch != null)
                {
                    AddParticipant(anyMatch);
                    ResetFilters_Click(null, null);
                    e.Handled = true;
                }
            }
        }

        // --- 2. Όταν είσαι ΜΕΣΑ στη λίστα (Enter ή Escape) ---
        private void SuggestionList_KeyDown(object sender, KeyEventArgs e)
        {
            var listbox = sender as ListBox;
            if (listbox == null) return;

            if (e.Key == Key.Enter)
            {
                if (listbox.SelectedItem != null)
                {
                    string selectedValue = listbox.SelectedItem.ToString();
                    SelectParticipantByValue(selectedValue);
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.Escape)
            {
                e.Handled = true;
                TextBox targetBox = null;

                if (listbox == LstNameSuggestions) targetBox = TxtName;
                else if (listbox == LstPhoneSuggestions) targetBox = TxtPhone;
                else if (listbox == LstAfmSuggestions) targetBox = TxtAfm;

                if (targetBox != null)
                {
                    targetBox.Focus();
                    targetBox.CaretIndex = targetBox.Text.Length;
                }
            }
        }

        private void SuggestionList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var listbox = sender as ListBox;
            if (listbox != null && listbox.SelectedItem != null)
            {
                string selectedValue = listbox.SelectedItem.ToString();
                SelectParticipantByValue(selectedValue);
            }
        }

        private void SelectParticipantByValue(string value)
        {
            var source = RecordsView.SourceCollection as IEnumerable<PersonRecord>;
            if (source == null) return;

            // Βρίσκουμε τον πελάτη
            var person = source.FirstOrDefault(p =>
                (p.Επωνυμία != null && p.Επωνυμία == value) ||
                (p.ΑΦΜ != null && p.ΑΦΜ == value) ||
                (p.Τηλέφωνο != null && p.Τηλέφωνο == value));

            if (person != null)
            {
                // ΕΙΔΙΚΗ ΛΟΓΙΚΗ ΓΙΑ ΑΦΜ
                // Αν το value που ήρθε είναι το ΑΦΜ του πελάτη και είμαστε στο πεδίο του ΑΦΜ
                if (value == person.ΑΦΜ && IsAfmPopupOpen)
                {
                    AfmQuery = person.ΑΦΜ; // Συμπληρώνει το TextBox
                    IsAfmPopupOpen = false; // Κλείνει το popup
                    TxtAfm.Focus();         // Δίνει focus πίσω στο TextBox για το επόμενο Enter
                    TxtAfm.CaretIndex = TxtAfm.Text.Length; // Πάει τον κέρσορα στο τέλος
                }
                else
                {
                    // Για Όνομα και Τηλέφωνο, κάνει απευθείας προσθήκη
                    AddParticipant(person);
                    ResetFilters_Click(null, null);
                }
            }
        }

        // --- 3. ΔΙΟΡΘΩΜΕΝΗ LoadGifts ΓΙΑ ΥΠΟΛΟΓΙΣΜΟ STOCK ---
        public void LoadGifts()
        {
            try
            {
                // 1. Υπολογισμός Χρησιμοποιημένων από το Grid
                var allRecords = RecordsView?.SourceCollection as IEnumerable<PersonRecord>;
                var usedGifts = SelectedParticipants?
                    .Where(p => !string.IsNullOrEmpty(p.SelectedGift))
                    .GroupBy(p => p.SelectedGift)
                    .ToDictionary(g => g.Key, g => g.Count()) ?? new Dictionary<string, int>();

                // 2. Διάβασμα Excel
                string fileToRead = File.Exists(tempGiftPath) ? tempGiftPath : originalGiftPath;
                if (!File.Exists(fileToRead)) return;

                var excelItems = new List<GiftRecord>();
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(fileToRead)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return;

                    for (int row = 2; row <= ws.Dimension.End.Row; row++)
                    {
                        var name = ws.Cells[row, 1].Value?.ToString()?.Trim();
                        int totalQty = 0;
                        if (ws.Cells[row, 2].Value != null) int.TryParse(ws.Cells[row, 2].Value.ToString(), out totalQty);

                        if (!string.IsNullOrEmpty(name))
                        {
                            int used = usedGifts.ContainsKey(name) ? usedGifts[name] : 0;
                            excelItems.Add(new GiftRecord { Eidos = name, Remaining = totalQty - used });
                        }
                    }
                }

                // 3. Ενημέρωση της GiftsList ΧΩΡΙΣ Clear() που χαλάει το Dropdown
                Application.Current.Dispatcher.Invoke(() =>
                {
                    // Αν η λίστα είναι άδεια, γέμισέ την
                    if (GiftsList.Count <= 1)
                    {
                        GiftsList.Clear();
                        GiftsList.Add(new GiftRecord { Eidos = "Επιλογή Δώρου", Remaining = 9999 });
                        foreach (var item in excelItems) GiftsList.Add(item);
                    }
                    // Αντί να σβήσεις τα πάντα, ενημέρωσε μόνο τα υπάρχοντα
                    foreach (var excelItem in excelItems)
                    {
                        // Βρες το δώρο στη μνήμη
                        var existing = GiftsList.FirstOrDefault(g => g.Eidos == excelItem.Eidos);

                        if (existing != null)
                        {
                            // Απλά άλλαξε το νούμερο, μην πειράξεις το αντικείμενο
                            existing.Remaining = excelItem.Remaining;
                        }
                        else
                        {
                            // Αν είναι καινούργιο δώρο που δεν υπήρχε πριν, πρόσθεσέ το
                            GiftsList.Add(excelItem);
                        }
                    }
                });
            }
            catch { }
        }

        private void SaveToTemp()
        {
            try
            {
                if (SelectedParticipants == null) return;

                string tempPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Templates", "temp_data.xlsx");
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("TempData");
                    string[] headers = { "Συμμετέχων", "Επωνυμία", "ΑΦΜ", "Τηλέφωνο 1", "Τηλέφωνο 2", "ΔΩΡΟ" };
                    for (int i = 0; i < headers.Length; i++) ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var item in SelectedParticipants)
                    {
                        ws.Cells[row, 1].Value = item.Comments;
                        ws.Cells[row, 2].Value = item.Επωνυμία;
                        ws.Cells[row, 3].Value = item.ΑΦΜ;
                        ws.Cells[row, 4].Value = item.Τηλέφωνο;
                        ws.Cells[row, 5].Value = item.Τηλέφωνο2;
                        ws.Cells[row, 6].Value = item.SelectedGift;
                        row++;
                    }
                    package.SaveAs(new FileInfo(tempPath));
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
                    // Ορίζουμε τις 6 στήλες που ζήτησες
                    string[] h = { "Συμμετέχων", "Επωνυμία", "ΑΦΜ", "Τηλέφωνο 1", "Τηλέφωνο 2", "ΔΩΡΟ" };
                    for (int i = 0; i < h.Length; i++) ws.Cells[1, i + 1].Value = h[i];

                    int r = 2;
                    foreach (var item in records)
                    {
                        ws.Cells[r, 1].Value = item.Comments;
                        ws.Cells[r, 2].Value = item.Επωνυμία;
                        ws.Cells[r, 3].Value = item.ΑΦΜ;
                        ws.Cells[r, 4].Value = item.Τηλέφωνο;
                        ws.Cells[r, 5].Value = item.Τηλέφωνο2;
                        ws.Cells[r, 6].Value = item.SelectedGift;
                        r++;
                    }
                    package.Save();
                }
                MessageBox.Show("Η εξαγωγή ολοκληρώθηκε!");
            }
        }

        private void Aade_Click(object sender, RoutedEventArgs e)
        {
            // Παίρνουμε το ΑΦΜ από το κεντρικό TextBox
            string afmToSearch = TxtAfm.Text.Trim();

            if (string.IsNullOrEmpty(afmToSearch))
            {
                MessageBox.Show("Παρακαλώ εισάγετε ένα ΑΦΜ πρώτα.");
                return;
            }

            // 1. Άνοιγμα του παραθύρου ΑΑΔΕ
            var aadeWin = new AadeWindow(afmToSearch);
            aadeWin.Owner = this;

            if (aadeWin.ShowDialog() == true)
            {
                // 2. Δημιουργία της εγγραφής με τα στοιχεία που ήρθαν από το AadeWindow
                var newPerson = new PersonRecord
                {
                    Επωνυμία = aadeWin.ResultName, // Χρησιμοποιούμε τα Properties που φτιάξαμε
                    ΑΦΜ = aadeWin.ResultAfm,
                    Comments = "", // Αποθηκεύουμε τη διεύθυνση στα σχόλια
                    Selected = true
                };

                // 3. Προσθήκη στην κύρια λίστα (για να υπάρχει στη "βάση" μας)
                var source = RecordsView.SourceCollection as List<PersonRecord>;
                source?.Add(newPerson);

                // 4. ΕΙΣΑΓΩΓΗ ΣΤΟΝ ΠΙΝΑΚΑ ΣΥΜΜΕΤΕΧΟΝΤΩΝ
                // Αυτή η μέθοδος θα κάνει τον έλεγχο για διπλότυπα ΚΑΙ θα ανοίξει το Popup των δώρων
                AddParticipant(newPerson);

                // 5. Καθαρισμός των φίλτρων αναζήτησης για να είμαστε έτοιμοι για τον επόμενο
                ResetFilters_Click(null, null);
            }
        }

        protected override void OnClosing(CancelEventArgs e) { SaveToTemp(); base.OnClosing(e); }

        private void GiftInfo_Click(object sender, RoutedEventArgs e)
        {
            // Βρίσκουμε ποιος πελάτης αντιστοιχεί στη γραμμή που πατήθηκε το κουμπί
            var button = sender as Button;
            var selectedPerson = button?.DataContext as PersonRecord;

            if (selectedPerson != null)
            {
                // Καλούμε τη νέα κοινή μέθοδο
                OpenGiftSelectionForPerson(selectedPerson);
            }
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
            _nameQuery = "";
            _phoneQuery = "";
            _afmQuery = "";

            OnPropertyChanged(nameof(NameQuery));
            OnPropertyChanged(nameof(PhoneQuery));
            OnPropertyChanged(nameof(AfmQuery));

            IsNamePopupOpen = false;
            IsPhonePopupOpen = false;
            IsAfmPopupOpen = false;

            RefreshFilter();
        }

        private void RemoveParticipant_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var person = button?.DataContext as PersonRecord;

            if (person != null)
            {
                // 1. Αφαίρεση από τη λίστα που ελέγχει τα διπλότυπα (ΑΥΤΟ ΕΛΕΙΠΕ)
                if (SelectedParticipants.Contains(person))
                {
                    SelectedParticipants.Remove(person);
                }

                // 2. Αποεπιλογή του πελάτη
                person.Selected = false;

                // 3. Μηδενισμός του δώρου για να επιστρέψει στο stock
                person.SelectedGift = null;

                // 4. Αποθήκευση και Ανανέωση
                SaveToTemp();
                RefreshFilter(); // Θα εξαφανίσει τη γραμμή από το DataGrid
                LoadGifts();     // Θα ενημερώσει τα αποθέματα

                RecordsGrid.Items.Refresh();
            }
        }

        private void BtnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var newData = ExcelLoader.Load(openFileDialog.FileName);
                    if (newData != null && newData.Count > 0)
                    {
                        RecordsView = CollectionViewSource.GetDefaultView(newData);
                        RecordsView.Filter = FilterRecords;

                        SelectedRecordsView = CollectionViewSource.GetDefaultView(newData);
                        SelectedRecordsView.Filter = (item) => ((PersonRecord)item).Selected;

                        SaveToTemp();
                        RefreshFilter();
                        LoadGifts(); // Φορτώνουμε και τα δώρα
                        MessageBox.Show("Το αρχείο φορτώθηκε και αποθηκεύτηκε ως η νέα λίστα εργασίας!", "Επιτυχία", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Το αρχείο δεν περιέχει έγκυρα δεδομένα ή είναι κενό.", "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Σφάλμα κατά τη φόρτωση: " + ex.Message, "Σφάλμα", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnQuickAdd_Click(object sender, RoutedEventArgs e)
        {
            // Παίρνουμε την τρέχουσα λίστα από το View
            var existingRecords = RecordsView?.SourceCollection as IEnumerable<PersonRecord>;

            // Περνάμε τη λίστα στον constructor του QuickAddWindow
            var quickAddWin = new QuickAddWindow(TxtName.Text, existingRecords);
            quickAddWin.Owner = this;

            if (quickAddWin.ShowDialog() == true)
            {
                var newPerson = new PersonRecord
                {
                    Επωνυμία = quickAddWin.FullName,
                    Τηλέφωνο = quickAddWin.Phone,
                    ΑΦΜ = quickAddWin.Afm,
                    Selected = true
                };

                if (RecordsView.SourceCollection is List<PersonRecord> sourceList)
                {
                    sourceList.Add(newPerson);
                }

                AddParticipant(newPerson);
                NameQuery = "";
                RefreshFilter();
            }
        }
        private string RemoveAccents(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            // Πίνακας αντικατάστασης για ελληνικά
            char[] accented = { 'Ά', 'Έ', 'Ή', 'Ί', 'Ό', 'Ύ', 'Ώ', 'ά', 'έ', 'ή', 'ί', 'ό', 'ύ', 'ώ', 'ϊ', 'ΐ', 'ϋ', 'ΰ' };
            char[] plain = { 'Α', 'Ε', 'Η', 'Ι', 'Ο', 'Υ', 'Ω', 'α', 'ε', 'η', 'ι', 'ο', 'υ', 'ω', 'ι', 'ι', 'υ', 'υ' };

            string result = text;
            for (int i = 0; i < accented.Length; i++)
            {
                result = result.Replace(accented[i], plain[i]);
            }

            return result.ToUpper().Trim(); // Επιστρέφει κεφαλαία χωρίς τόνους
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            // 1. Επιλογή αρχείου
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.Title = "Επιλέξτε το νέο αρχείο πελατών";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    // 2. Φόρτωση των δεδομένων από το επιλεγμένο Excel (χρησιμοποιώντας τον νέο ExcelLoader)
                    var newData = ExcelLoader.Load(openFileDialog.FileName);

                    if (newData != null && newData.Count > 0)
                    {
                        // 3. Ενημέρωση των Views της εφαρμογής
                        RecordsView = CollectionViewSource.GetDefaultView(newData);
                        RecordsView.Filter = FilterRecords;

                        SelectedRecordsView = CollectionViewSource.GetDefaultView(newData);
                        SelectedRecordsView.Filter = (item) => ((PersonRecord)item).Selected;

                        // 4. ΕΝΗΜΕΡΩΣΗ ΤΗΣ ΛΙΣΤΑΣ ΣΥΜΜΕΤΕΧΟΝΤΩΝ (Για τον κάτω πίνακα)
                        // Καθαρίζουμε την τρέχουσα μνήμη και προσθέτουμε όσους έχουν ήδη δώρο
                        SelectedParticipants.Clear();
                        foreach (var person in newData.Where(p => p.Selected))
                        {
                            SelectedParticipants.Add(person);
                        }

                        // 5. ΑΜΕΣΗ ΑΠΟΘΗΚΕΥΣΗ στο temp_data.xlsx
                        SaveToTemp();

                        // 6. ΑΝΑΝΕΩΣΗ UI
                        RefreshFilter();
                        LoadGifts(); // Επανυπολογισμός αποθεμάτων με βάση τα δώρα που κάναμε εισαγωγή

                        MessageBox.Show($"Η εισαγωγή ολοκληρώθηκε!\nΦορτώθηκαν {newData.Count} πελάτες και {SelectedParticipants.Count} συμμετοχές.",
                                        "Επιτυχία", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Το αρχείο φαίνεται να είναι κενό ή δεν έχει τη σωστή μορφή.",
                                        "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Σφάλμα κατά την εισαγωγή: " + ex.Message, "Σφάλμα", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

    }

    // --- ΔΙΟΡΘΩΜΕΝΗ ΚΛΑΣΗ GiftRecord (ΜΕ Remaining) ---
    public class GiftRecord : INotifyPropertyChanged
    {
        public string Eidos { get; set; }

        private int _remaining;
        public int Remaining
        {
            get => _remaining;
            set { _remaining = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Remaining))); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class GiftStockConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string giftName = value as string;
            if (string.IsNullOrEmpty(giftName) || giftName == "Επιλογή Δώρου") return "";

            var item = MainWindow.AppInstance?.GiftsList?.FirstOrDefault(g => g.Eidos == giftName);
            return item != null ? item.Remaining.ToString() : "0";
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => null;
    }
}