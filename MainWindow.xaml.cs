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
                _statusFilterIndex = value;
                OnPropertyChanged(nameof(StatusFilterIndex));
                RefreshFilter(); // Ανανέωση της λίστας μόλις αλλάξει η επιλογή
            }
        }

        private bool _forcingKiosk;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GoKioskFullScreen();

            // Χτίζουμε τη διαδρομή
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = System.IO.Path.Combine(baseDir, "Assets", "Templates", "Template.xlsx");

            // Έλεγχος αν υπάρχει
            if (File.Exists(defaultPath))
            {
                LoadData(defaultPath);
            }
            else
            {
                // ΤΩΡΑ ΘΑ ΣΟΥ ΠΕΙ ΤΟΝ ΛΟΓΟ:
                MessageBox.Show($"Δεν βρέθηκε το αρχείο αυτόματης φόρτωσης.\n\n" +
                                $"Έψαξα σε αυτή τη διαδρομή:\n{defaultPath}\n\n" +
                                $"Σιγουρέψου ότι στα Properties του αρχείου στο Visual Studio " +
                                $"το 'Copy to Output Directory' είναι 'Copy always'.",
                                "Το αρχείο λείπει", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Κουμπί για χειροκίνητη επιλογή άλλου αρχείου
        private async void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

                if (openFileDialog.ShowDialog() == true)
                {
                    string path = openFileDialog.FileName;
                    // Κλείδωσε το UI αν θες (π.χ. το κουμπί φόρτωσης)
                    // BtnLoad.IsEnabled = false; 

                    // 2. Τρέξε τη βαριά δουλειά σε άλλο Thread
                    List<PersonRecord> rawList = null;

                    await Task.Run(() =>
                    {
                        // Αυτό τώρα τρέχει στο background και δεν παγώνει το παράθυρο
                        rawList = ExcelLoader.Load(path);
                    });

                    // BtnLoad.IsEnabled = true;

                    if (rawList == null || rawList.Count == 0)
                    {
                        MessageBox.Show("Δεν βρέθηκαν εγγραφές.");
                        return;
                    }

                    RecordsView = CollectionViewSource.GetDefaultView(rawList);
                    RecordsView.Filter = FilterRecords;

                    MessageBox.Show($"Φορτώθηκαν {rawList.Count} εγγραφές επιτυχώς!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Σφάλμα: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- Η ΚΕΝΤΡΙΚΗ ΜΕΘΟΔΟΣ ΦΟΡΤΩΣΗΣ ---
        // Αυτή κάνει όλη τη δουλειά και καλείται και από το Loaded και από το Click
        private void LoadData(string filePath)
        {
            try
            {
                var rawList = ExcelLoader.Load(filePath);

                if (rawList == null || rawList.Count == 0)
                {
                    // Αν είναι αυτόματη φόρτωση, ίσως δεν θες μήνυμα λάθους, 
                    // αλλά εδώ το αφήνω για να ξέρεις αν απέτυχε.
                    MessageBox.Show("Το αρχείο είναι κενό ή δεν φορτώθηκε σωστά.");
                    return;
                }

                RecordsView = CollectionViewSource.GetDefaultView(rawList);
                RecordsView.Filter = FilterRecords;

                // Προαιρετικό: Ένα μήνυμα ότι φορτώθηκαν (μπορείς να το σχολιάσεις αν σε ενοχλεί στην εκκίνηση)
                // MessageBox.Show($"Φορτώθηκαν επιτυχώς {rawList.Count} εγγραφές!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Σφάλμα κατά τη φόρτωση του αρχείου: {ex.Message}");
            }
        }

        // --- Η Λογική του Φιλτραρίσματος (Διορθωμένη με IndexOf) ---
        private bool FilterRecords(object item)
        {
            var person = item as PersonRecord;
            if (person == null) return false;

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
                if (string.IsNullOrEmpty(person.ΑΦΜ) ||
                    person.ΑΦΜ.IndexOf(AfmQuery, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            // 3. Φίλτρο Τηλεφώνου
            if (!string.IsNullOrWhiteSpace(PhoneQuery))
            {
                string cleanQuery = NormalizeDigits(PhoneQuery);
                string p1 = NormalizeDigits(person.Τηλέφωνο);
                string p2 = NormalizeDigits(person.Τηλέφωνο2);

                bool match1 = !string.IsNullOrEmpty(p1) && p1.Contains(cleanQuery);
                bool match2 = !string.IsNullOrEmpty(p2) && p2.Contains(cleanQuery);

                if (!match1 && !match2) return false;
            }

            // --- ΦΙΛΤΡΟ ΚΑΤΑΣΤΑΣΗΣ ---
            if (StatusFilterIndex == 1) // Θέλουμε μόνο τα Checked
            {
                if (!person.Selected) return false;
            }
            else if (StatusFilterIndex == 2) // Θέλουμε μόνο τα Unchecked
            {
                if (person.Selected) return false;
            }

            return true;
        }

        private void RefreshFilter()
        {
            if (RecordsView != null)
            {
                RecordsView.Refresh();
            }
        }

        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return new string(s.Where(char.IsDigit).ToArray());
        }

        private void RecordsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
        }

        protected override void OnStateChanged(EventArgs e)
        {
            base.OnStateChanged(e);
            if (!_forcingKiosk) GoKioskFullScreen();
        }

        private void GoKioskFullScreen()
        {
            if (_forcingKiosk) return;

            try
            {
                _forcingKiosk = true;
                var handle = new WindowInteropHelper(this).Handle;
                var screen = WinForms.Screen.FromHandle(handle);
                var b = screen.Bounds;

                WindowState = WindowState.Normal;
                Left = b.Left;
                Top = b.Top;
                Width = b.Width;
                Height = b.Height;
                Topmost = true;
            }
            finally
            {
                _forcingKiosk = false;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void Aade_Click(object sender, RoutedEventArgs e) { }
        // 1. Ανοίγει το μενού όταν πατάς το κουμπί
        private void OpenExportMenu_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn != null && btn.ContextMenu != null)
            {
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                btn.ContextMenu.IsOpen = true;
            }
        }

        // 2. Επιλογή: Εξαγωγή ΟΛΩΝ (ανεξαρτήτως αν είναι τικαρισμένα ή όχι)
        private void ExportAll_Click(object sender, RoutedEventArgs e)
        {
            var allRecords = RecordsView?.SourceCollection as IEnumerable<PersonRecord>;
            if (allRecords == null) return;

            // Τα παίρνουμε όλα σε λίστα
            var listToExport = allRecords.ToList();

            ExportToExcel(listToExport, "All");
        }

        // 3. Επιλογή: Εξαγωγή των UNCHECKED (όσα δεν έχουν τικ)
        private void ExportUnchecked_Click(object sender, RoutedEventArgs e)
        {
            // Σιγουρεύουμε ότι το Grid έχει σώσει τυχόν αλλαγές της τελευταίας στιγμής
            RecordsGrid.CommitEdit();
            RecordsGrid.CommitEdit();

            var allRecords = RecordsView?.SourceCollection as IEnumerable<PersonRecord>;
            if (allRecords == null) return;

            // Φιλτράρουμε όπου Selected == false
            var listToExport = allRecords.Where(r => r.Selected == false).ToList();

            ExportToExcel(listToExport, "Unchecked");
        }

        // --- ΒΟΗΘΗΤΙΚΗ ΜΕΘΟΔΟΣ (Κάνει την πραγματική δουλειά) ---
        private void ExportToExcel(List<PersonRecord> records, string suffix)
        {
            if (records.Count == 0)
            {
                MessageBox.Show("Δεν βρέθηκαν εγγραφές.", "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Ρύθμιση Άδειας (για EPPlus 7)
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"PYLON_Export_{suffix}_{DateTime.Now:yyyyMMdd}.xlsx",
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

                        // --- 1. ΕΠΙΚΕΦΑΛΙΔΕΣ (Προστέθηκε η "Επιλογή" στην αρχή) ---
                        string[] headers = {
                    "Επιλογή", "Επωνυμία", "Επαφές - Α.Φ.Μ", "Τηλέφωνο 1", "Διακριτικός Τίτλος",
                    "Έγινε Επίδειξη", "Πόλεις (Μεγέθυνση) - Όνομα", "Επαφές - Ημερομηνία 1",
                    "E-mail 1", "E-mail 2", "Τηλέφωνο 2", "GDPR", "Παρουσίαση TWO",
                    "Παλιός Πελάτης", "Επαφές - Σχόλιο", "ΠΡΟΓΡΑΜΜΑ"
                };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            ws.Cells[1, i + 1].Value = headers[i];
                        }

                        // --- 2. ΓΕΜΙΣΜΑ ΔΕΔΟΜΕΝΩΝ ---
                        int row = 2;
                        foreach (var item in records)
                        {
                            // Στήλη 1: Αν είναι επιλεγμένο ή όχι
                            ws.Cells[row, 1].Value = item.Selected ? "ΝΑΙ" : "ΟΧΙ";

                            // Οι υπόλοιπες στήλες μετατοπίστηκαν κατά +1
                            ws.Cells[row, 2].Value = item.Επωνυμία;
                            ws.Cells[row, 3].Value = item.ΑΦΜ;
                            ws.Cells[row, 4].Value = item.Τηλέφωνο;
                            // Κενά πεδία...
                            ws.Cells[row, 11].Value = item.Τηλέφωνο2;

                            row++;
                        }

                        // --- 3. PRO ΜΟΡΦΟΠΟΙΗΣΗ ---
                        var dataRange = ws.Cells[1, 1, row - 1, headers.Length];

                        // Δημιουργία Table
                        var table = ws.Tables.Add(dataRange, "PylonData");
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                        table.ShowFilter = true;

                        // Γραμματοσειρά
                        ws.Cells.Style.Font.Name = "Segoe UI";
                        ws.Cells.Style.Font.Size = 10;

                        // Κεντράρισμα στηλών (Προστέθηκε η στήλη 1)
                        ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Επιλογή
                        ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // ΑΦΜ
                        ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Τηλ 1
                        ws.Column(11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Τηλ 2

                        // AutoFit και "αέρας"
                        ws.Cells.AutoFitColumns();
                        for (int i = 1; i <= headers.Length; i++)
                        {
                            ws.Column(i).Width = ws.Column(i).Width + 2;
                        }

                        // Freeze Panes
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