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
        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

                if (openFileDialog.ShowDialog() == true)
                {
                    LoadData(openFileDialog.FileName); // Καλεί την ίδια μέθοδο
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
        private void ExportExcel_Click(object sender, RoutedEventArgs e) { }
    }
}