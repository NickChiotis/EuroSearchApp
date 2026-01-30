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

namespace EuroSearchApp
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private bool _startupLoaded;

        private DataTable _table;
        public DataView RecordsView { get; private set; }

        private string _nameQuery = "";
        public string NameQuery
        {
            get { return _nameQuery; }
            set
            {
                _nameQuery = value ?? "";
                OnPropertyChanged(nameof(NameQuery));
                ApplyRowFilter();
            }
        }

        private string _phoneQuery = "";
        public string PhoneQuery
        {
            get { return _phoneQuery; }
            set
            {
                _phoneQuery = value ?? "";
                OnPropertyChanged(nameof(PhoneQuery));
                ApplyRowFilter();
            }
        }

        private string _afmQuery = "";
        public string AfmQuery
        {
            get { return _afmQuery; }
            set
            {
                _afmQuery = value ?? "";
                OnPropertyChanged(nameof(AfmQuery));
                ApplyRowFilter();
            }
        }

        // Kiosk fullscreen lock
        private bool _forcingKiosk;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            _table = new DataTable();
            _table.CaseSensitive = false;

            RecordsView = _table.DefaultView; // DataView
            OnPropertyChanged(nameof(RecordsView));
        }

        // Loaded="Window_Loaded"
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GoKioskFullScreen();

            if (_startupLoaded) return;
            _startupLoaded = true;

            try
            {
                LoadExcelIntoGridFromAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Σφάλμα φόρτωσης Excel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Click="LoadExcel_Click" (reload)
        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LoadExcelIntoGridFromAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Σφάλμα φόρτωσης", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadExcelIntoGridFromAssets()
        {
            const string assetPath = "Assets/Templates/Template.xlsx";

            // Resource -> temp path
            string tempExcelPath = ExcelAssetHelper.ExtractExcelFromAssetsToTemp(assetPath);

            // true: 1η γραμμή headers, false: όλα data
            _table = ExcelAnyLoader.LoadAllToDataTable(tempExcelPath, true);
            _table.CaseSensitive = false;

            RecordsView = _table.DefaultView;
            OnPropertyChanged(nameof(RecordsView));

            ApplyRowFilter();
        }

        // AutoGeneratingColumn="RecordsGrid_AutoGeneratingColumn"
        // Κλειδώνει όλες τις auto στήλες (εκτός από checkbox που το έχεις fixed στο XAML)
        private void RecordsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == "Επιλογή")
            {
                e.Cancel = true; // να μη διπλασιαστεί
                return;
            }

            e.Column.IsReadOnly = true;
        }

        private void ApplyRowFilter()
        {
            if (RecordsView == null) return;
            if (RecordsView.Table == null) return;

            var table = RecordsView.Table;

            string nameCol = ResolveColumn(table, new[] { "Όνομα", "Ονομα", "Name" });
            string phoneCol = ResolveColumn(table, new[] { "Τηλέφωνο", "Τηλεφωνο", "Phone" });
            string afmCol = ResolveColumn(table, new[] { "ΑΦΜ", "Α.Φ.Μ", "AFM", "Vat" });

            string nameQ = (NameQuery ?? "").Trim();
            string phoneQ = (PhoneQuery ?? "").Trim();
            string afmQ = (AfmQuery ?? "").Trim();

            var parts = new System.Collections.Generic.List<string>();

            if (!string.IsNullOrWhiteSpace(nameQ) && !string.IsNullOrWhiteSpace(nameCol))
            {
                string v = EscapeLike(nameQ);
                parts.Add(string.Format("[{0}] LIKE '%{1}%'", nameCol, v));
            }

            if (!string.IsNullOrWhiteSpace(phoneQ) && !string.IsNullOrWhiteSpace(phoneCol))
            {
                // προσέγγιση: αφαιρεί space/./-/+/( ) από την τιμή για να ταιριάζει πιο συχνά
                string v = EscapeLike(NormalizeDigits(phoneQ));
                if (!string.IsNullOrEmpty(v))
                    parts.Add(string.Format("{0} LIKE '%{1}%'", NormalizeExpr(phoneCol), v));
            }

            if (!string.IsNullOrWhiteSpace(afmQ) && !string.IsNullOrWhiteSpace(afmCol))
            {
                string v = EscapeLike(NormalizeDigits(afmQ));
                if (!string.IsNullOrEmpty(v))
                    parts.Add(string.Format("{0} LIKE '%{1}%'", NormalizeExpr(afmCol), v));
            }

            RecordsView.RowFilter = (parts.Count == 0) ? "" : string.Join(" AND ", parts.ToArray());
        }

        private static string ResolveColumn(DataTable table, string[] candidates)
        {
            foreach (var c in candidates)
            {
                if (table.Columns.Contains(c)) return c;
            }
            return null;
        }

        private static string EscapeLike(string input)
        {
            if (input == null) return "";
            // escape ' and LIKE special chars
            string s = input.Replace("'", "''");
            s = s.Replace("[", "[[]");
            s = s.Replace("%", "[%]");
            s = s.Replace("_", "[_]");
            s = s.Replace("]", "[]]");
            return s;
        }

        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return new string(s.Where(char.IsDigit).ToArray());
        }

        private static string NormalizeExpr(string colName)
        {
            // REPLACE chain για “καθάρισμα” σε RowFilter
            // Προσοχή: RowFilter δουλεύει με single quotes
            string x = string.Format("[{0}]", colName);
            x = string.Format("REPLACE({0}, ' ', '')", x);
            x = string.Format("REPLACE({0}, '-', '')", x);
            x = string.Format("REPLACE({0}, '+', '')", x);
            x = string.Format("REPLACE({0}, '(', '')", x);
            x = string.Format("REPLACE({0}, ')', '')", x);
            x = string.Format("REPLACE({0}, '.', '')", x);
            x = string.Format("REPLACE({0}, '/', '')", x);
            return x;
        }

        // Kiosk enforce
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
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }

        private void Aade_Click(object sender, RoutedEventArgs e) { }
        private void ExportExcel_Click(object sender, RoutedEventArgs e) { }
    }
}