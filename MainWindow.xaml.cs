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

namespace EuroSearchApp
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        // -----------------------------
        // Data + filtering (search)
        // -----------------------------
        public ObservableCollection<PersonRecord> AllRecords { get; } =
            new ObservableCollection<PersonRecord>();

        public ICollectionView RecordsView { get; }

        private string _nameQuery = "";
        public string NameQuery
        {
            get => _nameQuery;
            set
            {
                _nameQuery = value ?? "";
                OnPropertyChanged(nameof(NameQuery));
                RecordsView.Refresh();
            }
        }

        private string _phoneQuery = "";
        public string PhoneQuery
        {
            get => _phoneQuery;
            set
            {
                _phoneQuery = value ?? "";
                OnPropertyChanged(nameof(PhoneQuery));
                RecordsView.Refresh();
            }
        }

        private string _afmQuery = "";
        public string AfmQuery
        {
            get => _afmQuery;
            set
            {
                _afmQuery = value ?? "";
                OnPropertyChanged(nameof(AfmQuery));
                RecordsView.Refresh();
            }
        }

        // -----------------------------
        // Kiosk fullscreen lock
        // -----------------------------
        private bool _forcingKiosk;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            RecordsView = CollectionViewSource.GetDefaultView(AllRecords);
            RecordsView.Filter = FilterRecord;
        }

        // Καλείται από XAML: Loaded="Window_Loaded"
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GoKioskFullScreen();
        }

        // Καλείται από XAML: Click="Close_Click"
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Καλείται από XAML: Click="LoadExcel_Click"
        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Επίλεξε Excel αρχείο"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                var list = ExcelLoader.Load(dlg.FileName);

                AllRecords.Clear();
                foreach (var item in list)
                    AllRecords.Add(item);

                RecordsView.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Σφάλμα φόρτωσης", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Αν κάτι προσπαθήσει να αλλάξει κατάσταση/μέγεθος, το ξαναφέρνουμε σε kiosk fullscreen
        protected override void OnStateChanged(EventArgs e)
        {
            base.OnStateChanged(e);

            if (!_forcingKiosk)
                GoKioskFullScreen();
        }

        private void GoKioskFullScreen()
        {
            if (_forcingKiosk) return;

            try
            {
                _forcingKiosk = true;

                // Βρες την οθόνη που βρίσκεται το window
                var handle = new WindowInteropHelper(this).Handle;
                var screen = WinForms.Screen.FromHandle(handle);

                // Full bounds (περιλαμβάνει taskbar area)
                var b = screen.Bounds;

                // Για πραγματικό fullscreen: κανονική κατάσταση + manual sizing στα bounds
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

        private bool FilterRecord(object obj)
        {
            var x = obj as PersonRecord;
            if (x == null) return false;

            var nameQ = (NameQuery ?? "").Trim();
            var phoneQ = NormalizeDigits((PhoneQuery ?? "").Trim());
            var afmQ = NormalizeDigits((AfmQuery ?? "").Trim());

            bool okName = string.IsNullOrWhiteSpace(nameQ) ||
                          ((x.Ονομα ?? "").IndexOf(nameQ, StringComparison.OrdinalIgnoreCase) >= 0);

            bool okPhone = string.IsNullOrWhiteSpace(phoneQ) ||
                           NormalizeDigits(x.Τηλέφωνο).Contains(phoneQ);

            bool okAfm = string.IsNullOrWhiteSpace(afmQ) ||
                         NormalizeDigits(x.ΑΦΜ).Contains(afmQ);

            return okName && okPhone && okAfm;
        }

        private static string NormalizeDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return new string(s.Where(char.IsDigit).ToArray());
        }

        // -----------------------------
        // INotifyPropertyChanged
        // -----------------------------
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void Aade_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
