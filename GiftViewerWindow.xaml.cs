using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.IO;
using OfficeOpenXml; // Χρειάζεται για την αποθήκευση του νέου δώρου

namespace EuroSearchApp
{
    public partial class GiftViewerWindow : Window
    {
        // Η επιλογή που θα επιστραφεί στο κυρίως παράθυρο
        public string SelectedGiftName { get; private set; }

        // Η λίστα που περιέχει τα δεδομένα (Όνομα, Απόθεμα, κτλ)
        public List<GiftStockItem> GiftItems { get; set; } = new List<GiftStockItem>();

        // Path για το Excel (για την προσθήκη νέου δώρου)
        private string giftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "EuroGifts.xlsx");

        public GiftViewerWindow()
        {
            InitializeComponent();
        }

        // --- 1. ΔΙΑΧΕΙΡΙΣΗ ΠΛΗΚΤΡΩΝ (ESCAPE) ---
        // ΑΥΤΗ Η ΜΕΘΟΔΟΣ ΕΛΕΙΠΕ ΚΑΙ ΣΟΥ ΕΒΓΑΖΕ ΤΟ ΣΦΑΛΜΑ
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // Αν είναι ανοιχτό το Overlay προσθήκης, κλείσε μόνο αυτό
                if (AddOverlay != null && AddOverlay.Visibility == Visibility.Visible)
                {
                    AddOverlay.Visibility = Visibility.Collapsed;
                    return;
                }

                // Αλλιώς κλείσε το παράθυρο
                this.Close();
            }
        }

        // --- 2. ΑΝΑΖΗΤΗΣΗ (Search) ---
        private void TxtGiftSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = TxtGiftSearch.Text.Trim();

            if (string.IsNullOrEmpty(filter))
            {
                // Αν δεν γράφει τίποτα, δείξε τα όλα
                GiftsGrid.ItemsSource = GiftItems;
            }
            else
            {
                // Αν γράφει, φίltrare τη λίστα στη μνήμη
                var filtered = GiftItems.Where(x => x.Name != null &&
                                               x.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                GiftsGrid.ItemsSource = filtered;
            }
        }

        // --- 3. ΕΠΙΛΟΓΗ ΔΩΡΟΥ (Double Click) ---
        private void GiftsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GiftsGrid.SelectedItem is GiftStockItem item)
            {
                // Έλεγχος αν έχει μείνει απόθεμα
                if (item.RemainingQty <= 0)
                {
                    var result = MessageBox.Show(
                        $"Το δώρο '{item.Name}' έχει εξαντληθεί (Διαθέσιμα: 0).\nΘέλετε να το επιλέξετε παρόλα αυτά;",
                        "Εξαντλημένο Δώρο",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (result == MessageBoxResult.No) return;
                }

                SelectedGiftName = item.Name;
                DialogResult = true; // Επιστρέφει true στο MainWindow
                Close();
            }
        }

        // --- 4. ΔΙΑΓΡΑΦΗ ΔΩΡΟΥ (Κάδος) ---
        private void BtnDeleteRow_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var item = button.DataContext as GiftStockItem;

            if (item != null)
            {
                if (MessageBox.Show($"Είσαι σίγουρος ότι θέλεις να διαγράψεις το δώρο: {item.Name};\n(Θα αφαιρεθεί μόνο από τη λίστα προβολής)",
                    "Επιβεβαίωση", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    GiftItems.Remove(item);
                    RefreshGrid();
                }
            }
        }

        // --- 5. ΚΑΘΑΡΙΣΜΟΣ ΕΠΙΛΟΓΗΣ (Κουμπί Χ) ---
        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            SelectedGiftName = null; // Κανένα δώρο
            DialogResult = true;
            Close();
        }

        // --- 6. ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΔΩΡΟΥ (Overlay) ---
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (AddOverlay != null)
            {
                AddOverlay.Visibility = Visibility.Visible;
                if (TxtNewGift != null)
                {
                    TxtNewGift.Text = "";
                    TxtNewGift.Focus();
                }
            }
        }

        private void BtnCancelAdd_Click(object sender, RoutedEventArgs e)
        {
            if (AddOverlay != null) AddOverlay.Visibility = Visibility.Collapsed;
        }

        private void BtnSaveNew_Click(object sender, RoutedEventArgs e)
        {
            string newName = TxtNewGift.Text.Trim();
            if (string.IsNullOrEmpty(newName)) return;

            try
            {
                // 1. Το προσθέτουμε στη λίστα (Προσωρινά με 0 απόθεμα, ή όσο θες)
                var newItem = new GiftStockItem
                {
                    Name = newName,
                    TotalQty = 0,
                    UsedQty = 0
                };

                GiftItems.Insert(0, newItem); // Προσθήκη στην αρχή
                RefreshGrid();

                // 2. Προσπάθεια αποθήκευσης στο Excel (Προαιρετικό)
                SaveGiftToExcel(newName);

                if (AddOverlay != null) AddOverlay.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά την αποθήκευση: " + ex.Message);
            }
        }

        // --- ΒΟΗΘΗΤΙΚΕΣ ΜΕΘΟΔΟΙ ---

        private void RefreshGrid()
        {
            GiftsGrid.ItemsSource = null; // Reset για να δει τις αλλαγές
            GiftsGrid.ItemsSource = GiftItems;
        }

        private void SaveGiftToExcel(string giftName)
        {
            if (!File.Exists(giftPath)) return;

            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(giftPath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension != null)
                    {
                        int lastRow = ws.Dimension.End.Row;
                        ws.Cells[lastRow + 1, 1].Value = giftName;
                        ws.Cells[lastRow + 1, 2].Value = 0; // Αρχική ποσότητα 0
                        package.Save();
                    }
                }
            }
            catch { /* Αγνοούμε λάθη excel αν είναι ανοιχτό */ }
        }
    }

    // --- Η ΚΛΑΣΗ ΜΟΝΤΕΛΟΥ (ΕΔΩ ΕΙΝΑΙ Η ΣΩΣΤΗ ΘΕΣΗ - ΕΚΤΟΣ ΤΟΥ WINDOW) ---
    public class GiftStockItem
    {
        public string Name { get; set; }        // Όνομα Δώρου
        public int TotalQty { get; set; }       // Αρχικά (από Excel)
        public int UsedQty { get; set; }        // Χρησιμοποιημένα (από Πελάτες)

        // Αυτό υπολογίζει αυτόματα το υπόλοιπο
        public int RemainingQty => TotalQty - UsedQty;

        // Βοηθητικό για να κοκκινίζει η γραμμή στο XAML αν τελειώνει
        public bool IsLowStock => RemainingQty <= 0;
    }
}