using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OfficeOpenXml; // Απαραίτητο για το EPPlus

namespace EuroSearchApp
{
    public partial class GiftViewerWindow : Window
    {
        public string SelectedGiftName { get; private set; }

        // Το Path για το αρχείο των δώρων
        private string giftPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Gifts", "temp_gift.xlsx");

        public GiftViewerWindow()
        {
            InitializeComponent();
        }

        // Αναζήτηση
        private void TxtGiftSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var dv = GiftsGrid.ItemsSource as DataView;
            if (dv == null) return;

            string query = TxtGiftSearch.Text.Trim().Replace("'", "''");

            if (string.IsNullOrEmpty(query))
            {
                dv.RowFilter = "";
            }
            else
            {
                try
                {
                    string colName = dv.Table.Columns[0].ColumnName;
                    dv.RowFilter = $"[{colName}] LIKE '{query}%'";
                }
                catch
                {
                    dv.RowFilter = "";
                }
            }
        }

        // Επιλογή με διπλό κλικ
        private void GiftsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var row = GiftsGrid.SelectedItem as DataRowView;
            if (row != null)
            {
                SelectedGiftName = row[0].ToString();
                this.DialogResult = true;
                this.Close();
            }
        }

        // --- ΛΕΙΤΟΥΡΓΙΕΣ ΚΟΥΜΠΙΩΝ ---

        // Κουμπί (-) : Καθαρισμός (Χωρίς Δώρο)
        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            SelectedGiftName = ""; // Επιστρέφει κενό
            this.DialogResult = true; // Κλείνει το παράθυρο και ενημερώνει το Grid
            this.Close();
        }

        // Κουμπί (+) : Εμφάνιση παραθύρου προσθήκης
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            TxtNewGift.Text = "";
            AddOverlay.Visibility = Visibility.Visible;
            TxtNewGift.Focus();
        }

        // Ακύρωση Προσθήκης
        private void BtnCancelAdd_Click(object sender, RoutedEventArgs e)
        {
            AddOverlay.Visibility = Visibility.Collapsed;
        }

        // Αποθήκευση Νέου Δώρου
        private void BtnSaveNew_Click(object sender, RoutedEventArgs e)
        {
            string newGift = TxtNewGift.Text.Trim();
            if (string.IsNullOrEmpty(newGift))
            {
                MessageBox.Show("Παρακαλώ εισάγετε ονομασία δώρου.");
                return;
            }

            try
            {
                // 1. Αποθήκευση στο Excel
                SaveGiftToExcel(newGift);

                // 2. Ενημέρωση του Πίνακα (DataView) τοπικά
                var view = GiftsGrid.ItemsSource as DataView;
                if (view != null)
                {
                    DataRow newRow = view.Table.NewRow();
                    newRow[0] = newGift;
                    view.Table.Rows.InsertAt(newRow, 0); // Το βάζουμε στην κορυφή (μπροστά)
                }

                // 3. Κλείσιμο του Overlay
                AddOverlay.Visibility = Visibility.Collapsed;
                TxtGiftSearch.Text = ""; // Καθαρισμός φίλτρου
            }
            catch (Exception ex)
            {
                MessageBox.Show("Σφάλμα κατά την αποθήκευση: " + ex.Message);
            }
        }

        // --- ΛΕΙΤΟΥΡΓΙΑ ΔΙΑΓΡΑΦΗΣ ---
        private void BtnDeleteRow_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Είστε σίγουροι ότι θέλετε να διαγράψετε αυτό το δώρο;", "Επιβεβαίωση", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                return;

            var button = sender as Button;
            var rowView = button.DataContext as DataRowView;

            if (rowView != null)
            {
                string giftToDelete = rowView[0].ToString();

                try
                {
                    // 1. Διαγραφή από Excel
                    DeleteGiftFromExcel(giftToDelete);

                    // 2. Διαγραφή από τον Πίνακα (UI)
                    rowView.Delete();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Σφάλμα κατά τη διαγραφή: " + ex.Message);
                }
            }
        }

        // --- ΜΕΘΟΔΟΙ EXCEL ---

        private void SaveGiftToExcel(string giftName)
        {
            if (!File.Exists(giftPath)) return;

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(giftPath)))
            {
                var ws = package.Workbook.Worksheets[0];
                if (ws.Dimension == null)
                {
                    ws.Cells[1, 1].Value = "ΕΙΔΟΣ";
                    ws.Cells[2, 1].Value = giftName;
                }
                else
                {
                    // Ελέγχουμε αν υπάρχει ήδη
                    bool exists = false;
                    for (int r = 2; r <= ws.Dimension.End.Row; r++)
                    {
                        if (ws.Cells[r, 1].Value?.ToString() == giftName)
                        {
                            exists = true;
                            break;
                        }
                    }

                    if (!exists)
                    {
                        int lastRow = ws.Dimension.End.Row;
                        ws.Cells[lastRow + 1, 1].Value = giftName;
                        package.Save();
                    }
                }
            }
        }

        private void DeleteGiftFromExcel(string giftName)
        {
            if (!File.Exists(giftPath)) return;

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(giftPath)))
            {
                var ws = package.Workbook.Worksheets[0];
                if (ws.Dimension == null) return;

                for (int r = 2; r <= ws.Dimension.End.Row; r++)
                {
                    if (ws.Cells[r, 1].Value?.ToString() == giftName)
                    {
                        ws.DeleteRow(r); // Διαγραφή της γραμμής
                        package.Save();
                        return;
                    }
                }
            }
        }
    }
}