using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input; // Απαραίτητο για το MouseButtonEventArgs

namespace EuroSearchApp
{
    public partial class GiftViewerWindow : Window
    {
        public string SelectedGiftName { get; private set; }

        public GiftViewerWindow()
        {
            InitializeComponent();
        }

        // Η μέθοδος αναζήτησης διορθωμένη
        private void TxtGiftSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Παίρνουμε το DataView από τον πίνακα
            var dv = GiftsGrid.ItemsSource as DataView;

            if (dv == null) return;

            // Καθαρίζουμε το κείμενο αναζήτησης
            string query = TxtGiftSearch.Text.Trim().Replace("'", "''");

            if (string.IsNullOrEmpty(query))
            {
                dv.RowFilter = ""; // Αν είναι άδειο, δείξε τα πάντα
            }
            else
            {
                try
                {
                    // Παίρνουμε το όνομα της 1ης στήλης (ΕΙΔΟΣ)
                    string colName = dv.Table.Columns[0].ColumnName;

                    // Αλλάζουμε το φίλτρο: 
                    // Αφαιρώντας το % από την αρχή, η αναζήτηση γίνεται "Starts With"
                    dv.RowFilter = $"[{colName}] LIKE '{query}%'";
                }
                catch
                {
                    // Σε περίπτωση σφάλματος, επαναφορά
                    dv.RowFilter = "";
                }
            }
        }

        // Διπλό κλικ για επιλογή
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
    }
}