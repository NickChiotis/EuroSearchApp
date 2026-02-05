using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input; // Απαραίτητο για το MouseButtonEventArgs

namespace EuroSearchApp
{
    public partial class GiftViewerWindow : Window
    {
        private DataView _fullView;

        // Ιδιότητα που θα κρατάει το όνομα του δώρου για να το διαβάσει το MainWindow
        public string SelectedGiftName { get; private set; }

        public GiftViewerWindow()
        {
            InitializeComponent();
        }

        // Αυτό καλείται από το MainWindow για να γεμίσει τα δεδομένα στον πίνακα
        public void SetData(DataView dv)
        {
            _fullView = dv;
            GiftsGrid.ItemsSource = _fullView;
        }

        // Λειτουργία αναζήτησης μέσα στον viewer
        private void TxtGiftSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_fullView == null) return;

            string query = TxtGiftSearch.Text.Replace("'", "''"); // Ασφάλεια για ειδικούς χαρακτήρες

            try
            {
                // Προσπαθεί να κάνει φιλτράρισμα στη στήλη "ΕΙΔΟΣ"
                _fullView.RowFilter = $"[ΕΙΔΟΣ] LIKE '%{query}%'";
            }
            catch
            {
                // Αν η στήλη έχει κενά ή άλλο όνομα, ψάχνει στην πρώτη διαθέσιμη στήλη
                if (_fullView.Table.Columns.Count > 0)
                {
                    string colName = _fullView.Table.Columns[0].ColumnName;
                    _fullView.RowFilter = $"[{colName}] LIKE '%{query}%'";
                }
            }
        }

        // Λειτουργία Quick Pick: Με διπλό κλικ επιλέγεται το δώρο και κλείνει το παράθυρο
        private void GiftsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Παίρνουμε τη γραμμή που έγινε το κλικ
            var row = GiftsGrid.SelectedItem as DataRowView;

            if (row != null)
            {
                // Αποθηκεύουμε το κείμενο της πρώτης στήλης (το όνομα του δώρου)
                SelectedGiftName = row[0].ToString();

                // Θέτουμε το DialogResult σε true για να ξέρει το MainWindow ότι έγινε επιλογή
                this.DialogResult = true;
                this.Close();
            }
        }
    }
}