using System.Windows;

namespace EuroSearchApp
{
    public partial class LeadInfoWindow : Window
    {
        public LeadInfoWindow()
        {
            InitializeComponent();
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            // ΕΛΕΓΧΟΣ: Μόνο για το Φαρμακείο (2) και το Τηλέφωνο (4)
            if (string.IsNullOrWhiteSpace(TxtPharmacy.Text) ||
                string.IsNullOrWhiteSpace(TxtPhone.Text))
            {
                MessageBox.Show("Το Φαρμακείο και το Τηλέφωνο είναι υποχρεωτικά πεδία!");
                return;
            }

            // Αν όλα είναι οκ, κλείνει το παράθυρο και επιστρέφει true
            this.DialogResult = true;
        }

        // Helper properties για να παίρνεις τις τιμές εύκολα από το MainWindow
        public string Promoter => ComboPromoter.SelectionBoxItem?.ToString();
        public string Pharmacy => TxtPharmacy.Text;
        public string City => TxtCity.Text;
        public string Phone => TxtPhone.Text;
        public string Email => TxtEmail.Text;
        public string Program => TxtProgram.Text;
        public string Client => ComboClient.SelectionBoxItem?.ToString();
        public string Presentation => ComboPresentation.SelectionBoxItem?.ToString();
        public string Sales => ComboSales.SelectionBoxItem?.ToString();
        public string Notes => TxtNotes.Text;
    }
}