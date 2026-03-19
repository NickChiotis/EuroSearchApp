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
            // Έλεγχος για τα υποχρεωτικά πεδία 2 έως 9
            if (string.IsNullOrWhiteSpace(TxtPharmacy.Text) ||
                string.IsNullOrWhiteSpace(TxtCity.Text) ||
                string.IsNullOrWhiteSpace(TxtPhone.Text) ||
                string.IsNullOrWhiteSpace(TxtEmail.Text) ||
                string.IsNullOrWhiteSpace(TxtProgram.Text) ||
                string.IsNullOrWhiteSpace(TxtClient.Text) ||
                string.IsNullOrWhiteSpace(TxtPresentation.Text) ||
                string.IsNullOrWhiteSpace(TxtSales.Text))
            {
                MessageBox.Show("Τα πεδία 2 έως 9 είναι υποχρεωτικά!");
                return;
            }
            this.DialogResult = true; // Προχωράμε
        }
    }
}