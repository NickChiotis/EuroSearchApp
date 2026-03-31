using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EuroSearchApp
{
    public partial class LeadInfoWindow : Window
    {
        public LeadInfoWindow()
        {
            System.Windows.Media.RenderOptions.ProcessRenderMode = System.Windows.Interop.RenderMode.SoftwareOnly;
            InitializeComponent();


            // ΦΟΡΤΩΣΗ: Αν υπάρχει αποθηκευμένος Promoter, επέλεξέ τον
            string lastIdx = Properties.Settings.Default.LastPromoter;
            if (!string.IsNullOrEmpty(lastIdx))
            {
                int index = int.Parse(lastIdx);
                // Βεβαιωνόμαστε ότι ο index υπάρχει ακόμα στη λίστα
                if (index < ComboPromoter.Items.Count)
                {
                    ComboPromoter.SelectedIndex = index;
                }
            }
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            // Χρησιμοποιούμε SelectedIndex > 0 για να προσπεράσουμε το Placeholder στη θέση 0
            bool isPromoterEmpty = ComboPromoter.SelectedIndex <= 0;
            bool isPharmacyEmpty = string.IsNullOrWhiteSpace(TxtPharmacy.Text);
            bool isPhoneEmpty = string.IsNullOrWhiteSpace(TxtPhone.Text);

            if (isPromoterEmpty || isPharmacyEmpty || isPhoneEmpty)
            {
                MessageBox.Show("Ο Promoter, το Φαρμακείο και το Τηλέφωνο είναι υποχρεωτικά πεδία!");
                return;
            }

            Properties.Settings.Default.LastPromoter = ComboPromoter.SelectedIndex.ToString();
            Properties.Settings.Default.Save(); // Αυτό το γράφει στο δίσκο

            this.DialogResult = true;
        }

        // Helper properties για να παίρνεις τις τιμές εύκολα από το MainWindow
        public string Promoter => ComboPromoter.SelectionBoxItem?.ToString();
        public string Pharmacy => TxtPharmacy.Text;
        public string City => TxtCity.Text;
        public string Phone => TxtPhone.Text;
        public string Email => TxtEmail.Text;
        public string Program => ComboProgram.SelectionBoxItem?.ToString();
        public string Client => ComboClient.SelectionBoxItem?.ToString();
        public string Presentation => ComboPresentation.SelectionBoxItem?.ToString();
        public string Sales => ComboSales.SelectionBoxItem?.ToString();
        public string Notes => TxtNotes.Text;

        private void QuickNote_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn)
            {
                string noteText = "";

                // Αν το κουμπί έχει TextBlock μέσα του (τα μεγάλα κουμπιά)
                if (btn.Content is TextBlock tb)
                {
                    noteText = tb.Text;
                }
                // Αν το κουμπί έχει απλό κείμενο (τα μικρά κουμπιά)
                else
                {
                    noteText = btn.Content.ToString();
                }

                // Καθαρίζουμε το "+" και προσθέτουμε στο TextBox
                noteText = noteText.Replace("+ ", "");

                if (string.IsNullOrWhiteSpace(TxtNotes.Text))
                    TxtNotes.Text = noteText;
                else
                    TxtNotes.Text += ", " + noteText;

                // Cursor στο τέλος
                TxtNotes.Focus();
                TxtNotes.SelectionStart = TxtNotes.Text.Length;
            }
        }

        private void TxtPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Ελέγχει αν ο χαρακτήρας που πατήθηκε είναι ψηφίο (0-9)
            // Αν ΔΕΝ είναι αριθμός, τότε e.Handled = true (ακυρώνει το γράψιμο)
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static bool IsTextAllowed(string text)
        {
            // Χρησιμοποιούμε Regex για να επιτρέψουμε μόνο αριθμούς
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[^0-9]+");
            return !regex.IsMatch(text);
        }

        private void TxtPhone_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsTextAllowed(text))
                {
                    e.CancelCommand(); // Ακυρώνει την επικόλληση αν περιέχει γράμματα
                }
            }
            else
            {
                e.CancelCommand();
            }
        }
    }
}