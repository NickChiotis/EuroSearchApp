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
using System.Windows.Shapes;
using EuroSearchApp.Models;

namespace EuroSearchApp
{
    /// <summary>
    /// Interaction logic for QuickAddWindow.xaml
    /// </summary>
    public partial class QuickAddWindow : Window
    {
        public string FullName { get; set; }
        public string Phone { get; set; }
        public string Afm { get; set; }

        // Η λίστα με τους υπάρχοντες πελάτες για τον έλεγχο
        private IEnumerable<PersonRecord> _existingRecords;

        public QuickAddWindow(string initialName, IEnumerable<PersonRecord> existingRecords)
        {
            InitializeComponent();
            TxtNewName.Text = initialName;
            _existingRecords = existingRecords;
            TxtNewPhone.Focus();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            // Καθαρίζουμε τα κενά από την αρχή και το τέλος
            string name = TxtNewName.Text.Trim();
            string phone = TxtNewPhone.Text.Trim();
            string afm = TxtNewAfm.Text.Trim();

            // 1. Βασικός έλεγχος κενών
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(phone))
            {
                MessageBox.Show("Το Ονοματεπώνυμο και το Τηλέφωνο είναι υποχρεωτικά!");
                return;
            }

            // 2. ΕΛΕΓΧΟΣ ΔΙΠΛΟΤΥΠΩΝ (ΒΕΛΤΙΩΜΕΝΟΣ)
            if (_existingRecords != null)
            {
                // Ετοιμάζουμε τις "καθαρές" μορφή των δεδομένων εισόδου για σύγκριση
                string cleanInputName = RemoveAccents(name);
                string cleanInputPhone = NormalizeDigits(phone);

                bool exists = _existingRecords.Any(p =>
                {
                    // Έλεγχος Ονόματος (Χωρίς τόνους και κεφαλαία)
                    bool nameMatch = !string.IsNullOrEmpty(p.Επωνυμία) &&
                                     RemoveAccents(p.Επωνυμία) == cleanInputName;

                    // Έλεγχος Τηλεφώνου (Συγκρίνουμε μόνο τα ψηφία)
                    // Ελέγχουμε και το Τηλέφωνο 1 και το Τηλέφωνο 2 (αν υπάρχει)
                    string pPhone1 = NormalizeDigits(p.Τηλέφωνο);
                    string pPhone2 = NormalizeDigits(p.Τηλέφωνο2);

                    bool phoneMatch = (pPhone1 == cleanInputPhone) ||
                                      (!string.IsNullOrEmpty(pPhone2) && pPhone2 == cleanInputPhone);

                    // Έλεγχος ΑΦΜ (Ακριβής αντιστοιχία)
                    bool afmMatch = !string.IsNullOrEmpty(afm) &&
                                    p.ΑΦΜ != null &&
                                    p.ΑΦΜ.Trim() == afm;

                    return nameMatch || phoneMatch || afmMatch;
                });

                if (exists)
                {
                    MessageBox.Show("Αυτός ο Πελάτης υπάρχει ήδη (βρέθηκε ίδιο Όνομα, Τηλέφωνο ή ΑΦΜ)!",
                                    "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            // Αν όλα είναι OK
            FullName = name;
            Phone = phone;
            Afm = afm;
            this.DialogResult = true;
        }

        // --- ΒΟΗΘΗΤΙΚΕΣ ΜΕΘΟΔΟΙ (Αντίγραψέ τις μέσα στην κλάση του παραθύρου) ---

        private string RemoveAccents(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            // Απλή αντικατάσταση για ταχύτητα
            string normalized = text.ToUpper();
            string[] accents = { "Ά", "Έ", "Ή", "Ί", "Ό", "Ύ", "Ώ", "Ϊ", "Ϋ" };
            string[] plain = { "Α", "Ε", "Η", "Ι", "Ο", "Υ", "Ω", "Ι", "Υ" };

            for (int i = 0; i < accents.Length; i++)
            {
                normalized = normalized.Replace(accents[i], plain[i]);
            }
            return normalized;
        }

        private string NormalizeDigits(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            // Κρατάμε μόνο τους αριθμούς (διώχνουμε κενά, παύλες, κλπ)
            return new string(text.Where(char.IsDigit).ToArray());
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => this.DialogResult = false;
    }
}
