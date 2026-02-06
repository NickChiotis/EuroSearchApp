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
            string name = TxtNewName.Text.Trim();
            string phone = TxtNewPhone.Text.Trim();
            string afm = TxtNewAfm.Text.Trim();

            // 1. Βασικός έλεγχος κενών
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(phone))
            {
                MessageBox.Show("Το Ονοματεπώνυμο και το Τηλέφωνο είναι υποχρεωτικά!");
                return;
            }

            // 2. ΕΛΕΓΧΟΣ ΔΙΠΛΟΤΥΠΩΝ
            if (_existingRecords != null)
            {
                bool exists = _existingRecords.Any(p =>
                    (p.Επωνυμία != null && p.Επωνυμία.Equals(name, StringComparison.OrdinalIgnoreCase)) ||
                    (p.Τηλέφωνο != null && p.Τηλέφωνο == phone) ||
                    (!string.IsNullOrEmpty(afm) && p.ΑΦΜ != null && p.ΑΦΜ == afm)
                );

                if (exists)
                {
                    MessageBox.Show("Αυτός ο Πελάτης υπάρχει ήδη (ίδιο Όνομα, Τηλέφωνο ή ΑΦΜ)!", "Προσοχή", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            FullName = name;
            Phone = phone;
            Afm = afm;
            this.DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => this.DialogResult = false;
    }
}
