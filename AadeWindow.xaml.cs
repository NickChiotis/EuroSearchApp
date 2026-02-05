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
using EuroSearchApp.Services;

namespace EuroSearchApp
{
    public partial class AadeWindow : Window
    {
        // Εδώ θα αποθηκεύσουμε τα αποτελέσματα για να τα πάρει το κεντρικό παράθυρο
        public string FetchedName { get; private set; }
        public string FetchedAddress { get; private set; }

        public AadeWindow(string currentAfm)
        {
            InitializeComponent();
            TxtSearchAfm.Text = currentAfm; // Βάζουμε το ΑΦΜ που μας έστειλε το κεντρικό παράθυρο
        }

        private async void FetchAade_Click(object sender, RoutedEventArgs e)
        {
            // 1. Εμφάνιση Spinner και αρχικό χρώμα
            LoadingSpinner.Visibility = Visibility.Visible;
            TxtStatus.Text = "Γίνεται σύνδεση...";
            TxtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEF3C7"));

            string afm = TxtSearchAfm.Text;
            var result = await System.Threading.Tasks.Task.Run(() => AadeService.GetDetails(afm));

            // 2. Όταν τελειώσει, κρύβουμε τον Spinner
            LoadingSpinner.Visibility = Visibility.Collapsed;

            if (result.Success)
            {
                TxtResultName.Text = result.Name;
                TxtResultAddress.Text = result.Address;

                TxtStatus.Text = "Επιτυχής ανάκτηση στοιχείων!";
                TxtStatus.Foreground = Brushes.Green; // Ή όποιο χρώμα θέλεις για την επιτυχία
            }
            else
            {
                TxtStatus.Text = result.ErrorMessage;
                TxtStatus.Foreground = Brushes.Red;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Αποθηκεύουμε αυτά που γράφουν τα TextBox στα Properties
            FetchedName = TxtResultName.Text;
            FetchedAddress = TxtResultAddress.Text;

            // Επιστρέφουμε "True" στο κεντρικό παράθυρο (σημαίνει ΟΚ)
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Πρόσθεσε αυτό μέσα στην κλάση του παραθύρου
        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }
    }
}
