using System.Windows;
using EuroSearchApp.Services;

namespace EuroSearchApp
{
    public partial class AadeSettingsWindow : Window
    {
        public AadeSettingsWindow()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            // Φορτώνουμε τις αποθηκευμένες ρυθμίσεις στα κουτάκια
            TxtUser.Text = Properties.Settings.Default.AadeUser;
            TxtPass.Password = Properties.Settings.Default.AadePass;
            TxtMyAfm.Text = Properties.Settings.Default.MyAfm;
        }

        private async void TestConnection_Click(object sender, RoutedEventArgs e)
        {
            LblStatus.Text = "Γίνεται δοκιμή...";
            LblStatus.Foreground = System.Windows.Media.Brushes.Blue;

            // Προσωρινή αποθήκευση για το τεστ (χωρίς Save)
            Properties.Settings.Default.AadeUser = TxtUser.Text;
            Properties.Settings.Default.AadePass = TxtPass.Password;
            Properties.Settings.Default.MyAfm = TxtMyAfm.Text;

            // Δοκιμαστικό ΑΦΜ (π.χ. Υπουργείο Οικονομικών 090165560 ή ΓΓΠΣ 999977386)
            string testAfm = "090165560";

            var result = await System.Threading.Tasks.Task.Run(() => AadeService.GetDetails(testAfm));

            if (result.Success)
            {
                LblStatus.Text = $"ΕΠΙΤΥΧΙΑ! Η σύνδεση λειτουργεί.\nΒρέθηκε: {result.Name}";
                LblStatus.Foreground = System.Windows.Media.Brushes.Green;
            }
            else
            {
                LblStatus.Text = $"ΑΠΟΤΥΧΙΑ: {result.ErrorMessage}";
                LblStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Αποθήκευση μόνιμα
            Properties.Settings.Default.AadeUser = TxtUser.Text;
            Properties.Settings.Default.AadePass = TxtPass.Password;
            Properties.Settings.Default.MyAfm = TxtMyAfm.Text;

            Properties.Settings.Default.Save(); // <--- Αυτό γράφει στο δίσκο

            MessageBox.Show("Οι ρυθμίσεις αποθηκεύτηκαν!", "Επιτυχία", MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            // Επαναφορά (Reload) σε περίπτωση που πειράξαμε κάτι στο Test αλλά πατήσαμε Cancel
            Properties.Settings.Default.Reload();
            DialogResult = false;
            Close();
        }
    }
}