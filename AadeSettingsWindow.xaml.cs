using System;
using System.Windows;
using System.Windows.Input;
using EuroSearchApp.Services; // Σιγουρέψου ότι έχεις αυτό το using

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
            // Φορτώνουμε μόνο User & Pass
            TxtUser.Text = Properties.Settings.Default.AadeUser;
            TxtPass.Password = Properties.Settings.Default.AadePass;
        }

        private async void TestConnection_Click(object sender, RoutedEventArgs e)
        {
            // Προετοιμασία UI
            LoadingSpinner.Visibility = Visibility.Visible;
            LblStatus.Text = "Γίνεται σύνδεση...";
            LblStatus.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#D97706"));

            // Προσωρινή ανάθεση κωδικών για τη δοκιμή
            Properties.Settings.Default.AadeUser = TxtUser.Text;
            Properties.Settings.Default.AadePass = TxtPass.Password;

            string testAfm = "999977386";

            try
            {
                var result = await System.Threading.Tasks.Task.Run(() => AadeService.GetDetails(testAfm));

                LoadingSpinner.Visibility = Visibility.Collapsed;

                if (result.Success)
                {
                    LblStatus.Text = "Επιτυχής σύνδεση!";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Green;
                }
                else
                {
                    LblStatus.Text = "Αποτυχία: " + result.ErrorMessage;
                    LblStatus.Foreground = System.Windows.Media.Brushes.Red;
                }
            }
            catch (Exception ex)
            {
                LoadingSpinner.Visibility = Visibility.Collapsed;
                LblStatus.Text = "Σφάλμα συστήματος.";
                LblStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Αποθηκεύουμε τα νέα στοιχεία
            Properties.Settings.Default.AadeUser = TxtUser.Text;
            Properties.Settings.Default.AadePass = TxtPass.Password;

            Properties.Settings.Default.Save();

            // Επιστρέφουμε true για να κάνει το MainWindow τη δοκιμή
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Reload();
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