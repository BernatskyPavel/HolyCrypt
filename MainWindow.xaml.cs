using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.IO;
using System.Text;
using System.Windows;

namespace HolyCryptv3 {
    /// <summary>
    /// Interaction logic for MainWindow2.xaml
    /// </summary>
    public partial class MainWindow2: MetroWindow {
        public MainWindow2() {

            Application.Current.Resources.Source = new Uri($"pack://application:,,,/Localization/Language.{Properties.General.Default.Language}.xaml");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            this.Height = Properties.General.Default.WindowHeight;
            this.Width = Properties.General.Default.WindowWidth;
            _ = _MainFrame.Navigate(new Pages.Home.HomePage());
        }

        private void MenuItem_About_Click(object sender, RoutedEventArgs e) {
            new Windows.About().Show();
        }

        private void MenuItem_Help_Click(object sender, RoutedEventArgs e) {
            //new Windows.Help().Show();
            if (!Directory.Exists("./Help")) {
                _ = Directory.CreateDirectory("./Help");
            }
            if (!File.Exists("./Help/Help.chm")) {
                File.WriteAllBytes("./Help/Help.chm", Properties.Resources.Help);
            }

            System.Windows.Forms.Help.ShowHelp(null, "./Help/Help.chm");
        }

        public async void ShowMyMessage(string? Title, string? Msg) {
            _ = await this.ShowMessageAsync(Title, Msg);
        }
        public async void ShowErrorMessage(string? Title, string? Msg) {
            _ = await this.ShowMessageAsync(Title, Msg, MessageDialogStyle.Affirmative, new MetroDialogSettings() {
                ColorScheme = MetroDialogColorScheme.Inverted,
            });
        }

        //private void MenuItem_Settings_Click(object sender, RoutedEventArgs e) {
        //    this.FirstFlyout.IsOpen = true;
        //}
    }
}
