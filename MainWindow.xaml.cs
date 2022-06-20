using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.IO;
using System.Text;
using System.Windows;

namespace StegoLine {
    /// <summary>
    /// Interaction logic for MainWindow2.xaml
    /// </summary>
    public partial class MainWindow: MetroWindow {
        public MainWindow() {

            Application.Current.Resources.Source = new Uri($"pack://application:,,,/Localization/Language.{Properties.General.Default.Language}.xaml");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            this.Height = Properties.General.Default.WindowHeight;
            this.Width = Properties.General.Default.WindowWidth;
            _ = _MainFrame.Navigate(new Pages.Home.HomePage());
            MinWValue.StringFormat = @"{0:F2} pt";
            MaxWValue.StringFormat = @"{0:F2} pt";
            MinAValue.StringFormat = @"{0:F3}%";
            MaxAValue.StringFormat = @"{0:F3}%";
        }

        private void MenuItem_About_Click(object sender, RoutedEventArgs e) {
            new Windows.About().Show();
        }

        private void MenuItem_Help_Click(object sender, RoutedEventArgs e) {
            System.Windows.Forms.Help.ShowHelp(null, "./Resources/Help.chm");
        }

        public async void ShowMyMessage(string? Title, string? Msg) {
            _ = await this.ShowMessageAsync(Title, Msg);
        }

        public async void ShowErrorMessage(string? Title, string? Msg) {
            _ = await this.ShowMessageAsync(Title, Msg, MessageDialogStyle.Affirmative, new MetroDialogSettings() {
                ColorScheme = MetroDialogColorScheme.Inverted,
            });
        }

        private void MenuItem_Settings_Click(object sender, RoutedEventArgs e) {
            if (!this.FirstFlyout.IsOpen) {
                Properties.General.Default.Reload();
                this.FirstFlyout.IsOpen = true;
            }
        }

        private void SaveSettingsBtn_Click(object sender, RoutedEventArgs e) {
            Properties.General.Default.Save();
        }
    }
}
