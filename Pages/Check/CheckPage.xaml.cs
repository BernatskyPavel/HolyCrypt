using StegoLine.Utils;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace StegoLine.Pages.Check {
    /// <summary>
    /// Interaction logic for CheckPage.xaml
    /// </summary>
    public partial class CheckPage: Page {
        private string ConcealedFilePath = string.Empty;

        public CheckPage() {
            InitializeComponent();
        }

        private void OpenConcealedFileBtn_Click(object sender, RoutedEventArgs e) {
            ConcealedFileStatus.Foreground = Brushes.Red;
            ConcealedFileStatus.Content = "Файл не выбран";
            this.ConcealedFilePath = String.Empty;

            var FileDialog = new Microsoft.Win32.OpenFileDialog {
                FileName = "Document",
                DefaultExt = ".docx",
                Filter = "Word Documents|*.doc;*.docx"
            };

            bool? DlgResult = FileDialog.ShowDialog();
            if (DlgResult == true) {
                this.ConcealedFilePath = FileDialog.FileName;
                ConcealedFileStatus.Foreground = Brushes.Green;
                ConcealedFileStatus.Content = "Файл открыт";
            }
        }

        private void FullCheckBtn_Click(object sender, RoutedEventArgs e) {
            if (string.IsNullOrEmpty(this.ConcealedFilePath)) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CntNotSelectedMsg"].ToString()
                );
                return;
            }

            string? HashCode = RevealUtils.GetHashCode(this.ConcealedFilePath);

            if (string.IsNullOrEmpty(HashCode)) {
                (Application.Current.MainWindow as MainWindow)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    Application.Current.Resources["CntErrorMissingHashMsg"].ToString()
                );
                return;
            }

            int OldWidthStep = Properties.Reveal.Default.OutlineWidthStep;
            int OldAlpha = Properties.Reveal.Default.OutlineAlphaStep;

            bool IsMsgOk = false;

            foreach (double tick in this.CheckBitsPerSymbolSlider.Ticks) {
                //int BitsPerSymbol = (int)this.CheckBitsPerSymbolSlider.Value;
                RevealUtils.CalculateOutlineAlphaStep((int)tick);
                RevealUtils.CalculateOutlineWidthStep((int)tick);
                List<byte> RawBytes = RevealUtils.GetRawBytes(this.ConcealedFilePath, (int)tick);
                byte[] MsgBytes = RevealUtils.ParseRawBytes(RawBytes, (int)tick);
                IsMsgOk |= CheckUtils.CheckHashCode(MsgBytes, HashCode);
                if (IsMsgOk)
                    break;
            }

            Properties.Reveal.Default.OutlineWidthStep = OldWidthStep;
            Properties.Reveal.Default.OutlineAlphaStep = OldAlpha;
            Properties.Reveal.Default.Save();

            if (IsMsgOk) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["RevealPageCheckHeader"].ToString(),
                    Application.Current.Resources["CheckMsgSuccessMsg"].ToString()
                );
            }
            else {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["RevealPageCheckHeader"].ToString(),
                    Application.Current.Resources["CheckMsgFailureMsg"].ToString()
                );
            }
        }

        private void PartlyCheckBtn_Click(object sender, RoutedEventArgs e) {
            if (string.IsNullOrEmpty(this.ConcealedFilePath)) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CntNotSelectedMsg"].ToString()
                );
                return;
            }

            string? HashCode = RevealUtils.GetHashCode(this.ConcealedFilePath);

            if (string.IsNullOrEmpty(HashCode)) {
                (Application.Current.MainWindow as MainWindow)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    Application.Current.Resources["CntErrorMissingHashMsg"].ToString()
                );
                return;
            }

            int OldWidthStep = Properties.Reveal.Default.OutlineWidthStep;
            int OldAlpha = Properties.Reveal.Default.OutlineAlphaStep;

            int BitsPerSymbol = (int)this.CheckBitsPerSymbolSlider.Value;
            RevealUtils.CalculateOutlineAlphaStep(BitsPerSymbol);
            RevealUtils.CalculateOutlineWidthStep(BitsPerSymbol);

            List<byte> RawBytes = RevealUtils.GetRawBytes(this.ConcealedFilePath, BitsPerSymbol);
            byte[] MsgBytes = RevealUtils.ParseRawBytes(RawBytes, BitsPerSymbol);

            if (MsgBytes == null) {
                (Application.Current.MainWindow as MainWindow)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    Application.Current.Resources["CheckEmptyMsg"].ToString()
                );
                return;
            }

            Properties.Reveal.Default.OutlineWidthStep = OldWidthStep;
            Properties.Reveal.Default.OutlineAlphaStep = OldAlpha;
            Properties.Reveal.Default.Save();

            if (CheckUtils.CheckHashCode(MsgBytes, HashCode)) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["RevealPageCheckHeader"].ToString(),
                    Application.Current.Resources["CheckMsgSuccessMsg"].ToString()
                );
            }
            else {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["RevealPageCheckHeader"].ToString(),
                    Application.Current.Resources["CheckMsgFailureMsg"].ToString()
                );
            }
        }

        private void MenuButton_Click(object sender, RoutedEventArgs e) {
            if (this.NavigationService.CanGoBack) {
                this.NavigationService.GoBack();
            }
            else {
                _ = this.NavigationService.Navigate(new Home.HomePage());
            }

        }
    }
}
