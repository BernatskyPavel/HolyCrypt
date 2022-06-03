using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HolyCryptv3.Utils;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HolyCryptv3.Pages.Conceal {
    /// <summary>
    /// Interaction logic for ConcealMsgPage.xaml
    /// </summary>
    public partial class ConcealMsgPage: Page {

        private readonly Encoding MsgEncoding;
        private int MsgBitsCounter                  = 0;
        public bool IsNextBtnEnable { get; set; } = false;

        public ConcealMsgPage() {
            this.MsgEncoding = GeneralUtils.GetEncoding(1251);
            InitializeComponent();
        }

        public ConcealMsgPage(int EncodingPage) {
            this.MsgEncoding = GeneralUtils.GetEncoding(EncodingPage);
            InitializeComponent();
        }

        public ConcealMsgPage(Encoding MsgEncoding) {
            this.MsgEncoding = MsgEncoding;
            InitializeComponent();
        }

        private void OpenMsgBtn_Click(object sender, RoutedEventArgs e) {
            var FileDialog = new Microsoft.Win32.OpenFileDialog {
                FileName = "Document",
                DefaultExt = ".docx",
                Filter = "Documents|*.doc;*.docx;*.txt;*.pdf",
                CheckPathExists = true,
                CheckFileExists = true,
                Multiselect = false,
            };

            bool? DlgResult = FileDialog.ShowDialog();
            if (DlgResult == true) {
                string MsgFilePath = FileDialog.FileName;

                int LastDotPosition = MsgFilePath.LastIndexOf('.');
                if (LastDotPosition == -1) {
                    (Application.Current.MainWindow as MainWindow2)?.ShowErrorMessage(
                        Application.Current.Resources["ErrorBoxHeader"].ToString(),
                        Application.Current.Resources["OpenFileExtFailureMsg"].ToString()
                    );
                    return;
                }

                string FileExtension = MsgFilePath[(LastDotPosition+1)..];
                try {

                    switch (FileExtension) {
                        case "txt": {
                                this.MsgTextBox.Text = File.ReadAllText(MsgFilePath, Encoding.GetEncoding(1251));
                            }
                            break;
                        case "pdf": {
                                using PdfReader Reader = new(MsgFilePath);
                                PdfDocument PdfDoc = new(Reader);
                                string Text = string.Empty;
                                for (int Page = 1; Page <= PdfDoc.GetNumberOfPages(); Page++) {
                                    Text += PdfTextExtractor.GetTextFromPage(PdfDoc.GetPage(Page));
                                }
                                Reader.Close();
                                this.MsgTextBox.Text = Text;
                            }
                            break;
                        case "doc":
                        case "docx": {
                                using WordprocessingDocument Document = WordprocessingDocument.Open(MsgFilePath, false);
                                Body? DocumentBody = Document.MainDocumentPart?.Document.Body;
                                this.MsgTextBox.Text = DocumentBody?.InnerText;
                            }
                            break;
                        default:
                            return;
                    }
                }
                catch (Exception ex) {
                    (Application.Current.MainWindow as MainWindow2)?.ShowErrorMessage(
                        Application.Current.Resources["ErrorBoxHeader"].ToString(),
                        $"{Application.Current.Resources["FileOpenFailureMsg"]}\n{ex.Message}"
                    );

                    return;
                }

                this.MsgBitsTextBox.Clear();
                this.MsgBitsTextBox.Text = ConcealUtils.ToBinaryString(this.MsgEncoding, this.MsgTextBox.Text ?? "");

                this.MsgBitsCounter = MsgBitsTextBox.Text.Length;
                BitsCounterLabel.Text = this.MsgBitsCounter.ToString();
                this.IsNextBtnEnable = true;
            }
        }

        private void MsgTextBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e) {
            this.MsgBitsTextBox.Text = ConcealUtils.ToBinaryString(this.MsgEncoding, this.MsgTextBox.Text);
            this.MsgBitsCounter = MsgBitsTextBox.Text.Length;
            BitsCounterLabel.Text = this.MsgBitsCounter.ToString();
            //this.BitsCounterLabel.Content = this.MsgBitsTextBox.Text.Length;
            this.IsNextBtnEnable = this.MsgTextBox.Text.Length != 0;
        }

        private void ClearTextNextButton_Click(object sender, RoutedEventArgs e) {
            if (!this.IsNextBtnEnable) {
                (Application.Current.MainWindow as MainWindow2)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["MsgNotInputedMsg"].ToString()
                );
                return;
            }

            if (this.NavigationService.CanGoForward) {
                this.NavigationService.GoForward();
            }
            else {
                _ = this.NavigationService.Navigate(new ConcealCntPage(
                    new ConcealUtils.MsgInfo(this.MsgTextBox.Text, this.MsgEncoding, this.MsgBitsCounter)));
            }
        }

        private void MsgTextBox_TextChanged(object sender, TextChangedEventArgs e) {
            this.MsgBitsTextBox.Text = ConcealUtils.ToBinaryString(this.MsgEncoding, this.MsgTextBox.Text);
            this.MsgBitsCounter = MsgBitsTextBox.Text.Length;
            BitsCounterLabel.Text = this.MsgBitsCounter.ToString();
            //this.BitsCounterLabel.Content = this.MsgBitsTextBox.Text.Length;
            this.IsNextBtnEnable = this.MsgTextBox.Text.Length != 0;
        }
    }
}
