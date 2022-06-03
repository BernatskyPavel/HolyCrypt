using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HolyCryptv3.Utils;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace HolyCryptv3.Pages.Reveal {
    /// <summary>
    /// Interaction logic for RevealPage.xaml
    /// </summary>
    public partial class RevealPage: Page {
        private readonly Encoding MsgEncoding;
        private string ConcealedFilePath = string.Empty;
        public RevealPage() {
            this.MsgEncoding = GeneralUtils.GetEncoding(1251);
            InitializeComponent();
            RevealUtils.CalculateOutlineAlphaStep((int)this.DecodeBitsPerSymbolSlider.Value);
            RevealUtils.CalculateOutlineWidthStep((int)this.DecodeBitsPerSymbolSlider.Value);
        }

        public RevealPage(int EncodingPage) {
            this.MsgEncoding = GeneralUtils.GetEncoding(EncodingPage);
            InitializeComponent();
        }

        public RevealPage(Encoding MsgEncoding) {
            this.MsgEncoding = MsgEncoding;
            InitializeComponent();
        }


        private void DecodeBitsPerSymbolSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) {
            RevealUtils.CalculateOutlineAlphaStep((int)e.NewValue);
            RevealUtils.CalculateOutlineWidthStep((int)e.NewValue);
        }

        private void RevealMsgBtn_Click(object sender, RoutedEventArgs e) {
            this.DecodeStatusKey.Foreground = Brushes.Red;

            if (string.IsNullOrEmpty(this.ConcealedFilePath)) {
                (Application.Current.MainWindow as MainWindow2)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CntNotSelectedMsg"].ToString()
                );
                return;
            }

            int Size = (int)this.DecodeBitsPerSymbolSlider.Value;

            byte[] MsgBytes;
            try {
                List<byte> RawBytes = RevealUtils.GetRawBytes(this.ConcealedFilePath, Size);
                MsgBytes = RevealUtils.ParseRawBytes(RawBytes, Size);
            }
            catch (Exception ex) {
                (Application.Current.MainWindow as MainWindow2)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    $"{Application.Current.Resources["CntErrorMsg"]}\n{ex.Message}"
                );
                return;
            }

            string? HashCode = RevealUtils.GetHashCode(this.ConcealedFilePath);
            this.RevealStatusValue.Visibility = Visibility.Hidden;
            if (null != HashCode && CheckUtils.CheckHashCode(MsgBytes, HashCode)) {
                (Application.Current.MainWindow as MainWindow2)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CheckMsgSuccessMsg"].ToString()
                );
            }
            else {
                (Application.Current.MainWindow as MainWindow2)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CheckMsgFailureMsg"].ToString()
                );
            }

            RevealedText.Text = string.Join("", this.MsgEncoding.GetString(MsgBytes));
            this.DecodeStatusKey.Foreground = Brushes.Green;
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
            try {
                if (DlgResult == true) {
                    this.ConcealedFilePath = FileDialog.FileName;
                    ConcealedFileStatus.Foreground = Brushes.Green;
                    ConcealedFileStatus.Content = "Файл открыт";
                }
            }
            catch (Exception ex) {
                (Application.Current.MainWindow as MainWindow2)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    $"{Application.Current.Resources["FileOpenFailureMsg"]}\n{ex.Message}"
                );
                return;
            }

        }

        private void SaveRevealedMsgBtn_Click(object sender, RoutedEventArgs e) {

            if (string.IsNullOrEmpty(RevealedText.Text)) {
                (Application.Current.MainWindow as MainWindow2)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["RevealMsgEmpty"].ToString()
                );
                return;
            }

            Stream FileStream;
            Microsoft.Win32.SaveFileDialog FileDialog = new() {
                Filter = "Text file (*.txt)|*.txt|Word document (*.doc(x))|*.docx;*.doc|PDF file (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,
            };

            bool? DlgResult = FileDialog.ShowDialog();

            if (DlgResult == true) {

                int LastDotPosition = FileDialog.SafeFileName.LastIndexOf('.');
                if (LastDotPosition == -1) {
                    (Application.Current.MainWindow as MainWindow2)?.ShowErrorMessage(
                        Application.Current.Resources["ErrorBoxHeader"].ToString(),
                        Application.Current.Resources["OpenFileExtFailureMsg"].ToString()
                    );
                    return;
                }

                string FileExtension = FileDialog.SafeFileName[(LastDotPosition+1)..];

                if ((FileStream = FileDialog.OpenFile()) != null) {
                    try {
                        switch (FileExtension) {
                            case "txt": {
                                    FileStream.Write(Encoding.GetEncoding(1251).GetBytes(this.RevealedText.Text));
                                }
                                break;
                            case "pdf": {
                                    PdfWriter Writer = new(FileStream);
                                    PdfDocument PdfDoc = new(Writer);
                                    iText.Layout.Document Document = new(PdfDoc);
                                    iText.Layout.Element.Text PdfText = new(this.RevealedText.Text);
                                    _ = PdfText.SetFont(PdfFontFactory.CreateFont($"{Environment.GetFolderPath(Environment.SpecialFolder.Fonts)}/verdana.ttf", "CP1251"));
                                    _ = Document.Add(new iText.Layout.Element.Paragraph(PdfText));
                                    Document.Close();
                                }
                                break;
                            case "doc":
                            case "docx": {
                                    using WordprocessingDocument Document = WordprocessingDocument.Create(FileStream, WordprocessingDocumentType.Document);
                                    MainDocumentPart MainPart = Document.AddMainDocumentPart();
                                    MainPart.Document = new Document();
                                    Body DocumentBody = MainPart.Document.AppendChild(new Body());
                                    Paragraph NewParagraph = DocumentBody.AppendChild(new Paragraph());
                                    Run NewRun = NewParagraph.AppendChild(new Run());
                                    _ = NewRun.AppendChild(new Text(this.RevealedText.Text));
                                }
                                break;
                            default:
                                return;
                        }
                        FileStream.Close();
                    }
                    catch (Exception ex) {
                        (Application.Current.MainWindow as MainWindow2)?.ShowErrorMessage(
                            Application.Current.Resources["ErrorBoxHeader"].ToString(),
                            $"{Application.Current.Resources["FileOpenFailureMsg"]}\n{ex.Message}"
                        );
                        return;
                    }

                }
            }
        }

        private void MenuButton_Click(object sender, RoutedEventArgs e) {
            if (this.NavigationService.CanGoBack) {
                this.NavigationService.GoBack();
            }
            else {
                _ = this.NavigationService.Navigate(new HolyCryptv3.Pages.Home.HomePage());
            }
        }
    }
}
