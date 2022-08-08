using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using StegoLine.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace StegoLine.Pages.Conceal {
    /// <summary>
    /// Interaction logic for ConcealCntPage.xaml
    /// </summary>
    public partial class ConcealCntPage: Page {
        private string ContainerFilePath                = string.Empty;
        private readonly string IgnoredSymbolsList      = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~–\\s\n]";
        private int ContainerSymbolsCounter             = 0;
        public bool IsChecked                           = false;
        private readonly ConcealUtils.MsgInfo MsgInfo;

        public ConcealCntPage(ConcealUtils.MsgInfo Info) {
            this.MsgInfo = Info;
            InitializeComponent();
            ConcealUtils.CalculateOutlineWidthStep((int)this.BitsPerSymbolSlider.Value);
            ConcealUtils.CalculateOutlineAlphaStep((int)this.BitsPerSymbolSlider.Value);
        }

        private void OpenContainerBtn_Click(object sender, RoutedEventArgs e) {

            var FileDialog = new Microsoft.Win32.OpenFileDialog {
                FileName = "Document",
                DefaultExt = ".docx",
                Filter = "Documents|*.doc;*.docx",
                CheckPathExists = true,
                CheckFileExists = true,
                Multiselect = false,
            };

            if (FileDialog.ShowDialog() == true) {
                try {
                    this.ContainerFilePath = FileDialog.FileName;
                    string Container = string.Empty;
                    using WordprocessingDocument Documnet = WordprocessingDocument.Open(this.ContainerFilePath, false);
                    Body? DocumentBody = Documnet.MainDocumentPart?.Document.Body;
                    this.ContainerTextBox.Text = DocumentBody?.InnerText;
                }
                catch (Exception ex) {
                    (Application.Current.MainWindow as MainWindow)?.ShowErrorMessage(
                        Application.Current.Resources["ErrorBoxHeader"].ToString(),
                        $"{Application.Current.Resources["FileOpenFailureMsg"]}\n{ex.Message}"
                    );
                    return;
                }
            }
        }

        private void ContainerCheckBtn_Click(object sender, RoutedEventArgs e) {
            this.CntCheckResultLabel.Visibility = Visibility.Hidden;
            this.ContainerCheckLabel.Visibility = Visibility.Hidden;

            if (string.IsNullOrEmpty(this.ContainerFilePath)) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CntNotSelectedMsg"].ToString()
                );
                return;
            }

            Regex RegExp = new(this.IgnoredSymbolsList);
            this.ContainerSymbolsCounter = RegExp.Replace(this.ContainerTextBox.Text, string.Empty).Length;
            {
                bool CheckResult = this.ContainerSymbolsCounter >= this.MsgInfo.BinaryMsgLength / (int)this.BitsPerSymbolSlider.Value;
                this.ContainerCheckLabel.Content = this.ContainerSymbolsCounter.ToString();
                this.CntCheckResultLabel.Content =
                    CheckResult ?
                    Application.Current.Resources["ConcealCntPageCheckS"].ToString() :
                    Application.Current.Resources["ConcealCntPageCheckF"].ToString();
                this.CntCheckResultLabel.Foreground = CheckResult ? Brushes.Green : Brushes.Red;
                //this.ConcealBtn.IsEnabled = CheckResult;
                this.CntCheckResultLabel.Visibility = Visibility.Visible;
                this.ContainerCheckLabel.Visibility = Visibility.Visible;
                IsChecked = CheckResult;
                if (!CheckResult) {
                    (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                       Application.Current.Resources["InfoBoxHeader"].ToString(),
                       Application.Current.Resources["CheckFailureInfoMsg"].ToString()
                    );
                }
            }
        }

        private void ConcealBtn_Click(object sender, RoutedEventArgs e) {

            if (string.IsNullOrEmpty(this.ContainerFilePath)) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CntNotSelectedMsg"].ToString()
                );
                return;
            }

            if (!this.IsChecked) {
                (Application.Current.MainWindow as MainWindow)?.ShowMyMessage(
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    Application.Current.Resources["CntNotCheckedMsg"].ToString()
                );
                return;
            }

            ConcealStatusLabel.Foreground = Brushes.Red;
            string MsgBits = this.MsgInfo.Text;
            bool IsConcealingActive = true;

            Queue<(int, int)>? MsgBitsQueue = ConcealUtils.ParseMsg(this.MsgInfo.Text, new ConcealUtils.ParseConfig (
                this.MsgInfo.MsgEncoding,
                (int)this.BitsPerSymbolSlider.Value,
                this.ContainerSymbolsCounter,
                this.MsgInfo.BinaryMsgLength
            ));

            String MsgHashCode = GeneralUtils.HashCode(this.MsgInfo.Text, this.MsgInfo.MsgEncoding);

            if (null == MsgBitsQueue || MsgBitsQueue.Count == 0) {
                (Application.Current.MainWindow as MainWindow)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    Application.Current.Resources["MsgParseErrorMsg"].ToString()
                );
                return;
            }

            try {

                using WordprocessingDocument Origin =
                         WordprocessingDocument.Open(this.ContainerFilePath, false);
                string[] PathParts = this.ContainerFilePath.Split('.');
                using WordprocessingDocument Document =
                     (WordprocessingDocument)Origin.Clone($"{PathParts[0]}_stg_{DateTime.Now:yyyy_MM_ddTHH_mm_ssZ}.{PathParts[1]}", true);

                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();


                try {
                    Document.MainDocumentPart?.Document.AddNamespaceDeclaration(Properties.General.Default.NamespacePrefix, Properties.General.Default.NamespaceUri);
                }
                catch (Exception) {}

                DocumentBody.SetAttribute(
                        new OpenXmlAttribute(
                            Properties.General.Default.NamespacePrefix,
                            Properties.General.Default.HashAttributeName,
                            Properties.General.Default.NamespaceUri,
                            MsgHashCode
                ));



                var ParagraphList = DocumentBody.ChildElements.Where(child => child is Paragraph);
                if (ParagraphList == null) {
                    return;
                }
                foreach (Paragraph Paragraph in ParagraphList) {
                    if (Paragraph == null) {
                        continue;
                    }

                    var ParagraphChildList = Paragraph.ChildElements.ToList();

                    Paragraph.RemoveAllChildren();

                    foreach (var ParagraphChild in ParagraphChildList) {
                        if (ParagraphChild is not Run || !IsConcealingActive) {
                            _ = Paragraph.AppendChild(ParagraphChild);
                            continue;
                        }

                        if (MsgBitsQueue.Count == 0) {
                            _ = Paragraph.AppendChild(ParagraphChild);
                            IsConcealingActive = false;
                            continue;
                        }

                        if (ParagraphChild is not Run SelectedRun) {
                            continue;
                        }

                        int RunTxtLen = SelectedRun.InnerText.Length,
                            RunClrTxtLen = SelectedRun.InnerText.Where(ch => !this.IgnoredSymbolsList.Contains(ch)).Count();

                        if (RunClrTxtLen == 0) {
                            _ = Paragraph.AppendChild(ParagraphChild);
                            continue;
                        }

                        Queue<(int, int)> BitsRunFit = new();
                        while (RunClrTxtLen > 0 && MsgBitsQueue.Count != 0) {
                            (int Repeats, int BitsValue) = MsgBitsQueue.Dequeue();
                            int PortionSize = Repeats;
                            if (RunClrTxtLen < Repeats) {
                                MsgBitsQueue = new Queue<(int, int)>(MsgBitsQueue.Prepend((Repeats - RunClrTxtLen, BitsValue)));
                                PortionSize = RunClrTxtLen;
                            }
                            RunClrTxtLen -= PortionSize;
                            BitsRunFit.Enqueue((PortionSize, BitsValue));
                        }
                        bool IsSymbolIgnored = false;
                        string RunTxt = SelectedRun.InnerText;
                        while (BitsRunFit.Count > 0) {
                            Run RunCopy = (Run)SelectedRun.CloneNode(true);
                            RunCopy.RemoveAllChildren<Text>();
                            (int Repeats, int BitsValue) Bits = BitsRunFit.Dequeue();
                            int SymbolsLen = 0;

                            char[]? Chars = null;
                            if (!IsSymbolIgnored) {
                                Chars = RunTxt.TakeWhile(ch => {
                                    SymbolsLen += 1;
                                    IsSymbolIgnored = this.IgnoredSymbolsList.Contains(ch);
                                    return SymbolsLen <= Bits.Repeats && !IsSymbolIgnored;
                                }).ToArray();
                            }
                            else {
                                Chars = RunTxt.TakeWhile(ch => {
                                    SymbolsLen += 1;
                                    return this.IgnoredSymbolsList.Contains(ch);
                                }).ToArray();
                                RunTxt = RunTxt.Remove(0, SymbolsLen - 1);
                                _ = RunCopy.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(Chars) });

                                _ = Paragraph.AppendChild(RunCopy);
                                BitsRunFit = new Queue<(int, int)>(BitsRunFit.Prepend(Bits));
                                IsSymbolIgnored = false;
                                continue;
                            }

                            if (IsSymbolIgnored && SymbolsLen == 1) {
                                BitsRunFit = new Queue<(int, int)>(BitsRunFit.Prepend(Bits));
                                continue;
                            }
                            SymbolsLen += Chars.Length == RunTxt.Length ? 1 : 0;
                            RunTxt = RunTxt.Remove(0, SymbolsLen - 1);
                            _ = RunCopy.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(Chars) });
                            TextOutlineEffect? OutlineEffects = ConcealUtils.GetOutlineObj(Bits.BitsValue, (int)this.BitsPerSymbolSlider.Value);

                            RunCopy.RunProperties ??= new RunProperties();
                            if (OutlineEffects != null)
                                RunCopy.RunProperties.TextOutlineEffect = OutlineEffects;
                            _ = Paragraph.AppendChild(RunCopy);

                            if (Chars.Length != Bits.Repeats) {
                                BitsRunFit = new Queue<(int, int)>(BitsRunFit.Prepend((Bits.Repeats - Chars.Length, Bits.BitsValue)));
                            }
                        }

                        if (BitsRunFit.Count == 0 && RunTxt.Length != 0 && RunClrTxtLen == 0) {
                            Run RunCopy = (Run)SelectedRun.CloneNode(true);
                            RunCopy.RemoveAllChildren<Text>();
                            _ = RunCopy.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(RunTxt) });
                            _ = Paragraph.AppendChild(RunCopy);
                        }
                        else if (MsgBitsQueue.Count == 0 && RunClrTxtLen > 0) {
                            IsConcealingActive = false;
                            Run RunCopy = (Run)SelectedRun.CloneNode(true);
                            RunCopy.RemoveAllChildren<Text>();
                            _ = RunCopy.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(RunTxt) });
                            _ = Paragraph.AppendChild(RunCopy);
                        }

                    }
                }
                Document.Save();
                ConcealStatusLabel.Foreground = Brushes.Green;
            }
            catch (Exception ex) {
                (Application.Current.MainWindow as MainWindow)?.ShowErrorMessage(
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    $"{Application.Current.Resources["ConcealFailureMsg"]}\n{ex.Message}"
                );
                return;
            }
        }

        private void BitsPerSymbolSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) {
            ConcealUtils.CalculateOutlineAlphaStep((int)e.NewValue);
            ConcealUtils.CalculateOutlineWidthStep((int)e.NewValue);
        }

        private void ContainerPrevButton_Click(object sender, RoutedEventArgs e) {
            if (this.NavigationService.CanGoBack)
                this.NavigationService.GoBack();
        }

        private void MenuButton_Click(object sender, RoutedEventArgs e) {
            _ = this.NavigationService.Navigate(new StegoLine.Pages.Home.HomePage());
        }
    }
}
