using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace HolyCryptv3 {
    partial class MainWindow: Window {

        private string IgnoredSymbolsList = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~–\\s]";

        #region Message
        private void OpenMsgBtn_Click(object sender, RoutedEventArgs e) {
            this.MsgErrorHeader.Visibility = Visibility.Hidden;
            this.MsgErrorLabel.Visibility = Visibility.Hidden;
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
                this.MsgFilePath = FileDialog.FileName;

                int LastDotPosition = this.MsgFilePath.LastIndexOf('.');
                if (LastDotPosition == -1) {
                    return;
                }

                string FileExtension = this.MsgFilePath.Substring(LastDotPosition+1);

                switch (FileExtension) {
                    case "txt":
                        try {
                            this.MsgTextBox.Text = File.ReadAllText(this.MsgFilePath, Encoding.GetEncoding(1251));
                        }
                        catch (Exception ex) {
                            this.MsgErrorHeader.Visibility = Visibility.Visible;
                            this.MsgErrorLabel.Visibility = Visibility.Visible;
                            this.MsgErrorLabel.Content = ex.Message;
                            return;
                        }
                        break;
                    case "pdf":
                        try {
                            using (PdfReader Reader = new PdfReader(this.MsgFilePath)) {
                                PdfDocument PdfDoc = new PdfDocument(Reader);
                                string Text = string.Empty;
                                for (int Page = 1; Page <= PdfDoc.GetNumberOfPages(); Page++) {
                                    Text += PdfTextExtractor.GetTextFromPage(PdfDoc.GetPage(Page));
                                }
                                Reader.Close();
                                this.MsgTextBox.Text = Text;
                            }
                        }
                        catch (Exception ex) {
                            this.MsgErrorHeader.Visibility = Visibility.Visible;
                            this.MsgErrorLabel.Visibility = Visibility.Visible;
                            this.MsgErrorLabel.Content = ex.Message;
                            return;
                        }
                        break;
                    case "doc":
                    case "docx":
                        try {
                            using (WordprocessingDocument Document = WordprocessingDocument.Open(this.MsgFilePath, false)) {
                                Body? DocumentBody = Document.MainDocumentPart?.Document.Body;
                                this.MsgTextBox.Text = DocumentBody?.InnerText;
                            }
                        }
                        catch (Exception ex) {
                            this.MsgErrorHeader.Visibility = Visibility.Visible;
                            this.MsgErrorLabel.Visibility = Visibility.Visible;
                            this.MsgErrorLabel.Content = ex.Message;
                            return;
                        }
                        break;
                    default:
                        return;
                }

                this.MsgBitsTextBox.Clear();
                this.MsgBitsTextBox.Text = ToBinaryString(Encoding.GetEncoding(1251), this.MsgTextBox.Text ?? "");

                this.MsgBitsCounter = MsgBitsTextBox.Text.Length;
                BitsCounterLabel.Content = this.MsgBitsCounter;
                ContainerTestBtn.IsEnabled = false;
                ContainerCheckLabel.Visibility = Visibility.Hidden;
                ClearTextNextButton.IsEnabled = true;
            }
        }
        private void MsgTextBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e) {
            this.MsgBitsTextBox.Text = ToBinaryString(Encoding.GetEncoding(1251), this.MsgTextBox.Text);
            this.BitsCounterLabel.Content = this.MsgBitsTextBox.Text.Length;
        }
        #endregion

        #region Container
        private void OpenContainerBtn_Click(object sender, RoutedEventArgs e) {
            this.ContainerErrorHeader.Visibility = Visibility.Hidden;
            this.ContainerErrorLabel.Visibility = Visibility.Hidden;
            this.WordCountButton.IsEnabled = false;
            this.ContainerSymbolsCounterLabel.Content = string.Empty;
            this.ContainerCheckLabel.Content = string.Empty;
            this.ContainerTestBtn.IsEnabled = false;
            this.ContainerTest.IsEnabled = false;
            var FileDialog = new Microsoft.Win32.OpenFileDialog {
                FileName = "Document",
                DefaultExt = ".docx",
                Filter = "Documents|*.doc;*.docx",
                CheckPathExists = true,
                CheckFileExists = true,
                Multiselect = false,
            };
            bool? DlgResult = FileDialog.ShowDialog();
            if (DlgResult == true) {
                this.ContainerFilePath = FileDialog.FileName;
                try {
                    string Container = string.Empty;
                    using (WordprocessingDocument Documnet = WordprocessingDocument.Open(this.ContainerFilePath, false)) {
                        Body? DocumentBody = Documnet.MainDocumentPart?.Document.Body;
                        this.ContainerTextBox.Text = DocumentBody?.InnerText;
                    }
                }
                catch (Exception ex) {
                    this.ContainerErrorHeader.Visibility = Visibility.Visible;
                    this.ContainerErrorLabel.Visibility = Visibility.Visible;
                    this.ContainerErrorLabel.Content = ex.Message;
                    return;
                }

                WordCountButton.IsEnabled = true;
                ContainerTestBtn.IsEnabled = false;
                ContainerCheckLabel.Visibility = Visibility.Hidden;
                //ContainerNextButton.IsEnabled = false;
            }
        }

        private void CountSymbolsBtn_Click(object sender, RoutedEventArgs e) {
            //string IgnoredSymbolsList = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~–\\s]";
            Regex RegExp = new Regex(this.IgnoredSymbolsList);
            //string text = ContainerTextBox.Text;
            this.ContainerTextBox.Text = RegExp.Replace(this.ContainerTextBox.Text, string.Empty);
            //int FilteredTextLen = this.ContainerTextBox.Text.Length;
            this.ContainerSymbolsCounter = this.ContainerTextBox.Text.Length;
            this.ContainerSymbolsCounterLabel.Content = this.ContainerSymbolsCounter;
            this.ContainerTestBtn.IsEnabled = true;
        }

        private void ContainerTestBtn_Click(object sender, RoutedEventArgs e) {
            ContainerCheckLabel.Content = "Не подходит!";
            ContainerCheckLabel.Foreground = Brushes.Red;
            ContainerTest.IsEnabled = false;
            ContainerCheckLabel.Visibility = Visibility.Visible;
            if (this.ContainerSymbolsCounter >= this.MsgBitsCounter / 2) {
                ContainerCheckLabel.Content = "Подходит!";
                ContainerCheckLabel.Foreground = Brushes.Green;
                ContainerTest.IsEnabled = true;
            }
        }

        private void ConcealBtn_Click(object sender, RoutedEventArgs e) {
            ConcealStatusLabel.Foreground = Brushes.Red;
            string MsgBits = this.MsgBitsTextBox.Text;
            bool IsConcealingActive = true;
            //string IgnoredSymbolsList = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~– ]";

            Queue<(int, string)> MsgBitsQueue = this.parseBitString(MsgBits, 2);

            using (WordprocessingDocument Document =
                         WordprocessingDocument.Open(this.ContainerFilePath, true)) {

                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();
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
                        if (!(ParagraphChild is Run) || !IsConcealingActive) {
                            Paragraph.AppendChild(ParagraphChild);
                            continue;
                        }

                        if (MsgBitsQueue.Count == 0) {
                            Paragraph.AppendChild(ParagraphChild);
                            IsConcealingActive = false;
                            continue;
                        }

                        Run? SelectedRun = ParagraphChild as Run;
                        if (SelectedRun == null) {
                            continue;
                        }

                        int RunTxtLen = SelectedRun.InnerText.Length,
                            RunClrTxtLen = SelectedRun.InnerText.Where(ch => !this.IgnoredSymbolsList.Contains(ch)).Count();
                        
                        if (RunClrTxtLen == 0) {
                            Paragraph.AppendChild(ParagraphChild);
                            continue;
                        }

                        Queue<(int, string)> BitsRunFit = new Queue<(int, string)>();
                        while (RunClrTxtLen > 0 && MsgBitsQueue.Count != 0) {
                            (int Repeats, string Pattern) BitsPortion = MsgBitsQueue.Dequeue();
                            int PortionSize = BitsPortion.Repeats;
                            if (RunClrTxtLen < BitsPortion.Repeats) {
                                MsgBitsQueue = new Queue<(int, string)>(MsgBitsQueue.Prepend((BitsPortion.Repeats - RunClrTxtLen, BitsPortion.Pattern)));
                                PortionSize = RunClrTxtLen;
                            }
                            RunClrTxtLen -= PortionSize;
                            BitsRunFit.Enqueue((PortionSize, BitsPortion.Pattern));
                        }
                        bool IsSymbolIgnored = false;
                        string RunTxt = SelectedRun.InnerText;
                        while (BitsRunFit.Count > 0) {
                            Run RunCopy = (Run)SelectedRun.CloneNode(true);
                            RunCopy.RemoveAllChildren<Text>();
                            (int Repeats, string Pattern) Bits = BitsRunFit.Dequeue();
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
                                RunCopy.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(Chars) });

                                Paragraph.AppendChild(RunCopy);
                                BitsRunFit = new Queue<(int, string)>(BitsRunFit.Prepend(Bits));
                                IsSymbolIgnored = false;
                                continue;
                            }

                            if (IsSymbolIgnored && SymbolsLen == 1) {
                                BitsRunFit = new Queue<(int, string)>(BitsRunFit.Prepend(Bits));
                                continue;
                            }
                            SymbolsLen += Chars.Length == RunTxt.Length ? 1 : 0;
                            RunTxt = RunTxt.Remove(0, SymbolsLen - 1);
                            RunCopy.AddChild(new Text(new string(Chars)));
                            TextOutlineEffect? OutlineEffects = getOutline(Bits.Pattern);

                            RunCopy.RunProperties = RunCopy.RunProperties ?? new RunProperties();
                            RunCopy.RunProperties.TextOutlineEffect = OutlineEffects;
                            Paragraph.AppendChild(RunCopy);

                            if (Chars.Length != Bits.Repeats) {
                                BitsRunFit = new Queue<(int, string)>(BitsRunFit.Prepend((Bits.Repeats - Chars.Length, Bits.Pattern)));
                            }
                        }

                        if (BitsRunFit.Count == 0 && RunTxt.Length != 0 && RunClrTxtLen == 0) {
                            Run RunCopy = (Run)SelectedRun.CloneNode(true);
                            RunCopy.RemoveAllChildren<Text>();
                            RunCopy.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(RunTxt) });
                            Paragraph.AppendChild(RunCopy);
                        }
                        else if (MsgBitsQueue.Count == 0 && RunClrTxtLen > 0) {
                            IsConcealingActive = false;
                            Run RunCopy = (Run)SelectedRun.CloneNode(true);
                            RunCopy.RemoveAllChildren<Text>();
                            RunCopy.AddChild(new Text(RunTxt));
                            Paragraph.AppendChild(RunCopy);
                        }

                    }
                }
                Document.Save();
                ConcealStatusLabel.Foreground = Brushes.Green;
            }
        }
        #endregion
    }
}
