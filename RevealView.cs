using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace HolyCryptv3 {
    partial class MainWindow: Window {

        private string ConcealedFilePath = string.Empty;

        private void OpenConcealedFileBtn_Click(object sender, RoutedEventArgs e) {
            ConcealedFileStatus.Foreground = Brushes.Red;
            ConcealedFileStatus.Content = "Файл не выбран";
            DecodeButton.IsEnabled = false;
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
                DecodeButton.IsEnabled = true;
            }
        }
                
        private void RevealMsgBtn_Old_Click(object sender, RoutedEventArgs e) {
            this.RevealStatusValue.Content = "Неверный или поврежденный файл!";
            this.RevealStatusValue.Foreground = Brushes.Red;

            if (string.IsNullOrEmpty(this.ConcealedFilePath))
                return;

            string MsgBits = string.Empty;
            bool IsMsgEnded = false;
            int ErrorCounter = 0;

            using (WordprocessingDocument Document =
             WordprocessingDocument.Open(this.ConcealedFilePath, false)) {

                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();

                foreach (OpenXmlElement BodyElement in DocumentBody.ChildElements) {
                    if (IsMsgEnded) {
                        break;
                    }
                    if (BodyElement is Paragraph) {
                        foreach (OpenXmlElement ParagraphChild in BodyElement.ChildElements) {
                            if (ParagraphChild is Run) {
                                Run ChildRun = (Run)ParagraphChild;
                                string BitPattern = string.Empty;

                                TextOutlineEffect? OutlineEffect = ChildRun.RunProperties?.TextOutlineEffect;
                                int? Width = OutlineEffect?.LineWidth?.Value;
                                int? Alpha = OutlineEffect?
                                    .GetFirstChild<SolidColorFillProperties>()?
                                    .GetFirstChild<SchemeColor>()?
                                    .GetFirstChild<Alpha>()?.Val?.Value;

                                switch ((Width, Alpha)) {
                                    case (635, 96000):
                                        BitPattern = "00";
                                        ErrorCounter = 0;
                                        break;
                                    case (1270, 96000):
                                        BitPattern = "10";
                                        ErrorCounter = 0;
                                        break;
                                    case (635, 100000):
                                        BitPattern = "01";
                                        ErrorCounter = 0;
                                        break;
                                    case (1270, 100000):
                                        BitPattern = "11";
                                        ErrorCounter = 0;
                                        break;
                                    case (635, 99000):
                                        IsMsgEnded = true;
                                        break;
                                    default:
                                        ErrorCounter += 1;
                                        BitPattern = string.Empty;
                                        break;
                                }

                                if (ErrorCounter == 5) {
                                    break;
                                }

                                if (!string.IsNullOrEmpty(BitPattern)) {
                                    for (int j = 0; j < ChildRun.InnerText.Length; j++) {
                                        MsgBits += BitPattern;
                                    }
                                }
                                if (IsMsgEnded) {
                                    break;
                                }
                            }
                        }

                    }
                }

            }

            if (ErrorCounter != 0 && !IsMsgEnded) {
                this.RevealStatusValue.Content = "Неверный или поврежденный файл!";
                this.RevealStatusValue.Foreground = Brushes.Red;
                return;
            }

            byte[] ByteArray = new byte[MsgBits.Length / 8];

            BitArray MsgBitsArray = new BitArray(MsgBits.Length);

            int i = 0;

            foreach (char Bit in MsgBits.Reverse()) {
                if (Bit == '1') {
                    MsgBitsArray.Set(i, true);
                }
                i++;
            }

            MsgBitsArray.CopyTo(ByteArray, 0);
            //Encoding Win1251Encoding = Encoding.GetEncoding(1251);
            RevealedText.Text = string.Join("", this.Encoding.GetString(ByteArray).Reverse());
            this.RevealStatusValue.Content = "Успешно извлечено!";
            this.RevealStatusValue.Foreground = Brushes.Green;
            this.SaveRevealedMsgBtn.IsEnabled = true;
        }

        private void RevealMsgBtn_Click(object sender, RoutedEventArgs e) {
            this.RevealStatusValue.Content = "Неверный или поврежденный файл!";
            this.RevealStatusValue.Foreground = Brushes.Red;

            if (string.IsNullOrEmpty(this.ConcealedFilePath))
                return;

            string MsgBits = string.Empty;
            bool IsMsgEnded = false;
            int ErrorCounter = 0;

            int Size = (int)this.DecodeBitsPerSymbolSlider.Value;
            int HalfSize = Size / 2;
            int mask = (int)(Math.Pow(2, HalfSize) - 1);
            int Steps = sizeof(byte) * 8 / Size;
            int Index = 0;
            List<byte> _bytes = new List<byte>();
            string? HashCode = null;


            using (WordprocessingDocument Document =
             WordprocessingDocument.Open(this.ConcealedFilePath, false)) {

                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();

                HashCode = DocumentBody.GetAttribute(this.HashAttributeName, this.NamespaceUri).Value;

                foreach (OpenXmlElement BodyElement in DocumentBody.ChildElements) {
                    if (IsMsgEnded) {
                        break;
                    }
                    if (ErrorCounter == 5) {
                        break;
                    }
                    if (BodyElement is Paragraph) {
                        foreach (OpenXmlElement ParagraphChild in BodyElement.ChildElements) {
                            if (ParagraphChild is Run) {

                                Run ChildRun = (Run)ParagraphChild;
                                string BitPattern = string.Empty;

                                TextOutlineEffect? OutlineEffect = ChildRun.RunProperties?.TextOutlineEffect;
                                if (null == OutlineEffect) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                try {
                                    OutlineEffect.GetAttribute(this.RunAttributeName, this.NamespaceUri);
                                }
                                catch(KeyNotFoundException) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                //if (OutlineEffect.GetAttribute(this.RunAttributeName, this.NamespaceUri) != null) {
                                    
                                //}

                                int? Width = OutlineEffect?.LineWidth?.Value;
                                int? Alpha = OutlineEffect?
                                    .GetFirstChild<SolidColorFillProperties>()?
                                    .GetFirstChild<SchemeColor>()?
                                    .GetFirstChild<Alpha>()?.Val?.Value;

                                if (null == Alpha || null == Width) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                int LowerBits = this.GetBitsFromOutlineWidth(Width.Value);
                                int HigherBits = this.GetBitsFromOutlineAlpha(Alpha.Value);

                                int Bits = (HigherBits << HalfSize) | LowerBits;

                                _bytes.Add((byte)Bits);
                                //ISteps += 1;
                                ErrorCounter = 0;


                                //switch ((Width, Alpha)) {
                                //    case (635, 96000):
                                //        BitPattern = "00";
                                //        ErrorCounter = 0;
                                //        break;
                                //    case (1270, 96000):
                                //        BitPattern = "10";
                                //        ErrorCounter = 0;
                                //        break;
                                //    case (635, 100000):
                                //        BitPattern = "01";
                                //        ErrorCounter = 0;
                                //        break;
                                //    case (1270, 100000):
                                //        BitPattern = "11";
                                //        ErrorCounter = 0;
                                //        break;
                                //    case (635, 99000):
                                //        IsMsgEnded = true;
                                //        break;
                                //    default:
                                //        ErrorCounter += 1;
                                //        BitPattern = string.Empty;
                                //        break;
                                //}

                                

                                //if (!string.IsNullOrEmpty(BitPattern)) {
                                //    for (int j = 0; j < ChildRun.InnerText.Length; j++) {
                                //        MsgBits += BitPattern;
                                //    }
                                //}
                                //if (IsMsgEnded) {
                                //    break;
                                //}
                            }
                        }

                    }
                }
            }

            int ByteLen = _bytes.Count * Size / 8;
            byte[] bytes = new byte[ByteLen];

            //_bytes.Reverse();
            
            foreach (byte[] parts in _bytes.Chunk(Steps)) {
                byte NewByte = 0;
                for (int _i = 0; _i < Math.Min(Steps, parts.Length); _i++) {
                    NewByte = (byte)(NewByte | (parts[_i] << (Size * _i)));
                }
                if (Index < ByteLen) {
                    bytes[Index++] = NewByte;
                }
            }

            //if (ErrorCounter != 0 && !IsMsgEnded) {
            //    this.RevealStatusValue.Content = "Неверный или поврежденный файл!";
            //    this.RevealStatusValue.Foreground = Brushes.Red;
            //    return;
            //}

            //byte[] ByteArray = new byte[MsgBits.Length / 8];

            //BitArray MsgBitsArray = new BitArray(MsgBits.Length);

            //int i = 0;

            //foreach (char Bit in MsgBits.Reverse()) {
            //    if (Bit == '1') {
            //        MsgBitsArray.Set(i, true);
            //    }
            //    i++;
            //}

            //MsgBitsArray.CopyTo(ByteArray, 0);
            //Encoding Win1251Encoding = Encoding.GetEncoding(1251);

            //if (null != HashCode && this.CheckHashCode(bytes.Reverse().ToArray(), HashCode)) {
            //    MessageBox.Show("Hash GOOD!");
            //}
            

            RevealedText.Text = string.Join("", this.Encoding.GetString(bytes.Reverse().ToArray()));
            this.RevealStatusValue.Content = "Успешно извлечено!";
            this.RevealStatusValue.Foreground = Brushes.Green;
            this.SaveRevealedMsgBtn.IsEnabled = true;
        }

        private void MenuButton_Click(object sender, RoutedEventArgs e) {
            this.TabMenu.SelectedIndex = 0;
        }

        private void SaveRevealedMsgBtn_Click(object sender, RoutedEventArgs e) {
            Stream FileStream;
            Microsoft.Win32.SaveFileDialog FileDialog = new Microsoft.Win32.SaveFileDialog{
                Filter = "Text file (*.txt)|*.txt|Word document (*.doc(x))|*.docx;*.doc|PDF file (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,
            };

            bool? DlgResult = FileDialog.ShowDialog();

            if (DlgResult == true) {

                int LastDotPosition = FileDialog.SafeFileName.LastIndexOf('.');
                if (LastDotPosition == -1) {
                    this.MsgErrorHeader.Visibility = Visibility.Visible;
                    this.MsgErrorLabel.Visibility = Visibility.Visible;
                    this.MsgErrorLabel.Content = "Некорректное имя файла!";
                    return;
                }

                string FileExtension = FileDialog.SafeFileName.Substring(LastDotPosition+1);

                if ((FileStream = FileDialog.OpenFile()) != null) {

                    switch (FileExtension) {
                        case "txt":
                            try {
                                FileStream.Write(Encoding.GetEncoding(1251).GetBytes(this.RevealedText.Text));
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
                                PdfWriter Writer = new PdfWriter(FileStream);
                                PdfDocument PdfDoc = new PdfDocument(Writer);
                                iText.Layout.Document Document = new iText.Layout.Document(PdfDoc);
                                iText.Layout.Element.Text PdfText = new iText.Layout.Element.Text(this.RevealedText.Text);
                                PdfText.SetFont(PdfFontFactory.CreateFont("C:/Windows/Fonts/verdana.ttf", "CP1251"));
                                Document.Add(new iText.Layout.Element.Paragraph(PdfText));
                                Document.Close();
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
                                using (WordprocessingDocument Document = WordprocessingDocument.Create(FileStream, WordprocessingDocumentType.Document)) {
                                    MainDocumentPart MainPart = Document.AddMainDocumentPart();
                                    MainPart.Document = new Document();
                                    Body DocumentBody = MainPart.Document.AppendChild(new Body());
                                    Paragraph NewParagraph = DocumentBody.AppendChild(new Paragraph());
                                    Run NewRun = NewParagraph.AppendChild(new Run());
                                    NewRun.AppendChild(new Text(this.RevealedText.Text));
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

                    FileStream.Close();
                }            
            }
        }
    }
}
