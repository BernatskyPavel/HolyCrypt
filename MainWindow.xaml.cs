using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2010.Word;
using iText;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

namespace HolyCryptv3 {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow: Window {
        private string message_file_path        = string.Empty;
        private string container_file_path      = string.Empty;
        private string encoded_file_path        = string.Empty;
        private int lettersCount                = 0;
        private int bitsCount                   = 0;

        private Dictionary<string, (int, int)> bits_params = new Dictionary<string, (int, int)> {
            {"0000", (96000,635)},  {"1000", (98000,915)},
            {"0001", (96250,670)},  {"1001", (98250,950)},
            {"0010", (96500,705)},  {"1010", (98500,985)},
            {"0011", (96750,740)},  {"1011", (98750,1020)},
            {"0100", (97000,775)},  {"1100", (99000,1055)},
            {"0101", (97250,810)},  {"1101", (99250,1090)},
            {"0110", (97500,845)},  {"1110", (99500,1125)},
            {"0111", (97750,880)},  {"1111", (99750,1160)},
            {  "XX", (100000,635)}
        };

        private Dictionary<string, (int, int)> outlines = new Dictionary<string, (int, int)> {
            {"00", (96000,635)},
            {"01", (100000,635)},
            {"10", (96000,1270)},
            {"11", (100000,1270)},
            {"XX", (99000,635)},
        };

        private TextOutlineEffect? outline_basis = null;
        private TextOutlineEffect? finish_outline = null;

        static string ToBinaryString(Encoding encoding, string text) {
            return string.Join("", encoding.GetBytes(text).Select(n => Convert.ToString(n, 2).PadLeft(8, '0')));
        }

        private TextOutlineEffect? getOutline(string bits) {
            string key = bits;

            if (!this.outlines.ContainsKey(bits)) {
                key = "XX";
            }

            (int, int) values = this.outlines[key];

            TextOutlineEffect? effect = (TextOutlineEffect?)this.outline_basis?.CloneNode(true);

            if (effect == null) {
                return effect;
            }

            SchemeColor? scheme = new SchemeColor {
                Val = SchemeColorValues.ExtraSchemeColor1,
            };

            scheme.AddChild(new Alpha {
                Val = values.Item1
            });

            effect.AddChild(new SolidColorFillProperties(scheme));
            effect.AddChild(new PresetLineDashProperties {
                Val = PresetLineDashValues.Solid
            });
            effect.AddChild(new BevelEmpty());
            effect.LineWidth = values.Item2;
            return effect;
        }

        public MainWindow() {

            this.outline_basis = new TextOutlineEffect {
                CapType = LineCapValues.Round,
                Alignment = PenAlignmentValues.Center,
                Compound = CompoundLineValues.Simple,
            };

            this.finish_outline = this.getOutline("XX");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }

        private Queue<(int, string)> parseBitString(string bit_string, int size) {
            var list = new Queue<(int, string)>();

            List<string> parts = bit_string.Chunk(size).Select(arr => new string(arr)).ToList();
            string prev = string.Empty, temp = string.Empty;
            int count = 0;

            for (int i = 0; i < parts.Count; i++) {
                temp = parts[i];
                if (prev == string.Empty) {
                    count = 1;
                    prev = temp;
                }
                else {
                    if (prev == temp) {
                        count += 1;
                    }
                    else {
                        list.Enqueue((count, prev));
                        count = 1;
                    }
                    prev = temp;
                }
            }

            list.Enqueue((count, prev));
            list.Enqueue((1, "XX"));
            return list;
        }

        private void GeneralButton_Click(object sender, RoutedEventArgs e) {
            Button button = (Button)sender;
            switch (button.Name) {
                case "CipherButton":
                    TabMenu.SelectedIndex = 1;
                    break;
                case "DecipherButton":
                    TabMenu.SelectedIndex = 2;
                    break;
                default:
                    break;
            }
        }

        private void CipherButton_Click(object sender, RoutedEventArgs e) {

        }

        private void DecipherButton_Click(object sender, RoutedEventArgs e) {

        }

        private void OpenFileButton2_Click(object sender, RoutedEventArgs e) {
            this.ErrorHeader.Visibility = Visibility.Hidden;
            this.ErrorLabel.Visibility = Visibility.Hidden;
            var dialog = new Microsoft.Win32.OpenFileDialog {
                FileName = "Document",
                DefaultExt = ".docx",
                Filter = "Documents|*.doc;*.docx;*.txt;*.pdf",
                CheckPathExists = true,
                CheckFileExists = true,
                Multiselect = false,
            };

            bool? result = dialog.ShowDialog();
            if (result == true) {
                this.message_file_path = dialog.FileName;

                int last_dot = this.message_file_path.LastIndexOf('.');
                if (last_dot == -1) {
                    return;
                }

                string ext = this.message_file_path.Substring(last_dot+1);

                switch (ext) {
                    case "txt":
                        try {
                            this.ClearTextBox.Text = File.ReadAllText(this.message_file_path, Encoding.GetEncoding(1251));
                        }
                        catch (Exception ex) {
                            this.ErrorHeader.Visibility = Visibility.Visible;
                            this.ErrorLabel.Visibility = Visibility.Visible;
                            this.ErrorLabel.Content = ex.Message;
                            return;
                        }
                        break;
                    case "pdf":
                        try {
                            using (PdfReader reader = new PdfReader(this.message_file_path)) {
                                PdfDocument doc = new PdfDocument(reader);
                                string text = string.Empty;
                                for (int page = 1; page <= doc.GetNumberOfPages(); page++) {
                                    text += PdfTextExtractor.GetTextFromPage(doc.GetPage(page));
                                }
                                reader.Close();
                                this.ClearTextBox.Text = text;
                            }
                        }
                        catch (Exception ex) {
                            this.ErrorHeader.Visibility = Visibility.Visible;
                            this.ErrorLabel.Visibility = Visibility.Visible;
                            this.ErrorLabel.Content = ex.Message;
                            return;
                        }
                        break;
                    case "doc":
                    case "docx":
                        try {
                            using (WordprocessingDocument document = WordprocessingDocument.Open(this.message_file_path, false)) {
                                Body? body = document.MainDocumentPart?.Document.Body;
                                this.ClearTextBox.Text = body?.InnerText;
                            }
                        }
                        catch (Exception ex) {
                            this.ErrorHeader.Visibility = Visibility.Visible;
                            this.ErrorLabel.Visibility = Visibility.Visible;
                            this.ErrorLabel.Content = ex.Message;
                            return;
                        }
                        break;
                    default:
                        return;
                }

                this.BinaryClearTextBox.Clear();
                this.BinaryClearTextBox.Text = ToBinaryString(Encoding.GetEncoding(1251), this.ClearTextBox.Text ?? "");

                this.bitsCount = BinaryClearTextBox.Text.Length;
                BitsCounterLabel.Content = this.bitsCount;
                ContainerCheckButton.IsEnabled = false;
                ContainerCheckLabel.Visibility = Visibility.Hidden;
                ClearTextNextButton.IsEnabled = true;
            }
        }

        private void OpenFile2Button2_Click(object sender, RoutedEventArgs e) {
            this.Error2Header.Visibility = Visibility.Hidden;
            this.Error2Label.Visibility = Visibility.Hidden;
            this.WordCountButton.IsEnabled = false;
            this.WordCounterLabel.Content = string.Empty;
            this.ContainerCheckLabel.Content = string.Empty;
            this.ContainerCheckButton.IsEnabled = false;
            this.ContainerTest.IsEnabled = false;
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document";
            dialog.DefaultExt = ".docx";
            dialog.Filter = "Documents|*.doc;*.docx";
            dialog.CheckPathExists = true;
            dialog.CheckFileExists = true;
            dialog.Multiselect = false;

            bool? result = dialog.ShowDialog();
            if (result == true) {
                this.container_file_path = dialog.FileName;
                try {
                    string container = string.Empty;
                    using (WordprocessingDocument document = WordprocessingDocument.Open(this.container_file_path, false)) {
                        Body? body = document.MainDocumentPart?.Document.Body;
                        this.ContainerTextBox.Text = body?.InnerText;
                    }
                }
                catch (Exception ex) {
                    this.Error2Header.Visibility = Visibility.Visible;
                    this.Error2Label.Visibility = Visibility.Visible;
                    this.Error2Label.Content = ex.Message;
                    return;
                }

                WordCountButton.IsEnabled = true;
                ContainerCheckButton.IsEnabled = false;
                ContainerCheckLabel.Visibility = Visibility.Hidden;
                //ContainerNextButton.IsEnabled = false;
            }
        }

        private void WordCountButton2_Click(object sender, RoutedEventArgs e) {
            string pattern = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~–\\s]";
            Regex regex = new Regex(pattern);
            string text = ContainerTextBox.Text;
            this.ContainerTextBox.Text = regex.Replace(text, string.Empty);
            int letters = this.ContainerTextBox.Text.Length;
            this.lettersCount = letters;
            WordCounterLabel.Content = this.lettersCount;
            ContainerCheckButton.IsEnabled = true;
            //ContainerNextButton.IsEnabled = false;
        }

        private void ContainerCheckButton2_Click(object sender, RoutedEventArgs e) {
            ContainerCheckLabel.Content = "Не подходит!";
            ContainerCheckLabel.Foreground = Brushes.Red;
            //ContainerNextButton.IsEnabled = false;
            ContainerTest.IsEnabled = false;
            ContainerCheckLabel.Visibility = Visibility.Visible;
            if (this.lettersCount >= this.bitsCount / 2) {
                ContainerCheckLabel.Content = "Подходит!";
                ContainerCheckLabel.Foreground = Brushes.Green;
                //ContainerNextButton.IsEnabled = true;
                ContainerTest.IsEnabled = true;
            }
        }

        private void ClearTextNextButton_Click(object sender, RoutedEventArgs e) {
            CipherTabs.SelectedIndex = 1;
        }

        private void ContainerPrevButton_Click(object sender, RoutedEventArgs e) {
            CipherTabs.SelectedIndex = 0;
        }

        private void ContainerNextButton_Click(object sender, RoutedEventArgs e) {

        }

        private void ContainerTest2_Click(object sender, RoutedEventArgs e) {
            TestLabel.Foreground = Brushes.Red;
            List<(int, string)> bits_info = new List<(int, string)> ();
            string encoded_msg = this.BinaryClearTextBox.Text;
            bool is_active = true;
            string ignored_chars = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~– ]";

            Queue<(int, string)> bits_queue = parseBitString(encoded_msg, 2);

            using (WordprocessingDocument wordprocessingDocument =
                         WordprocessingDocument.Open(this.container_file_path, true)) {

                Body body = wordprocessingDocument.MainDocumentPart?.Document.Body??new Body();
                var paragraphs = body.ChildElements.Where(child => child is Paragraph);
                if (paragraphs == null) {
                    return;
                }
                foreach (Paragraph paragraph in paragraphs) {
                    if (paragraph == null) {
                        continue;
                    }

                    var childs = paragraph.ChildElements.ToList();

                    paragraph.RemoveAllChildren();

                    foreach (var child in childs) {
                        if (!(child is Run) || !is_active) {
                            paragraph.AppendChild(child);
                            continue;
                        }

                        if (bits_queue.Count == 0) {
                            paragraph.AppendChild(child);
                            is_active = false;
                            continue;
                        }

                        Run? current_run = child as Run;
                        if (current_run == null) {
                            continue;
                        }

                        int run_txt_len = current_run.InnerText.Length,
                run_clean_txt_len = current_run.InnerText.Where(ch => !ignored_chars.Contains(ch)).Count();
                        if (run_clean_txt_len == 0) {
                            paragraph.AppendChild(child);
                            continue;
                        }

                        Queue<(int, string)> fit = new Queue<(int, string)>();
                        while (run_clean_txt_len > 0 && bits_queue.Count != 0) {
                            (int, string) part = bits_queue.Dequeue();
                            int work_size = part.Item1;
                            if (run_clean_txt_len < part.Item1) {
                                bits_queue = new Queue<(int, string)>(bits_queue.Prepend((part.Item1 - run_clean_txt_len, part.Item2)));
                                work_size = run_clean_txt_len;
                            }
                            run_clean_txt_len -= work_size;
                            fit.Enqueue((work_size, part.Item2));
                        }
                        bool is_ignored = false;
                        string text = current_run.InnerText;
                        while (fit.Count > 0) {
                            Run temp = (Run)current_run.CloneNode(true);
                            temp.RemoveAllChildren<Text>();
                            var part = fit.Dequeue();
                            int len = 0;

                            char[]? chars = null;
                            if (!is_ignored) {
                                chars = text.TakeWhile(ch => {
                                    len += 1;
                                    is_ignored = ignored_chars.Contains(ch);
                                    return len <= part.Item1 && !is_ignored;
                                }).ToArray();
                            }
                            else {
                                chars = text.TakeWhile(ch => {
                                    len += 1;
                                    return ignored_chars.Contains(ch);
                                }).ToArray();
                                text = text.Remove(0, len - 1);
                                temp.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(chars) });

                                paragraph.AppendChild(temp);
                                fit = new Queue<(int, string)>(fit.Prepend(part));
                                is_ignored = false;
                                continue;
                            }

                            if (is_ignored && len == 1) {
                                fit = new Queue<(int, string)>(fit.Prepend(part));
                                continue;
                            }
                            len += chars.Length == text.Length ? 1 : 0;
                            text = text.Remove(0, len - 1);
                            temp.AddChild(new Text(new string(chars)));
                            TextOutlineEffect? outline = getOutline(part.Item2);

                            temp.RunProperties = temp.RunProperties ?? new RunProperties();
                            temp.RunProperties.TextOutlineEffect = outline;
                            paragraph.AppendChild(temp);

                            if (chars.Length != part.Item1) {
                                fit = new Queue<(int, string)>(fit.Prepend((part.Item1 - chars.Length, part.Item2)));
                            }
                        }

                        if (fit.Count == 0 && text.Length != 0 && run_clean_txt_len == 0) {
                            Run temp = (Run)current_run.CloneNode(true);
                            temp.RemoveAllChildren<Text>();
                            temp.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(text) });
                            paragraph.AppendChild(temp);
                        }
                        else if (bits_queue.Count == 0 && run_clean_txt_len > 0) {
                            is_active = false;
                            Run temp = (Run)current_run.CloneNode(true);
                            temp.RemoveAllChildren<Text>();
                            temp.AddChild(new Text(text));
                            paragraph.AppendChild(temp);
                        }

                    }
                }
                wordprocessingDocument.Save();
                TestLabel.Foreground = Brushes.Green;
            }
        }

        private void OpenFile3Button_Click(object sender, RoutedEventArgs e) {
            EncodedFileStatus.Foreground = Brushes.Red;
            EncodedFileStatus.Content = "Файл не выбран";
            DecodeButton.IsEnabled = false;
            this.encoded_file_path = String.Empty;

            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document";
            dialog.DefaultExt = ".docx";
            dialog.Filter = "Word Documents|*.doc;*.docx";

            bool? result = dialog.ShowDialog();
            if (result == true) {
                this.encoded_file_path = dialog.FileName;
                EncodedFileStatus.Foreground = Brushes.Green;
                EncodedFileStatus.Content = "Файл открыт";
                DecodeButton.IsEnabled = true;
            }
        }

        private void DecodeButton2_Click(object sender, RoutedEventArgs e) {
            this.DecodeStatusValue.Content = "Неверный или поврежденный файл!";
            this.DecodeStatusValue.Foreground = Brushes.Red;

            if (string.IsNullOrEmpty(this.encoded_file_path))
                return;

            string encoded_msg = string.Empty;
            bool is_ended = false;
            int error_counter = 0;

            using (WordprocessingDocument wordprocessingDocument =
             WordprocessingDocument.Open(this.encoded_file_path, false)) {

                Body body = wordprocessingDocument.MainDocumentPart?.Document.Body??new Body();

                foreach (OpenXmlElement el in body.ChildElements) {
                    if (is_ended) {
                        break;
                    }
                    if (el is Paragraph) {
                        foreach (OpenXmlElement pel in el.ChildElements) {
                            if (pel is Run) {
                                Run run = (Run)pel;
                                string code = string.Empty;

                                TextOutlineEffect? effect = run.RunProperties?.TextOutlineEffect;
                                int? width = effect?.LineWidth?.Value;
                                int? alpha = effect?
                                    .GetFirstChild<SolidColorFillProperties>()?
                                    .GetFirstChild<SchemeColor>()?
                                    .GetFirstChild<Alpha>()?.Val?.Value;

                                switch ((width, alpha)) {
                                    case (635, 96000):
                                        code = "00";
                                        error_counter = 0;
                                        break;
                                    case (1270, 96000):
                                        code = "10";
                                        error_counter = 0;
                                        break;
                                    case (635, 100000):
                                        code = "01";
                                        error_counter = 0;
                                        break;
                                    case (1270, 100000):
                                        code = "11";
                                        error_counter = 0;
                                        break;
                                    case (635, 99000):
                                        is_ended = true;
                                        break;
                                    default:
                                        error_counter += 1;
                                        code = string.Empty;
                                        break;
                                }

                                if (error_counter == 5) {
                                    break;
                                }

                                if (!string.IsNullOrEmpty(code)) {
                                    for (int j = 0; j < run.InnerText.Length; j++) {
                                        encoded_msg += code;
                                    }
                                }
                                if (is_ended) {
                                    break;
                                }
                            }
                        }

                    }
                }

            }

            if (error_counter != 0 && !is_ended) {
                this.DecodeStatusValue.Content = "Неверный или поврежденный файл!";
                this.DecodeStatusValue.Foreground = Brushes.Red;
                return;
            }

            byte[] bytes = new byte[encoded_msg.Length / 8];

            BitArray bits = new BitArray(encoded_msg.Length);

            int i = 0;

            foreach (char bit in encoded_msg.Reverse()) {
                if (bit == '1') {
                    bits.Set(i, true);
                }
                i++;
            }

            bits.CopyTo(bytes, 0);
            Encoding win1251 = Encoding.GetEncoding(1251);
            DecodedText.Text = string.Join("", win1251.GetString(bytes).Reverse());
            this.DecodeStatusValue.Content = "Успешно извлечено!";
            this.DecodeStatusValue.Foreground = Brushes.Green;
        }

        private void MenuButton_Click(object sender, RoutedEventArgs e) {
            this.TabMenu.SelectedIndex = 0;
        }

        private void ClearTextBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e) {
            this.BinaryClearTextBox.Text = ToBinaryString(Encoding.GetEncoding(1251), this.ClearTextBox.Text);
            this.BitsCounterLabel.Content = this.BinaryClearTextBox.Text.Length;
        }
    }
}

