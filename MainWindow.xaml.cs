using DocumentFormat.OpenXml.Office2010.Word;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace HolyCryptv3 {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow: Window {
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

        private TextOutlineEffect? OutlineBase = null;
        private TextOutlineEffect? CompleteOutline = null;

        

        public MainWindow() {

            this.OutlineBase = new TextOutlineEffect {
                CapType = LineCapValues.Round,
                Alignment = PenAlignmentValues.Center,
                Compound = CompoundLineValues.Simple,
            };

            this.CompleteOutline = this.getOutline("XX");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }

        

        private void ClearTextNextButton_Click(object sender, RoutedEventArgs e) {
            ConcealTabControl.SelectedIndex = 1;
        }

        private void ContainerPrevButton_Click(object sender, RoutedEventArgs e) {
            ConcealTabControl.SelectedIndex = 0;
        }

        private void ContainerNextButton_Click(object sender, RoutedEventArgs e) {

        }

        //private void ContainerTest2_Click(object sender, RoutedEventArgs e) {
        //    TestLabel.Foreground = Brushes.Red;
        //    List<(int, string)> bits_info = new List<(int, string)> ();
        //    string encoded_msg = this.MsgBitsTextBox.Text;
        //    bool is_active = true;
        //    string ignored_chars = "[!\"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~– ]";

        //    Queue<(int, string)> bits_queue = parseBitString(encoded_msg, 2);

        //    using (WordprocessingDocument wordprocessingDocument =
        //                 WordprocessingDocument.Open(this.container_file_path, true)) {

        //        Body body = wordprocessingDocument.MainDocumentPart?.Document.Body??new Body();
        //        var paragraphs = body.ChildElements.Where(child => child is Paragraph);
        //        if (paragraphs == null) {
        //            return;
        //        }
        //        foreach (Paragraph paragraph in paragraphs) {
        //            if (paragraph == null) {
        //                continue;
        //            }

        //            var childs = paragraph.ChildElements.ToList();

        //            paragraph.RemoveAllChildren();

        //            foreach (var child in childs) {
        //                if (!(child is Run) || !is_active) {
        //                    paragraph.AppendChild(child);
        //                    continue;
        //                }

        //                if (bits_queue.Count == 0) {
        //                    paragraph.AppendChild(child);
        //                    is_active = false;
        //                    continue;
        //                }

        //                Run? current_run = child as Run;
        //                if (current_run == null) {
        //                    continue;
        //                }

        //                int run_txt_len = current_run.InnerText.Length,
        //        run_clean_txt_len = current_run.InnerText.Where(ch => !ignored_chars.Contains(ch)).Count();
        //                if (run_clean_txt_len == 0) {
        //                    paragraph.AppendChild(child);
        //                    continue;
        //                }

        //                Queue<(int, string)> fit = new Queue<(int, string)>();
        //                while (run_clean_txt_len > 0 && bits_queue.Count != 0) {
        //                    (int, string) part = bits_queue.Dequeue();
        //                    int work_size = part.Item1;
        //                    if (run_clean_txt_len < part.Item1) {
        //                        bits_queue = new Queue<(int, string)>(bits_queue.Prepend((part.Item1 - run_clean_txt_len, part.Item2)));
        //                        work_size = run_clean_txt_len;
        //                    }
        //                    run_clean_txt_len -= work_size;
        //                    fit.Enqueue((work_size, part.Item2));
        //                }
        //                bool is_ignored = false;
        //                string text = current_run.InnerText;
        //                while (fit.Count > 0) {
        //                    Run temp = (Run)current_run.CloneNode(true);
        //                    temp.RemoveAllChildren<Text>();
        //                    var part = fit.Dequeue();
        //                    int len = 0;

        //                    char[]? chars = null;
        //                    if (!is_ignored) {
        //                        chars = text.TakeWhile(ch => {
        //                            len += 1;
        //                            is_ignored = ignored_chars.Contains(ch);
        //                            return len <= part.Item1 && !is_ignored;
        //                        }).ToArray();
        //                    }
        //                    else {
        //                        chars = text.TakeWhile(ch => {
        //                            len += 1;
        //                            return ignored_chars.Contains(ch);
        //                        }).ToArray();
        //                        text = text.Remove(0, len - 1);
        //                        temp.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(chars) });

        //                        paragraph.AppendChild(temp);
        //                        fit = new Queue<(int, string)>(fit.Prepend(part));
        //                        is_ignored = false;
        //                        continue;
        //                    }

        //                    if (is_ignored && len == 1) {
        //                        fit = new Queue<(int, string)>(fit.Prepend(part));
        //                        continue;
        //                    }
        //                    len += chars.Length == text.Length ? 1 : 0;
        //                    text = text.Remove(0, len - 1);
        //                    temp.AddChild(new Text(new string(chars)));
        //                    TextOutlineEffect? outline = getOutline(part.Item2);

        //                    temp.RunProperties = temp.RunProperties ?? new RunProperties();
        //                    temp.RunProperties.TextOutlineEffect = outline;
        //                    paragraph.AppendChild(temp);

        //                    if (chars.Length != part.Item1) {
        //                        fit = new Queue<(int, string)>(fit.Prepend((part.Item1 - chars.Length, part.Item2)));
        //                    }
        //                }

        //                if (fit.Count == 0 && text.Length != 0 && run_clean_txt_len == 0) {
        //                    Run temp = (Run)current_run.CloneNode(true);
        //                    temp.RemoveAllChildren<Text>();
        //                    temp.AddChild(new Text { Space = SpaceProcessingModeValues.Preserve, Text = new string(text) });
        //                    paragraph.AppendChild(temp);
        //                }
        //                else if (bits_queue.Count == 0 && run_clean_txt_len > 0) {
        //                    is_active = false;
        //                    Run temp = (Run)current_run.CloneNode(true);
        //                    temp.RemoveAllChildren<Text>();
        //                    temp.AddChild(new Text(text));
        //                    paragraph.AppendChild(temp);
        //                }

        //            }
        //        }
        //        wordprocessingDocument.Save();
        //        TestLabel.Foreground = Brushes.Green;
        //    }
        //}
    }
}

