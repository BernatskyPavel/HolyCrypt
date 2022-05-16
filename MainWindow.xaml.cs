using DocumentFormat.OpenXml.Office2010.Word;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace HolyCryptv3 {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow: Window {

        private int MinOutlineWidth = 635;
        private int MaxOutlineWidth = 1270;
        private (int Encode, int Decode) OutlineWidthSteps = (0,0);
        private int MinOutlineAlpha = 96000;
        private int MaxOutlineAlpha = 100000;
        private (int Encode, int Decode) OutlineAlphaSteps = (0,0);
        private readonly Encoding Encoding;

        private Dictionary<string, (int, int)> outlines = new Dictionary<string, (int, int)> {
            {"00", (96000,635)},
            {"01", (100000,635)},
            {"10", (96000,1270)},
            {"11", (100000,1270)},
            {"XX", (99000,635)},
        };

        private readonly TextOutlineEffect? OutlineBase;
        //private TextOutlineEffect? CompleteOutline = null;

        public MainWindow() {

            this.OutlineBase = new TextOutlineEffect {
                CapType = LineCapValues.Round,
                Alignment = PenAlignmentValues.Center,
                Compound = CompoundLineValues.Simple,
            };

            this.OutlineBase.AddNamespaceDeclaration("w41", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");

            //this.CompleteOutline = this.getOutline("XX");
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.Encoding = Encoding.GetEncoding(1251);

            InitializeComponent();
            this.CalculateOutlineAlphaStep((int)this.BitsPerSymbolSlider.Value, Type.BOTH);
            this.CalculateOutlineWidthStep((int)this.BitsPerSymbolSlider.Value, Type.BOTH);
        }

        private void ClearTextNextButton_Click(object sender, RoutedEventArgs e) {
            ConcealTabControl.SelectedIndex = 1;
        }

        private void ContainerPrevButton_Click(object sender, RoutedEventArgs e) {
            ConcealTabControl.SelectedIndex = 0;
        }

        private void BitsPerSymbolSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) {
            Type regime = ((Slider)sender).Name == "BitsPerSymbolSlider" ? Type.ENCODE : Type.DECODE;
            this.CalculateOutlineAlphaStep((int)e.NewValue, regime);
            this.CalculateOutlineWidthStep((int)e.NewValue, regime);
        }
    }
}

