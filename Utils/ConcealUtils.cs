using DocumentFormat.OpenXml.Office2010.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StegoLine.Utils {
    public class ConcealUtils {

        private static readonly TextOutlineEffect OutlineBase;

        static ConcealUtils() {
            OutlineBase = new TextOutlineEffect {
                CapType = LineCapValues.Round,
                Alignment = PenAlignmentValues.Center,
                Compound = CompoundLineValues.Simple,
            };

            OutlineBase.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                Properties.General.Default.NamespacePrefix,
                Properties.General.Default.RunAttributeName,
                Properties.General.Default.NamespaceUri,
                "True"
            ));
        }

        public class ParseConfig {
            public Encoding Encoding {
                get;
            }
            public int PartSize {
                get;
            }
            public int ContainerLength {
                get;
            }
            public int MsgLength {
                get;
            }

            public ParseConfig(Encoding encoding, int partSize, int containerLength, int msgLength) {
                this.Encoding = encoding;
                this.PartSize = partSize;
                this.ContainerLength = containerLength;
                this.MsgLength = msgLength;
            }
        }
        public class MsgInfo {
            public string Text {
                get;
            }
            public Encoding MsgEncoding {
                get;
            }
            public int BinaryMsgLength {
                get;
            }

            public MsgInfo(string msg, Encoding msgEncoding, int binaryMsgLength) {
                this.Text = msg;
                this.MsgEncoding = msgEncoding;
                this.BinaryMsgLength = binaryMsgLength;
            }
        }

        public enum Features {
            WIDTH,
            ALPHA
        }

        public static string ToBinaryString(Encoding encoding, string text) {
            return string.Join("", encoding.GetBytes(text).Select(n => Convert.ToString(n, 2).PadLeft(8, '0')));
        }

        public static Queue<(int, int)>? ParseMsg(string Msg, ParseConfig Config) {

            if (Config == null) {
                return null;
            }

            var list = new Queue<(int, int)>();

            byte[] bytes = Config.Encoding.GetBytes(Msg);
            int mask = (int)(Math.Pow(2, Config.PartSize) - 1);

            int Steps = sizeof(byte) * 8 / Config.PartSize;

            if (Config.ContainerLength < Config.MsgLength / Config.PartSize) {
                return null;
            }

            if (Config.MsgLength == 0) {
                return null;
            }

            int MaxRand = Config.ContainerLength / (Config.MsgLength / Config.PartSize);

            foreach (byte _byte in bytes.Reverse()) {
                int TempBuff = _byte;
                int Position = new Random().Next(0, MaxRand);
                int Before = Position, After = MaxRand - Position;
                for (int i = 0; i < Steps; i++) {
                    if (Before != 0) {
                        list.Enqueue((Before, -1));
                    }
                    int BitsPart = TempBuff & mask;
                    list.Enqueue((1, BitsPart));
                    TempBuff >>= Config.PartSize;
                    if (After != 0) {
                        list.Enqueue((After, -1));
                    }
                }
            }

            return list;
        }

        public static TextOutlineEffect? GetOutlineObj(int bits, int size) {

            if (bits == -1) {
                return null;
            }

            TextOutlineEffect? effect = (TextOutlineEffect?)OutlineBase?.CloneNode(true);

            if (effect == null) {
                return effect;
            }

            SchemeColor? scheme = new() {
                Val = SchemeColorValues.ExtraSchemeColor1,
            };

            int alpha = GetValueFromFullBits(bits, size, Features.ALPHA);
            if (alpha == -1) {
                return null;
            }

            _ = scheme.AddChild(new Alpha {
                Val = alpha
            });

            _ = effect.AddChild(new SolidColorFillProperties(scheme));
            _ = effect.AddChild(new PresetLineDashProperties {
                Val = PresetLineDashValues.Solid
            });
            _ = effect.AddChild(new BevelEmpty());

            int width = GetValueFromFullBits(bits, size, Features.WIDTH);

            if (width == -1) {
                return null;
            }

            effect.LineWidth = width;
            return effect;
        }

        public static int GetValueFromFullBits(int bits, int size, Features prop) {
            int HalfSize = size / 2;
            int mask = (int)(Math.Pow(2, HalfSize) - 1);
            return prop switch {
                Features.WIDTH => GetOutlineWidthValue(bits & mask),
                Features.ALPHA => GetOutlineAlphaValue((bits & (mask << HalfSize)) >> HalfSize),
                _ => -1,
            };
        }

        private static int GetOutlineWidthValue(int bits) {
            return Properties.General.Default.MinOutlineWidth + Properties.Conceal.Default.OutlineWidthStep * bits;
        }

        private static int GetOutlineAlphaValue(int bits) {
            return Properties.General.Default.MinOutlineAlpha + Properties.Conceal.Default.OutlineAlphaStep * bits;
        }

        public static void CalculateOutlineWidthStep(int size) {
            int TempBuff = (int)Math.Pow(2, size / 2) - 1;
            int Step = (Properties.General.Default.MaxOutlineWidth - Properties.General.Default.MinOutlineWidth) / TempBuff;
            Properties.Conceal.Default.OutlineWidthStep = Step;
            Properties.Conceal.Default.Save();
        }

        public static void CalculateOutlineAlphaStep(int size) {
            int TempBuff = (int)Math.Pow(2, size / 2) - 1;
            int Step = (Properties.General.Default.MaxOutlineAlpha - Properties.General.Default.MinOutlineAlpha) / TempBuff;
            Properties.Conceal.Default.OutlineAlphaStep = Step;
            Properties.Conceal.Default.Save();
        }
    }
}
