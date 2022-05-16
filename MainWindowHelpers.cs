using DocumentFormat.OpenXml.Office2010.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace HolyCryptv3
{
    partial class MainWindow : Window {

        private Random Rand = new Random();
        enum Properties {
            WIDTH,
            ALPHA
        }
        enum Type {
            ENCODE,
            DECODE,
            BOTH
        }

        static string ToBinaryString(Encoding encoding, string text) {
            return string.Join("", encoding.GetBytes(text).Select(n => Convert.ToString(n, 2).PadLeft(8, '0')));
        }

        private TextOutlineEffect? getOutline(string bits) {
            string key = bits;

            if (!this.outlines.ContainsKey(bits)) {
                key = "XX";
            }

            (int, int) values = this.outlines[key];

            TextOutlineEffect? effect = (TextOutlineEffect?)this.OutlineBase?.CloneNode(true);

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

        private TextOutlineEffect? GetOutlineObj(int bits, int size) {

            if (bits == -1) {
                return null;
            }

            TextOutlineEffect? effect = (TextOutlineEffect?)this.OutlineBase?.CloneNode(true);

            if (effect == null) {
                return effect;
            }

            SchemeColor? scheme = new SchemeColor {
                Val = SchemeColorValues.ExtraSchemeColor1,
            };

            int alpha = this.GetValueFromFullBits(bits, size, Properties.ALPHA);
            if (alpha == -1) {
                return null;
            }

            scheme.AddChild(new Alpha {
                Val = alpha 
            });

            effect.AddChild(new SolidColorFillProperties(scheme));
            effect.AddChild(new PresetLineDashProperties {
                Val = PresetLineDashValues.Solid
            });
            effect.AddChild(new BevelEmpty());
            
            int width = this.GetValueFromFullBits(bits, size, Properties.WIDTH);

            if (width == -1) {
                return null;
            }

            effect.LineWidth = width;
            return effect;
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

        private Queue<(int, int)>? ParseMsg(string Msg, int Size) {
            var list = new Queue<(int, int)>();

            byte[] bytes = this.Encoding.GetBytes(Msg);
            int mask = (int)(Math.Pow(2, Size) - 1);

            int Steps = sizeof(byte) * 8 / Size;

            if (this.ContainerSymbolsCounter < this.MsgBitsCounter / Size) {
                return null;
            }

            int MaxRand = this.ContainerSymbolsCounter / (this.MsgBitsCounter / Size);
            
            foreach (byte _byte in bytes.Reverse()) {
                int TempBuff = _byte;
                int Position = this.Rand.Next(0, MaxRand);
                int Before = Position, After = MaxRand - Position;
                for (int i = 0; i < Steps; i++) {
                    if (Before != 0) {
                        list.Enqueue((Before, -1));
                    }
                    int BitsPart = TempBuff & mask;
                    list.Enqueue((1, BitsPart));
                    TempBuff >>= Size;
                    if (After != 0) {
                        list.Enqueue((After, -1));
                    }
                }
            }

            return list;
        }

        private void CalculateOutlineWidthStep(int size, Type regime) {
            int TempBuff = (int)Math.Pow(2, size / 2) - 1;
            int Step = (this.MaxOutlineWidth - this.MinOutlineWidth) / TempBuff;
            switch (regime) {
                case Type.ENCODE:
                    this.OutlineWidthSteps.Encode = Step;
                    break;
                case Type.DECODE:
                    this.OutlineWidthSteps.Decode = Step;
                    break;
                case Type.BOTH:
                    this.OutlineWidthSteps.Encode = Step;
                    this.OutlineWidthSteps.Decode = Step;
                    break;
            }
        }

        private void CalculateOutlineAlphaStep(int size, Type regime) {
            int TempBuff = (int)Math.Pow(2, size / 2) - 1;
            int Step = (this.MaxOutlineAlpha - this.MinOutlineAlpha) / TempBuff;
            switch (regime) {
                case Type.ENCODE:
                    this.OutlineAlphaSteps.Encode = Step;
                    break;
                case Type.DECODE:
                    this.OutlineAlphaSteps.Decode = Step;
                    break;
                case Type.BOTH:
                    this.OutlineAlphaSteps.Encode = Step;
                    this.OutlineAlphaSteps.Decode = Step;
                    break;
            }

        }

        private int GetValueFromFullBits(int bits, int size, Properties prop) {
            int HalfSize = size / 2;
            int mask = (int)(Math.Pow(2, HalfSize) - 1);
            switch (prop) {
                case Properties.WIDTH:
                    return this.GetOutlineWidthValue(bits & mask);
                case Properties.ALPHA:
                    return this.GetOutlineAlphaValue((bits & (mask << HalfSize)) >> HalfSize);
            }
            return -1;
        }

        private int GetOutlineWidthValue(int bits) {
            return this.MinOutlineWidth + this.OutlineWidthSteps.Encode * bits;
        }

        private int GetOutlineAlphaValue(int bits) {
            return this.MinOutlineAlpha + this.OutlineAlphaSteps.Encode * bits;
        }

        private int GetBitsFromOutlineAlpha(int alpha) {
            return (int)Math.Ceiling((double)(alpha - this.MinOutlineAlpha) / this.OutlineAlphaSteps.Decode);
        }

        private int GetBitsFromOutlineWidth(int width) {
            return (int)Math.Ceiling((double)(width - this.MinOutlineWidth) / this.OutlineWidthSteps.Decode);
        }
    }
}
