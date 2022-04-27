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
    }
}
