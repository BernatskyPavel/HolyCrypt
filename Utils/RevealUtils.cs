using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace StegoLine.Utils {
    public static class RevealUtils {
        public static void CalculateOutlineWidthStep(int size) {
            int TempBuff = (int)Math.Pow(2, size / 2) - 1;
            int Step = (Properties.General.Default.MaxOutlineWidth - Properties.General.Default.MinOutlineWidth) / TempBuff;
            Properties.Reveal.Default.OutlineWidthStep = Step;
            Properties.Reveal.Default.Save();
        }

        public static void CalculateOutlineAlphaStep(int size) {
            int TempBuff = (int)Math.Pow(2, size / 2) - 1;
            int Step = (Properties.General.Default.MaxOutlineAlpha - Properties.General.Default.MinOutlineAlpha) / TempBuff;
            Properties.Reveal.Default.OutlineAlphaStep = Step;
            Properties.Reveal.Default.Save();
        }

        public static int GetBitsFromOutlineAlpha(int alpha) {
            return (int)Math.Ceiling((double)(alpha - Properties.General.Default.MinOutlineAlpha) / Properties.Reveal.Default.OutlineAlphaStep);
        }

        public static int GetBitsFromOutlineWidth(int width) {
            return (int)Math.Ceiling((double)(width - Properties.General.Default.MinOutlineWidth) / Properties.Reveal.Default.OutlineWidthStep);
        }

        public static string? GetHashCode(string FilePath) {
            if (string.IsNullOrEmpty(FilePath))
                return null;

            try {
                using WordprocessingDocument Document = WordprocessingDocument.Open(FilePath, false);
                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();
                return DocumentBody.GetAttribute(Properties.General.Default.HashAttributeName, Properties.General.Default.NamespaceUri).Value;
            }
            catch (Exception) {
                return null;
            }
        }

        public static (byte[], string?) RevealMsg(string FilePath, int BitsPerSymbol) {
            bool IsMsgEnded = false;
            int ErrorCounter = 0;

            int HalfSize = BitsPerSymbol / 2;
            int Steps = sizeof(byte) * 8 / BitsPerSymbol;
            int Index = 0;
            List<byte> _bytes = new();
            string? HashCode = null;


            using (WordprocessingDocument Document =
             WordprocessingDocument.Open(FilePath, false)) {

                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();

                HashCode = DocumentBody.GetAttribute(Properties.General.Default.HashAttributeName, Properties.General.Default.NamespaceUri).Value;

                foreach (OpenXmlElement BodyElement in DocumentBody.ChildElements) {
                    if (IsMsgEnded || ErrorCounter == 5) {
                        break;
                    }
                    //if (ErrorCounter == 5) {
                    //    break;
                    //}
                    if (BodyElement is Paragraph) {
                        foreach (OpenXmlElement ParagraphChild in BodyElement.ChildElements) {
                            if (ParagraphChild is Run ChildRun) {
                                string BitPattern = string.Empty;

                                TextOutlineEffect? OutlineEffect = ChildRun.RunProperties?.TextOutlineEffect;
                                if (null == OutlineEffect) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                try {
                                    _ = OutlineEffect.GetAttribute(Properties.General.Default.RunAttributeName, Properties.General.Default.NamespaceUri);
                                }
                                catch (KeyNotFoundException) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                int? Width = OutlineEffect?.LineWidth?.Value;
                                int? Alpha = OutlineEffect?
                                    .GetFirstChild<SolidColorFillProperties>()?
                                    .GetFirstChild<SchemeColor>()?
                                    .GetFirstChild<Alpha>()?.Val?.Value;

                                if (null == Alpha || null == Width) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                int LowerBits = GetBitsFromOutlineWidth(Width.Value);
                                int HigherBits = GetBitsFromOutlineAlpha(Alpha.Value);

                                int Bits = (HigherBits << HalfSize) | LowerBits;

                                _bytes.Add((byte)Bits);
                                ErrorCounter = 0;
                            }
                        }

                    }
                }
            }

            int ByteLen = _bytes.Count * BitsPerSymbol / 8;
            byte[] bytes = new byte[ByteLen];

            //_bytes.Reverse();

            foreach (byte[] parts in _bytes.Chunk(Steps)) {
                byte NewByte = 0;
                for (int _i = 0; _i < Math.Min(Steps, parts.Length); _i++) {
                    NewByte = (byte)(NewByte | (parts[_i] << (BitsPerSymbol * _i)));
                }
                if (Index < ByteLen) {
                    bytes[Index++] = NewByte;
                }
            }

            return (bytes.Reverse().ToArray(), HashCode);
        }

        public static List<byte> GetRawBytes(string FilePath, int BitsPerSymbol) {
            bool IsMsgEnded = false;
            int ErrorCounter = 0;

            int HalfSize = BitsPerSymbol / 2;
            List<byte> _bytes = new();
            //string? HashCode = null;


            using (WordprocessingDocument Document =
             WordprocessingDocument.Open(FilePath, false)) {

                Body DocumentBody = Document.MainDocumentPart?.Document.Body??new Body();

                //                HashCode = DocumentBody.GetAttribute(Properties.General.Default.HashAttributeName, Properties.General.Default.NamespaceUri).Value;

                foreach (OpenXmlElement BodyElement in DocumentBody.ChildElements) {
                    if (IsMsgEnded) {
                        break;
                    }
                    if (ErrorCounter == 50) {
                        break;
                    }
                    if (BodyElement is Paragraph) {
                        foreach (OpenXmlElement ParagraphChild in BodyElement.ChildElements) {
                            if (ParagraphChild is Run ChildRun) {

                                TextOutlineEffect? OutlineEffect = ChildRun.RunProperties?.TextOutlineEffect;
                                if (null == OutlineEffect) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                try {
                                    _ = OutlineEffect.GetAttribute(Properties.General.Default.RunAttributeName, Properties.General.Default.NamespaceUri);
                                }
                                catch (KeyNotFoundException) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                int? Width = OutlineEffect?.LineWidth?.Value;
                                int? Alpha = OutlineEffect?
                                    .GetFirstChild<SolidColorFillProperties>()?
                                    .GetFirstChild<SchemeColor>()?
                                    .GetFirstChild<Alpha>()?.Val?.Value;

                                if (null == Alpha || null == Width) {
                                    ErrorCounter += 1;
                                    continue;
                                }

                                int LowerBits = GetBitsFromOutlineWidth(Width.Value);
                                int HigherBits = GetBitsFromOutlineAlpha(Alpha.Value);

                                int Bits = (HigherBits << HalfSize) | LowerBits;

                                _bytes.Add((byte)Bits);
                                ErrorCounter = 0;
                            }
                        }

                    }
                }
            }

            return _bytes;
        }

        public static byte[] ParseRawBytes(List<byte> RawBytes, int BitsPerSymbol) {
            int Steps = sizeof(byte) * 8 / BitsPerSymbol;
            int ByteLen = RawBytes.Count * BitsPerSymbol / 8;
            byte[] bytes = new byte[ByteLen];
            int Index = 0;

            foreach (byte[] parts in RawBytes.Chunk(Steps)) {
                byte NewByte = 0;
                for (int _i = 0; _i < Math.Min(Steps, parts.Length); _i++) {
                    NewByte = (byte)(NewByte | (parts[_i] << (BitsPerSymbol * _i)));
                }
                if (Index < ByteLen) {
                    bytes[Index++] = NewByte;
                }
            }

            return bytes.Reverse().ToArray();
        }
    }
}
