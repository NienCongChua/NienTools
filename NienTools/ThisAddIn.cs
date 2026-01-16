using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelDna.Integration;

namespace NienTools
{
    public partial class ThisAddIn
    {
        // -------------------------
        // VNTools-like Excel functions (non-accented names)
        // -------------------------

        [ExcelFunction(Name = "DocTien", Description = "Chuyển số thành chữ tiền Việt (tên gốc: =VND)")]
        public static string DocTien(
            object So,
            object CoDonVi = null,
            object VietHoa = null,
            object DonViTien = null,
            object DonViXu = null)
        {
            try
            {
                if (So == null) return "";

                double value;
                if (!TryGetDouble(So, out value)) return "#VALUE!";

                // Default behavior: show unit (đồng) and capitalize first letter for accounting documents
                bool coDonVi = IsMissing(CoDonVi) ? true : ToBool(CoDonVi, true);
                bool vietHoa = IsMissing(VietHoa) ? true : ToBool(VietHoa, false);
                string donViTien = DonViTien as string ?? "đồng";
                string donViXu = DonViXu as string ?? "xu";

                string sign = value < 0 ? "âm " : "";
                value = Math.Abs(value);

                long intPart = (long)Math.Floor(value);
                int decimalPart = (int)Math.Round((value - intPart) * 100, MidpointRounding.AwayFromZero);

                // carry: 1.999 -> 2.00
                if (decimalPart >= 100)
                {
                    intPart += 1;
                    decimalPart -= 100;
                }

                string words = intPart == 0 ? "không" : NumberToVietnamese(intPart);
                if (string.IsNullOrWhiteSpace(words)) words = "không";

                if (coDonVi)
                {
                    if (decimalPart == 0)
                    {
                        words = $"{words} {donViTien} chẵn";
                    }
                    else
                    {
                        string xuWords = decimalPart == 0 ? "" : NumberToVietnamese(decimalPart);
                        words = $"{words} {donViTien} {xuWords} {donViXu}";
                    }
                }
                else
                {
                    if (decimalPart > 0)
                    {
                        words = $"{words} phẩy {DecimalPartToWords(decimalPart)}";
                    }
                }

                words = sign + words;
                if (vietHoa)
                {
                    words = CapitalizeFirst(words);
                }

                return words;
            }
            catch
            {
                return "#ERROR";
            }
        }

        [ExcelFunction(Name = "DocSo", Description = "Đọc số thuần (không thêm đơn vị) (tên gốc: =VND/UNI without unit)")]
        public static string DocSo(
            object So,
            object VietHoa = null)
        {
            try
            {
                if (So == null) return "";

                double value;
                if (!TryGetDouble(So, out value)) return "#VALUE!";

                // Default: capitalize first letter for accounting convenience
                bool vietHoa = IsMissing(VietHoa) ? true : ToBool(VietHoa, false);
                string sign = value < 0 ? "âm " : "";
                value = Math.Abs(value);

                long intPart = (long)Math.Floor(value);
                int decimalPart = (int)Math.Round((value - intPart) * 100);

                string words = intPart == 0 ? "không" : NumberToVietnamese(intPart);
                if (decimalPart > 0)
                {
                    words = $"{words} phẩy {DecimalPartToWords(decimalPart)}";
                }

                words = sign + words;
                if (vietHoa)
                {
                    words = CapitalizeFirst(words);
                }

                return words;
            }
            catch
            {
                return "#ERROR";
            }
        }

        [ExcelFunction(Name = "DocTienAnh", Description = "Chuyển số thành chữ tiếng Anh (tên gốc: =USD)")]
        public static string DocTienAnh(
            object So,
            object CoDollar = null)
        {
            try
            {
                if (So == null) return "";

                double value;
                if (!TryGetDouble(So, out value)) return "#VALUE!";

                bool coDollar = IsMissing(CoDollar) ? true : ToBool(CoDollar, true);

                string sign = value < 0 ? "Minus " : "";
                value = Math.Abs(value);

                long dollars = (long)Math.Floor(value);
                int cents = (int)Math.Round((value - dollars) * 100, MidpointRounding.AwayFromZero);

                // carry if rounding produced 100 cents
                if (cents >= 100)
                {
                    dollars += 1;
                    cents -= 100;
                }

                string words = NumberToEnglish(dollars);
                if (coDollar)
                {
                    words = $"{words} {(dollars == 1 ? "Dollar" : "Dollars")}";
                }

                if (cents > 0)
                {
                    // read cents as words, then 'cent' (singular as requested)
                    var centWords = NumberToEnglish(cents);
                    words += $" and {centWords} cent";
                }

                // Capitalize first character of the whole result
                words = CapitalizeFirst((sign + words).Trim());

                return words;
            }
            catch
            {
                return "#ERROR";
            }
        }

        [ExcelFunction(Name = "DocUnicode", Description = "Đọc số theo bảng mã Unicode (tên gốc: =UNI)")]
        public static string DocUnicode(object So)
        {
            // Unicode is default in .NET strings; behave like DocSo without unit
            return DocSo(So, 0);
        }

        [ExcelFunction(Name = "DocVNI", Description = "Đọc theo bảng mã VNI (tên gốc: =VNI)")]
        public static string DocVNI(object So)
        {
            // Convert Unicode Vietnamese output to a VNI-like representation for legacy fonts.
            string unicode = DocSo(So, 0);
            if (string.IsNullOrEmpty(unicode)) return unicode;
            return ConvertUnicodeToVNI(unicode);
        }

        [ExcelFunction(Name = "ChuThanhSo", Description = "Chuyển chữ thành số (giới hạn - simple parser)")]
        public static object ChuThanhSo(string chu)
        {
            // Implement a simple parser that handles common forms like "mot trieu hai tram..."
            if (string.IsNullOrWhiteSpace(chu)) return 0;
            try
            {
                long result;
                if (TryParseVietnameseNumber(chu, out result))
                {
                    return result;
                }
                return "#N/A";
            }
            catch
            {
                return "#ERROR";
            }
        }

        [ExcelFunction(Name = "VietHoaDauCau", Description = "Viết hoa đầu câu")]
        public static string VietHoaDauCau(string chuoi)
        {
            if (string.IsNullOrEmpty(chuoi)) return chuoi;
            return CapitalizeFirst(chuoi);
        }

        [ExcelFunction(Name = "TachHoTen", Description = "Tách họ tên (Họ; Tên đệm; Tên) - trả về 'Ho|Dem|Ten'")]
        public static string TachHoTen(string hoTen)
        {
            if (string.IsNullOrWhiteSpace(hoTen)) return "||";
            var parts = hoTen.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1) return $"| |{parts[0]}";
            if (parts.Length == 2) return $"{parts[0]}||{parts[1]}";

            string ho = parts[0];
            string ten = parts[parts.Length - 1];
            string dem = string.Join(" ", parts.Skip(1).Take(parts.Length - 2));
            return $"{ho}|{dem}|{ten}";
        }

        // -------------------------
        // Helper methods
        // -------------------------

        private static bool IsMissing(object o)
        {
            if (o == null) return true;
            // ExcelDna provides ExcelMissing/ExcelEmpty singletons; handle common missing markers
            try
            {
                if (o == ExcelDna.Integration.ExcelMissing.Value) return true;
            }
            catch { }
            try
            {
                if (o == ExcelDna.Integration.ExcelEmpty.Value) return true;
            }
            catch { }
            if (o == System.Type.Missing) return true;
            if (o is DBNull) return true;
            return false;
        }

        private static bool TryGetDouble(object o, out double value)
        {
            value = 0;

            var scalar = GetScalarFromExcelArg(o);
            if (scalar == null) return false;

            // Numeric primitives
            switch (scalar)
            {
                case double d: value = d; return true;
                case float f: value = f; return true;
                case int i: value = i; return true;
                case long l: value = l; return true;
                case decimal m: value = (double)m; return true;
                case bool b: value = b ? 1 : 0; return true;
            }

            // String parse: supports "123", 1,234.56, 1.234,56, spaces...
            if (scalar is string s0)
            {
                var s = s0.Trim();

                // Remove wrapping quotes: "123" -> 123
                if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
                    s = s.Substring(1, s.Length - 2).Trim();

                // If user provided an explicit empty text (""), treat as zero
                if (string.IsNullOrWhiteSpace(s))
                {
                    value = 0;
                    return true;
                }

                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out value)) return true;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;

                // normalize thousand/decimal separators
                int lastDot = s.LastIndexOf('.');
                int lastComma = s.LastIndexOf(',');

                s = s.Replace(" ", "");
                if (lastDot >= 0 && lastComma >= 0)
                {
                    if (lastDot > lastComma) s = s.Replace(",", "");            // 1,234.56
                    else s = s.Replace(".", "").Replace(",", ".");              // 1.234,56
                }
                else
                {
                    // if only comma exists, treat as decimal separator in many VN locales
                    if (lastComma >= 0 && lastDot < 0) s = s.Replace(",", ".");
                }

                return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
            }

            try
            {
                value = Convert.ToDouble(scalar, CultureInfo.CurrentCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }


        private static bool ToBool(object o, bool defaultValue)
        {
            if (o == null) return defaultValue;
            if (o is bool) return (bool)o;
            if (o is int) return (int)o != 0;
            if (o is double) return Math.Abs((double)o) > 0.000001;
            var s = o.ToString().Trim().ToLowerInvariant();
            if (s == "true" || s == "1") return true;
            if (s == "false" || s == "0") return false;
            return defaultValue;
        }

        private static object GetScalarFromExcelArg(object o)
        {
            if (o == null) return null;

            // Excel-DNA missing/empty
            if (ReferenceEquals(o, ExcelMissing.Value) || ReferenceEquals(o, ExcelEmpty.Value))
                return null;

            // Excel error
            if (o is ExcelError) return null;

            // If it's a reference to a cell/range -> get the value of the top-left cell
            if (o is ExcelReference r)
            {
                try
                {
                    // xlfGetCell with 5 typically returns the value of the cell
                    // (works well for single cell; for range we take top-left)
                    var v = XlCall.Excel(XlCall.xlfGetCell, 5, r);
                    return v;
                }
                catch
                {
                    // fallback: try evaluate
                    try { return XlCall.Excel(XlCall.xlfEvaluate, r); }
                    catch { return null; }
                }
            }

            // If a range is passed as array, take top-left using array lower bounds
            if (o is object[,] a2)
            {
                if (a2.Length == 0) return null;
                int r0 = a2.GetLowerBound(0);
                int c0 = a2.GetLowerBound(1);
                var v = a2[r0, c0];
                return v;
            }

            if (o is object[] a1)
                return a1.Length > 0 ? a1[0] : null;

            return o;
        }


        private static string CapitalizeFirst(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            try
            {
                var vi = CultureInfo.GetCultureInfo("vi-VN");
                var first = s.Substring(0, 1);
                var rest = s.Length > 1 ? s.Substring(1) : string.Empty;
                return vi.TextInfo.ToUpper(first) + rest;
            }
            catch
            {
                return char.ToUpper(s[0]) + s.Substring(1);
            }
        }

        private static string DecimalPartToWords(int decimalPart)
        {
            // decimalPart is assumed 0..99, produce words for each digit (for "phẩy" representation)
            if (decimalPart == 0) return "";
            if (decimalPart < 10) return NumberToVietnamese(decimalPart);
            return NumberToVietnamese(decimalPart);
        }

        private static string[] unitNames = { "", "nghìn", "triệu", "tỷ" };

        // Ensure small numbers are handled immediately
        private static string NumberToVietnamese(long number)
        {
            if (number == 0) return "không";
            if (number < 0) return "âm " + NumberToVietnamese(Math.Abs(number));

            // Quick path for numbers < 1000
            if (number < 1000)
            {
                return ConvertBelowOneThousand((int)number);
            }

            var sb = new StringBuilder();

            long billions = number / 1000000000;
            long millions = (number % 1000000000) / 1000000;
            long thousands = (number % 1000000) / 1000;
            int rest = (int)(number % 1000);

            if (billions > 0)
            {
                sb.Append(ConvertBelowOneThousand((int)billions));
                sb.Append(" tỷ");
            }
            if (millions > 0)
            {
                if (sb.Length > 0) sb.Append(" ");
                sb.Append(ConvertBelowOneThousand((int)millions));
                sb.Append(" triệu");
            }
            if (thousands > 0)
            {
                if (sb.Length > 0) sb.Append(" ");
                sb.Append(ConvertBelowOneThousand((int)thousands));
                sb.Append(" nghìn");
            }
            if (rest > 0)
            {
                if (sb.Length > 0) sb.Append(" ");
                sb.Append(ConvertBelowOneThousand(rest));
            }

            var result = sb.ToString().Replace("  ", " ").Trim();
            if (string.IsNullOrWhiteSpace(result))
                return ConvertBelowOneThousand((int)number % 1000);
            return result;
        }

        private static string ConvertBelowOneThousand(int num)
        {
            // num in 0..999
            if (num == 0) return "không";
            int hundreds = num / 100;
            int tens = (num % 100) / 10;
            int units = num % 10;
            var parts = new List<string>();
            if (hundreds > 0)
            {
                parts.Add(DigitToWord(hundreds) + " trăm");
                if (tens == 0 && units > 0)
                    parts.Add("lẻ");
            }
            if (tens > 0)
            {
                if (tens == 1)
                    parts.Add("mười");
                else
                    parts.Add(DigitToWord(tens) + " mươi");
            }
            if (units > 0)
            {
                string unitWord = UnitWordForPosition(units, tens);
                parts.Add(unitWord);
            }
            return string.Join(" ", parts).Replace("  ", " ").Trim();
        }

        private static string DigitToWord(int d)
        {
            switch (d)
            {
                case 0: return "không";
                case 1: return "một";
                case 2: return "hai";
                case 3: return "ba";
                case 4: return "bốn";
                case 5: return "năm";
                case 6: return "sáu";
                case 7: return "bảy";
                case 8: return "tám";
                case 9: return "chín";
                default: return "";
            }
        }

        private static string UnitWordForPosition(int unit, int tens)
        {
            if (unit == 1)
            {
                if (tens == 0 || tens == 1) return "một";
                return "mốt"; // 21, 31 => "mốt"
            }
            if (unit == 5)
            {
                if (tens == 0) return "năm";
                return "lăm"; // 15, 25 => "lăm"
            }
            return DigitToWord(unit);
        }

        private static string NumberToEnglish(long number)
        {
            if (number == 0) return "zero";
            if (number < 0) return "minus " + NumberToEnglish(Math.Abs(number));

            var parts = new List<string>();
            var scales = new[] { "", "thousand", "million", "billion", "trillion" };
            int scale = 0;
            while (number > 0)
            {
                int chunk = (int)(number % 1000);
                if (chunk > 0)
                {
                    var chunkWords = ChunkToEnglishFast(chunk);
                    if (!string.IsNullOrEmpty(scales[scale])) chunkWords += " " + scales[scale];
                    parts.Insert(0, chunkWords);
                }
                number /= 1000;
                scale++;
            }
            return string.Join(" ", parts).Trim();
        }

        private static string ChunkToEnglishFast(int n)
        {
            // n in 1..999
            var below20 = new[] { "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen" };
            var tens = new[] { "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

            var sb = new StringBuilder();
            if (n >= 100)
            {
                sb.Append(below20[n / 100] + " hundred");
                n %= 100;
                if (n > 0) sb.Append(" ");
            }
            if (n >= 20)
            {
                sb.Append(tens[n / 10]);
                if (n % 10 > 0) sb.Append("-" + below20[n % 10]);
            }
            else if (n > 0)
            {
                sb.Append(below20[n]);
            }
            return sb.ToString();
        }

        // Keep old ChunkToEnglish name for compatibility if referenced elsewhere
        private static string ChunkToEnglish(int n)
        {
            return ChunkToEnglishFast(n);
        }

        private static string ConvertUnicodeToVNI(string input)
        {
            // Basic mapping for common Vietnamese vowels -> VNI-style sequences.
            // This is a simplified mapping and may not cover all cases.

            // Use a simple replace table to convert Unicode diacritics to VNI-style sequences.
            var replaces = new Dictionary<string, string>
            {
                {"à","a`"},{"á","a'"},{"ả","a?"},{"ã","a~"},{"ạ","a."},
                {"ă","aw"},{"ằ","aw`"},{"ắ","aw'"},{"ẳ","aw?"},{"ẵ","aw~"},{"ặ","aw."},
                {"â","aa"},{"ầ","aa`"},{"ấ","aa'"},{"ẩ","aa?"},{"ẫ","aa~"},{"ậ","aa."},

                {"è","e`"},{"é","e'"},{"ẻ","e?"},{"ẽ","e~"},{"ẹ","e."},
                {"ê","ee"},{"ề","ee`"},{"ế","ee'"},{"ể","ee?"},{"ễ","ee~"},{"ệ","ee."},

                {"ì","i`"},{"í","i'"},{"ỉ","i?"},{"ĩ","i~"},{"ị","i."},

                {"ò","o`"},{"ó","o'"},{"ỏ","o?"},{"õ","o~"},{"ọ","o."},
                {"ô","oo"},{"ồ","oo`"},{"ố","oo'"},{"ổ","oo?"},{"ỗ","oo~"},{"ộ","oo."},
                {"ơ","ow"},{"ờ","ow`"},{"ớ","ow'"},{"ở","ow?"},{"ỡ","ow~"},{"ợ","ow."},

                {"ù","u`"},{"ú","u'"},{"ủ","u?"},{"ũ","u~"},{"ụ","u."},
                {"ư","uw"},{"ừ","uw`"},{"ứ","uw'"},{"ử","uw?"},{"ữ","uw~"},{"ự","uw."},

                {"ỳ","y`"},{"ý","y'"},{"ỷ","y?"},{"ỹ","y~"},{"ỵ","y."},

                // Uppercase
                {"À","A`"},{"Á","A'"},{"Ả","A?"},{"Ã","A~"},{"Ạ","A."},
                {"Ă","AW"},{"Ằ","AW`"},{"Ắ","AW'"},{"Ẳ","AW?"},{"Ẵ","AW~"},{"Ặ","AW."},
                {"Â","AA"},{"Ầ","AA`"},{"Ấ","AA'"},{"Ẩ","AA?"},{"Ẫ","AA~"},{"Ậ","AA."},

                {"È","E`"},{"É","E'"},{"Ẻ","E?"},{"Ẽ","E~"},{"Ẹ","E."},
                {"Ê","EE"},{"Ề","EE`"},{"Ế","EE'"},{"Ể","EE?"},{"Ễ","EE~"},{"Ệ","EE."},

                {"Ì","I`"},{"Í","I'"},{"Ỉ","I?"},{"Ĩ","I~"},{"Ị","I."},

                {"Ò","O`"},{"Ó","O'"},{"Ỏ","O?"},{"Õ","O~"},{"Ọ","O."},
                {"Ô","OO"},{"ồ","oo`"},{"ố","oo'"},{"ổ","oo?"},{"ỗ","oo~"},{"ộ","oo."},
                {"Ơ","OW"},{"Ờ","OW`"},{"Ớ","OW'"},{"Ở","OW?"},{"Ỏ","OW~"},{"Ợ","OW."},

                {"Ù","U`"},{"Ú","U'"},{"Ủ","U?"},{"Ũ","U~"},{"Ụ","U."},
                {"Ư","UW"},{"Ừ","UW`"},{"Ứ","UW'"},{"Ử","UW?"},{"Ữ","UW~"},{"Ự","UW."},

                {"Ỳ","Y`"},{"Ý","Y'"},{"Ỷ","Y?"},{"Ỹ","Y~"},{"Ỵ","Y."}
            };

            var sb = new StringBuilder(input);
            foreach (var kv in replaces)
            {
                sb.Replace(kv.Key, kv.Value);
            }
            return sb.ToString();
        }

        private static bool TryParseVietnameseNumber(string input, out long value)
        {
            // Very simplified parser: handles numbers up to trillions for common forms.
            // Supports words without diacritics like "mot", "hai", "trieu", "ty", "nghin", "tram", "muoi".
            value = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;

            var tokens = input.ToLowerInvariant().Replace("-", " ").Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            long current = 0;
            long total = 0;
            bool negative = false;

            foreach (var t in tokens)
            {
                if (t == "âm" || t == "am") { negative = true; continue; }

                if (WordToNumber.ContainsKey(t))
                {
                    current += WordToNumber[t];
                }
                else if (t == "mươi" || t == "muoi" )
                {
                    // treat as multiply by 10 when appropriate
                    if (current == 0) current = 1;
                    current *= 10;
                }
                else if (t == "trăm" || t == "tram")
                {
                    if (current == 0) current = 1;
                    current *= 100;
                }
                else if (t == "nghìn" || t == "nghin")
                {
                    total += current * 1000;
                    current = 0;
                }
                else if (t == "triệu" || t == "trieu")
                {
                    total += current * 1000000;
                    current = 0;
                }
                else if (t == "tỷ" || t == "ty")
                {
                    total += current * 1000000000;
                    current = 0;
                }
                else if (t == "chẵn" || t == "chan")
                {
                    // ignore
                }
                else if (t == "phẩy" || t == "phay")
                {
                    // stop at decimal point for this simple parser
                    break;
                }
                else
                {
                    // unknown token - ignore for now
                }
            }

            total += current;
            value = negative ? -total : total;
            return true;
        }

        private static readonly Dictionary<string, long> WordToNumber = new Dictionary<string, long>
        {
            {"khong",0},{"không",0},{"mot",1},{"một",1},{"hai",2},{"ba",3},{"bon",4},{"bốn",4},{"nam",5},{"năm",5},
            {"sau",6},{"sáu",6},{"bay",7},{"bảy",7},{"tam",8},{"tám",8},{"chin",9},{"chín",9},
            {"muoi",10},{"mười",10},{"mươi",10},{"tram",100},{"trăm",100},
            {"nghin",1000},{"nghìn",1000},{"trieu",1000000},{"triệu",1000000},{"ty",1000000000},{"tỷ",1000000000}
        };

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            
        }
        
        #endregion
    }
}
