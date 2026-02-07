using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperTools
{
    internal class Helper : Functions
    {
        #region Helper mothods
        // -------------------------
        // Helper methods
        // -------------------------
        public static bool IsMissing(object o)
        {
            if (o == null) return true;
            try { if (o == ExcelDna.Integration.ExcelMissing.Value) return true; } catch { }
            try { if (o == ExcelDna.Integration.ExcelEmpty.Value) return true; } catch { }
            if (o == System.Type.Missing) return true;
            if (o is DBNull) return true;
            return false;
        }

        public static bool TryGetDouble(object o, out double value)
        {
            value = 0;
            var scalar = GetScalarFromExcelArg(o);
            if (scalar == null) return false;
            switch (scalar)
            {
                case double d: value = d; return true;
                case float f: value = f; return true;
                case int i: value = i; return true;
                case long l: value = l; return true;
                case decimal m: value = (double)m; return true;
                case bool b: value = b ? 1 : 0; return true;
            }
            if (scalar is string s0)
            {
                var s = s0.Trim();
                if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
                    s = s.Substring(1, s.Length - 2).Trim();
                if (string.IsNullOrWhiteSpace(s))
                {
                    value = 0;
                    return true;
                }
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out value)) return true;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;
                int lastDot = s.LastIndexOf('.');
                int lastComma = s.LastIndexOf(',');
                s = s.Replace(" ", "");
                if (lastDot >= 0 && lastComma >= 0)
                {
                    if (lastDot > lastComma) s = s.Replace(",", "");
                    else s = s.Replace(".", "").Replace(",", ".");
                }
                else
                {
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

        public static bool ToBool(object o, bool defaultValue)
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
            if (ReferenceEquals(o, ExcelMissing.Value) || ReferenceEquals(o, ExcelEmpty.Value))
                return null;
            if (o is ExcelError) return null;
            if (o is ExcelReference r)
            {
                try
                {
                    var v = XlCall.Excel(XlCall.xlfGetCell, 5, r);
                    return v;
                }
                catch
                {
                    try { return XlCall.Excel(XlCall.xlfEvaluate, r); }
                    catch { return null; }
                }
            }
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

        public static string CapitalizeFirst(string s)
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

        public static string DecimalPartToWords(int decimalPart)
        {
            if (decimalPart == 0) return "";
            if (decimalPart < 10) return NumberToVietnamese(decimalPart);
            return NumberToVietnamese(decimalPart);
        }

        public static string[] unitNames = { "", "nghìn", "triệu", "tỷ" };
        public static string NumberToVietnamese(long number, string thousandName = "nghìn")
        {
            if (number == 0) return "không";
            if (number < 0) return "âm " + NumberToVietnamese(Math.Abs(number), thousandName);
            // Định nghĩa các đơn vị lớn hơn
            string[] units = { "", thousandName, "triệu", "tỷ", "nghìn tỷ", "triệu tỷ", "tỷ tỷ" };
            List<string> parts = new List<string>();
            int unitIndex = 0;
            while (number > 0 && unitIndex < units.Length)
            {
                int chunk = (int)(number % 1000);
                if (chunk > 0)
                {
                    string chunkText = ConvertBelowOneThousand(chunk);
                    if (!string.IsNullOrEmpty(units[unitIndex]))
                        chunkText += " " + units[unitIndex];
                    parts.Insert(0, chunkText);
                }
                number /= 1000;
                unitIndex++;
            }
            // Nếu số còn lại lớn hơn 0, tiếp tục thêm "tỷ" cho mỗi 3 số
            while (number > 0)
            {
                int chunk = (int)(number % 1000);
                if (chunk > 0)
                {
                    string chunkText = ConvertBelowOneThousand(chunk) + " tỷ";
                    parts.Insert(0, chunkText);
                }
                number /= 1000;
            }
            var result = string.Join(" ", parts).Replace("  ", " ").Trim();
            if (string.IsNullOrWhiteSpace(result))
                return ConvertBelowOneThousand((int)number % 1000);
            return result;
        }

        public static string ConvertBelowOneThousand(int num)
        {
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
            return string.Join(" ", parts).Replace(" ", " ").Trim();
        }

        public static string DigitToWord(int d)
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

        public static string UnitWordForPosition(int unit, int tens)
        {
            if (unit == 1)
            {
                if (tens == 0 || tens == 1) return "một";
                return "mốt";
            }
            if (unit == 5)
            {
                if (tens == 0) return "năm";
                return "lăm";
            }
            return DigitToWord(unit);
        }

        public static string NumberToEnglish(long number)
        {
            if (number == 0) return "zero";
            if (number < 0) return "minus " + NumberToEnglish(Math.Abs(number));
            var parts = new List<string>();
            // Hỗ trợ đến sextillion (10^21), có thể mở rộng thêm nếu cần
            var scales = new[] { "", "thousand", "million", "billion", "trillion", "quadrillion", "quintillion", "sextillion" };
            int scale = 0;
            while (number > 0 && scale < scales.Length)
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
            // Nếu số còn lại lớn hơn 0, tiếp tục thêm "billion" cho mỗi 3 số
            while (number > 0)
            {
                int chunk = (int)(number % 1000);
                if (chunk > 0)
                {
                    var chunkWords = ChunkToEnglishFast(chunk) + " billion";
                    parts.Insert(0, chunkWords);
                }
                number /= 1000;
            }
            return string.Join(" ", parts).Trim();
        }

        public static string ChunkToEnglishFast(int n)
        {
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

        private static readonly Dictionary<string, long> WordToNumber = new Dictionary<string, long>
        {
            {"khong",0},{"không",0},{"mot",1},{"một",1},{"hai",2},{"ba",3},{"bon",4},{"bốn",4},{"nam",5},{"năm",5},
            {"sau",6},{"sáu",6},{"bay",7},{"bảy",7},{"tam",8},{"tám",8},{"chin",9},{"chín",9},
            {"muoi",10},{"mười",10},{"mươi",10},{"tram",100},{"trăm",100},
            {"ngan",1000},{"nghìn",1000},{"trieu",1000000},{"triệu",1000000},{"ty",1000000000},{"tỷ",1000000000}
        };
        #endregion
    }
}
