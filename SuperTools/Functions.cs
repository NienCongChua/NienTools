using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperTools
{
    public class Functions
    {


        #region REMOVEACCENT Function
        [ExcelFunction(Name = "REMOVEACCENT", Description = "Loại bỏ dấu tiếng Việt khỏi chuỗi")]
        public static string REMOVEACCENT(
            [ExcelArgument(Description ="Chuỗi cần loại bỏ dấu")] object input)
        {
            try
            {
                if (input == null) return "";
                string str = input.ToString();
                var normalizedString = str.Normalize(NormalizationForm.FormD);
                var stringBuilder = new StringBuilder(capacity: normalizedString.Length);
                foreach (var c in normalizedString)
                {
                    var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                    if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                    {
                        stringBuilder.Append(c);
                    }
                }
                var cleaned = stringBuilder.ToString().Normalize(NormalizationForm.FormC).Trim();
                cleaned = cleaned.Replace('đ', 'd').Replace('Đ', 'D');
                return cleaned;
            }
            catch
            {
                return "#VALUE!";
            }
        }
        #endregion

        #region VND Function
        [ExcelFunction(Name = "VND", Description = "Chuyển số thành chữ tiền Việt (VND)")]
        public static string VND(
            [ExcelArgument(Description = "Số cần chuyển đổi (bắt buộc)")] object number,
            [ExcelArgument(Description = "Đọc tiền hay đọc số (mặc định: true <Đọc tiền>)")] object showDong = null,
            [ExcelArgument(Description = "Sử dụng đơn vị mặc định (nghìn/ngàn)? (mặc định: true <Nghìn>)")] object dv = null)
        {
            try
            {
                if (number == null) return "";
                double value;
                if (!Helper.TryGetDouble(number, out value)) return "#VALUE!";
                bool coDonVi = Helper.IsMissing(showDong) ? true : Helper.ToBool(showDong, true);
                bool useDefaultDv = Helper.IsMissing(dv) ? true : Helper.ToBool(dv, true);
                bool vietHoa = true;
                string donViTien = "đồng";
                string donViXu = "xu";
                string sign = value < 0 ? "âm " : "";
                value = Math.Abs(value);
                long intPart = (long)Math.Floor(value);
                int decimalPart = (int)Math.Round((value - intPart) * 100, MidpointRounding.AwayFromZero);
                if (decimalPart >= 100)
                {
                    intPart += 1;
                    decimalPart -= 100;
                }
                string words;
                if (useDefaultDv)
                {
                    words = intPart == 0 ? "không" : Helper.NumberToVietnamese(intPart, "nghìn");
                }
                else
                {
                    words = intPart == 0 ? "không" : Helper.NumberToVietnamese(intPart, "ngàn");
                }
                if (string.IsNullOrWhiteSpace(words)) words = "không";
                if (coDonVi)
                {
                    if (decimalPart == 0)
                    {
                        words = $"{words} {donViTien} chẵn";
                    }
                    else
                    {
                        string xuWords = decimalPart == 0 ? "" : Helper.NumberToVietnamese(decimalPart);
                        words = $"{words} {donViTien} {xuWords} {donViXu}";
                    }
                }
                else
                {
                    if (decimalPart > 0)
                    {
                        words = $"{words} phẩy {Helper.DecimalPartToWords(decimalPart)}";
                    }
                }
                words = sign + words;
                if (vietHoa)
                {
                    words = Helper.CapitalizeFirst(words);
                }
                return words + '.';
            }
            catch
            {
                return "#ERROR";
            }
        }
        #endregion

        #region USD Function
        [ExcelFunction(Name = "USD", Description = "Chuyển số thành chữ tiếng Anh (USD)")]
        public static string USD(
            [ExcelArgument(Description = "Số cần chuyển đổi (bắt buộc)")] object number,
            [ExcelArgument(Description = "Đọc tiền hay đọc số (mặc định: true <Đọc tiền>)")] object showDollar = null)
        {
            try
            {
                if (number == null) return "";
                double value;
                if (!Helper.TryGetDouble(number, out value)) return "#VALUE!";
                bool coDollar = Helper.IsMissing(showDollar) ? true : Helper.ToBool(showDollar, true);
                string sign = value < 0 ? "Minus " : "";
                value = Math.Abs(value);
                long dollars = (long)Math.Floor(value);
                int cents = (int)Math.Round((value - dollars) * 100, MidpointRounding.AwayFromZero);
                if (cents >= 100)
                {
                    dollars += 1;
                    cents -= 100;
                }
                string words = Helper.NumberToEnglish(dollars);
                if (coDollar)
                {
                    words = $"{words} {(dollars == 1 ? "dollar" : "dollars")}";
                }
                if (cents > 0)
                {
                    var centWords = Helper.NumberToEnglish(cents);
                    if (coDollar)
                    {
                        words += $" and {centWords} {(cents == 1 ? "cent" : "cents")}";
                    }
                    else
                    {
                        if (dollars > 0)
                            words += $" and {centWords}";
                        else
                            words = centWords;
                    }
                }
                words = Helper.CapitalizeFirst((sign + words).Trim());
                return words + '.';
            }
            catch
            {
                return "#ERROR";
            }
        }
        #endregion
    }
}
