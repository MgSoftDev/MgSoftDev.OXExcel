using System.Text.RegularExpressions;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Helpers.Extensions
{
    internal static class ExcelRefence
    {
        internal static string ToReferenceAlfa(this uint value)
        {
            var aaNum = 26;
            var doblesAa = value % aaNum == 0 ? (value / aaNum) - 1 : value / aaNum;
            var aa = string.Empty;
            if (doblesAa > 0)
                aa = Convert.ToChar(64 + doblesAa).ToString();
            var singleA = value - (doblesAa * aaNum);
            var a = Convert.ToChar(64 + singleA).ToString();

            return aa + a;
        }

        internal static uint ToColIndex(this string value )
        {
            value = value.ToUpper();
            var anum = Convert.ToUInt32(value[0]) - 64;
            var result = (uint)(value.Length == 1 ? anum : (anum * 27) + Convert.ToInt16(value[1]) - 65);
            return result;
        }

        internal static uint GetRow(this string excelDeference)
        {//A5   XY99  ZDC125545
            var reg = Regex.Match(excelDeference, @"\d*\Z");
            if (reg.Success && reg.Value != "")
                return Convert.ToUInt32(reg.Value);
            return 1;
        }
        internal static uint GetCol(this string excelDeference)
        {//A5   XY99  ZDC125545
            var reg = Regex.Match(excelDeference, @"^[A-Z a-z]*");
            if (reg.Success && reg.Value != "")
                return reg.Value.ToColIndex();
            return 1;
        }

        internal static OxRangeEntity ToRange(this string value)
        {
            if(string.IsNullOrEmpty(value) || !value.Contains(":")) return null;
            var sp = value.Split(new[] {":"}, StringSplitOptions.RemoveEmptyEntries).ToList();
            return sp.Count != 2 ? null : new OxRangeEntity(sp[0].GetCol(), sp[0].GetRow(), sp[1].GetCol(), sp[1].GetRow());
        }

    }
}
