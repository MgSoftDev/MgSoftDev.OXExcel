namespace MgSoftDev.OXExcel.Helpers.Extensions
{
    public static class ColorHelper
    {
        public static string ToHexFormat(this System.Drawing.Color color )
        {
            var r = $"{color.R:x}".Length == 1 ? "0" + $"{color.R:x}" : $"{color.R:x}";
            var g = $"{color.G:x}".Length == 1 ? "0" + $"{color.G:x}" : $"{color.G:x}";
            var b = $"{color.B:x}".Length == 1 ? "0" + $"{color.B:x}" : $"{color.B:x}";
            var a = $"{color.A:x}".Length == 1 ? "0" + $"{color.A:x}" : $"{color.A:x}";
            return a+r+g+b;
        }
    }
}
