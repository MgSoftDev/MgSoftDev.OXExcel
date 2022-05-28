using DocumentFormat.OpenXml;

namespace MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions
{
    internal static class OpenXmlNativeTypes
    {
        internal static BooleanValue ToBooleanValue(this bool value) =>  BooleanValue.FromBoolean(value);
        // envazar encapsular englobar  
        internal static List<T> CreateList<T>(this T value) => new List<T>() {value};

        internal static UInt32Value ToUInt32Value(this long value) => UInt32Value.FromUInt32((uint) value);
        internal static UInt32Value ToUInt32Value(this uint value) => UInt32Value.FromUInt32(value);
        internal static UInt32Value ToUInt32Value(this int value) => UInt32Value.FromUInt32((uint)value);

        
    }
}
