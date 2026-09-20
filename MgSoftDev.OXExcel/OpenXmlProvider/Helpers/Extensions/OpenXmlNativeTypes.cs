using DocumentFormat.OpenXml;

namespace MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions
{
    internal static class OpenXmlNativeTypes
    {
        internal static BooleanValue ToBooleanValue(this bool value) =>  BooleanValue.FromBoolean(value);

        /// <summary>Zoom de una vista: el esquema lo admite entre 10 y 400, fuera de ahí el atributo no se escribe.</summary>
        internal static UInt32Value ToZoomValue(this uint value) => value >= 10 && value <= 400 ? UInt32Value.FromUInt32(value) : null;
        // envazar encapsular englobar  
        internal static List<T> CreateList<T>(this T value) => new List<T>() {value};

        internal static UInt32Value ToUInt32Value(this long value) => UInt32Value.FromUInt32((uint) value);
        internal static UInt32Value ToUInt32Value(this uint value) => UInt32Value.FromUInt32(value);
        internal static UInt32Value ToUInt32Value(this int value) => UInt32Value.FromUInt32((uint)value);

        
    }
}
