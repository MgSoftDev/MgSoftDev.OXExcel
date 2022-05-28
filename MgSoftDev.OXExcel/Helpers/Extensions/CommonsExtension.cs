using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.OpenXmlProvider;

namespace MgSoftDev.OXExcel.Helpers.Extensions
{
    public static class CommonsExtension
    {
        internal static OxCellTypeValues ToCellTypeValues(this Type value )
        {
            if(value == null) return OxCellTypeValues.String;
            if(value == typeof(string)) return OxCellTypeValues.String;
            if(value == typeof(bool) || value == typeof(bool?)) return OxCellTypeValues.String;
            return OxCellTypeValues.Number;
        }

        internal static string GetPropertiName(this MemberExpression value)
        {
            if (value == null) return "";
            var res = "";
            res += value.Member.Name;
            var nextLevel =value.Expression as MemberExpression;
            if (nextLevel == null) return res;
            if(nextLevel.NodeType== ExpressionType.MemberAccess)
                res = nextLevel.GetPropertiName() + "."+ res;
            return res;
        }

        public static T GetAttribute<T>(this ICustomAttributeProvider provider)
            where T : Attribute => (T)provider.GetCustomAttributes(typeof(T), true).FirstOrDefault();

        ///<summary>Cast of methods Anonimus Ej: GrottyHacks.Cast(weaklyTyped,new { Fruit="", Topping="" });</summary>
        public static T Cast<T>(object target, T example)
        {
            return (T)target;
        }

        internal static string ToExcelValue(this object value)
        {


            if (value == null)
                return "";
            if (value is DateTime)
                return ((DateTime) value).ToOADate().ToString(CultureInfo.InvariantCulture);
            if (value is TimeSpan)
                return new DateTime(((TimeSpan)value).Ticks).ToOADate().ToString(CultureInfo.InvariantCulture);
            if (value is string)
                return ((string) value).ToString(Const.CultureData);
            if (value is bool)
                return ((bool) value).ToString(Const.CultureData);

            if (value is short)
                return ((short) value).ToString(Const.CultureData);
            if (value is int)
                return ((int) value).ToString(Const.CultureData);
            if (value is long)
                return ((long) value).ToString(Const.CultureData);

            if (value is ushort)
                return ((ushort) value).ToString(Const.CultureData);
            if (value is uint)
                return ((uint)value).ToString(Const.CultureData);
            if (value is ulong)
                return ((ulong)value).ToString(Const.CultureData);

            if (value is double)
                return ((double) value).ToString(Const.CultureData);
            if (value is float)
                return ((float) value).ToString(Const.CultureData);
            if (value is decimal)
                return ((decimal) value).ToString(Const.CultureData);

            return value.ToString();

        }

        public static T Clone<T>(this T source)
        {
            if (!typeof(T).IsSerializable)
            {
                throw new ArgumentException("The type must be serializable.", "source");
            }

            // Don't serialize a null object, simply return the default for that object
            if (ReferenceEquals(source, null))
            {
                return default(T);
            }

            IFormatter formatter = new BinaryFormatter();
            Stream stream = new MemoryStream();
            using (stream)
            {
                formatter.Serialize(stream, source);
                stream.Seek(0, SeekOrigin.Begin);
                return (T)formatter.Deserialize(stream);
            }
        }

        public static string GetPropertyVal(this object value, string propertiPath, object defaultValue = null)
        {
            var path = propertiPath.Split('.').ToList();
            var val = value;
            object obj = null;
            foreach (var pt in path)
            {
                obj = null;
                var p = val?.GetType().GetProperties().FirstOrDefault(f => f.Name == pt);
                if (p == null) break;
                var pObject = val.GetType().InvokeMember(p.Name, BindingFlags.GetProperty, null, val, null);
                val = pObject;
                obj = pObject;
            }
            return obj == null ? defaultValue.ToExcelValue() : obj.ToExcelValue();
        }

        public static List<PropertyInfo> GetProperties(this object entity) => entity?.GetType().GetProperties().ToList() ?? new List<PropertyInfo>();


        internal static OxCellFormartEntity Combine(this OxCellFormartEntity value, OxCellFormartEntity value2)
        {
            if (value == null) return value2;
            if (value2 == null) return value;
            value.Borders = value.Borders ?? value2.Borders;
            if (value.Borders != null)
            {
                value.Borders.Bottom = value.Borders.Bottom ?? value2.Borders.Bottom;
                value.Borders.Diagonal = value.Borders.Diagonal ?? value2.Borders.Diagonal;
                value.Borders.Left = value.Borders.Left ?? value2.Borders.Left;
                value.Borders.Right = value.Borders.Right ?? value2.Borders.Right;
                value.Borders.Top = value.Borders.Top ?? value2.Borders.Top;
            }
            value.Alignment = value.Alignment ?? value2.Alignment;
            value.Fill = value.Fill ?? value2.Fill;
            if (value.Fill != null)
            {
                value.Fill.GradientFill = value.Fill.GradientFill ?? value2.Fill.GradientFill;
                value.Fill.PatternFill = value.Fill.PatternFill ?? value2.Fill.PatternFill;
            }
            value.Font = value.Font ?? value2.Font;
            value.NumberFormat = value.NumberFormat ?? value2.NumberFormat;
           
            return value;
        }
    }
}

