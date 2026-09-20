using System.Collections.Concurrent;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.OpenXmlProvider;
using Newtonsoft.Json;

namespace MgSoftDev.OXExcel.Helpers.Extensions
{
    public static class CommonsExtension
    {
        internal static OxCellTypeValues ToCellTypeValues(this Type value)
        {
            if (value == null) return OxCellTypeValues.String;
            if (value == typeof( string )) return OxCellTypeValues.String;
            if (value == typeof( bool ) || value == typeof( bool? )) return OxCellTypeValues.String;

            return OxCellTypeValues.Number;
        }

        internal static string GetPropertiName(this MemberExpression value)
        {
            if (value == null) return "";

            var res = "";
            res += value.Member.Name;
            var nextLevel = value.Expression as MemberExpression;

            if (nextLevel == null) return res;

            if (nextLevel.NodeType == ExpressionType.MemberAccess) res = nextLevel.GetPropertiName() + "." + res;

            return res;
        }

        public static T GetAttribute< T >(this ICustomAttributeProvider provider)
            where T : Attribute=>
            ( T ) provider.GetCustomAttributes(typeof( T ), true).FirstOrDefault();

        ///<summary>Cast of methods Anonimus Ej: GrottyHacks.Cast(weaklyTyped,new { Fruit="", Topping="" });</summary>
        public static T Cast< T >(object target, T example) { return ( T ) target; }

        internal static string ToExcelValue(this object value)
        {
            if (value == null) return "";
            if (value is DateTime) return ( ( DateTime ) value ).ToOADate().ToString(CultureInfo.InvariantCulture);
            if (value is TimeSpan) return new DateTime(( ( TimeSpan ) value ).Ticks).ToOADate().ToString(CultureInfo.InvariantCulture);
            if (value is string) return ( ( string ) value ).ToString(Const.CultureData);
            if (value is bool) return ( ( bool ) value ).ToString(Const.CultureData);

            if (value is short) return ( ( short ) value ).ToString(Const.CultureData);
            if (value is int) return ( ( int ) value ).ToString(Const.CultureData);
            if (value is long) return ( ( long ) value ).ToString(Const.CultureData);

            if (value is ushort) return ( ( ushort ) value ).ToString(Const.CultureData);
            if (value is uint) return ( ( uint ) value ).ToString(Const.CultureData);
            if (value is ulong) return ( ( ulong ) value ).ToString(Const.CultureData);

            if (value is double) return ( ( double ) value ).ToString(Const.CultureData);
            if (value is float) return ( ( float ) value ).ToString(Const.CultureData);
            if (value is decimal) return ( ( decimal ) value ).ToString(Const.CultureData);

            return value.ToString();
        }

        //public static T Clone<T>(this T source)
        //{
        //    if (!typeof(T).IsSerializable)
        //    {
        //        throw new ArgumentException("The type must be serializable.", "source");
        //    }

        //    // Don't serialize a null object, simply return the default for that object
        //    if (ReferenceEquals(source, null))
        //    {
        //        return default(T);
        //    }

        //    IFormatter formatter = new BinaryFormatter();
        //    Stream stream = new MemoryStream();
        //    using (stream)
        //    {
        //        formatter.Serialize(stream, source);
        //        stream.Seek(0, SeekOrigin.Begin);
        //        return (T)formatter.Deserialize(stream);
        //    }
        //}
        public static T Clone< T >(this T source)
        {
            // Don't serialize a null object, simply return the default for that object
            if (Object.ReferenceEquals(source, null))
            {
                return default( T );
            }

            var deserializeSettings = new JsonSerializerSettings {ObjectCreationHandling = ObjectCreationHandling.Replace};
            var serializeSettings   = new JsonSerializerSettings {ReferenceLoopHandling  = ReferenceLoopHandling.Ignore};

            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(source, serializeSettings), deserializeSettings);
        }

        public static string GetPropertyVal(this object value, string propertiPath, object defaultValue = null)
        {
            var    path = propertiPath.Split('.');
            var    val  = value;
            object obj  = null;

            foreach ( var pt in path )
            {
                obj = null;
                if (val == null) break;

                var getter = GetPropertyGetter(val.GetType(), pt);

                if (getter == null) break;

                var pObject = getter(val);
                val = pObject;
                obj = pObject;
            }

            return obj == null ? defaultValue.ToExcelValue() : obj.ToExcelValue();
        }

        /// <summary>
        /// Lectores de propiedad compilados y cacheados por tipo. Antes cada celda hacía GetProperties() más
        /// InvokeMember, que es el camino de reflexión más lento y asigna un arreglo nuevo en cada llamada.
        /// </summary>
        private static readonly ConcurrentDictionary<(Type Type, string Property), Func<object, object>> _PropertyGetters =
            new ConcurrentDictionary<(Type Type, string Property), Func<object, object>>();

        private static Func<object, object> GetPropertyGetter(Type type, string property)
        {
            return _PropertyGetters.GetOrAdd(( type, property ), key =>
            {
                var info = key.Type.GetProperties().FirstOrDefault(f => f.Name == key.Property);

                if (info == null || !info.CanRead) return null;

                var parameter = Expression.Parameter(typeof( object ), "instance");
                var body      = Expression.Convert(Expression.Property(Expression.Convert(parameter, key.Type), info), typeof( object ));

                return Expression.Lambda<Func<object, object>>(body, parameter).Compile();
            });
        }

        public static List<PropertyInfo> GetProperties(this object entity)=>entity?.GetType().GetProperties().ToList() ?? new List<PropertyInfo>();


        /// <summary>
        /// Copia de un formato sin pasar por JSON. Solo se duplican las partes que <see cref="Combine"/> modifica en
        /// sitio (Borders y Fill); Font, Alignment y NumberFormat se reemplazan completos, así que se comparten.
        /// Pensado para el camino de tablas, donde se ejecuta una vez por celda.
        /// </summary>
        internal static OxCellFormartEntity CloneFast(this OxCellFormartEntity value)
        {
            if (value == null) return null;

            return new OxCellFormartEntity
            {
                NumberFormat = value.NumberFormat,
                Font         = value.Font,
                Alignment    = value.Alignment,
                Fill         = value.Fill == null
                                   ? null
                                   : new OxFillEntity { PatternFill = value.Fill.PatternFill, GradientFill = value.Fill.GradientFill },
                Borders = value.Borders == null
                              ? null
                              : new OxBorderEntity
                              {
                                  Bottom       = value.Borders.Bottom,
                                  Top          = value.Borders.Top,
                                  Right        = value.Borders.Right,
                                  Left         = value.Borders.Left,
                                  Diagonal     = value.Borders.Diagonal,
                                  DiagonalDown = value.Borders.DiagonalDown,
                                  DiagonalUp   = value.Borders.DiagonalUp,
                                  Outline      = value.Borders.Outline
                              }
            };
        }

        /// <summary>Copia de una definición de fila sin pasar por JSON; el formato se comparte porque no se modifica.</summary>
        internal static OxRowEntity CloneFast(this OxRowEntity value)
        {
            if (value == null) return null;

            return new OxRowEntity
            {
                Collapsed    = value.Collapsed,
                CustomFormat = value.CustomFormat,
                CustomHeight = value.CustomHeight,
                Height       = value.Height,
                Hidden       = value.Hidden,
                OutlineLevel = value.OutlineLevel,
                RowIndex     = value.RowIndex,
                ShowPhonetic = value.ShowPhonetic,
                ThickBot     = value.ThickBot,
                ThickTop     = value.ThickTop,
                Format       = value.Format
            };
        }

        internal static OxCellFormartEntity Combine(this OxCellFormartEntity value, OxCellFormartEntity value2)
        {
            if (value  == null) return value2;
            if (value2 == null) return value;

            value.Borders = value.Borders ?? value2.Borders;

            if (value.Borders != null && value2.Borders != null)
            {
                value.Borders.Bottom   = value.Borders.Bottom   ?? value2.Borders.Bottom;
                value.Borders.Diagonal = value.Borders.Diagonal ?? value2.Borders.Diagonal;
                value.Borders.Left     = value.Borders.Left     ?? value2.Borders.Left;
                value.Borders.Right    = value.Borders.Right    ?? value2.Borders.Right;
                value.Borders.Top      = value.Borders.Top      ?? value2.Borders.Top;
            }

            value.Alignment = value.Alignment ?? value2.Alignment;
            value.Fill      = value.Fill      ?? value2.Fill;

            if (value.Fill != null && value2.Fill != null)
            {
                value.Fill.GradientFill = value.Fill.GradientFill ?? value2.Fill.GradientFill;
                value.Fill.PatternFill  = value.Fill.PatternFill  ?? value2.Fill.PatternFill;
            }

            value.Font         = value.Font         ?? value2.Font;
            value.NumberFormat = value.NumberFormat ?? value2.NumberFormat;

            return value;
        }
    }
}
