using System.Linq.Expressions;
using System.Reflection;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Table;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxTableColumnsFactory<T> 
    {
        internal readonly List<OxTableColumnsEntity> TableColumns;

        internal OxTableColumnsFactory()
        {
            TableColumns = new List<OxTableColumnsEntity>();
        }

        public OxTableColumnFactory Add<T2>( Expression<Func<T, T2>> propertyLambda) 
        {
            var me = propertyLambda.Body as MemberExpression;
            if (me == null)
            {
                throw new ArgumentException(
                    "You must pass a lambda of the form: '() => Class.Property' or '() => object.Property'");
            }
            
            var c = new OxTableColumnFactory(me.GetPropertiName(),me.Type,me.Type.ToCellTypeValues(), TableColumns.Count + 1);
            c.Header(me.GetPropertiName());

            c.ExtractAttributes( me.Member as PropertyInfo );
            

            TableColumns.Add(c.TableColumn);
            return c;
        }

        public OxTableColumnFactory Add(string propertyName)
        {
            var c = new OxTableColumnFactory(propertyName, typeof(string), OxCellTypeValues.String, TableColumns.Count +1);
            c.Header(propertyName);
            TableColumns.Add(c.TableColumn);
            return c;
        }
    }

    public class OxTableColumnsFactory
    {
        internal readonly List<OxTableColumnsEntity> TableColumns;

        internal OxTableColumnsFactory()
        {
            TableColumns = new List<OxTableColumnsEntity>();
        }

        public OxTableColumnFactory Add( string propertyName )
        {
            var c = new OxTableColumnFactory( propertyName, typeof( string ), OxCellTypeValues.String, TableColumns.Count +1 );
            c.Header( propertyName );
            TableColumns.Add( c.TableColumn );

            return c;
        }
    }

}
