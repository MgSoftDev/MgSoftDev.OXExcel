using System.Data;
using System.Reflection;
using System.Reflection.Emit;
using MgSoftDev.OXExcel.Entities.Dynamic;

namespace MgSoftDev.OXExcel.Helpers.Extensions
{
    public static class  DataTableExtencion
    {
        public static List<DynamicEntity> ToDynamicList(this DataTable table)
        {
            var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(new AssemblyName("DynamicAssembly"), AssemblyBuilderAccess.Run);
            var moduleBuilder   = assemblyBuilder.DefineDynamicModule("Dynamic.dll");
            var typeBuilder     = moduleBuilder.DefineType(Guid.NewGuid().ToString());
            typeBuilder.SetParent(typeof(DynamicEntity));

            var i = 0;
            foreach (DataColumn col in table.Columns)
            {

                var propertyBuilder = typeBuilder.DefineProperty(col.ColumnName.ToCleanColumn(), System.Reflection.PropertyAttributes.None, col.GetDataType(), Type.EmptyTypes);

                var getMethodBuilder = typeBuilder.DefineMethod("get_" + col.ColumnName.ToCleanColumn(), MethodAttributes.Public, CallingConventions.HasThis, col.GetDataType(), Type.EmptyTypes);
                var getter = getMethodBuilder.GetILGenerator();
                getter.Emit(OpCodes.Ldarg_0);
                getter.Emit(OpCodes.Ldstr, col.ColumnName.ToCleanColumn());
                getter.Emit(OpCodes.Callvirt, typeof(DynamicEntity).GetMethod("Get", BindingFlags.Instance | BindingFlags.NonPublic).MakeGenericMethod(col.GetDataType()));
                getter.Emit(OpCodes.Ret);
                propertyBuilder.SetGetMethod(getMethodBuilder);
                i++;
            }
            var type =  typeBuilder.CreateTypeInfo().AsType();

            var result = new List<DynamicEntity>();
            foreach (DataRow row in table.Rows)
            {
                var data = new Dictionary<string, object>();
                for (i = 0; i < table.Columns.Count; i++)
                    data.Add(table.Columns[i].ColumnName.ToCleanColumn(), row[i]);
                var child = (DynamicEntity)Activator.CreateInstance(type);
                child.SetProperties( data);
                result.Add(child);
            }
            return result;
        }

        public static Type GetDataType(this DataColumn column)
        {
            var result = column.DataType;
            if (!column.AllowDBNull) return result;
            if (column.DataType == typeof(int))
                result = typeof(int?);
            else if (column.DataType == typeof(Int16))
                result = typeof(Int16?);
            else if (column.DataType == typeof(Int32))
                result = typeof(Int32?);
            else if (column.DataType == typeof(Int64))
                result = typeof(Int64?);
            else if (column.DataType == typeof(bool))
                result = typeof(bool?);
            else if (column.DataType == typeof(Boolean))
                result = typeof(Boolean?);
            else if (column.DataType == typeof(Double))
                result = typeof(Double?);
            else if (column.DataType == typeof(float))
                result = typeof(float?);
            else if (column.DataType == typeof(decimal))
                result = typeof(decimal?);
            else if (column.DataType == typeof(uint))
                result = typeof(uint?);
            else if (column.DataType == typeof(UInt16))
                result = typeof(UInt16?);
            else if (column.DataType == typeof(UInt32))
                result = typeof(UInt32?);
            else if (column.DataType == typeof(UInt64))
                result = typeof(UInt64?);
            else if (column.DataType == typeof(short))
                result = typeof(short?);
            else if (column.DataType == typeof(ushort))
                result = typeof(ushort?);
            else if (column.DataType == typeof(DateTime))
                result = typeof(DateTime?);
            else if (column.DataType == typeof(TimeSpan))
                result = typeof(TimeSpan?);
            return result;
        }

        private static string ToCleanColumn(this string value) { return value.Replace(".", "").Replace(" ",""); }

    }
}
