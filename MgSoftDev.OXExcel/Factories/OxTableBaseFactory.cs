using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Table;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxTableBaseFactory
    {
        internal readonly OxTableEntity Table;

        internal OxTableBaseFactory(object data, uint col, uint row)
        {
            Table = new OxTableEntity();
            Inicializar(data,col,row);
        }
        internal OxTableBaseFactory(object data, string col, uint row)
        {
            Table = new OxTableEntity();
            Inicializar(data, col.ToColIndex(), row);
        }
        internal OxTableBaseFactory(object data, string reference)
        {
            Table = new OxTableEntity();
            Inicializar(data, reference.GetCol(), reference.GetRow());
        }

        private void Inicializar(object data, uint col, uint row)
        {
            Table.DataCollection = (data as IEnumerable<object>)?.ToList() ?? new List<object>();
            Table.Column = col;
            Table.Row = row;
            Table.Columns = new List<OxTableColumnsEntity>();
            RowDefinition(new OxRowFactory(0));
            TableType(OxTableType.Excel);
            TableStyle(new OxTableStyleInfoFactory());
            HideAutoFilter(false);
            AutoGenerateColumns(false);
        }

        public OxTableBaseFactory TableType(OxTableType value)
        {
            Table.TableType = value;
            return this;
        }
        public OxTableBaseFactory Name(string value)
        {
            Table.TableName = value;
            return this;
        }
        public OxTableBaseFactory HideAutoFilter(bool value = true)
        {
            Table.AutoFilter = !value;
            return this;
        }
        public OxTableBaseFactory AutoGenerateColumns(bool value = true)
        {
            Table.AutoGenerateColumns = value;
            return this;
        }
        public OxTableBaseFactory TotalsRowShown(bool value = true)
        {
            Table.TotalsRowShow = value;
            return this;
        }
        public OxTableBaseFactory RowDefinition(OxRowFactory rowDefinitions)
        {
            Table.RowDefinition = rowDefinitions.Row;
            return this;
        }
        public OxRowFactory RowDefinition()
        {
            var res = new OxRowFactory(0);
            Table.RowDefinition = res.Row;
            return res;
        }
        public OxTableBaseFactory RowDefinition(Action<OxRowFactory> rowAction)
        {
            var r = new OxRowFactory(0);
            rowAction(r);
            Table.RowDefinition = r.Row;
            return this;
        }
        public OxTableBaseFactory RowDefinitionTemplate(Func<OxTableRowDefinitionTemplateEntity,OxRowFactory> rowAction)
        {
            Table.RowDefinitionTemplate = rowAction;
            return this;
        }


        #region TableStyle Info
        public OxTableBaseFactory TableStyle(OxTableStyleInfoFactory tableStyleDefinitions)
        {
            Table.TableStyleInfo = tableStyleDefinitions.TableStyle;
            return this;
        }
        public OxTableStyleInfoFactory TableStyle()
        {
            var res = new OxTableStyleInfoFactory();
            Table.TableStyleInfo = res.TableStyle;
            return res;
        }

        public OxTableBaseFactory TableStyle(Action<OxTableStyleInfoFactory> tableStyleAction)
        {
            var r = new OxTableStyleInfoFactory();
            tableStyleAction(r);
            Table.TableStyleInfo = r.TableStyle;
            return this;
        }
        #endregion


    }
    
}
