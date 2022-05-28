using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using MgSoftDev.OXExcel.Attributes;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Entities.Table;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxTableColumnFactory 
    {
        internal readonly OxTableColumnsEntity TableColumn;

        internal OxTableColumnFactory(string propertypath, Type originType, OxCellTypeValues cellType, int order)
        {
            TableColumn = new OxTableColumnsEntity
            {
                PropertyPath = propertypath,
                OriginType = originType,
                CellTypeValue = cellType,
                Order = order
            };
            DefaultNumberFormat();
            Size(11);
        }
        public OxTableColumnFactory Order(int order )
        {
            TableColumn.Order = order;
            return this;
        }

        internal OxTableColumnFactory ExtractAttributes(PropertyInfo propertyInfo)
        {
            
            var attr = (DisplayAttribute)propertyInfo.GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();

            if (attr != null)
            {
                Header(attr.Name);
                Order(attr.Order);
            }
            var attr2 = propertyInfo.GetAttribute<DisplayNameAttribute>();
            if (attr2 != null) Header(attr2.DisplayName);

            var colAttr = propertyInfo.GetAttribute<OxColumnAttribute>();
            if (colAttr != null)
            {
                if (!string.IsNullOrEmpty(colAttr.Header)) Header(colAttr.Header);
                Format(colAttr.CellFormart);
                if (colAttr.CellTypeValue != OxCellTypeValues.Default) CellType(colAttr.CellTypeValue);
                DefaultFormulaVal(colAttr.DefaultFormulaValue);
                DefaultVal(colAttr.DefaultValue);
                HeaderFormat(colAttr.HeaderCellFormart);
                IsFormula(colAttr.IsFormula);
                ShowPhonetic(colAttr.ShowPhonetic);
                if (colAttr.Size > 0) Size(colAttr.Size);

                if( colAttr.Order != int.MinValue ) Order( colAttr.Order );
            }

            
            return this;
        }

        public OxTableColumnFactory IsFormula(bool value = true)
        {
            TableColumn.IsFormula = value;
            return this;
        }

        public OxTableColumnFactory DefaultFormulaVal(string value)
        {
            TableColumn.DefaultFormulaValue = value;
            return this;
        }
        public OxTableColumnFactory DefaultVal(object value)
        {
            TableColumn.DefaultValue = value;
            return this;
        }
        public OxTableColumnFactory Size(uint value)
        {
            TableColumn.Size = value;
            return this;
        }
        public OxTableColumnFactory Type(Type value)
        {
            TableColumn.OriginType = value;
            CellType(TableColumn.OriginType.ToCellTypeValues());
            DefaultNumberFormat();
            return this;
        }
        public OxTableColumnFactory CellType(OxCellTypeValues value)
        {
             TableColumn.CellTypeValue = value;
            return this;
        }
        public OxTableColumnFactory ShowPhonetic(bool value = true)
        {
             TableColumn.ShowPhonetic = value;
            return this;
        }

        public OxTableColumnFactory HyperlinkTemplate(Func<OxTableColumnHyperlinkTemplateEntity,OxHyperlinkEntity> hyperlinkFunc )
        {
            TableColumn.HyperlinkTemplate = hyperlinkFunc;
            return this;
        }

        public OxTableColumnFactory TemplateValue(Func<OxTableColumnTemplateEntity, object> template)
        {
            TableColumn.TemplateValue = template;
            return this;
        }

        #region Header
        public OxTableColumnFactory Header(string value)
        {
            TableColumn.Header = value;
            return this;
        }
        public OxTableColumnFactory HeaderFormat(OxCellFormartFactory value)
        {
            TableColumn.HeaderCellFormart = value?.Format;            
            return this;
        }
        public OxCellFormartFactory HeaderFormat()
        {
            var f = new OxCellFormartFactory();
            TableColumn.HeaderCellFormart = f.Format;
            return f;
        }
        public OxTableColumnFactory HeaderFormat(Action<OxCellFormartFactory> formatAction)
        {
            var f = new OxCellFormartFactory();
            formatAction(f);
            TableColumn.HeaderCellFormart = f.Format;
            return this;
        }

        #endregion

        #region Filter
        public OxTableColumnFactory Filter(IEnumerable<string> valueElements)
        {
            TableColumn.CustomColumnFilter = null;
            TableColumn.ColumnFilter = valueElements.Select(s => new OxColumnFilterEntity() {Val = s}).ToList();
            return this;
        }
        public OxTableColumnFactory Filter(object value, OxFilterOperators filterOperator)
        {
            TableColumn.ColumnFilter = null;
            TableColumn.CustomColumnFilter = new OxCustomColumnFilterEntity() { Val = value.ToExcelValue(), Operator = filterOperator, Condition = OxCustomFilterCondition.None};
            return this;
        }
        public OxTableColumnFactory Filter(object value, OxFilterOperators filterOperator, OxCustomFilterCondition condition, object value2, OxFilterOperators filterOperator2)
        {
            TableColumn.ColumnFilter = null;
            TableColumn.CustomColumnFilter = new OxCustomColumnFilterEntity() { Val = value.ToExcelValue(), Operator = filterOperator, Condition = condition, Val2 = value2.ToExcelValue(), Operator2 = filterOperator2};
            return this;
        }
        #endregion

        #region togtal rows

        public OxTableColumnFactory TotalRow(string totalRowLabel )
        {
            TableColumn.TotalRow = TableColumn.TotalRow ?? new OxTableColumnTotalRowEntity { IncludeHidden = false, RowFormula = TotalsRowFormulas.None };
            TableColumn.TotalRow.TotalsRowLabel = totalRowLabel;
            return this;
        }
        public OxTableColumnFactory TotalRow(TotalsRowFormulas totalRowFormula, bool includeHidenValue )
        {
            TableColumn.TotalRow = TableColumn.TotalRow ?? new OxTableColumnTotalRowEntity { IncludeHidden = false, RowFormula = TotalsRowFormulas.None };
            TableColumn.TotalRow.IncludeHidden = includeHidenValue;
            TableColumn.TotalRow.RowFormula = totalRowFormula;
            return this;
        }
       

        public OxTableColumnFactory TotalRow(string formula, bool includeHidenValue)
        {
            TableColumn.TotalRow = TableColumn.TotalRow ?? new OxTableColumnTotalRowEntity { IncludeHidden = false, RowFormula = TotalsRowFormulas.None };
            TableColumn.TotalRow.IncludeHidden =includeHidenValue;
            TableColumn.TotalRow.RowFormula = TotalsRowFormulas.Custom;
            TableColumn.TotalRow.CustomFormula =formula;
            return this;
        }

        public OxTableColumnFactory TotalRowFormat(OxCellFormartFactory value)
        {
            TableColumn.TotalRow = TableColumn.TotalRow?? new OxTableColumnTotalRowEntity { IncludeHidden = false, RowFormula = TotalsRowFormulas.None };
            TableColumn.TotalRow.CellFormart = value.Format;
            //DefaultNumberFormat();
            return this;
        }
        public OxCellFormartFactory TotalRowFormat()
        {
            TableColumn.TotalRow = TableColumn.TotalRow ?? new OxTableColumnTotalRowEntity { IncludeHidden = false, RowFormula = TotalsRowFormulas.None  };
            var f = new OxCellFormartFactory(TableColumn.TotalRow.CellFormart?? new OxCellFormartEntity());
            TableColumn.TotalRow.CellFormart = f.Format;
            //DefaultNumberFormat();
            return f;
        }
        public OxTableColumnFactory TotalRowFormat(Action<OxCellFormartFactory> formatAction)
        {
            TableColumn.TotalRow = TableColumn.TotalRow ?? new OxTableColumnTotalRowEntity { IncludeHidden = false, RowFormula = TotalsRowFormulas.None };
            var f = new OxCellFormartFactory(TableColumn.TotalRow.CellFormart ?? new OxCellFormartEntity());
            formatAction(f);
            TableColumn.TotalRow.CellFormart = f.Format;
            //DefaultNumberFormat();
            return this;
        }

        #endregion


        #region format
        public OxTableColumnFactory Format(OxCellFormartFactory value)
        {
            TableColumn.CellFormart = value?.Format;
            DefaultNumberFormat();
            return this;
        }
        public OxCellFormartFactory Format()
        {
            var f = new OxCellFormartFactory();
            TableColumn.CellFormart = f.Format;
            DefaultNumberFormat();
            return f;
        }
        public OxTableColumnFactory Format(Action<OxCellFormartFactory> formatAction)
        {
            var f = new OxCellFormartFactory();
            formatAction(f);
            TableColumn.CellFormart = f.Format;
            DefaultNumberFormat();
            return this;
        }

        public OxTableColumnFactory TemplateFormat(Func<OxTableColumnTemplateEntity, OxCellFormartFactory> template)
        {
            TableColumn.TemplateFormat = template;
            return this;
        }

        private void DefaultNumberFormat()
        {
            if (TableColumn.OriginType == null) return;
            string numberFormat = null;
            if (TableColumn.OriginType == typeof(DateTime) || TableColumn.OriginType == typeof(DateTime?))
                numberFormat = "dd/MM/yyyy hh:mm:ss";
            else if (TableColumn.OriginType == typeof(TimeSpan) || TableColumn.OriginType == typeof(TimeSpan?))
                numberFormat = "hh:mm:ss";
            else return;
            
            if (TableColumn.CellFormart == null)
                Format().NumberFormat(numberFormat);
            else if(TableColumn.CellFormart.NumberFormat == null)
                TableColumn.CellFormart.NumberFormat = new OxNumberFormatEntity {FormatCode = numberFormat};            
        }

        #endregion




    }
    
}
