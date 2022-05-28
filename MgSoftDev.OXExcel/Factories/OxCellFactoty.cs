using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxCellFactory
    {
        internal readonly OxCellEntity Cell;

        public OxCellFactory(uint col ,uint row )
        {
            Cell = new OxCellEntity() { Row = row, Column = col};
            Inicializar();
        }
        public OxCellFactory(string col , uint row)
        {
            Cell = new OxCellEntity() { Row = row, Column = col.ToColIndex() };
            Inicializar();
        }
        public OxCellFactory(string cellReference)
        {
            Cell = new OxCellEntity() { Row = cellReference.GetRow(), Column = cellReference.GetCol() };
            Inicializar();
        }

        private void Inicializar()
        {
            Value(null);
            Cell.Formula = null;
            Phonetic(false);
            Cell.CellFormart = null;
            Cell.OriginType = null;
        }

        public OxCellFactory Value(object value)
        {
            Cell.Value = value.ToExcelValue();
            Cell.OriginType = value== null?typeof(string): value.GetType();
            CellType(Cell.OriginType.ToCellTypeValues());
            DefaultNumberFormat();
            return this;
        }

        public OxCellFactory Hyperlink(Uri url, string toolTip)
        {
            Cell.Hyperlink = new OxHyperlinkEntity(url, toolTip) { Column =Cell.Column, Row=Cell.Row};
            return this;
        }
        public OxCellFactory Hyperlink(string excelReference, string toolTip)
        {
            Cell.Hyperlink = new OxHyperlinkEntity(excelReference,toolTip) { Column = Cell.Column, Row = Cell.Row };
            return this;
        }

        public OxCellFactory Formula(string value)
        {
            Cell.Formula = new  OxCellFormulaEntity() {Formula = value};
            return this;
        }
        public OxCellFactory Margen(uint column, uint row)
        {
            Cell.MargenReference = $"{Cell.Column.ToReferenceAlfa()}{Cell.Row}:{(Cell.Column+column).ToReferenceAlfa()}{Cell.Row+row}";
            return this;
        }
        public OxCellFactory CellType(OxCellTypeValues typeValue)
        {
            Cell.CellTypeValue = typeValue;
            return this;
        }
        public OxCellFactory Phonetic(bool value = true)
        {
            Cell.ShowPhonetic = value;
            return this;
        }
        public OxCellFactory Format(OxCellFormartFactory value)
        {
            Cell.CellFormart = value.Format;
            DefaultNumberFormat();
            return this;
        }
        public OxCellFormartFactory Format()
        {
            var f = new OxCellFormartFactory();
            Cell.CellFormart = f.Format;
            DefaultNumberFormat();
            return f;
        }
        public OxCellFactory Format(Action<OxCellFormartFactory> formatAction)
        {
            var f = new OxCellFormartFactory();
            formatAction(f);
            Cell.CellFormart = f.Format;
            DefaultNumberFormat();
            return this;
        }

        #region Private

        private void DefaultNumberFormat()
        {
            if (Cell.OriginType == null) return;
            string numberFormat = null;
            if (Cell.OriginType == typeof(DateTime) || Cell.OriginType == typeof(DateTime?))
                numberFormat = "dd/MM/yyyy hh:mm:ss";
            else if (Cell.OriginType == typeof(TimeSpan) || Cell.OriginType == typeof(TimeSpan?))
                numberFormat = "hh:mm:ss";
            else return;
            
            if (Cell.CellFormart == null)
                Format().NumberFormat(numberFormat);
            else if(Cell.CellFormart.NumberFormat == null)
                Cell.CellFormart.NumberFormat = new OxNumberFormatEntity {FormatCode = numberFormat};            
        }

        
        #endregion

    }
}
