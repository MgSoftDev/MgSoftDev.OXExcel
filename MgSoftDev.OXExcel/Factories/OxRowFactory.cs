using MgSoftDev.OXExcel.Entities.ColsRowsCells;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxRowFactory
    {
        internal readonly OxRowEntity Row;

        public OxRowFactory(uint rowIndex)
        {
            Row = new OxRowEntity {Format = null, RowIndex = rowIndex};
            
            Collapsed(false);
            CustomFormat(false);
            CustomHeight(false);
            Hidden(false);
            OutlineLevel(0);
            ShowPhonetic(false);
            ThickBot(false);
            ThickTop(false);
        }

        internal OxRowFactory(OxRowEntity value)
        {
            Row = value;
        }

        public OxRowFactory Collapsed(bool value = true)
        {
            Hidden(true);
            return this;
        }
        internal OxRowFactory CustomFormat(bool value = true)
        {
            Row.CustomFormat = value;
            return this;
        }
        private OxRowFactory CustomHeight(bool value = true)
        {
            Row.CustomHeight = value;
            return this;
        }
        public OxRowFactory Hidden(bool value = true)
        {
            Row.Hidden = value;
            return this;
        }
        public OxRowFactory Height(double value)
        {
            Row.Height = value;
            CustomHeight();
            return this;
        }
        public OxRowFactory OutlineLevel(byte value)
        {
            Row.OutlineLevel = value;
            return this;
        }
        public OxRowFactory ShowPhonetic(bool value = true)
        {
            Row.ShowPhonetic = value;
            return this;
        }
        public OxRowFactory ThickBot(bool value = true)
        {
            Row.ThickBot = value;
            return this;
        }
        public OxRowFactory ThickTop(bool value = true)
        {
            Row.ThickTop = value;
            return this;
        }

        public OxRowFactory Format(OxCellFormartFactory value)
        {
            Row.Format = value.Format;
            CustomFormat();
            return this;
        }
        public OxCellFormartFactory Format()
        {
            var f = Row.Format== null ?new OxCellFormartFactory(): new OxCellFormartFactory(Row.Format);
            Row.Format = f.Format;
            CustomFormat();
            return f;
        }
        public OxRowFactory Format(Action<OxCellFormartFactory> formatAction)
        {
            var f = Row.Format == null ? new OxCellFormartFactory() : new OxCellFormartFactory(Row.Format);
            formatAction(f);
            Row.Format = f.Format;
            CustomFormat();
            return this;
        }
    }
}
