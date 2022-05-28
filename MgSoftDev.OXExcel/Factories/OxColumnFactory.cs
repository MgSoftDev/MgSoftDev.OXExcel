using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxColumnFactory
    {
        internal readonly OxColumnEntity Column;

        public OxColumnFactory(uint fromColumn, uint toColumn)
        {
            Column = new OxColumnEntity()
            {
             Min   = fromColumn,
             Max = toColumn,
             Format = null
            };
            BestFit();
            OutlineLevel(0);
            CollapsedOutlining(false);
            Hidden(false);
            Phonetic(false);
           
        }
        public OxColumnFactory(string fromColumn, string toColumn)
        {
            Column = new OxColumnEntity()
            {
                Min = fromColumn.ToColIndex(),
                Max = toColumn.ToColIndex(),
                Format = null
            };
            BestFit();
            OutlineLevel(0);
            CollapsedOutlining(false);
            Hidden(false);
            Phonetic(false);
        }

        public OxColumnFactory BestFit()
        {
            Column.CustomWidth = true;
            Column.BestFit = true;
            return this;
        }
        public OxColumnFactory OutlineLevel(byte outlineLevel)
        {
            Column.OutlineLevel = outlineLevel;
            return this;
        }
        public OxColumnFactory CollapsedOutlining(bool collapsed = true)
        {
            Hidden(true);
            return this;
        }
        public OxColumnFactory Hidden(bool hidden = true)
        {
            Column.Hidden = hidden;
            return this;
        }

        public OxColumnFactory Width(double width)
        {
            Column.CustomWidth = true;
            Column.BestFit = false;
            Column.Width = width;
            return this;
        }

        public OxColumnFactory Phonetic(bool phonetic)
        {
            Column.Phonetic = phonetic;
            return this;
        }

        public OxColumnFactory Format(OxCellFormartFactory value)
        {
            Column.Format = value.Format;
            return this;
        }
        public OxCellFormartFactory Format()
        {
            var f = new OxCellFormartFactory();
            Column.Format =f.Format;
            return f;
        }
        public OxColumnFactory Format(Action<OxCellFormartFactory> formatAction)
        {
            var f = new OxCellFormartFactory();
            formatAction(f);
            Column.Format = f.Format;
            return this;
        }




    }
}
