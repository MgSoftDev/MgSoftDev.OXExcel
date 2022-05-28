using MgSoftDev.OXExcel.Entities.ColsRowsCells;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxColumnsFactory
    {
        internal readonly List<OxColumnEntity> Columns;

        public OxColumnsFactory()
        {
            Columns = new List<OxColumnEntity>();
        }

        public OxColumnFactory Add(uint fromColumn, uint toColumn)
        {
            var r = new OxColumnFactory(fromColumn, toColumn);
            Columns.Add(r.Column);
            return r;
        }

        public OxColumnFactory Add(string fromColumn, string toColumn)
        {
            var r = new OxColumnFactory(fromColumn, toColumn);
            Columns.Add(r.Column);
            return r;
        }
    }
}
