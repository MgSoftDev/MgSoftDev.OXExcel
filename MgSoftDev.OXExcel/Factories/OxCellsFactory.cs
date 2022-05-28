using MgSoftDev.OXExcel.Entities.ColsRowsCells;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxCellsFactory
    {
        internal readonly List<OxCellEntity> Cells;

        public OxCellsFactory()
        {
            Cells = new List<OxCellEntity>();
        }

        public OxCellFactory Add(uint col, uint row)
        {
            var f = new OxCellFactory(col, row);
            Cells.Add(f.Cell);
            return f;
        }

        public OxCellFactory Add(string col, uint row)
        {
            var f = new OxCellFactory(col, row);
            Cells.Add(f.Cell);
            return f;
        }
        public OxCellFactory Add(string cellReference)
        {
            var f = new OxCellFactory(cellReference);
            Cells.Add(f.Cell);
            return f;
        }
    }
}
