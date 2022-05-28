using MgSoftDev.OXExcel.Entities.ColsRowsCells;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxRowsFactory
    {
        internal readonly List<OxRowEntity> Rows;

        public OxRowsFactory()
        {
            Rows = new List<OxRowEntity>();
        }

        public OxRowFactory Add(uint rowIndex)
        {
            var ant = Rows.FirstOrDefault(e => e.RowIndex == rowIndex);
            if (ant != null) Rows.Remove(ant);

            var r = new OxRowFactory(rowIndex);            
            Rows.Add(r.Row);
            return r;
        }
    }
}
