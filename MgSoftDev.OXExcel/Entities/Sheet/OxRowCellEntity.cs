using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Interface;
using MgSoftDev.OXExcel.OpenXmlProvider;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    public class OxRowsCellCollection 
    {
        public SortedDictionary<uint, OxRowCellsEntity> Rows { get; set; } = new SortedDictionary<uint, OxRowCellsEntity>();
        private readonly object _Lock = new object();


        public void AddOrRemplace( IReferenceRow row)
        {
            lock( _Lock )
            {
                if( Rows.TryGetValue( row.RowIndex, out var val ) ) val.Row = row;
                else Add(row);
            }
        }


        public void Add(IReferenceRow row)
        {
            lock (_Lock)
            {
                if (!Rows.ContainsKey(row.RowIndex))
                {
                    Rows.Add(row.RowIndex, new OxRowCellsEntity(){Row  = row});
                    UpdateRowMinMax(row.RowIndex);
                }
            }

        }

        public OxRowCellsEntity AddAndGet(IReferenceRow row)
        {
            lock (_Lock)
            {
                if (!Rows.ContainsKey(row.RowIndex))
                {
                    var item = new OxRowCellsEntity() { Row = row };
                    Rows.Add(row.RowIndex,item );
                    UpdateRowMinMax(row.RowIndex);
                    return item;
                }

                Rows.TryGetValue(row.RowIndex, out var val);
                return val;
            }

        }

        public OxRowCellsEntity GetValue(uint rowIndex)
        {
            lock (_Lock)
            {
                Rows.TryGetValue(rowIndex, out var val);

                return val;
            }
        }

        public void AddCell(IReferenceCell cell)
        {
            lock (_Lock)
            {
                GetValue(cell.Row).Add(cell);
            }

        }

        public void Clear()
        {
            Rows.Clear();
        }
         private void UpdateRowMinMax(uint rowIndex)
        {

            Const.MinRowIndex= rowIndex < Const.MinRowIndex ? rowIndex : Const.MinRowIndex;
            Const.MaxRowIndex = rowIndex > Const.MaxRowIndex ? rowIndex : Const.MaxRowIndex;
        }

    }

    public class OxRowCellsEntity
    {
        public IReferenceRow Row { get; set; }
        public SortedDictionary<uint, IReferenceCell> Cells { get; set; } = new SortedDictionary<uint, IReferenceCell>();
        private readonly object _Lock = new object();

        public void Add(IReferenceCell cell)
        {
            lock (_Lock)
            {
                if (!Cells.ContainsKey(cell.Column))
                {
                    Cells.Add(cell.Column, cell );
                    UpdateCellMinMax(cell.Column);

                    if (cell is OxCellEntity cc && cc.MargenReference != null)
                    {
                        if (!Const.margetCells.Contains(cc.MargenReference))
                            Const.margetCells.Add(cc.MargenReference);
                    }

                }
            }

        }


        public IReferenceCell GetValue(uint rowIndex)
        {
            lock (_Lock)
            {
                Cells.TryGetValue(rowIndex, out var val);

                return val;
            }
        }
        private void UpdateCellMinMax(uint cellIndex)
        {

            Const.MinCellIndex = cellIndex < Const.MinCellIndex ? cellIndex : Const.MinCellIndex;
            Const.MaxCellIndex = cellIndex > Const.MaxCellIndex ? cellIndex : Const.MaxCellIndex;
        }
    }


}
