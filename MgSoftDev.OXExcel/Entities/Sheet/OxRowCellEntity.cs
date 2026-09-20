using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Interface;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    /// <summary>
    /// Filas y celdas de UNA hoja, junto con la extension que ocupan, sus celdas combinadas y sus hipervinculos.
    /// Todo esto es propio de la hoja: cuando vivia en Const, que es del documento, las combinaciones y los
    /// vinculos de una hoja se repetian en las siguientes y el rango declarado iba creciendo entre hojas.
    /// </summary>
    public class OxRowsCellCollection 
    {
        public SortedDictionary<uint, OxRowCellsEntity> Rows { get; set; } = new SortedDictionary<uint, OxRowCellsEntity>();
        private readonly object _Lock = new object();

        internal uint MinRowIndex { get; private set; } = uint.MaxValue;
        internal uint MaxRowIndex { get; private set; }
        internal uint MinCellIndex { get; private set; } = uint.MaxValue;
        internal uint MaxCellIndex { get; private set; }

        /// <summary>Referencias de las celdas combinadas de la hoja, sin repetir.</summary>
        internal List<string> MergeReferences { get; } = new List<string>();

        /// <summary>Hipervinculos de la hoja; se llenan al escribir las celdas.</summary>
        internal List<OxHyperlinkEntity> Hyperlinks { get; } = new List<OxHyperlinkEntity>();


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
                    Rows.Add(row.RowIndex, new OxRowCellsEntity(){Row  = row, Owner = this});
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
                    var item = new OxRowCellsEntity() { Row = row, Owner = this };
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
                Rows?.Clear();
                MergeReferences.Clear();
                Hyperlinks.Clear();
                MinRowIndex  = uint.MaxValue;
                MaxRowIndex  = 0;
                MinCellIndex = uint.MaxValue;
                MaxCellIndex = 0;
        }

        /// <summary>Crece el rango de filas de la hoja. Las tablas en streaming lo usan para declarar el rango antes de escribirlas.</summary>
        internal void UpdateRowMinMax(uint rowIndex)
        {
            if (rowIndex < MinRowIndex) MinRowIndex = rowIndex;
            if (rowIndex > MaxRowIndex) MaxRowIndex = rowIndex;
        }

        internal void UpdateCellMinMax(uint cellIndex)
        {
            if (cellIndex < MinCellIndex) MinCellIndex = cellIndex;
            if (cellIndex > MaxCellIndex) MaxCellIndex = cellIndex;
        }

        internal void AddMergeReference(string reference)
        {
            if (!MergeReferences.Contains(reference)) MergeReferences.Add(reference);
        }
    }

    public class OxRowCellsEntity
    {
        public IReferenceRow Row { get; set; }
        public SortedDictionary<uint, IReferenceCell> Cells { get; set; } = new SortedDictionary<uint, IReferenceCell>();
        private readonly object _Lock = new object();

        /// <summary>Hoja de la fila; ahi se registran la extension y las celdas combinadas.</summary>
        internal OxRowsCellCollection Owner { get; set; }

        public void Add(IReferenceCell cell)
        {
            lock (_Lock)
            {
                if (!Cells.ContainsKey(cell.Column))
                {
                    Cells.Add(cell.Column, cell );
                    Owner?.UpdateCellMinMax(cell.Column);

                    if (cell is OxCellEntity cc && cc.MargenReference != null)
                        Owner?.AddMergeReference(cc.MargenReference);
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
    }


}
