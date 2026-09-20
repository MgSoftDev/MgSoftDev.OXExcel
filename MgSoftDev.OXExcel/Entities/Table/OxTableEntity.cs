using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Factories;

namespace MgSoftDev.OXExcel.Entities.Table
{
    internal class OxTableEntity
    {
        public uint Row { get; set; }
        public uint Column { get; set; }
        public string TableName { get; set; }
        public List<OxTableColumnsEntity> Columns { get; set; }
        public List<object> DataCollection { get; set; }

        /// <summary>Filas entregadas de una en una mientras se escribe la hoja; cuando está, sustituye a DataCollection.</summary>
        public IEnumerable<object> DataSource { get; set; }

        /// <summary>Total de filas declarado por quien arma la tabla; evita recorrer DataSource antes de escribir.</summary>
        public long? DeclaredRowsCount { get; set; }
        public OxRowEntity RowDefinition { get; set; }
        public Func<OxTableRowDefinitionTemplateEntity, OxRowFactory> RowDefinitionTemplate { get; set; }
        public OxTableType TableType { get; set; }
        internal uint RowsCounts { get; set; } = 0;
        public bool AutoFilter { get; set; }
        public bool AutoGenerateColumns { get; set; }
        public bool TotalsRowShow { get; set; }
        public OxTableStyleInfoEntity TableStyleInfo { get; set; }
    }
}
