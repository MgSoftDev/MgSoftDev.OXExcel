using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Entities.Interface;

namespace MgSoftDev.OXExcel.Entities.ColsRowsCells
{
    [Serializable]
    internal class OxRowEntity : IReferenceRow
    {
        public bool Collapsed { get; set; }
        public bool CustomFormat { get; set; }
        public bool CustomHeight { get; set; }
        public double Height { get; set; }
        public bool Hidden { get; set; }
        public byte OutlineLevel { get; set; }
        public uint RowIndex { get; set; }
        public bool ShowPhonetic { get; set; }
        public bool ThickBot { get; set; }
        public bool ThickTop { get; set; }
        public OxCellFormartEntity Format { get; set; }
    }
}
