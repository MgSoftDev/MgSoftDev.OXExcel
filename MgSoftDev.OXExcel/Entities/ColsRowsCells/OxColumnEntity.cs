using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Entities.ColsRowsCells
{
    internal class OxColumnEntity
    {
        public bool BestFit { get; set; }
        public bool Collapsed { get; set; }
        public bool CustomWidth { get; set; }
        public double Width { get; set; }
        public bool Hidden { get; set; }
        public uint Max { get; set; }
        public uint Min { get; set; }
        public byte OutlineLevel { get; set; }
        public bool Phonetic { get; set; }
        public OxCellFormartEntity Format { get; set; }
    }
}
