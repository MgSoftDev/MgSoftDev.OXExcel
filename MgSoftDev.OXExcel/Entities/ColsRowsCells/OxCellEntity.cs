using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Entities.Interface;

namespace MgSoftDev.OXExcel.Entities.ColsRowsCells
{
    internal class OxCellEntity : IReferenceCell
    {
        public uint Row { get; set; }
        public uint Column { get; set; }
        public string Value { get; set; }
        public Type OriginType { get; set; }
        public OxCellFormulaEntity Formula { get; set; }
        public OxCellTypeValues CellTypeValue { get; set; }
        public bool ShowPhonetic { get; set; }
        public OxCellFormartEntity CellFormart { get; set; }
        public string MargenReference { get; set; }
        public OxHyperlinkEntity Hyperlink { get; set; }





    }
}
