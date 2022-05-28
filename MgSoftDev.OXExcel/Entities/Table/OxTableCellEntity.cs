using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Entities.Interface;

namespace MgSoftDev.OXExcel.Entities.Table
{
    internal class OxTableCellEntity : IReferenceCell
    {
        public uint Row { get; set; }
        public uint Column { get; set; }
        public int Value { get; set; }
       // public int OriginType { get; set; }
        public OxCellFormulaEntity Formula { get; set; }
        public OxCellTypeValues CellTypeValue { get; set; }
        public bool ShowPhonetic { get; set; }
        public OxCellFormartEntity CellFormart { get; set; }
       // public string MargenReference { get; set; }
        public OxHyperlinkEntity Hyperlink { get; set; }





    }

}
