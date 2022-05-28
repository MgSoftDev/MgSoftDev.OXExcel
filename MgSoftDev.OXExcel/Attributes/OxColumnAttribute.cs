using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Factories;

namespace MgSoftDev.OXExcel.Attributes
{
    public class OxColumnAttribute : Attribute
    {
        public string               Header              { get; set; }
        public uint                 Size                { get; set; }
        public string               DefaultFormulaValue { get; set; } = "";
        public object               DefaultValue        { get; set; }
        public OxCellTypeValues     CellTypeValue       { get; set; }
        public bool                 ShowPhonetic        { get; set; }
        public bool                 IsFormula           { get; set; }
        public OxCellFormartFactory CellFormart         { get; set; }
        public OxCellFormartFactory HeaderCellFormart   { get; set; }
        public int                  Order               { get; set; } = int.MinValue;

    }

    
}
