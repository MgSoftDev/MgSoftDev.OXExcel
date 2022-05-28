using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Entities.Table
{
    internal class OxTableColumnTotalRowEntity
    {
        public string CustomFormula { get; set; }
        public string TotalsRowLabel { get; set; }
        public TotalsRowFormulas RowFormula { get; set; }
        public bool IncludeHidden { get; set; }
        public OxCellFormartEntity CellFormart { get; set; }
    }
}
