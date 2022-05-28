using MgSoftDev.OXExcel.Factories;

namespace MgSoftDev.OXExcel.Entities.Table
{
    public class OxTableColumnTemplateEntity
    {
        public uint TableRowIndex { get; set; }
        public uint SheetRowIndex { get; set; }
        public object CellValue { get; set; }
        public object Row { get; set; }
        public List<object> MasterData { get; set; }
        public OxCellFormartFactory Format { get; set; }
    }
}
