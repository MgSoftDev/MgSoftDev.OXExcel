namespace MgSoftDev.OXExcel.Entities.Table
{
    public class OxTableColumnHyperlinkTemplateEntity
    {
        public uint TableRowIndex { get; set; }
        public uint SheetRowIndex { get; set; }
        public object CellValue { get; set; }
        public object Row { get; set; }
        public List<object> MasterData { get; set; }
    }
}
