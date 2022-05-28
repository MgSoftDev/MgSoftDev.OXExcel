using MgSoftDev.OXExcel.Factories;

namespace MgSoftDev.OXExcel.Entities.Table
{
    public class OxTableRowDefinitionTemplateEntity
    {
        public uint TableRowIndex { get; set; }
        public uint SheetRowIndex { get; set; }
        public OxRowFactory RowDefinition { get; set; }
        public object Rows { get; set; }
        public List<object> MasterData { get; set; }

    }
}
