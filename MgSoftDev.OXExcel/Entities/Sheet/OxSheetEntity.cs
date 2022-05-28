using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Table;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    internal class OxSheetEntity
    {
        public string TabName { get; set; }
        public OxSheetViewEntity SheetView { get; set; }
        public OxSheetVisibilities SheetVisibility { get; set; }
        public OxPageMarginsEntity PageMargins { get; set; }
        public OxSheetPropertiesEntity SheetProperties { get; set; }
        public List<OxColumnEntity> Columns { get; set; }
        public OxPageSetupEntity PageSetup { get; set; }

        
        public OxRowsCellCollection RowsCellsList { get; set; }
       

        public List<OxImageEntity> Images { get; set; }
        public List<OxTableEntity> Tables { get; set; }
        public Uri BackgroundImage { get; set; }
        
    }
}
