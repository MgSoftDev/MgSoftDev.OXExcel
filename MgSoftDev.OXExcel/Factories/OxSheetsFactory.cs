using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxSheetsFactory
    {
        internal readonly List<OxSheetEntity> Sheets;

        internal OxSheetsFactory(List<OxSheetEntity> sheets)
        {
            Sheets = sheets;
        }

        public OxSheetFactory Add(string tabName)
        {
            var r = new OxSheetFactory(tabName);
            Sheets.Add(r.Sheet);
            return r;
        }
    }
}
