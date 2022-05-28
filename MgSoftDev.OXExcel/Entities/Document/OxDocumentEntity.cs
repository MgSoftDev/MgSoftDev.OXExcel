using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Entities.Document
{
    internal class OxDocumentEntity
    {
        public List<OxSheetEntity> Sheets { get; set; }
        public OxCalculationEntity Calculation { get; set; }
        public OxDocumentTypes DocumentType { get; set; }
        public OxPackagePropertiesEntity PackageProperties { get; set; }
    }
}
