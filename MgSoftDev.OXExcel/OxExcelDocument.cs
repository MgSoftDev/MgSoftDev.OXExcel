using System.Globalization;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Document;
using MgSoftDev.OXExcel.Entities.Sheet;
using MgSoftDev.OXExcel.Factories;
using MgSoftDev.OXExcel.OpenXmlProvider;

namespace MgSoftDev.OXExcel
{
    public class OxExcelDocument : IDisposable
    {
        internal readonly OxDocumentEntity Document;

        public OxExcelDocument()
        {
            Document = new OxDocumentEntity
            {
                Sheets = new List<OxSheetEntity>(),
                Calculation = null,
                PackageProperties = null
            };
            DocumentType(OxDocumentTypes.Workbook);
        }


        #region Document Properties 
        public OxExcelDocument DocumentType(OxDocumentTypes documentType)
        {
            Document.DocumentType = documentType;
            return this;
        }
        public OxExcelDocument DataCultureInfo(CultureInfo culture)
        {
            Const.CultureData = culture;
            return this;
        }

        public OxExcelDocument Calculation(Action<OxCalculationFactory> calculationAction)
        {
            var f = new OxCalculationFactory();
            calculationAction(f);
            Document.Calculation = f.Calculation;
            return this;
        }
        public OxExcelDocument Calculation(OxCalculationFactory calculation)
        {
            Document.Calculation = calculation.Calculation;
            return this;
        }
        public OxExcelDocument PackageProperties(Action<OxPackagePropertiesFactory> packagePropertiesAction)
        {
            var f = new OxPackagePropertiesFactory();
            packagePropertiesAction(f);
            Document.PackageProperties = f.PackageProperties;
            return this;
        }
        public OxExcelDocument PackageProperties(OxPackagePropertiesFactory packageProperties)
        {
            Document.PackageProperties = packageProperties.PackageProperties;
            return this;
        }

        public void Save(string filePath)
        {

            Thread.CurrentThread.CurrentCulture = Const.CultureData;
            var docXml = new OpenXmlExcelProvider(Document);
            docXml.Build(filePath);
            Dispose();
        }

        #endregion

        #region Child Properties 
        public OxExcelDocument AddSheet(Action<OxSheetsFactory> sheetsAction )
        {
            sheetsAction(new OxSheetsFactory(Document.Sheets));
            return this;
        }
        public OxSheetFactory AddSheet(string tabName)
        {
            var sf = new OxSheetFactory(tabName);
            Document.Sheets.Add(sf.Sheet);
            return sf;
        }
        public OxExcelDocument AddSheet(OxSheetFactory sheet)
        {
            Document.Sheets.Add(sheet.Sheet);
            return this;
        }
        public OxExcelDocument AddSheet(IEnumerable<OxSheetFactory> sheets)
        {
            Document.Sheets.AddRange(sheets.Select(s=>s.Sheet));
            return this;
        }
        #endregion

        public void Dispose()
        {
            // clean Table DataCollection
            Document.Sheets.ForEach(d =>
            {
                Const.Clean();
                d.RowsCellsList.Clear();

                GC.SuppressFinalize(d.RowsCellsList);

                d.RowsCellsList = null;
                d.Tables.ForEach(t =>
                {
                    t.DataCollection.Clear();
                    GC.SuppressFinalize(t.DataCollection);
                    t.DataCollection = null;
                });
            });
            GC.SuppressFinalize(Document);
            GC.Collect();
        }
    }
}
