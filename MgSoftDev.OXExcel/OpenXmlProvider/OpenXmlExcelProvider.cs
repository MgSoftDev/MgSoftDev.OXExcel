using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Document;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Entities.Sheet;
using MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions;
using MgSoftDev.OXExcel.OpenXmlProvider.Models;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal partial class OpenXmlExcelProvider
    {

        private readonly OxDocumentEntity _Doc;
        

        public OpenXmlExcelProvider(OxDocumentEntity doc)
        {
            _Doc = doc;
            Const.Formats = new List<OxCellFormartEntity>();
            Const.StringShareds = new List<string>() { "" };
            Const.Hyperlinks = new List<OxHyperlinkEntity>();
            Const.UniqueValuesList = new UniqueList<string>();
            Const.TypeList = new UniqueList<Type>();
        }

        internal void Build(string pathFile)=> CreatePackage(pathFile);
        internal void Build(Stream stream)  => CreatePackage(stream);

        #region Package
        private void CreatePackage(string filePath)
        {
            using (var package = SpreadsheetDocument.Create(filePath, _Doc.DocumentType.ToSpreadsheetDocumentType(), true))
            {
                CreateParts(package);
            }

        }
        private void CreatePackage(Stream stream)
        {
            using (var package = SpreadsheetDocument.Create(stream, _Doc.DocumentType.ToSpreadsheetDocumentType(), true))
            {
                CreateParts(package);
            }
        }

        #endregion
        #region Dinamicas
        private void CreateParts(SpreadsheetDocument document)
        {
            var workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);
            GenerateThemePart1Content(workbookPart1);

            _Doc.Sheets.ForEach(ox =>
            {
                var worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("SId" + (_Doc.Sheets.IndexOf(ox) + 1).ToString());
                
                GenerateWorksheetPartContent2(worksheetPart1, ox);
                
                //region Images
                if (ox.Images != null && ox.Images.Count > 0)
                {
                    var drawPart = worksheetPart1.AddNewPart<DrawingsPart>("rId1");
                    
                    GenerateDrawingsPart1Content(drawPart,ox.Images);
                    
                    var imgIndex = 1;
                    ox.Images.Select(s => s.Id).Distinct().ToList().ForEach(id =>
                    {
                        var img     = ox.Images.First(_=>_.Id == id);
                        var imgPart = drawPart.AddNewPart<ImagePart>($"image/{ img.Extension}", "rId" + imgIndex++);
                        GenerateImagePart1Content(imgPart, img);
                    });
                }
                // end images
                if (ox.BackgroundImage != null)
                {
                    var imgPart = worksheetPart1.AddNewPart<ImagePart>($"image/{ new FileInfo(ox.BackgroundImage.AbsolutePath).Extension.Replace(".","")}", "rId2");
                    GenerateImagePart1Content(imgPart, new OxImageEntity(){Uri = ox.BackgroundImage.AbsolutePath });
                }
                var hyIndex = 0;
                Const.Hyperlinks.Where(w=> w.Uri!= null).ToList().ForEach(f => worksheetPart1.AddHyperlinkRelationship(f.Uri, true, "rId"+ hyIndex++));                
            });
            
            GenerateSharedStringTablePart1Content(workbookPart1);
            GenerateWorkbookStylesPart1Content(workbookPart1);
            GenerateExtendedFilePropertiesPart1Content(document);

            SetPackageProperties(document);
        }

        private void GenerateSharedStringTablePart1Content(WorkbookPart workbookPart1)
        {
            var xw = OpenXmlWriter.Create(workbookPart1.AddNewPart<SharedStringTablePart>("rId5"));
            xw.WriteStartElement(new SharedStringTable() { Count = (uint)Const.StringShareds.Count, UniqueCount = (uint)Const.StringShareds.Count });
            Const.StringShareds.ForEach(ss => xw.WriteElement(new SharedStringItem(new Text(ss))));
            xw.WriteEndElement();
            xw.Close();
        }

        private string GetSharedIndex(string shared)
        {
            if (shared == null) return "0";
            if (!Const.StringShareds.Exists(f => f.Equals(shared)))
                Const.StringShareds.Add(shared);
            return (Const.StringShareds.FindIndex(i => i.Equals(shared))).ToString();
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            if(_Doc.PackageProperties == null) return;
            document.PackageProperties.Creator = _Doc.PackageProperties.Creator;
            document.PackageProperties.Created = _Doc.PackageProperties.Created;            
            document.PackageProperties.Modified = _Doc.PackageProperties.Modified;
            document.PackageProperties.LastModifiedBy = _Doc.PackageProperties.LastModifiedBy;
            document.PackageProperties.Title = _Doc.PackageProperties.Title;
            document.PackageProperties.Version = _Doc.PackageProperties.Version;
        }
        private void GenerateExtendedFilePropertiesPart1Content(SpreadsheetDocument document)
        {
            var           xw          = OpenXmlWriter.Create(document.AddNewPart<ExtendedFilePropertiesPart>("rId4"));
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            xw.WriteStartElement(properties1);

            #region application
            xw.WriteElement(new Ap.Application() { Text = "Microsoft Excel" });
            #endregion

            #region documentSecurity1
            xw.WriteElement(new Ap.DocumentSecurity() { Text = "0" });
            #endregion

            #region scaleCrop1
            xw.WriteElement(new Ap.ScaleCrop() { Text = "false" });
            #endregion

            #region headingPairs1
            xw.WriteStartElement(new Ap.HeadingPairs());

            xw.WriteStartElement(new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U });

            xw.WriteStartElement(new Vt.Variant());
            xw.WriteElement(new Vt.VTLPSTR() { Text = "Hojas de cálculo" });
            xw.WriteEndElement();

            xw.WriteStartElement(new Vt.Variant());
            xw.WriteElement(new Vt.VTInt32() { Text = _Doc.Sheets.Count.ToString() });
            xw.WriteEndElement();

            xw.WriteEndElement();

            xw.WriteEndElement();
            #endregion

            #region titlesOfParts1
            xw.WriteStartElement(new Ap.TitlesOfParts());
            xw.WriteStartElement(new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (uint)_Doc.Sheets.Count });
            _Doc.Sheets.ForEach(st => xw.WriteElement(new Vt.VTLPSTR(st.TabName)));
            xw.WriteEndElement();
            xw.WriteEndElement();
            #endregion

            #region company1
            xw.WriteElement(new Ap.Company() { Text = _Doc.PackageProperties == null ? "" : _Doc.PackageProperties.Company });
            #endregion

            #region linksUpToDate1
            xw.WriteElement(new Ap.LinksUpToDate() { Text = "false" });
            #endregion

            #region sharedDocument1
            xw.WriteElement(new Ap.SharedDocument() { Text = "false" });
            #endregion

            #region hyperlinksChanged1
            xw.WriteElement(new Ap.HyperlinksChanged() { Text = "false" });
            #endregion

            #region applicationVersion1
            xw.WriteElement(new Ap.ApplicationVersion() { Text = "15.0300" });
            #endregion

            xw.WriteEndElement();
            xw.Close();
        }

        #endregion



    }
}
