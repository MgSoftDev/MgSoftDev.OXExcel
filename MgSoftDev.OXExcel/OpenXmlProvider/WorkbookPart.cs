using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal partial class OpenXmlExcelProvider
    {


        #region Dinamicas

        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart)
        {
            var workbookPart_Writer = OpenXmlWriter.Create(workbookPart);
            var workbook1           = new Workbook() {MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "x15"}};
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            var alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006");
            var alternateContentChoice1 = new AlternateContentChoice() {Requires = "x15"};
            alternateContentChoice1.Append(
                workbookPart.CreateUnknownElement(
                    "<x15ac:absPath xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\" url=\"C:\\Users\\MiguelFernando\\Desktop\\\" />"));
            alternateContent1.Append(alternateContentChoice1);
            workbook1.Append(new FileVersion()
            {
                ApplicationName = "xl",
                LastEdited = "6",
                LowestEdited = "6",
                BuildVersion = "14420"
            });
            workbook1.Append(new WorkbookProperties() {DefaultThemeVersion = 153222U});
            workbook1.Append(alternateContent1);
            workbook1.Append(
                new BookViews(new WorkbookView() {XWindow = 0, YWindow = 0, WindowWidth = 20490U, WindowHeight = 7755U}));
            workbook1.Append(new Sheets(_Doc.Sheets.ToSheet()));
            workbook1.Append(_Doc.Calculation.ToCalculationProperties());
            workbookPart_Writer.WriteElement(workbook1);
            workbookPart_Writer.Close();
        }

        #endregion

        
        
    }
}
