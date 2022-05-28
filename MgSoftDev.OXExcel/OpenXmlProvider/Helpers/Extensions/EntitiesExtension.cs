using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Document;
using MgSoftDev.OXExcel.Entities.Sheet;
using MgSoftDev.OXExcel.Entities.Table;
using MgSoftDev.OXExcel.Helpers.Extensions;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions
{
    internal static class EntitiesExtension
    {
        internal static CalculationProperties ToCalculationProperties(this OxCalculationEntity value)
        {
            var result = new CalculationProperties()
            {
                CalculationId = 152511U,
                CalculationMode = CalculateModeValues.Manual,
                CalculationOnSave = new BooleanValue(true)
            };
            if (value != null)
                result = new CalculationProperties()
                {
                    CalculationId = 152511U,
                    CalculationMode = value.CalculationMode.ToCalculateModeValues(),
                    CalculationOnSave = value.CalculationOnSave.ToBooleanValue()
                };
            return result;
        }

        internal static List<Sheet> ToSheet(this List<OxSheetEntity> sheets)
        {
            var res = new List<Sheet>();
            sheets.ToList().ForEach(ox => res.Add(new Sheet() { Name = ox.TabName, SheetId = (sheets.IndexOf(ox) + 1U).ToUInt32Value(), Id = $"SId{sheets.IndexOf(ox) + 1}", State = ox.SheetVisibility.ToSheetStateValues() }));
            return res;
        }

        internal static SheetProperties ToSheetProperties(this OxSheetPropertiesEntity value)
        {
            var res = new SheetProperties()
            {
                EnableFormatConditionsCalculation = value.EnableFormatConditionsCalculation,
                Published = value.Published,
                FilterMode = value.FilterMode,
            };
            if (value.OutlineProperties != null)
                res.OutlineProperties = new OutlineProperties()
                {
                    SummaryRight = value.OutlineProperties.SummaryRight,
                    SummaryBelow = value.OutlineProperties.SummaryBelow,
                    ApplyStyles = value.OutlineProperties.ApplyStyles,
                    ShowOutlineSymbols = value.OutlineProperties.ShowOutlineSymbols
                };
            if (value.TabColor != null)
                res.TabColor = new TabColor()
                {
                    Tint = value.TabColor.Tint,
                    Rgb = value.TabColor.Rgb.ToHexFormat(),
                    Auto = value.TabColor.Auto
                };

            return res;
        }

        internal static SheetView ToSheetView(this OxSheetViewEntity value)
        {
            if(value == null)
                return new SheetView() { ShowGridLines = true,TabSelected = false, WorkbookViewId =0U, ZoomScale =100U, ZoomScaleNormal  =100U};
            var res = new SheetView
            {
                WorkbookViewId = 0U,
                ShowFormulas = value.ShowFormulas.ToBooleanValue(),
                ShowGridLines = value.ShowGridLines.ToBooleanValue(),
                ShowOutlineSymbols = value.ShowOutlineSymbols.ToBooleanValue(),
                ShowRowColHeaders = value.ShowRowColHeaders.ToBooleanValue(),
                ShowRuler = value.ShowRuler.ToBooleanValue(),
                ShowWhiteSpace = value.ShowWhiteSpace.ToBooleanValue(),
                ShowZeros = value.ShowZeros.ToBooleanValue(),
                TabSelected = value.TabSelected.ToBooleanValue(),
                WindowProtection = value.WindowProtection.ToBooleanValue(),
                View = value.SheetView.ToSheetViewValues(),
                ZoomScale = value.ZoomScale,
                ZoomScaleNormal = value.ZoomScaleNormal,
                ZoomScalePageLayoutView = value.ZoomScalePageLayoutView,
                ZoomScaleSheetLayoutView = value.ZoomScaleSheetLayoutView
            };
            if(!string.IsNullOrEmpty(value.PaneFrozenReference))
                res.Append(new Pane() { HorizontalSplit = value.PaneFrozenReference.GetCol()-1, VerticalSplit = value.PaneFrozenReference.GetRow()-1, TopLeftCell = value.PaneFrozenReference, ActivePane = PaneValues.BottomRight, State = PaneStateValues.Frozen });
            return res;
        }

        internal static PageMargins ToPageMargins(this OxPageMarginsEntity value)
        {
            if (value == null)
                return new PageMargins()
                {
                    Left = 0.15D,
                    Right = 0.15D,
                    Top = 0.15D,
                    Bottom = 0.15D,
                    Header = 0.3D,
                    Footer = 0.3D
                };
            return new PageMargins()
            {
                Header = value.Header,
                Left = value.Left,
                Footer = value.Footer,
                Right = value.Right,
                Top = value.Top,
                Bottom = value.Bottom,
            };
        }

        internal static PageSetup ToPageSetup(this OxPageSetupEntity value)
        {
            return new PageSetup
            {
                Id = "rId1",
                PaperSize = value.PaperSize,
                FitToHeight = value.FitToHeight,
                FitToWidth = value.FitToWidth,
                PageOrder = value.PageOrder.ToPageOrderValues(),
                UsePrinterDefaults = value.UsePrinterDefaults,
                UseFirstPageNumber = value.UseFirstPageNumber,
                Scale = value.Scale,
                Draft = value.Draft,
                VerticalDpi = value.VerticalDpi,
                Copies = value.Copies,
                Orientation = value.PageSetupOrientation.ToOrientationValues(),
                HorizontalDpi = value.HorizontalDpi,
                BlackAndWhite = value.BlackAndWhite,
                FirstPageNumber = value.FirstPageNumber,
                Errors = value.PrintError.ToPrintErrorValues(),
                CellComments = value.PrintCellComments.ToCellCommentsValues(),

            };
        }


        internal static TableStyleInfo ToTableStyleInfo(this OxTableStyleInfoEntity value) => new TableStyleInfo {Name = value.Name, ShowFirstColumn = value.ShowFirstColumn, ShowLastColumn = value.ShowLastColumn, ShowRowStripes = value.ShowRowStripes , ShowColumnStripes = value.ShowColumnStripes};

        internal static TableColumn ToTableColumn(this OxTableColumnsEntity value, uint colIndex)
        {
            var res = new TableColumn()
            {
                Id = colIndex,
                Name = value.Header,
            };
            if (value.TotalRow == null) return res;
            res.TotalsRowLabel =value.TotalRow.RowFormula == TotalsRowFormulas.None? value.TotalRow.TotalsRowLabel: null;
            res.TotalsRowFunction = value.TotalRow.IncludeHidden?new EnumValue<TotalsRowFunctionValues>(TotalsRowFunctionValues.Custom) : value.TotalRow.RowFormula.ToTotalsRowFunctionValues();
            if (!string.IsNullOrEmpty(value.TotalRow.CustomFormula))
                res.Append(new TotalsRowFormula(value.TotalRow.CustomFormula) );
            else if(value.TotalRow.IncludeHidden)
                res.Append(new TotalsRowFormula(value.GetSubTotalFormula()));
            return res;
        }
    }

    
}
