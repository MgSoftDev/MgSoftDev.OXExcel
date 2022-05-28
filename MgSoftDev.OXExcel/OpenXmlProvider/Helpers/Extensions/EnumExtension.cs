using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions
{
    internal static class EnumExtension
    {
        internal static SpreadsheetDocumentType ToSpreadsheetDocumentType(this OxDocumentTypes docType)
        {
            var result = SpreadsheetDocumentType.Workbook;
            switch (docType)
            {
                case OxDocumentTypes.Template:
                    result = SpreadsheetDocumentType.Template;
                    break;
                case OxDocumentTypes.MacroEnabledWorkbook:
                    result = SpreadsheetDocumentType.MacroEnabledWorkbook;
                    break;
                case OxDocumentTypes.MacroEnabledTemplate:
                    result = SpreadsheetDocumentType.MacroEnabledTemplate;
                    break;
                case OxDocumentTypes.AddIn:
                    result = SpreadsheetDocumentType.AddIn;
                    break;
                case OxDocumentTypes.Workbook:
                    result = SpreadsheetDocumentType.Workbook;
                    break;
            }
            return result;
        }

        internal static CalculateModeValues ToCalculateModeValues(this OxCalculateModes value)
        {
            var result = CalculateModeValues.Auto;
            switch (value)
            {
                case OxCalculateModes.Manual:
                    result = CalculateModeValues.Manual;
                    break;
                case OxCalculateModes.Auto:
                    result = CalculateModeValues.Auto;
                    break;
                case OxCalculateModes.AutoNoTable:
                    result = CalculateModeValues.AutoNoTable;
                    break;
            }
            return result;
        }

        internal static SheetStateValues ToSheetStateValues(this OxSheetVisibilities value)
        {
            SheetStateValues res;
            switch (value)
            {
                case OxSheetVisibilities.Visible:
                    res = SheetStateValues.Visible;
                    break;
                case OxSheetVisibilities.Hidden:
                    res = SheetStateValues.Hidden;
                    break;
                case OxSheetVisibilities.VeryHidden:
                    res = SheetStateValues.VeryHidden;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static  SheetViewValues ToSheetViewValues(this OxSheetViews value)
        {
            SheetViewValues res;
            switch (value)
            {
                case OxSheetViews.PageLayout:
                    res = SheetViewValues.PageLayout;
                    break;
                case OxSheetViews.Normal:
                    res = SheetViewValues.Normal;
                    break;
                case OxSheetViews.PageBreakPreview:
                    res = SheetViewValues.PageBreakPreview;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static PageOrderValues ToPageOrderValues(this OxPageOrders value)
        {
            PageOrderValues res;
            switch (value)
            {
                case OxPageOrders.DownThenOver:
                    res = PageOrderValues.DownThenOver;
                    break;
                case OxPageOrders.OverThenDown:
                    res = PageOrderValues.OverThenDown;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static OrientationValues ToOrientationValues(this OxPageSetupOrientations values)
        {
            OrientationValues res;
            switch (values)
            {
                case OxPageSetupOrientations.Default:
                    res = OrientationValues.Default;
                    break;
                case OxPageSetupOrientations.Portrait:
                    res = OrientationValues.Portrait;
                    break;
                case OxPageSetupOrientations.Landscape:
                    res = OrientationValues.Landscape;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(values), values, null);
            }
            return res;
        }

        internal static PrintErrorValues ToPrintErrorValues(this OxPrintErrors value)
        {
            PrintErrorValues res;
            switch (value)
            {
                case OxPrintErrors.Blank:
                    res = PrintErrorValues.Blank;
                    break;
                case OxPrintErrors.Dash:
                    res = PrintErrorValues.Dash;
                    break;
                case OxPrintErrors.Displayed:
                    res = PrintErrorValues.Displayed;
                    break;
                case OxPrintErrors.Na:
                    res = PrintErrorValues.NA;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static CellCommentsValues ToCellCommentsValues(this OxPrintCellComments value)
        {
            CellCommentsValues res;
            switch (value)
            {
                case OxPrintCellComments.None:
                    res = CellCommentsValues.None;
                    break;
                case OxPrintCellComments.AsDisplayed:
                    res = CellCommentsValues.AsDisplayed;
                    break;
                case OxPrintCellComments.AtEnd:
                    res = CellCommentsValues.AtEnd;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static CellValues ToCellValues(this OxCellTypeValues value)
        {
            CellValues res;
            switch (value)
            {
                case OxCellTypeValues.Number:
                    res = CellValues.Number;
                    break;
                case OxCellTypeValues.Error:
                    res = CellValues.Error;
                    break;
                case OxCellTypeValues.SharedString:
                    res = CellValues.SharedString;
                    break;
                case OxCellTypeValues.String:
                    res = CellValues.String;
                    break;
                case OxCellTypeValues.InlineString:
                    res = CellValues.InlineString;
                    break;
                case OxCellTypeValues.Date:
                    res = CellValues.Date;
                    break;
                case OxCellTypeValues.Default:
                    res = CellValues.String;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static UnderlineValues ToUnderlineValues(this OxUnderlines value)
        {
            UnderlineValues res;
            switch (value)
            {
                case OxUnderlines.None:
                    res = UnderlineValues.None;
                    break;
                case OxUnderlines.Single:
                    res = UnderlineValues.Single;
                    break;
                case OxUnderlines.Double:
                    res = UnderlineValues.Double;
                    break;
                case OxUnderlines.SingleAccounting:
                    res = UnderlineValues.SingleAccounting;
                    break;
                case OxUnderlines.DoubleAccounting:
                    res = UnderlineValues.DoubleAccounting;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static VerticalAlignmentRunValues ToVerticalAlignmentRunValues(
            this OxVerticalAlignments value)
        {
            VerticalAlignmentRunValues res;
            switch (value)
            {
                case OxVerticalAlignments.Baseline:
                    res = VerticalAlignmentRunValues.Baseline;
                    break;
                case OxVerticalAlignments.Superscript:
                    res = VerticalAlignmentRunValues.Superscript;
                    break;
                case OxVerticalAlignments.Subscript:
                    res = VerticalAlignmentRunValues.Subscript;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static FontSchemeValues ToFontSchemeValues(this OxFontSchemes value)
        {
            FontSchemeValues res;
            switch (value)
            {
                case OxFontSchemes.None:
                    res = FontSchemeValues.None;
                    break;
                case OxFontSchemes.Major:
                    res = FontSchemeValues.Major;
                    break;
                case OxFontSchemes.Minor:
                    res = FontSchemeValues.Minor;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static PatternValues ToPatternValues(this OxPatterns value)
        {
            PatternValues res;
            switch (value)
            {
                case OxPatterns.None:
                    res = PatternValues.None;
                    break;
                case OxPatterns.Solid:
                    res = PatternValues.Solid;
                    break;
                case OxPatterns.MediumGray:
                    res = PatternValues.MediumGray;
                    break;
                case OxPatterns.DarkGray:
                    res = PatternValues.DarkGray;
                    break;
                case OxPatterns.LightGray:
                    res = PatternValues.LightGray;
                    break;
                case OxPatterns.DarkHorizontal:
                    res = PatternValues.DarkHorizontal;
                    break;
                case OxPatterns.DarkVertical:
                    res = PatternValues.DarkVertical;
                    break;
                case OxPatterns.DarkDown:
                    res = PatternValues.DarkDown;
                    break;
                case OxPatterns.DarkUp:
                    res = PatternValues.DarkUp;
                    break;
                case OxPatterns.DarkGrid:
                    res = PatternValues.DarkGrid;
                    break;
                case OxPatterns.DarkTrellis:
                    res = PatternValues.DarkTrellis;
                    break;
                case OxPatterns.LightHorizontal:
                    res = PatternValues.LightHorizontal;
                    break;
                case OxPatterns.LightVertical:
                    res = PatternValues.LightVertical;
                    break;
                case OxPatterns.LightDown:
                    res = PatternValues.LightDown;
                    break;
                case OxPatterns.LightUp:
                    res = PatternValues.LightUp;
                    break;
                case OxPatterns.LightGrid:
                    res = PatternValues.LightGrid;
                    break;
                case OxPatterns.LightTrellis:
                    res = PatternValues.LightTrellis;
                    break;
                case OxPatterns.Gray125:
                    res = PatternValues.Gray125;
                    break;
                case OxPatterns.Gray0625:
                    res = PatternValues.Gray0625;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static BorderStyleValues ToBorderStyleValues(this OxBorderStyles value)
        {
            BorderStyleValues res;
            switch (value)
            {
                case OxBorderStyles.None:
                    res = BorderStyleValues.None;
                    break;
                case OxBorderStyles.Thin:
                    res = BorderStyleValues.Thin;
                    break;
                case OxBorderStyles.Medium:
                    res = BorderStyleValues.Medium;
                    break;
                case OxBorderStyles.Dashed:
                    res = BorderStyleValues.Dashed;
                    break;
                case OxBorderStyles.Dotted:
                    res = BorderStyleValues.Dotted;
                    break;
                case OxBorderStyles.Thick:
                    res = BorderStyleValues.Thick;
                    break;
                case OxBorderStyles.Double:
                    res = BorderStyleValues.Double;
                    break;
                case OxBorderStyles.Hair:
                    res = BorderStyleValues.Hair;
                    break;
                case OxBorderStyles.MediumDashed:
                    res = BorderStyleValues.MediumDashed;
                    break;
                case OxBorderStyles.DashDot:
                    res = BorderStyleValues.DashDot;
                    break;
                case OxBorderStyles.MediumDashDot:
                    res = BorderStyleValues.MediumDashDot;
                    break;
                case OxBorderStyles.DashDotDot:
                    res = BorderStyleValues.DashDotDot;
                    break;
                case OxBorderStyles.MediumDashDotDot:
                    res = BorderStyleValues.MediumDashDotDot;
                    break;
                case OxBorderStyles.SlantDashDot:
                    res = BorderStyleValues.SlantDashDot;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static VerticalAlignmentValues ToVerticalAlignmentValues(this OxTextVerticalAlignments value)
        {
            VerticalAlignmentValues res;
            switch (value)
            {
                case OxTextVerticalAlignments.Top:
                    res = VerticalAlignmentValues.Top;
                    break;
                case OxTextVerticalAlignments.Center:
                    res = VerticalAlignmentValues.Center;
                    break;
                case OxTextVerticalAlignments.Bottom:
                    res = VerticalAlignmentValues.Bottom;
                    break;
                case OxTextVerticalAlignments.Justify:
                    res = VerticalAlignmentValues.Justify;
                    break;
                case OxTextVerticalAlignments.Distributed:
                    res = VerticalAlignmentValues.Distributed;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static HorizontalAlignmentValues ToHorizontalAlignmentValues(this OxTextHorizontalAlignments value)
        {
            HorizontalAlignmentValues res;
            switch (value)
            {
                case OxTextHorizontalAlignments.General:
                     res = HorizontalAlignmentValues.General;
                    break;
                case OxTextHorizontalAlignments.Left:
                     res = HorizontalAlignmentValues.Left;
                    break;
                case OxTextHorizontalAlignments.Center:
                     res = HorizontalAlignmentValues.Center;
                    break;
                case OxTextHorizontalAlignments.Right:
                     res = HorizontalAlignmentValues.Right;
                    break;
                case OxTextHorizontalAlignments.Fill:
                     res = HorizontalAlignmentValues.Fill;
                    break;
                case OxTextHorizontalAlignments.Justify:
                     res = HorizontalAlignmentValues.Justify;
                    break;
                case OxTextHorizontalAlignments.CenterContinuous:
                     res = HorizontalAlignmentValues.CenterContinuous;
                    break;
                case OxTextHorizontalAlignments.Distributed:
                     res = HorizontalAlignmentValues.Distributed;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static EnumValue<FilterOperatorValues> ToFilterOperatorValues(this OxFilterOperators value)
        {
            FilterOperatorValues res;
            switch (value)
            {
                case OxFilterOperators.Equal:
                    res = FilterOperatorValues.Equal;
                    break;
                case OxFilterOperators.LessThan:
                    res = FilterOperatorValues.LessThan;
                    break;
                case OxFilterOperators.LessThanOrEqual:
                    res = FilterOperatorValues.LessThanOrEqual;
                    break;
                case OxFilterOperators.NotEqual:
                    res = FilterOperatorValues.NotEqual;
                    break;
                case OxFilterOperators.GreaterThanOrEqual:
                    res = FilterOperatorValues.GreaterThanOrEqual;
                    break;
                case OxFilterOperators.GreaterThan:
                    res = FilterOperatorValues.GreaterThan;
                    break;
                case OxFilterOperators.StartWith:
                    return null;
                case OxFilterOperators.EndWith:
                    return null;
                case OxFilterOperators.Contrains:
                    return null;
                case OxFilterOperators.NotContrains:
                    return FilterOperatorValues.NotEqual;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
            return res;
        }

        internal static EnumValue<TotalsRowFunctionValues> ToTotalsRowFunctionValues(this TotalsRowFormulas value)
        {
            switch (value)
            {
                case TotalsRowFormulas.None:
                    return null;
                case TotalsRowFormulas.Sum:
                    return TotalsRowFunctionValues.Sum;
                case TotalsRowFormulas.Minimum:
                    return TotalsRowFunctionValues.Minimum;
                case TotalsRowFormulas.Maximum:
                    return TotalsRowFunctionValues.Maximum;
                case TotalsRowFormulas.Average:
                    return TotalsRowFunctionValues.Average;
                case TotalsRowFormulas.Count:
                    return TotalsRowFunctionValues.Count;
                case TotalsRowFormulas.CountNumbers:
                    return TotalsRowFunctionValues.CountNumbers;
                case TotalsRowFormulas.StandardDeviation:
                    return TotalsRowFunctionValues.StandardDeviation;
                case TotalsRowFormulas.Variance:
                    return TotalsRowFunctionValues.Variance;
                case TotalsRowFormulas.Custom:
                    return TotalsRowFunctionValues.Custom;
                
            }
            return null;
        }
    }
}
