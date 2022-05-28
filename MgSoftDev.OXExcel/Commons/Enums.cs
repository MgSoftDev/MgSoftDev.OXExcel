namespace MgSoftDev.OXExcel.Commons
{
    public enum OxSheetVisibilities
    {
        Visible,
        Hidden,
        VeryHidden
    }

    public enum OxPageSetupOrientations
    {
        Default,
        Portrait,
        Landscape
    }

    public enum OxPrintCellComments // CellCommentsValues
    {
        None,
        AsDisplayed,
        AtEnd
    }

    public enum OxPrintErrors //PrintErrorValues
    {
        Blank,
        Dash,
        Displayed,
        Na,
    }

    public enum OxPageOrders // PageOrderValues
    {
        DownThenOver,
        OverThenDown
    }


    public enum OxCalculateModes
    {
        Manual,
        Auto,
        AutoNoTable
    }

    public enum OxDocumentTypes
    {
        Template,
        MacroEnabledWorkbook,
        MacroEnabledTemplate,
        AddIn,
        Workbook
    }

    public enum OxSheetViews
    {
        PageLayout,
        Normal,
        PageBreakPreview
    }
    public enum OxPaperSizeDefault : uint
    {
        Letter=1,
        Legal = 5,
        Standard1 = 45,
        Standard2 = 16,
        Standard3 = 17,
        Standard4 = 46,
        Standard5 = 44,
        SuperA_A4 = 57,
        A2 = 66,
        A3 = 8,
        A3Extra = 63,
        A3ExtraTransverse = 68,
        A7Transverse = 67,
        A4 = 9,
        A4Extra = 53,
        A4Plus = 60,
        A4Transverse = 55,
        A4Small = 10, 
        A5 = 11,
        A5Extra = 64,
        A5Transverse = 61,
        SuperB_A3 = 58,
        B4 = 12,
        B5 = 13,
        B5Extra = 65,
        JisB5Transverse = 62,
        C = 42,
        D = 25,
        // continuar lista en https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.drawing.charts.pagesetup.aspx



    }


    public enum OxCellTypeValues
    {
        Default,
        String,
        Number,
        Error,
        SharedString,
        InlineString,
        ///<summary>This item is only available in Office2010</summary>
        Date
    }

    public enum OxUnderlines
    {
        None,
        Single,
        Double,
        SingleAccounting,
        DoubleAccounting,
    }

    public enum OxVerticalAlignments
    {
        Baseline,
        Superscript,
        Subscript
    }

    public enum OxFontSchemes
    {
        None,
        Major,
        Minor
    }

    public enum OxPatterns
    {
        None,
        Solid,
        MediumGray,
        DarkGray,
        LightGray,
        DarkHorizontal,
        DarkVertical,
        DarkDown,
        DarkUp,
        DarkGrid,
        DarkTrellis,
        LightHorizontal,
        LightVertical,
        LightDown,
        LightUp,
        LightGrid,
        LightTrellis,
        Gray125,
        Gray0625,
    }

    public enum OxGradients
    {
        Linear,
        Path
    }

    public enum OxBorderStyles
    {
        None,
        Thin,
        Medium,
        Dashed,
        Dotted,
        Thick,
        Double,
        Hair,
        MediumDashed,
        DashDot,
        MediumDashDot,
        DashDotDot,
        MediumDashDotDot,
        SlantDashDot
    }

    public enum OxTextHorizontalAlignments
    {
        General,
        Left,
        Center,
        Right,
        Fill,
        Justify,
        CenterContinuous,
        Distributed
    }

    public enum OxTextVerticalAlignments
    {
        Top,
        Center,
        Bottom,
        Justify,
        Distributed
    }

    public enum OxTableType
    {
        Excel,
        Skeleton,
    }

    public enum OxFilterOperators
    {
        Equal,
        LessThan,
        LessThanOrEqual,
        NotEqual,
        GreaterThanOrEqual,
        GreaterThan,

        StartWith,
        EndWith,
        Contrains,
        NotContrains
    }

    public enum OxCustomFilterCondition
    {
        And,
        Or,
        None
    }

    public enum TotalsRowFormulas
    {
        None,
        Sum,
        Minimum,
        Maximum,
        Average,
        Count,
        CountNumbers,
        StandardDeviation,
        Variance,
        Custom
    }
}
