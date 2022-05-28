using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    internal class OxSheetViewEntity
    {
        public bool ShowFormulas { get; set; }
        public bool ShowGridLines { get; set; }
        public bool ShowOutlineSymbols { get; set; }
        public bool ShowRowColHeaders { get; set; }
        public bool ShowRuler { get; set; }
        public bool ShowWhiteSpace { get; set; }
        public bool ShowZeros { get; set; }
        public bool TabSelected { get; set; }
        public OxSheetViews SheetView { get; set; }

        public bool WindowProtection { get; set; }
        public uint ZoomScale { get; set; }
        public uint ZoomScaleNormal { get; set; }
        public uint ZoomScalePageLayoutView { get; set; }
        public uint ZoomScaleSheetLayoutView { get; set; }
        public string PaneFrozenReference { get; set; }

    }
}
