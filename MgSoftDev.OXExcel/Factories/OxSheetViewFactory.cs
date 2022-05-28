using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Sheet;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxSheetViewFactory
    {
        internal readonly OxSheetViewEntity SheetView;

        public OxSheetViewFactory()
        {
            SheetView = new OxSheetViewEntity();
            ShowFormulas(false);
            HideGridLines(false);
            ShowOutlineSymbols(false);
            HideRowColHeaders(false);
            ShowRuler(false);
            ShowWhiteSpace(false);
            HideZeros(false);
            TabSelected(false);
            WindowProtection(false);
            ZoomScale(100);
            ZoomScaleNormal(0);
            ZoomScalePageLayoutView(0);
            ZoomScaleSheetLayoutView(0);
            ViewSheet(OxSheetViews.Normal);
        }
        

        public OxSheetViewFactory ShowFormulas(bool value = true)
        {
            SheetView.ShowFormulas = value;
            return this;
        }

        public OxSheetViewFactory HideGridLines(bool value = true)
        {
            SheetView.ShowGridLines = !value;
            return this;
        }

        public OxSheetViewFactory ShowOutlineSymbols(bool value = true)
        {
            SheetView.ShowOutlineSymbols = value;
            return this;
        }

        public OxSheetViewFactory HideRowColHeaders(bool value = true)
        {
            SheetView.ShowRowColHeaders = !value;
            return this;
        }

        public OxSheetViewFactory ShowRuler(bool value = true)
        {
            SheetView.ShowRuler = value;
            return this;
        }

        public OxSheetViewFactory ShowWhiteSpace(bool value = true)
        {
            SheetView.ShowWhiteSpace = value;
            return this;
        }

        public OxSheetViewFactory HideZeros(bool value = true)
        {
            SheetView.ShowZeros = !value;
            return this;
        }

        public OxSheetViewFactory TabSelected(bool value = true)
        {
            SheetView.TabSelected = value;
            return this;
        }
        

        public OxSheetViewFactory WindowProtection(bool value = true)
        {
            SheetView.WindowProtection = value;
            return this;
        }

        public OxSheetViewFactory ZoomScale(uint value)
        {
            SheetView.ZoomScale = value;
            return this;
        }

        public OxSheetViewFactory ZoomScaleNormal(uint value)
        {
            SheetView.ZoomScaleNormal = value;
            return this;
        }

        public OxSheetViewFactory ZoomScalePageLayoutView(uint value)
        {
            SheetView.ZoomScalePageLayoutView = value;
            return this;
        }

        public OxSheetViewFactory ZoomScaleSheetLayoutView(uint value)
        {
            SheetView.ZoomScaleSheetLayoutView = value;
            return this;
        }

        public OxSheetViewFactory ViewSheet(OxSheetViews value)
        {
            SheetView.SheetView = value;
            return this;

        }

        public OxSheetViewFactory PaneFrozen(uint col, uint row)
        {
            SheetView.PaneFrozenReference = col.ToReferenceAlfa() + row;
            return this;
        }

        public OxSheetViewFactory PaneFrozen(string col, uint row)
        {
            SheetView.PaneFrozenReference = col + row;
            return this;
        }

        public OxSheetViewFactory PaneFrozen(string excelReference)
        {
            SheetView.PaneFrozenReference = excelReference;
            return this;
        }
    }
}
