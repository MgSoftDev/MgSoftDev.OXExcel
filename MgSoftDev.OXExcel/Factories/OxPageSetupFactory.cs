using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxPageSetupFactory
    {

        internal readonly OxPageSetupEntity PageSetup;
        public OxPageSetupFactory()
        {
            PageSetup = new OxPageSetupEntity();
            Scale(100);
            Orientation(OxPageSetupOrientations.Default);
            BlackAndWhite(false);
            PrintCellComments(OxPrintCellComments.None);
            Copies(0);
            Draft(false);
            PrintError(OxPrintErrors.Na);
            UseFirstPageNumber(false);
            FirstPageNumber(1);
            FitToHeight(0);
            FitToWidth(0);
            HorizontalDpi(0);
            VerticalDpi(0);
            PageOrder(OxPageOrders.OverThenDown);
            UsePrinterDefaults(true);
            PaperSize(OxPaperSizeDefault.Letter);
            PaperHeight(0);
            PaperWidth(0);
        }

        public OxPageSetupFactory Scale(uint scale)
        {
            PageSetup.Scale = scale;
            return this;
        }
        public OxPageSetupFactory Orientation(OxPageSetupOrientations orientation)
        {
            PageSetup.PageSetupOrientation = orientation;
            return this;
        }
        public OxPageSetupFactory BlackAndWhite(bool blackAndWhite = true)
        {
            PageSetup.BlackAndWhite = blackAndWhite;
            return this;
        }
        public OxPageSetupFactory PrintCellComments(OxPrintCellComments printCellComment)
        {
            PageSetup.PrintCellComments = printCellComment;
            return this;
        }
        public OxPageSetupFactory Copies(uint copies)
        {
            PageSetup.Copies = copies;
            return this;
        }
        public OxPageSetupFactory Draft(bool draft)
        {
            PageSetup.Draft = draft;
            return this;
        }
public OxPageSetupFactory PrintError(OxPrintErrors printError )
        {
            PageSetup.PrintError = printError;
            return this;
        }

        public OxPageSetupFactory UseFirstPageNumber(bool useFirstPageNumber)
        {
            PageSetup.UseFirstPageNumber = useFirstPageNumber;
            return this;
        }

        public OxPageSetupFactory FirstPageNumber(uint firstPageNumber)
        {
            PageSetup.FirstPageNumber = firstPageNumber;
            return this;
        }
        public OxPageSetupFactory FitToHeight(uint fitToHeight)
        {
            PageSetup.FitToHeight = fitToHeight;
            return this;
        }

        public OxPageSetupFactory FitToWidth(uint fitToWidth)
        {
            PageSetup.FitToWidth = fitToWidth;
            return this;
        }

        public OxPageSetupFactory HorizontalDpi(uint horizontalDpi)
        {
            PageSetup.HorizontalDpi = horizontalDpi;
            return this;
        }

        public OxPageSetupFactory VerticalDpi(uint verticalDpi)
        {
            PageSetup.VerticalDpi = verticalDpi;
            return this;
        }


        public OxPageSetupFactory PageOrder(OxPageOrders pageOrder)
        {
            PageSetup.PageOrder = pageOrder;
            return this;
        }

        public OxPageSetupFactory UsePrinterDefaults(bool usePrinterDefaults)
        {
            PageSetup.UsePrinterDefaults = usePrinterDefaults;
            return this;
        }

        public OxPageSetupFactory PaperSize(OxPaperSizeDefault paperSize)
        {
            PageSetup.PaperSize = (uint)paperSize;
            return this;
        }

        public OxPageSetupFactory PaperHeight(uint paperHeight)
        {
            PageSetup.PaperHeight = paperHeight;
            return this;
        }

        public OxPageSetupFactory PaperWidth(uint paperWidth)
        {
            PageSetup.PaperWidth = paperWidth;
            return this;
        }




    }
}
