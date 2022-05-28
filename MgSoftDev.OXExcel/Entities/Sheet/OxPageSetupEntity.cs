using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    internal class OxPageSetupEntity
    {
        public uint Scale { get; set; }
        public OxPageSetupOrientations PageSetupOrientation { get; set; }
        public bool BlackAndWhite { get; set; }
        public OxPrintCellComments PrintCellComments { get; set; }
        public uint Copies { get; set; }
        public bool Draft { get; set; }
        public OxPrintErrors PrintError { get; set; }
        public uint FirstPageNumber { get; set; }
        public bool UseFirstPageNumber { get; set; }
        public uint FitToHeight { get; set; }
        public uint FitToWidth { get; set; }
        public uint HorizontalDpi { get; set; }
        public uint VerticalDpi { get; set; }
        public OxPageOrders PageOrder { get; set; }
        public bool UsePrinterDefaults { get; set; }
        public uint PaperSize { get; set; }        
        public uint PaperHeight { get; set; }
        public uint PaperWidth { get; set; }

    }
}
