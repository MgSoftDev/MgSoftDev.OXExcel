namespace MgSoftDev.OXExcel.Entities.ColsRowsCells
{
    public class OxHyperlinkEntity
    {
        internal Uri Uri { get; set; }
        internal string ToolTip { get; set; }
        internal uint Row { get; set; }
        internal uint Column { get; set; }
        internal string Location { get; set; }
        public OxHyperlinkEntity(Uri url, string toolTip)
        {
            ToolTip = toolTip;
            Uri = url;
        }

        public  OxHyperlinkEntity(string excelReference, string toolTip)
        {
            ToolTip = toolTip;
            Location = excelReference;
        }
    }
}
