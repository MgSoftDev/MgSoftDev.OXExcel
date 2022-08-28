using System.Drawing;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    internal class OxImageEntity
    {
        public string Id         { get; set; }
        public string Name       { get; set; }
        public string Uri        { get; set; }
        public byte[] ImageBytes { get; set; }
        public string Extension  { get; set; }

        public OxRangeEntity Range { get; set; }

        public RectangleF Rectangle { get; set; }
    }
}
