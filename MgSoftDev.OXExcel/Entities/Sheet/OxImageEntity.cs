using System.Drawing;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Sheet
{
    internal class OxImageEntity
    {
        public string Name { get; set; }
        public string Uri { get; set; }

        public OxRangeEntity Range { get; set; }

        public RectangleF Rectangle { get; set; }
    }
}
