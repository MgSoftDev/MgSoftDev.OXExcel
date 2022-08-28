using System.Drawing;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxImageFactory
    {
        internal readonly OxImageEntity Image;

        public OxImageFactory(OxRangeEntity range)
        {
            Image = new OxImageEntity
            {
              Range = range
            };
            Rectangle(new RectangleF(0, 0, 100, 100));
        }

        public OxImageFactory Url(string value)
        {
            Image.Id         = value;
            Image.Uri        = value;
            Image.Extension  = Path.GetExtension(value);
            Image.ImageBytes = null;
            return this;
        }
        public OxImageFactory ArrayBytes(byte[] bytes)
        {
            Image.Id         = Guid.NewGuid().ToString();
            Image.Uri        = null;
            Image.ImageBytes = bytes;
            Image.Extension  = "png";
            return this;
        }
        public OxImageFactory Name(string value)
        {
            Image.Name = value;
            return this;
        }
        public OxImageFactory Rectangle(RectangleF value)
        {
            Image.Rectangle = value;
            return this;
        }
    }
}
