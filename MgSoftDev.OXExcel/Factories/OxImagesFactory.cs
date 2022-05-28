using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxImagesFactory
    {
        internal readonly List<OxImageEntity> Images;

        public OxImagesFactory()
        {
            Images = new List<OxImageEntity>();
        }

        public OxImageFactory Add(OxRangeEntity range)
        {
            var res = new OxImageFactory(range);
            Images.Add(res.Image);
            return res;
        }

    }

}

