using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxAlignmentFactory
    {
        internal readonly OxAlignmentEntity Alignment;

        public OxAlignmentFactory()
        {
            Alignment = new OxAlignmentEntity();
            Horizontal(OxTextHorizontalAlignments.Left);
            Vertical(OxTextVerticalAlignments.Top);
            JustifyLastLine(false);
            ShrinkToFit(false);
            Rotation(0);
        }

        internal OxAlignmentFactory(OxAlignmentEntity alignment)
        {
            Alignment = alignment;
        }

        public OxAlignmentFactory Horizontal(OxTextHorizontalAlignments value)
        {
            Alignment.HorizontalAlignment = value;
            return this;
        }
        public OxAlignmentFactory Vertical(OxTextVerticalAlignments value)
        {
            Alignment.VerticalAlignment = value;
            return this;
        }
        public OxAlignmentFactory JustifyLastLine(bool value = true)
        {
            Alignment.JustifyLastLine = value;
            return this;
        }
        public OxAlignmentFactory ShrinkToFit(bool value = true)
        {
            Alignment.ShrinkToFit = value;
            return this;
        }
        public OxAlignmentFactory Rotation(uint value)
        {
            Alignment.Rotation = value;
            return this;
        }
    }
}
