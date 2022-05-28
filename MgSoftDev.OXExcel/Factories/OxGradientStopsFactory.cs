using System.Drawing;
using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxGradientStopsFactory
    {
        internal readonly List<OxGradientStopEntity> GradientFills;

        internal OxGradientStopsFactory(List<OxGradientStopEntity> gradientFills)
        {
            GradientFills = gradientFills;
        }

        public OxGradientStopsFactory Add(Color color, double position)
        {
            GradientFills.Add(new OxGradientStopEntity {Color = color,Position = position});
            return this;
        }
    }
}
