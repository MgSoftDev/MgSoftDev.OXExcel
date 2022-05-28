using MgSoftDev.OXExcel.Entities.Sheet;
using Color = System.Drawing.Color;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxSheetPropertiesFactory
    {
        internal readonly OxSheetPropertiesEntity Properties;

        public OxSheetPropertiesFactory()
        {
            Properties = new OxSheetPropertiesEntity();
            EnableFormatConditionsCalculation(false);
            FilterMode(false);
            Published(false);
            Properties.TabColor = null;
            Properties.OutlineProperties = null;
        }

        public OxSheetPropertiesFactory EnableFormatConditionsCalculation(bool value = true)
        {
            Properties.EnableFormatConditionsCalculation = value;
            return this;
        }
        public OxSheetPropertiesFactory FilterMode(bool value = true)
        {
            Properties.FilterMode = value;
            return this;
        }
        public OxSheetPropertiesFactory Published(bool value = true)
        {
            Properties.Published = value;
            return this;
        }

        public OxSheetPropertiesFactory TabColor(Color rgb, double tint=0.0)
        {
            Properties.TabColor = new OxTabColorEntity() {Auto = false,Rgb =rgb ,Tint = tint};
            return this;
        }

        public OxSheetPropertiesFactory OutlineProperties(OxOutlinePropertiesFactory outline)
        {
            Properties.OutlineProperties = outline.OutlineProperties;
            return this;
        }

        public OxSheetPropertiesFactory OutlineProperties(Action<OxOutlinePropertiesFactory> outlineAction)
        {
            var f = new OxOutlinePropertiesFactory();
            outlineAction(f);
            Properties.OutlineProperties = f.OutlineProperties;
            return this;
        }

        

    }
}
