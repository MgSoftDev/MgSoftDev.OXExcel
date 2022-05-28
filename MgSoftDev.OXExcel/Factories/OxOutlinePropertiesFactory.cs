using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxOutlinePropertiesFactory
    {
        internal readonly OxOutlinePropertiesEntity OutlineProperties;

        public OxOutlinePropertiesFactory()
        {
            OutlineProperties = new OxOutlinePropertiesEntity();
            ApplyStyles(false);
            OutlineSymbols(false);
            NotSummaryBelow(false);
            NotSummaryRight(false);
        }

        public OxOutlinePropertiesFactory ApplyStyles(bool apply = true)
        {
            OutlineProperties.ApplyStyles = apply;
            return this;
        }

        public OxOutlinePropertiesFactory OutlineSymbols(bool show = true)
        {
            OutlineProperties.ShowOutlineSymbols = show;
            return this;
        }
        public OxOutlinePropertiesFactory NotSummaryBelow(bool apply = false)
        {
            OutlineProperties.SummaryBelow = apply;
            return this;
        }
        public OxOutlinePropertiesFactory NotSummaryRight(bool apply = false)
        {
            OutlineProperties.SummaryRight = apply;
            return this;
        }


    }
}
