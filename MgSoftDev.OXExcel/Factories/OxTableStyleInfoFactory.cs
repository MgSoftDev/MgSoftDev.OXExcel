using MgSoftDev.OXExcel.Entities.Table;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxTableStyleInfoFactory
    {
        internal readonly OxTableStyleInfoEntity TableStyle;

        public OxTableStyleInfoFactory()
        {
            TableStyle = new OxTableStyleInfoEntity();
            Name("TableStyleMedium6");
            HideRowStripes(false);
        }

        public OxTableStyleInfoFactory Name(string value)
        {
            TableStyle.Name = value;
            return this;
        }

        public OxTableStyleInfoFactory ShowFirstColumn(bool value = true)
        {
            TableStyle.ShowFirstColumn = value;
            return this;
        }
        public OxTableStyleInfoFactory ShowLastColumn(bool value = true)
        {
            TableStyle.ShowLastColumn = value;
            return this;
        }
        public OxTableStyleInfoFactory HideRowStripes(bool value = true)
        {
            TableStyle.ShowRowStripes = !value;
            return this;
        }
        public OxTableStyleInfoFactory ShowColumnStripes(bool value = true)
        {
            TableStyle.ShowColumnStripes = value;
            return this;
        }

    }
}
