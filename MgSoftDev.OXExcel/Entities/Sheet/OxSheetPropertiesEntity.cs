namespace MgSoftDev.OXExcel.Entities.Sheet
{
    internal class OxSheetPropertiesEntity
    {
        public bool EnableFormatConditionsCalculation { get; set; }
        public bool FilterMode { get; set; }
        public bool Published { get; set; }

        public OxTabColorEntity TabColor { get; set; }
        public OxOutlinePropertiesEntity OutlineProperties { get; set; }


    }
}
