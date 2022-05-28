using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Factories;

namespace MgSoftDev.OXExcel.Entities.Table
{
    internal class OxTableColumnsEntity
    {
        public string Header { get; set; }
        public uint Size { get; set; }
        public string PropertyPath { get; set; }
        public Type OriginType { get; set; }
        public string DefaultFormulaValue { get; set; }
        public object DefaultValue { get; set; }
        public OxCellTypeValues CellTypeValue { get; set; }
        public bool ShowPhonetic{ get; set; }
        public OxCellFormartEntity CellFormart { get; set; }
        public OxCellFormartEntity HeaderCellFormart { get; set; }
        public bool IsFormula { get; set; }
        public Func<OxTableColumnTemplateEntity,object> TemplateValue { get; set; }


        public Func<OxTableColumnTemplateEntity, OxCellFormartFactory> TemplateFormat;
        public OxCustomColumnFilterEntity CustomColumnFilter { get; set; }
        public List<OxColumnFilterEntity> ColumnFilter { get; set; }
        public OxTableColumnTotalRowEntity TotalRow { get; set; }
        public Func<OxTableColumnHyperlinkTemplateEntity, OxHyperlinkEntity> HyperlinkTemplate { get; set; }

        public int Order { get; set; }
    }
}
