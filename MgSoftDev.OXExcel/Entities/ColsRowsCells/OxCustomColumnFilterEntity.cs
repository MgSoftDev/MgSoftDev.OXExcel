using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.ColsRowsCells
{
    internal class OxCustomColumnFilterEntity
    {
        public string Val { get; set; }
        public OxFilterOperators Operator { get; set; }
        public OxCustomFilterCondition Condition { get; set; }
        public string Val2 { get; set; }
        public OxFilterOperators Operator2 { get; set; }

    }
}
