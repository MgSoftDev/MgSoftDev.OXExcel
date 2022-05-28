using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.Commons
{
    public class OxRangeEntity
    {
        public uint FromColumn { get; set; }
        public uint ToColumn { get; set; }
        public uint FromRow { get; set; }
        public uint ToRow { get; set; }

        public OxRangeEntity(uint fromCol, uint fromRow, uint toCol, uint toROw )
        {
            FromColumn = fromCol;
            ToColumn = toCol;
            FromRow = fromRow;
            ToRow = toROw;
        }

        public OxRangeEntity(string fromCol, uint fromRow, string toCol, uint toROw)
        {
            FromColumn = fromCol.ToColIndex();
            ToColumn = toCol.ToColIndex();
            FromRow = fromRow;
            ToRow = toROw;
        }

        public OxRangeEntity(string excelrange)
        {
            var r = excelrange.ToRange();
            FromColumn = r.FromColumn;
            ToColumn = r.ToColumn;
            FromRow = r.FromRow;
            ToRow = r.ToRow;
        }

    }
}
