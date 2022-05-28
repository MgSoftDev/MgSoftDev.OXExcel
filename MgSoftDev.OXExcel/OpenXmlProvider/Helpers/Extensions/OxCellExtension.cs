using System.Text.RegularExpressions;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Table;

namespace MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions
{
    internal static class OxCellExtension
    {
        internal static uint GetRowSpanOrDefault(this IEnumerable<OxCellEntity> value)
        {
            var cells = value.ToList();
            return cells.Count>0 ? cells.Max(m => m.Row) : 1;
        }
        internal static string GetValueCleaned(this OxCellEntity value) => Regex.Replace(value.Value, @"[^\x09\x0A\x0D\x20-\xD7FF\xE00-\xFFFD\x10000-x10FFFF]", "");
        internal static string GetValueCleaned(this OxTableCellEntity value) => Regex.Replace(Const.UniqueValuesList.GetValue(value.Value) , @"[^\x09\x0A\x0D\x20-\xD7FF\xE00-\xFFFD\x10000-x10FFFF]", "");
        
       
    }
}
