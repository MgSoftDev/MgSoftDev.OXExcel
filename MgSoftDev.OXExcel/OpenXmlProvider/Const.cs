using System.Globalization;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.OpenXmlProvider.Models;

namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal  static class Const
    {
        internal static CultureInfo CultureData =  Thread.CurrentThread.CurrentCulture;
        internal static uint MinRowIndex = uint.MaxValue;
        internal static uint MaxRowIndex = 0;
        internal static uint MinCellIndex = uint.MaxValue;
        internal static uint MaxCellIndex = 0;

        internal static List<string> margetCells = new List<string>();

        internal static uint GlobalIndextable;


        internal static List<OxCellFormartEntity> Formats;
        internal static List<string> StringShareds;
        internal static List<OxHyperlinkEntity> Hyperlinks;
        internal static UniqueList<string> UniqueValuesList;
        internal static UniqueList<Type> TypeList;

        internal static void Clean()
        {
            margetCells?.Clear();
            Formats?.Clear();
            StringShareds?.Clear();
            Hyperlinks?.Clear();
            UniqueValuesList?.Clear();
            TypeList.Clear();
        }
    }
}
