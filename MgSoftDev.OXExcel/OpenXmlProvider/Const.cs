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

        /// <summary>Índice de cada formato distinto; evita recorrer la lista comparando formato por formato en cada celda.</summary>
        internal static Dictionary<OxCellFormartEntity, int> FormatIndexes = new Dictionary<OxCellFormartEntity, int>();
        internal static List<string> StringShareds;

        /// <summary>Índice de cada cadena compartida; evita recorrer la lista completa por celda.</summary>
        internal static Dictionary<string, int> StringSharedIndexes = new Dictionary<string, int>();
        internal static List<OxHyperlinkEntity> Hyperlinks;
        internal static UniqueList<string> UniqueValuesList;
        internal static UniqueList<Type> TypeList;

        /// <summary>
        /// Arma las tablas completas en memoria antes de escribir la hoja, como se hacía antes de la escritura en
        /// streaming. Es solo un respaldo; se expone en <see cref="OxExcelDocument.MaterializeTableRows"/>.
        /// </summary>
        internal static bool MaterializeTableRows;

        internal static void Clean()
        {
            margetCells?.Clear();
            Formats?.Clear();
            FormatIndexes?.Clear();
            StringShareds?.Clear();
            StringSharedIndexes?.Clear();
            Hyperlinks?.Clear();
            UniqueValuesList?.Clear();
            TypeList?.Clear();
        }
    }
}
