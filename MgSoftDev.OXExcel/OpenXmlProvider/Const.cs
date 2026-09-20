using System.Globalization;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.OpenXmlProvider.Models;

namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal  static class Const
    {
        internal static CultureInfo CultureData =  Thread.CurrentThread.CurrentCulture;
        // La extensión de la hoja, sus celdas combinadas y sus hipervínculos viven en OxRowsCellCollection,
        // porque son de cada hoja y no del documento.

        internal static uint GlobalIndextable;


        internal static List<OxCellFormartEntity> Formats;

        /// <summary>Índice de cada formato distinto; evita recorrer la lista comparando formato por formato en cada celda.</summary>
        internal static Dictionary<OxCellFormartEntity, int> FormatIndexes = new Dictionary<OxCellFormartEntity, int>();
        internal static List<string> StringShareds;

        /// <summary>Índice de cada cadena compartida; evita recorrer la lista completa por celda.</summary>
        internal static Dictionary<string, int> StringSharedIndexes = new Dictionary<string, int>();
        internal static UniqueList<string> UniqueValuesList;
        internal static UniqueList<Type> TypeList;

        /// <summary>
        /// Arma las tablas completas en memoria antes de escribir la hoja, como se hacía antes de la escritura en
        /// streaming. Es solo un respaldo; se expone en <see cref="OxExcelDocument.MaterializeTableRows"/>.
        /// </summary>
        internal static bool MaterializeTableRows;

        internal static void Clean()
        {
            Formats?.Clear();
            FormatIndexes?.Clear();
            StringShareds?.Clear();
            StringSharedIndexes?.Clear();
            UniqueValuesList?.Clear();
            TypeList?.Clear();
        }
    }
}
