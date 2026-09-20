using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Entities.Sheet;
using MgSoftDev.OXExcel.Entities.Table;
using MgSoftDev.OXExcel.Factories;
using MgSoftDev.OXExcel.Helpers.Extensions;
using MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using Picture = DocumentFormat.OpenXml.Spreadsheet.Picture;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;
using System.Drawing;

// ReSharper disable PossiblyMistakenUseOfParamsMethod


namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal partial class OpenXmlExcelProvider
    {
        
        

        private void GenerateWorksheetPartContent2(WorksheetPart worksheetPart1,  OxSheetEntity sheet)
        {
            var xw = OpenXmlWriter.Create(worksheetPart1);

            #region Worksheet
            var worksheet1 = new Worksheet();
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            xw.WriteStartElement(worksheet1);

            #region SheetProperties
            // se establesen las propiedas de la Sheet como OutLine, TabColor
            if (sheet.SheetProperties != null)            
                xw.WriteElement(sheet.SheetProperties.ToSheetProperties());
            #endregion

            #region SheetDimension
            // se definen la SheetDimension, es la parte donde se define desde que celda
            // y hasta que celda hay información
            
            
            // Add rows from tables, recorremos las tablas y extraemos las rows. Las tablas solo preparan aquí sus
            // encabezados y su extensión: las filas de datos se generan una a una dentro de <sheetData>, para no
            // tener millones de celdas vivas al mismo tiempo.
            var tableRows = new List<IEnumerable<OxRowCellsEntity>>();

            sheet.Tables.ForEach(f =>
            {
                if (Const.MaterializeTableRows) TableInsertOrUpdate(sheet, f);
                else if (TablePrepare(sheet, f)) tableRows.Add(TableDataRows(sheet, f, false));
            });

            var reference = "A1";

            if( sheet.RowsCellsList.Rows.Count != 0 )
                reference = Const.MinCellIndex.ToReferenceAlfa() + Const.MinRowIndex + ":" + Const.MaxCellIndex.ToReferenceAlfa() + Const.MaxRowIndex;
            
            xw.WriteElement(new SheetDimension() { Reference = reference });
            #endregion


            #region SheetViews
            // en esta parte se define las propiedades de visualizacion como:
            // GridLines, TabSelect, ShowFormulas, ShowRules,ZoomScale, etc.


            xw.WriteElement(new SheetViews(sheet.SheetView.ToSheetView()));            
            #endregion

            #region SheetFormatProperties

            xw.WriteElement(new SheetFormatProperties() { BaseColumnWidth = 10U, DefaultRowHeight = 15D, DyDescent = 0.25D });
            #endregion

            #region Columns
            // en esta parte se definen las propiedades de cada columna como:
            // Style, width, hidden, etc.


            if (sheet.Columns != null && sheet.Columns.Count > 0)
            {
                xw.WriteStartElement(new Columns());
                sheet.Columns.ForEach(c=>xw.WriteElement(ToColumn(c)));
                xw.WriteEndElement();
            }
            #endregion

            #region SheetData
            // en esta parte se crean las celdas y sus rows, basicamente se definen los valores de las celdas

            xw.WriteStartElement(new SheetData());
            
            if ( sheet.RowsCellsList.Rows != null)
            {
                CreateExcelRowsFromOxRows(xw, sheet, MergeRowSources(sheet.RowsCellsList.Rows.Values, tableRows));
            }


            xw.WriteEndElement();
            #endregion


            #region MergeCells
            // En esta parte es donde se combinan celdas

            if (Const.margetCells.Count > 0)
            {
                xw.WriteStartElement(new MergeCells() { Count = (uint)Const.margetCells.Count });
                Const.margetCells.ForEach(mc => xw.WriteElement(new MergeCell() { Reference = mc }));
                xw.WriteEndElement();
            }
            #endregion
            #region Hyperlinks
            // en esta parte sedan de alta los links que te llevan a una Url o a otra parte del documento

            if (Const.Hyperlinks.Count > 0)
            {
                xw.WriteStartElement(new Hyperlinks());
                var hyIndex = 0;
                Const.Hyperlinks.ForEach(f =>
                {
                    var link = new Hyperlink()
                    {
                        Tooltip = f.ToolTip,
                        Reference = f.Column.ToReferenceAlfa() + f.Row,
                    };
                    if (f.Uri != null)
                        link.Id = "rId" + hyIndex++;
                    else
                        link.Location = f.Location;
                    xw.WriteElement(link);
                });
                xw.WriteEndElement();
            }

            #endregion
            #region PageMargins
            // en esta parte se definen los margenes de las paginas

            xw.WriteElement(sheet.PageMargins.ToPageMargins());
            #endregion

           

            #region PageSetup
            // en esta parte de definen las propiedades de imprecion del documento
            if (sheet.PageSetup != null)
                xw.WriteElement(sheet.PageSetup.ToPageSetup());
            #endregion


            #region Cargar Imagenes
            // en esta parte se insertan las imagenes que se importaron en la clase:
            // OpenXmlExcelProvider.CreateParts()

            if (sheet.Images != null && sheet.Images.Count > 0)
                xw.WriteElement(new Drawing() { Id = "rId1" });
            #endregion

            #region BackGround
            // en esta parte se insertan la imagen de fondo que se importo en la clase:
            // OpenXmlExcelProvider.CreateParts()


            if (sheet.BackgroundImage != null)
                xw.WriteElement(new Picture() { Id = "rId2" });
            #endregion



            #region TableParts
            // en esta parte se definen el tipo de tabla y sus rangos de donde y hasta donde abarca 

            // Solo las tablas de tipo Excel generan su parte; si ninguna lo es no se escribe <tableParts>, porque
            // un conteo que no cuadra con los hijos hace que Excel repare el archivo.
            var excelTables = sheet.Tables?.Where(w => w.TableType == OxTableType.Excel).ToList() ?? new List<OxTableEntity>();

            if (excelTables.Count > 0)
            {
                xw.WriteStartElement(new TableParts() { Count = (uint)excelTables.Count });
                excelTables.ForEach(ot =>
                {
                    Const.GlobalIndextable++;
                    xw.WriteElement(new TablePart() {Id = "rIdt" + Const.GlobalIndextable});

                    var tableDefinitionPart1 = worksheetPart1.AddNewPart<TableDefinitionPart>("rIdt" + Const.GlobalIndextable);
                    var xwTable = OpenXmlWriter.Create(tableDefinitionPart1);
                    GenerateTableDefinitionPart1Content(xwTable, Const.GlobalIndextable, ot);
                    xwTable.Close();
                });
                xw.WriteEndElement();
            }
            #endregion

          
            xw.WriteEndElement();



            #endregion
            xw.Close();
        }

        #region Rows Cols cells

        /// <summary>
        /// Genera la tabla completa en memoria. Es el camino anterior a la escritura en streaming; se conserva como
        /// respaldo y se activa con <see cref="Const.MaterializeTableRows"/>.
        /// </summary>
        internal static void TableInsertOrUpdate(OxSheetEntity sheet, OxTableEntity table)
        {
            if (!TablePrepare(sheet, table)) return;

            // Las filas se van guardando en sheet.RowsCellsList; aquí solo hay que agotar el enumerador.
            foreach (var row in TableDataRows(sheet, table, true)) { }
        }

        /// <summary>
        /// Prepara la tabla: autogenera y ordena las columnas, cuenta las filas y escribe en la hoja el renglón de
        /// encabezados y el de totales. Las filas de datos no se generan aquí; las entrega <see cref="TableDataRows"/>
        /// una a la vez. Devuelve false cuando la tabla no tiene filas de datos que escribir.
        /// </summary>
        private static bool TablePrepare(OxSheetEntity sheet, OxTableEntity table)
        {
            // Autogenerate Columns
            if( table.AutoGenerateColumns && table?.DataCollection != null && table.Columns != null  )
            {
                var columns = new OxTableColumnsFactory();

                table.DataCollection?.FirstOrDefault()
                     .GetProperties()
                     .Where( w => w.PropertyType.Namespace == "System" && !table.Columns.Select( s => s.PropertyPath ).Contains( w.Name ) )
                     .ToList()
                     .ForEach( p =>columns.Add( p.Name ).Type( p.PropertyType ).Order(table.Columns.Count+ columns.TableColumns.Count).ExtractAttributes( p ) );

                table.Columns.AddRange( columns.TableColumns );

            }

            if (table?.DataCollection == null || table.Columns == null || table.Columns.Count == 0)
                return false;

            table.Columns = table.Columns.OrderBy( o => o.Order).ToList();
            table.RowsCounts = (uint)(table.DeclaredRowsCount ?? table.DataCollection.LongCount());

            // insert Row of Columns
            var d = table.RowDefinition.CloneFast();
            d.RowIndex     = table.Row;
            d.CustomFormat = false;
            d.Format       = null;
            var cRow = sheet.RowsCellsList.AddAndGet(d);

            OxRowCellsEntity tRow = null;
            // insertar row totales solo 1 si tiene totales activos
            if (table.TotalsRowShow)
            {

                d              = table.RowDefinition.CloneFast();
                d.RowIndex     = table.Row + table.RowsCounts + 1;
                d.CustomFormat = false;
                d.Format       = null;
                tRow = sheet.RowsCellsList.AddAndGet(d);
            }

            // inserta las celdas que van hacer de headers de la tabla
            var cIndex = 0U;
            table.Columns.ForEach(c =>
            {
                var cell = new OxCellEntity()
                {
                    Row = table.Row,
                    Column = table.Column + cIndex,
                    CellFormart = c.HeaderCellFormart,
                    CellTypeValue = OxCellTypeValues.SharedString,
                    OriginType = typeof(string),
                    ShowPhonetic = c.ShowPhonetic,
                    Value = c.Header,
                };
                cRow.Add(cell);

                //insert total rows
                if (table.TotalsRowShow && tRow!= null)
                {
                    var cellT = new OxCellEntity()
                    {
                        // el 1 es por la row de totales
                        Row = table.Row + table.RowsCounts + 1,
                        Column = table.Column + cIndex,
                        CellTypeValue = OxCellTypeValues.String,
                        OriginType = typeof(string),
                        ShowPhonetic = c.ShowPhonetic,
                        Value = "",
                    };
                    if (c.TotalRow != null)
                    {
                        cellT.CellFormart = c.TotalRow.CellFormart;
                        if (c.TotalRow.RowFormula == TotalsRowFormulas.None)
                            cellT.Value = c.TotalRow.TotalsRowLabel;
                        else
                            cellT.Formula = new OxCellFormulaEntity() {Formula = c.GetSubTotalFormula()};
                    }
                    tRow.Add(cellT);
                }

                cIndex++;
            });

            // Las filas de datos ya no pasan por RowsCellsList cuando se escriben en streaming, así que la última
            // fila de la tabla se registra aquí para que <dimension> siga saliendo completa.
            var lastRow = table.Row + table.RowsCounts + (table.TotalsRowShow ? 1U : 0U);
            if (lastRow > Const.MaxRowIndex) Const.MaxRowIndex = lastRow;

            return table.RowsCounts > 0;
        }

        /// <summary>
        /// Entrega las filas de datos de la tabla una por una. Con <paramref name="materialize"/> en true además las
        /// guarda en la hoja (comportamiento anterior); en false la fila se entrega y se descarta, que es lo que
        /// permite escribir tablas de millones de celdas con memoria constante.
        /// </summary>
        private static IEnumerable<OxRowCellsEntity> TableDataRows(OxSheetEntity sheet, OxTableEntity table, bool materialize)
        {
            var lista  = table.DataCollection;
            var source = table.DataSource ?? lista;
            var fcols  = table.Columns.Where(w => w.CustomColumnFilter != null || w.ColumnFilter != null).ToList();

            // El formato de una columna combinado con el de la fila es el mismo para todas sus celdas, así que se
            // calcula una sola vez. Antes se clonaba por celda con un ida y vuelta por JSON, que era el mayor
            // consumo de memoria y CPU de la librería. Solo las columnas con plantilla necesitan copia por celda.
            var baseFormats = table.RowDefinitionTemplate == null ? BuildColumnFormats(table, table.RowDefinition?.Format) : null;
            var rIndex      = 0U;

            try
            {
                foreach (var r in source)
                {
                    rIndex++;

                    // Con una tabla en streaming el rango de la hoja ya se escribió con el total declarado; una fila
                    // de más dejaría el archivo fuera de ese rango.
                    if (rIndex > table.RowsCounts)
                        throw new InvalidOperationException($"La tabla declaró {table.RowsCounts} filas y la secuencia entregó más.");

                    yield return AddTableRow(sheet, table, r, lista, rIndex, fcols, baseFormats, materialize);
                }
            }
            finally
            {
                // GC Lista and DataCollection
                lista.Clear();
            }
        }

        /// <summary>Formato de cada columna ya combinado con el de la fila; se reutiliza en todas las celdas de la columna.</summary>
        private static OxCellFormartEntity[] BuildColumnFormats(OxTableEntity table, OxCellFormartEntity rowFormat)
        {
            var formats = new OxCellFormartEntity[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++) formats[i] = table.Columns[i].CellFormart.CloneFast().Combine(rowFormat);

            return formats;
        }

        private static OxRowCellsEntity AddTableRow(OxSheetEntity sheet, OxTableEntity table, object r, List<object> lista, uint rIndex,
                                                    List<OxTableColumnsEntity> fcols, OxCellFormartEntity[] baseFormats, bool materialize)
        {
            var rowDeff = table.RowDefinition.CloneFast();
            if (table.RowDefinitionTemplate != null)
            {
                rowDeff = table.RowDefinitionTemplate(new OxTableRowDefinitionTemplateEntity
                {
                    Rows          = r,
                    RowDefinition = new OxRowFactory(rowDeff),
                    MasterData    = lista,
                    TableRowIndex = rIndex,
                    SheetRowIndex = table.Row + rIndex,
                }).Row;
            }

            #region Insert Rows

            rowDeff.RowIndex     = table.Row + rIndex;
            rowDeff.CustomFormat = false;
            if (fcols.Count > 0)
                rowDeff.Hidden = r.HiddenForFilter(fcols);

            #endregion

            var row     = materialize ? sheet.RowsCellsList.AddAndGet(rowDeff) : new OxRowCellsEntity { Row = rowDeff };
            var formats = baseFormats ?? BuildColumnFormats(table, rowDeff.Format);
            var cIndex  = 0U;

            for (var i = 0; i < table.Columns.Count; i++)
            {
                var c      = table.Columns[i];
                var format = formats[i];
                var val    = r.GetPropertyVal(c.PropertyPath, c.DefaultValue);
                OxHyperlinkEntity link = null;

                #region Templates

                // Las plantillas reciben el formato y pueden modificarlo, así que se les entrega una copia propia.
                if (c.TemplateValue != null || c.TemplateFormat != null) format = format.CloneFast();

                if (c.TemplateValue != null)
                    val =
                        c.TemplateValue(new OxTableColumnTemplateEntity
                        {
                            Format = new OxCellFormartFactory(format),
                            TableRowIndex = rIndex,
                            SheetRowIndex = table.Row + rIndex,
                            MasterData = lista,
                            Row = r,
                            CellValue = val
                        }).ToExcelValue();
                if (c.TemplateFormat != null)
                    format =
                        c.TemplateFormat(new OxTableColumnTemplateEntity
                        {
                            Format = new OxCellFormartFactory(format),
                            TableRowIndex = rIndex,
                            SheetRowIndex = table.Row + rIndex,
                            MasterData = lista,
                            Row = r,
                            CellValue = val
                        }).Format.Combine(format);
                if (c.HyperlinkTemplate != null)
                {
                    link = c.HyperlinkTemplate(new OxTableColumnHyperlinkTemplateEntity()
                    {
                        TableRowIndex = rIndex,
                        SheetRowIndex = table.Row + rIndex,
                        MasterData = lista,
                        Row = r,
                        CellValue = val
                    });
                    link.Row = table.Row + rIndex;
                    link.Column = table.Column + cIndex;
                }

                #endregion

                var cellType = c.CellTypeValue;
                if (c.TemplateCellType != null)
                    cellType =
                        c.TemplateCellType(new OxTableColumnTemplateEntity
                        {
                            Format = new OxCellFormartFactory(format),
                            TableRowIndex = rIndex,
                            SheetRowIndex = table.Row + rIndex,
                            MasterData = lista,
                            Row = r,
                            CellValue = val
                        });

                var cell = new OxTableCellEntity()
                {
                    Row = table.Row + rIndex,
                    Column = table.Column + cIndex,
                    CellFormart = format,
                    CellTypeValue = cellType,
                    ShowPhonetic = c.ShowPhonetic,
                    Value = Const.UniqueValuesList.Add(val),
                    Hyperlink = link
                };
                if (c.IsFormula)
                {
                    cell.Formula = new OxCellFormulaEntity() { Formula = val };
                    cell.Value = Const.UniqueValuesList.Add(c.DefaultFormulaValue);
                }
                row.Add(cell);
                cIndex++;
            }

            return row;
        }

        /// <summary>
        /// Une las filas que ya tiene la hoja con las que van generando las tablas, en orden ascendente de índice.
        /// Si dos fuentes caen en el mismo renglón se combinan sus celdas para no emitirlo dos veces; ganan las
        /// celdas de la hoja, igual que cuando todo se armaba en memoria.
        /// </summary>
        private static IEnumerable<OxRowCellsEntity> MergeRowSources(IEnumerable<OxRowCellsEntity> sheetRows, List<IEnumerable<OxRowCellsEntity>> tableRows)
        {
            if (tableRows == null || tableRows.Count == 0) return sheetRows;

            var sources = new List<IEnumerable<OxRowCellsEntity>> { sheetRows };
            sources.AddRange(tableRows);

            return MergeRowSourcesIterator(sources);
        }

        private static IEnumerable<OxRowCellsEntity> MergeRowSourcesIterator(List<IEnumerable<OxRowCellsEntity>> sources)
        {
            var enumerators = sources.Select(s => s.GetEnumerator()).ToList();

            try
            {
                var current = enumerators.Select(e => e.MoveNext() ? e.Current : null).ToList();

                while (true)
                {
                    var next = -1;
                    for (var i = 0; i < current.Count; i++)
                        if (current[i] != null && (next < 0 || current[i].Row.RowIndex < current[next].Row.RowIndex)) next = i;

                    if (next < 0) yield break;

                    var row = current[next];
                    current[next] = enumerators[next].MoveNext() ? enumerators[next].Current : null;

                    for (var i = 0; i < current.Count; i++)
                        while (current[i] != null && current[i].Row.RowIndex == row.Row.RowIndex)
                        {
                            foreach (var cell in current[i].Cells.Values) row.Add(cell);
                            current[i] = enumerators[i].MoveNext() ? enumerators[i].Current : null;
                        }

                    yield return row;
                }
            }
            finally
            {
                enumerators.ForEach(e => e.Dispose());
            }
        }


        /// <summary>Máximos de una hoja de Excel; pasarse de ahí genera un archivo que Excel no abre.</summary>
        private const uint MaxExcelRows = 1048576;

        private const uint MaxExcelColumns = 16384;

        private void CreateExcelRowsFromOxRows(OpenXmlWriter xw, OxSheetEntity sheet, IEnumerable<OxRowCellsEntity> rows)
        {

            var firstRow  = true;

            foreach( var r in rows )
            {
                    if (r.Row.RowIndex > MaxExcelRows)
                        throw new InvalidOperationException($"La hoja \"{sheet.TabName}\" llegó a la fila {r.Row.RowIndex} y Excel solo admite {MaxExcelRows} filas. Acorte el rango de datos o repártalos en varias hojas.");

               
                    if (!firstRow) xw.WriteEndElement();
                    xw.WriteStartElement(ToRow(r.Row as OxRowEntity, 1));
                    firstRow = false;
                
                foreach ( var c in r.Cells.Values )
                {
                    if (c.Column > MaxExcelColumns)
                        throw new InvalidOperationException($"La hoja \"{sheet.TabName}\" llegó a la columna {c.Column} y Excel solo admite {MaxExcelColumns} columnas.");

                    
                     
                    #region For cell


                    switch (c)
                    {
                        case OxCellEntity entity:

                            if (entity.Hyperlink != null) Const.Hyperlinks.Add(entity.Hyperlink);

                            try
                            {
                                xw.WriteElement(ToCell(entity));
                            }
                            catch (Exception e)
                            {
                                throw new Exception("Cheque su cadena (" + entity.Value + ")  Row=" + c.Row + "  Col=" + c.Column, e);
                            }

                            break;
                        case OxTableCellEntity entity:

                            if (entity.Hyperlink != null) Const.Hyperlinks.Add(entity.Hyperlink);

                            try
                            {
                                xw.WriteElement(ToCell(entity));
                            }
                            catch (Exception e)
                            {
                                throw new Exception("Cheque su cadena (" + Const.UniqueValuesList.GetValue(entity.Value) + ")  Row=" + c.Row + "  Col=" + c.Column, e);
                            }


                            break;
                    }



                    #endregion
                }
            }

           
            if (!firstRow) xw.WriteEndElement();

            //GC data
           
            
        }

        private Cell ToCell(OxCellEntity value)
        {
            var styleIndex = GetFormartIndex(value.CellFormart);
            if (styleIndex == null && value.Hyperlink != null)
                styleIndex = 1U;
            var res = new Cell
            {
                CellReference = value.Column.ToReferenceAlfa() + value.Row,
                DataType = value.CellTypeValue.ToCellValues(),
                CellValue =
                    value.CellTypeValue == OxCellTypeValues.SharedString
                        ? new CellValue(GetSharedIndex(value.GetValueCleaned()))
                        : new CellValue(value.GetValueCleaned()),
                StyleIndex = styleIndex.HasValue ? (UInt32Value)styleIndex : null,
                CellFormula = value.Formula != null
                    ? new CellFormula() { Text = value.Formula.Formula, CalculateCell = true }
                    : null,
                ShowPhonetic = value.ShowPhonetic ? (BooleanValue)value.ShowPhonetic : null,
            };
            return res;
        }

        private Cell ToCell(OxTableCellEntity value)
        {
            var styleIndex = GetFormartIndex( value.CellFormart);
            if (styleIndex == null && value.Hyperlink != null)
                styleIndex = 1U;
            var res = new Cell
            {
                    CellReference = value.Column.ToReferenceAlfa() + value.Row,
                    DataType      = value.CellTypeValue.ToCellValues(),
                    CellValue =
                            value.CellTypeValue == OxCellTypeValues.SharedString
                                    ? new CellValue(GetSharedIndex(value.GetValueCleaned()))
                                    : new CellValue(value.GetValueCleaned()),
                    StyleIndex = styleIndex.HasValue ? (UInt32Value)styleIndex : null,
                    CellFormula = value.Formula != null
                            ? new CellFormula() { Text = value.Formula.Formula, CalculateCell = true }
                            : null,
                    ShowPhonetic = value.ShowPhonetic ? (BooleanValue)value.ShowPhonetic : null,
            };
            return res;
        }



        private Column ToColumn(OxColumnEntity values)
        {
            var styleIndex = GetFormartIndex(values.Format);
            return new Column
            {
                BestFit = values.BestFit,
                Max = values.Max,
                Min = values.Min,
                OutlineLevel = values.OutlineLevel,
                Hidden = values.Hidden,
                Style = styleIndex.HasValue ? (UInt32Value)styleIndex : null,
                Width = values.Width,
                CustomWidth = values.CustomWidth,
                Collapsed = values.Collapsed,
                Phonetic = values.Phonetic
            };
        }
        private Row ToRow(OxRowEntity row, uint span)
        {
            var styleIndex = GetFormartIndex(row.Format);
            var res = new Row
            {
                RowIndex = row.RowIndex, 
                StyleIndex = styleIndex.HasValue ? (UInt32Value) styleIndex : null,
                CustomFormat = row.CustomFormat ,
                Spans = new ListValue<StringValue>() {InnerText = $"1:{span}"},
                DyDescent = 0.25D,
                Collapsed = row.Collapsed ? (BooleanValue) row.Collapsed : null,
                Hidden = row.Hidden ? (BooleanValue) row.Hidden : null,
                OutlineLevel = row.OutlineLevel > 0 ? (ByteValue) row.OutlineLevel : null,
                ShowPhonetic = row.ShowPhonetic ? (BooleanValue) row.ShowPhonetic : null,
                ThickBot = row.ThickBot ? (BooleanValue) row.ThickBot : null,
                ThickTop = row.ThickTop ? (BooleanValue) row.ThickTop : null
            };

            if (!row.CustomHeight) return res;
            res.CustomHeight = row.CustomHeight;
            res.Height = row.Height;

            return res;
        }

      

        #endregion
        

        #region Images Part
        // Generates content of drawingsPart1.
        private void GenerateDrawingsPart1Content(DrawingsPart drawPart, List<OxImageEntity> images )
        {
            var xw = OpenXmlWriter.Create(drawPart);
            var worksheetDrawing2 = new Xdr.WorksheetDrawing();
            worksheetDrawing2.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            xw.WriteStartElement(worksheetDrawing2);
            var imgDistinct = images.Select(s=>s.Id).Distinct().ToList();
            foreach(var img in images)
            {
                Bitmap bmp;
                    if(img.Uri!= null)
                    bmp = new Bitmap(img.Uri);
                    else
                    {
                        using var stream = new MemoryStream(img.ImageBytes);
                        bmp = new Bitmap(stream);
                    }

                var indexImgSheet = ( imgDistinct.IndexOf(img.Id) +1);
                #region TwoCellAnchor
                xw.WriteStartElement( new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell });

                #region FromMarker
                xw.WriteStartElement( new Xdr.FromMarker());
                xw.WriteElement(new Xdr.ColumnId((img.Range.FromColumn-1).ToString()));
                xw.WriteElement( new Xdr.ColumnOffset(img.Rectangle.X.ToString(CultureInfo.InvariantCulture)));
                xw.WriteElement( new Xdr.RowId((img.Range.FromRow - 1).ToString()));
                xw.WriteElement( new Xdr.RowOffset(img.Rectangle.Y.ToString(CultureInfo.InvariantCulture)));                
                xw.WriteEndElement();
                #endregion

                #region ToMarker
                xw.WriteStartElement(new Xdr.ToMarker());
                xw.WriteElement(new Xdr.ColumnId( img.Range.ToColumn.ToString()));
                xw.WriteElement( new Xdr.ColumnOffset(img.Rectangle.Bottom.ToString(CultureInfo.InvariantCulture)));
                xw.WriteElement( new Xdr.RowId( img.Range.ToRow.ToString()));
                xw.WriteElement( new Xdr.RowOffset( img.Rectangle.Right.ToString(CultureInfo.InvariantCulture)));
                xw.WriteEndElement();
                #endregion

                #region picture2
                xw.WriteStartElement( new Xdr.Picture());

                #region nonVisualPictureProperties2
                xw.WriteStartElement( new Xdr.NonVisualPictureProperties());
                xw.WriteElement( new Xdr.NonVisualDrawingProperties() { Id = (uint)(indexImgSheet + 1), Name = $"Imagen_{indexImgSheet}_{img.Name}"   });

                #region NonVisualPictureDrawingProperties
                xw.WriteStartElement( new Xdr.NonVisualPictureDrawingProperties());
                xw.WriteElement( new A.PictureLocks() { NoChangeAspect = true });
                
                xw.WriteEndElement();
                #endregion

                xw.WriteEndElement();
                #endregion

                #region blipFill2
                xw.WriteStartElement( new Xdr.BlipFill());

                #region blip2
                A.Blip blip2 = new A.Blip() { Embed = "rId" + indexImgSheet };
                blip2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                xw.WriteStartElement(blip2);

                #region BlipExtensionList
                xw.WriteStartElement( new A.BlipExtensionList());

                #region blipExtension2
                xw.WriteStartElement( new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" });
                A14.UseLocalDpi useLocalDpi2 = new A14.UseLocalDpi() { Val = false };
                useLocalDpi2.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                xw.WriteElement(useLocalDpi2);
                xw.WriteEndElement();
                #endregion

                xw.WriteEndElement();
                #endregion

                xw.WriteEndElement();
                #endregion

                #region Stretch
                xw.WriteStartElement( new A.Stretch());
                xw.WriteElement( new A.FillRectangle());
                xw.WriteEndElement();
                #endregion

                xw.WriteEndElement();
                #endregion

                #region ShapeProperties
                xw.WriteStartElement( new Xdr.ShapeProperties());
                #region Transform2D
                xw.WriteStartElement( new A.Transform2D());
                xw.WriteElement( new A.Offset() { X = 0, Y = 0L });
                xw.WriteElement(new A.Extents() { Cx =  bmp.Width * (long)(914400 / bmp.HorizontalResolution), Cy = bmp.Height * (long)(914400 / bmp.VerticalResolution) }); 
                xw.WriteEndElement();
                #endregion
                #region PresetGeometry
                xw.WriteStartElement( new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle });
                xw.WriteElement( new A.AdjustValueList());
                xw.WriteEndElement();
                #endregion
                xw.WriteEndElement();
                #endregion

                xw.WriteEndElement();
                #endregion

                #region ClientData
                xw.WriteElement(new Xdr.ClientData());
                #endregion

                xw.WriteEndElement();
                #endregion
                bmp.Dispose();
            }

            xw.WriteEndElement();
            xw.Close();
        }
        private void GenerateImagePart1Content(ImagePart imagePart1, OxImageEntity img)
        {
            if(img.Uri!= null)
            {
                using var st = new System.IO.FileStream(img.Uri, System.IO.FileMode.Open, System.IO.FileAccess.Read);

                imagePart1.FeedData(st);
            }
            else if (img.ImageBytes!= null)
            {
                using var st = new System.IO.MemoryStream(img.ImageBytes );

                imagePart1.FeedData(st);
            }
        }
        #endregion images part
        #region Tablas def
        private void GenerateTableDefinitionPart1Content(OpenXmlWriter xw, uint indexName, OxTableEntity table)
        {
            var reference = $"{table.Column.ToReferenceAlfa()}{table.Row}:{(table.Column + (uint)table.Columns.LongCount() - 1U).ToReferenceAlfa()}{table.Row + table.RowsCounts}";
            var referenceExtra = $"{table.Column.ToReferenceAlfa()}{table.Row}:{(table.Column + (uint)table.Columns.LongCount() - 1U).ToReferenceAlfa()}{table.Row + table.RowsCounts + (table.TotalsRowShow ? 1U : 0U)}";

            var hIndex = 0;
            var tName = string.IsNullOrEmpty(table.TableName) ? "Tabla" + indexName : table.TableName.Replace(" ", "") + "_";
            xw.WriteStartElement(new Table() { Id = indexName, Name = tName, DisplayName = tName, Reference = referenceExtra, TotalsRowCount = table.TotalsRowShow ? 1U : 0U });
            if (table.AutoFilter)
            {
                var autoFilter = new AutoFilter() { Reference = reference };
                var i = 0;
                table.Columns.ForEach(c =>
                {
                    if (c.CustomColumnFilter == null && c.ColumnFilter == null) return;
                    var filterColumn = new FilterColumn() { ColumnId = (uint)i };
                    if (c.CustomColumnFilter != null)
                    {
                        var customFilters = new CustomFilters();
                        if (c.CustomColumnFilter.Condition == OxCustomFilterCondition.And)
                            customFilters.And = true;
                        customFilters.Append(new CustomFilter() { Operator = c.CustomColumnFilter.Operator.ToFilterOperatorValues(), Val = c.CustomColumnFilter.Val.ApplyOperator(c.CustomColumnFilter.Operator) });
                        if (c.CustomColumnFilter.Condition != OxCustomFilterCondition.None)
                            customFilters.Append(new CustomFilter() { Operator = c.CustomColumnFilter.Operator2.ToFilterOperatorValues(), Val = c.CustomColumnFilter.Val2.ApplyOperator(c.CustomColumnFilter.Operator2) });
                        filterColumn.Append(customFilters);
                    }
                    else if (c.ColumnFilter != null)
                    {
                        var filters = new Filters();
                        c.ColumnFilter.ForEach(f => filters.Append(new Filter() { Val = f.Val }));
                        filterColumn.Append(filters);
                    }
                    autoFilter.Append(filterColumn);
                    i++;
                });
                xw.WriteElement(autoFilter);
            }
            xw.WriteStartElement(new TableColumns() { Count = (uint)table.Columns.Count });

            table.Columns.ForEach(h =>
            {
                xw.WriteElement(h.ToTableColumn((uint)(hIndex + 1)));
                hIndex++;
            });

            xw.WriteEndElement();
            xw.WriteElement(table.TableStyleInfo.ToTableStyleInfo());

            xw.WriteEndElement();
        }
        #endregion
    }
}
