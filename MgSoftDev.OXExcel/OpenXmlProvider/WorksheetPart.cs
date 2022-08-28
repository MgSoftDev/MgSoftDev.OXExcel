using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
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
        
        

        private void GenerateWorksheetPartContent2(WorksheetPart worksheetPart1, OpenXmlWriter xw, OxSheetEntity sheet)
        {
            

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
            
            
            // Add rows from tables, recorremos las tablas y extraemos las rows
            sheet.Tables.ForEach(f => TableInsertOrUpdate(sheet,f));

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
                CreateExcelRowsFromOxRows(xw, sheet);
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

            if (sheet.Tables != null && sheet.Tables.Count > 0)
            {
                xw.WriteStartElement(new TableParts() { Count = (uint)sheet.Tables.Count });
                sheet.Tables.ForEach(ot =>
                {
                    if (ot.TableType != OxTableType.Excel) return;

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


            Console.WriteLine($"Total _Formats = {Const.Formats.Count}");
            Console.WriteLine($"Total _StringShareds = {Const.StringShareds.Count}");
            Console.WriteLine($"Total _Hyperlinks = {Const.Hyperlinks.Count}");
            Console.WriteLine($"Total UniqueValuesList = {Const.UniqueValuesList}");


            #endregion
        }

        #region Rows Cols cells

        internal static void TableInsertOrUpdate(OxSheetEntity sheet, OxTableEntity table)
        {
            Console.WriteLine("Generando Rows de tablas V 5.1");
            
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
                return;

            table.Columns = table.Columns.OrderBy( o => o.Order).ToList();

            var lista = table.DataCollection;
            var rIndex = 0U;
            table.RowsCounts = (uint)table.DataCollection.LongCount();


            // insert Row of Columns
            var d = table.RowDefinition.Clone();
            d.RowIndex     = table.Row;
            d.CustomFormat = false;
            d.Format       = null;
            var cRow = sheet.RowsCellsList.AddAndGet(d);

            OxRowCellsEntity tRow = null;
            // insertar row totales solo 1 si tiene totales activos
            if (table.TotalsRowShow)
            {

                d              = table.RowDefinition.Clone();
                d.RowIndex     = (uint) (table.Row + lista.Count +1);
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

               

                // Estable se el tamaño de la columna por default es 11
                var col = new OxColumnFactory(table.Column + cIndex, table.Column + cIndex).Width(c.Size);
                sheet.Columns.Add(col.Column);

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


            
            // insert data cells
            var fcols = table.Columns.Where(w => w.CustomColumnFilter != null || w.ColumnFilter != null).ToList();

            var tasks = new List<Task>();
            foreach (var r in lista)
            {
                rIndex++;
                cIndex = 0U;
                tasks.Add( OxRowEntity(sheet, table, r, lista, rIndex, fcols,  cIndex));
                
            }
            Task.WaitAll(tasks.ToArray());
            //await Task.WhenAll(tasks);

            

           

            // GC Lista and DataCollection
           tasks.Clear();
           tasks = null;
            lista.Clear();
            lista = null;
           
            GC.Collect();
            Console.WriteLine("FIN Generando Rows de tablas");
           

        }

        private static Task OxRowEntity(OxSheetEntity sheet, OxTableEntity table, object r, List<object> lista,
                                              uint rIndex,
                                              List<OxTableColumnsEntity> fcols,  uint cIndex)
        {
            return Task.Factory.StartNew(() =>
            {
                try
                {
                    var rowDeff = table.RowDefinition.Clone();
                    if (table.RowDefinitionTemplate != null)
                    {
                        rowDeff = table.RowDefinitionTemplate(new OxTableRowDefinitionTemplateEntity
                        {
                            Rows = r,
                            RowDefinition = new OxRowFactory(rowDeff),
                            MasterData = lista,
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

                    var row = sheet.RowsCellsList.AddAndGet(rowDeff);



                    table.Columns.ForEach(c =>
                    {
                        var format = c.CellFormart.Clone().Combine(rowDeff.Format);
                        var val = r.GetPropertyVal(c.PropertyPath, c.DefaultValue);
                        OxHyperlinkEntity link = null;

                        #region Templates

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

                        var cell = new OxTableCellEntity()
                        {
                            Row = table.Row + rIndex,
                            Column = table.Column + cIndex,
                            CellFormart = format,
                            CellTypeValue = c.CellTypeValue,
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
                    });

                   

                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            });
        }

        private void CreateExcelRowsFromOxRows(OpenXmlWriter xw, OxSheetEntity sheet)
        {
            Console.WriteLine("Generando Doc para las rows y cels");

            var firstRow  = true;

            foreach( var r in sheet.RowsCellsList.Rows.Values )
            {
               
                    if (!firstRow) xw.WriteEndElement();
                    xw.WriteStartElement(ToRow(r.Row as OxRowEntity, 1));
                    firstRow = false;
                
                foreach ( var c in r.Cells.Values )
                {
                    
                     
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
                                throw new Exception("Cheque su cadena (" + entity.Value + ")  Row=" +c.Row + "  Col=" +
                                        c.Column);
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
                                throw new Exception("Cheque su cadena (" + Const.UniqueValuesList.GetValue(entity.Value) + ")  Row=" + c.Row + "  Col=" +
                                        c.Column);
                            }


                            break;
                    }



                    #endregion
                }
            }

           
            if (!firstRow) xw.WriteEndElement();

            //GC data
           
            GC.Collect();
            Console.WriteLine("END  Generando Doc para las rows y cels y se libera memoria ");
            
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
        private void GenerateDrawingsPart1Content(OpenXmlWriter xw, List<OxImageEntity> images )
        {
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
