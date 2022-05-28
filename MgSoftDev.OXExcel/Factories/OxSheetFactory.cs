using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Entities.Sheet;
using MgSoftDev.OXExcel.Entities.Table;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxSheetFactory
    {
        internal readonly OxSheetEntity Sheet;

        public  OxSheetFactory(string tabName)
        {
            Sheet = new OxSheetEntity
            {
                Columns = new List<OxColumnEntity>(),
                RowsCellsList = new OxRowsCellCollection(),
                Images = new List<OxImageEntity>(),
                Tables = new List<OxTableEntity>()
            };
            TabName(tabName);
            SheetVisibility(OxSheetVisibilities.Visible);
            Sheet.PageMargins = null;
            Sheet.PageSetup = null;
            Sheet.SheetView = null;
            
        }

        #region Document Properties 
        public OxSheetFactory TabName(string name)
        {
            Sheet.TabName = name;
            return this;
        }
        
        public OxSheetFactory SheetVisibility(OxSheetVisibilities visibility)
        {
            Sheet.SheetVisibility = visibility;
            return this;
        }
        
        public OxPageMarginsFactory PageMargins()
        {
            var res = new OxPageMarginsFactory();
            Sheet.PageMargins = res.PageMargins;
            return res;
        }

        public OxSheetFactory PageMargins(Action<OxPageMarginsFactory> pageMarginsAction)
        {
            var f = new OxPageMarginsFactory();
            pageMarginsAction(f);
            Sheet.PageMargins = f.PageMargins;
            return this;
        }

        public OxSheetFactory PageMargins(OxPageMarginsFactory pageMargins)
        {
            Sheet.PageMargins = pageMargins.PageMargins;
            return this;
        }


        public OxSheetFactory AddColumn( Action<OxColumnsFactory> columnsDefinitions)
        {
            var f = new OxColumnsFactory();
            columnsDefinitions(f);
            Sheet.Columns.AddRange(f.Columns);
            return this;
        }
        
        public OxSheetFactory AddColumn(OxColumnFactory colDefinitions)
        {
            Sheet.Columns.Add(colDefinitions.Column);
            return this;
        }

        public OxSheetFactory AddColumn(IEnumerable<OxColumnFactory> colDefinitions)
        {
            Sheet.Columns.AddRange(colDefinitions.Select(s=> s.Column));
            return this;
        }

        public OxSheetFactory AddRow(Action<OxRowsFactory> rowsDefinitions)
        {
            var f = new OxRowsFactory();
            rowsDefinitions(f);
            f.Rows.ForEach( r =>Sheet.RowsCellsList.AddOrRemplace( r ));

            return this;
        }

        public OxSheetFactory AddRow(OxRowFactory rowDefinitions)
        {
            Sheet.RowsCellsList.AddOrRemplace(rowDefinitions.Row);
            return this;
        }

        public OxSheetFactory AddRow(IEnumerable<OxRowFactory> rowDefinitions)
        {
            rowDefinitions.ToList().ForEach(r=>Sheet.RowsCellsList.AddOrRemplace(r.Row));
            return this;
        }

        public OxSheetFactory BackGroundImage(Uri image)
        {
            Sheet.BackgroundImage = image;
            return this;
        }

        public OxSheetPropertiesFactory SheetProperties()
        {
            var res =new  OxSheetPropertiesFactory();
            Sheet.SheetProperties = res.Properties;
            return res;
        }

        public OxSheetFactory SheetProperties(Action<OxSheetPropertiesFactory> sheetPropertiesAction)
        {
            var f = new OxSheetPropertiesFactory();
            sheetPropertiesAction(f);
            Sheet.SheetProperties = f.Properties;
            return this;
        }

        public OxSheetFactory SheetProperties(OxSheetPropertiesFactory sheetProperties)
        {
            Sheet.SheetProperties = sheetProperties.Properties;
            return this;
        }


        public OxSheetFactory PageSetup(Action<OxPageSetupFactory> pageSetupAction)
        {
            var f = new OxPageSetupFactory();
            pageSetupAction(f);
            Sheet.PageSetup = f.PageSetup;
            return this;
        }

        public OxSheetFactory PageSetup(OxPageSetupFactory pageSetup)
        {
            Sheet.PageSetup = pageSetup.PageSetup;
            return this;
        }
        public OxPageSetupFactory PageSetup() => new OxPageSetupFactory();


        public OxSheetViewFactory SheetView()
        {
            var res = new OxSheetViewFactory();
            Sheet.SheetView = res.SheetView;
            return res;
        }
        public OxSheetFactory SheetView(Action<OxSheetViewFactory> sheetViewAction)
        {
            var f = new OxSheetViewFactory();
            sheetViewAction(f);
            Sheet.SheetView = f.SheetView;
            return this;
        }

        public OxSheetFactory SheetView(OxSheetViewFactory sheetView)
        {
            Sheet.SheetView = sheetView.SheetView;
            return this;
        }

        public OxSheetFactory Cell( Action<OxCellsFactory> cellsAction )
        {
            var c = new OxCellsFactory();
            cellsAction( c );

            c.Cells.ForEach( f => AddRow( f.Row ) );

            c.Cells.ForEach( cc => Sheet.RowsCellsList.AddCell( cc ) );

            return this;
        }

        public OxSheetFactory Cell( OxCellFactory cell )
        {
            AddRow( cell.Cell.Row );
            Sheet.RowsCellsList.AddCell( cell.Cell );

            return this;
        }

        public OxSheetFactory Cell( IEnumerable<OxCellFactory> cells )
        {
            foreach( var c in cells )
            {
                AddRow( c.Cell.Row );
                Sheet.RowsCellsList.AddCell( c.Cell );
            }

            return this;
        }

        public OxCellFactory Cell( uint col, uint row )
        {
            var c = new OxCellFactory( col, row );
            AddRow( c.Cell.Row );
            Sheet.RowsCellsList.AddCell( c.Cell );

            return c;
        }

        public OxCellFactory Cell( string col, uint row )
        {
            var c = new OxCellFactory( col, row );
            AddRow( c.Cell.Row );
            Sheet.RowsCellsList.AddCell( c.Cell );

            return c;
        }

        public OxCellFactory Cell( string cellReference )
        {
            var c = new OxCellFactory( cellReference );
            AddRow( c.Cell.Row );
            Sheet.RowsCellsList.AddCell( c.Cell );

            return c;
        }




        public OxImageFactory AddImage(OxRangeEntity range)
        {
            var res = new OxImageFactory(range);
            Sheet.Images.Add(res.Image);
            return res;
        }

        public OxSheetFactory AddImage(Action<OxImagesFactory> imageAction)
        {
            var f = new OxImagesFactory();
            imageAction(f);
            Sheet.Images.AddRange(f.Images); 
            return this;
        }

        public OxSheetFactory AddImage(OxImageFactory image)
        {
            Sheet.Images.Add(image.Image);
            return this;
        }
        public OxSheetFactory AddImage(IEnumerable<OxImageFactory> images)
        {
            Sheet.Images.AddRange(images.Select(s=> s.Image));
            return this;
        }
        public OxSheetFactory AddTable(OxTableFactory table)
        {
            Sheet.Tables.Add(table.Table);
            return this;
        }
        public OxSheetFactory AddTable<T>(OxTableFactory<T> table)
        {
            Sheet.Tables.Add(table.Table);
            return this;
        }
        public OxSheetFactory AddTable<T>(IEnumerable<T> data, uint col, uint row, Action<OxTableFactory<T>> action)
        {
            var t = new OxTableFactory<T>(data.ToList(), col, row);
            action(t);
            Sheet.Tables.Add(t.Table);
            return this;
        }
        public OxSheetFactory AddTable<T>(IEnumerable<T> data, string col, uint row, Action<OxTableFactory<T>> action)
        {
            var t = new OxTableFactory<T>(data.ToList(), col, row);
            action(t);
            Sheet.Tables.Add(t.Table);
            return this;
        }
        public OxSheetFactory AddTable<T>(IEnumerable<T> data, string reference, Action<OxTableFactory<T>> action)
        {
            var t = new OxTableFactory<T>(data.ToList(), reference);
            action(t);
            Sheet.Tables.Add(t.Table);
            return this;
        }
        public OxSheetFactory AddTable(object data, uint col, uint row, Action<OxTableFactory> action)
        {
            var t = new OxTableFactory(data, col, row);
            action(t);
            Sheet.Tables.Add(t.Table);
            return this;
        }
        public OxSheetFactory AddTable(object data, string col, uint row, Action<OxTableFactory> action)
        {
            var t = new OxTableFactory(data, col, row);
            action(t);
            Sheet.Tables.Add(t.Table);
            return this;
        }
        public OxSheetFactory AddTable(object data, string reference, Action<OxTableFactory> action)
        {
            var t = new OxTableFactory(data, reference);
            action(t);
            Sheet.Tables.Add(t.Table);
            return this;
        }

        #endregion






        #region Child Properties 


        #endregion

        private void AddRow(uint index)
        {
            AddRow(new OxRowFactory(index));            
        }
    }
}
