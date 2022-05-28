namespace MgSoftDev.OXExcel.Factories
{
    public class OxTableFactory<T> : OxTableBaseFactory
    {

        public OxTableFactory(IEnumerable<T> data, uint col, uint row) : base(data,col,row)
        {
            
        }
        public OxTableFactory(IEnumerable<T> data, string col, uint row) : base(data, col, row)
        {
        }
        public OxTableFactory(IEnumerable<T> data, string reference) : base(data, reference)
        {
        }


        public OxTableFactory<T> Columns( Action<OxTableColumnsFactory<T>> colsAction)
        {
            var c = new OxTableColumnsFactory<T>();
            colsAction(c);
            Table.Columns = c.TableColumns;
            return this;
        }

    }

    public class OxTableFactory : OxTableBaseFactory
    {

        public OxTableFactory(object data, uint col, uint row) : base(data, col, row)
        {

        }
        public OxTableFactory(object data, string col, uint row) : base(data, col, row)
        {
        }
        public OxTableFactory(object data, string reference) : base(data, reference)
        {
        }


        public OxTableFactory Columns(Action<OxTableColumnsFactory> colsAction)
        {
            var c = new OxTableColumnsFactory();
            colsAction(c);
            Table.Columns = c.TableColumns;
            return this;
        }

    }
}
