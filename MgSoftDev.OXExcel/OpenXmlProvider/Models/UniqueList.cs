namespace MgSoftDev.OXExcel.OpenXmlProvider.Models
{
    /// <summary>
    /// Esta clase sirve para manejar una lista unica de elementos cada vez que se agrega un nuevo elemento se verifica si este no ha se
    /// agregado anteriormente de ser asi devuelve su index de lo contrario de agrega el nuevo elemento y se le genera un nuevo index
    /// </summary>
    /// <typeparam name="T1">El tipo de dato a almacenar</typeparam>
    internal class UniqueList<T1>
    {
        private readonly Dictionary<T1,int> _KeyValueDictionary = new Dictionary<T1, int>();
        private readonly Dictionary<int,T1> _ValueKeyDictionary = new Dictionary<int, T1>();
        private readonly object _Lock = new object();
        private int _Index;

        public int Add(T1 value)
        {
            lock( _Lock )
            {
                if (!_KeyValueDictionary.ContainsKey(value))
                {
                    _KeyValueDictionary.Add(value, _Index);

                    _ValueKeyDictionary.Add(_Index, value);

                    _Index++;
                    return _Index - 1;
                }

                _KeyValueDictionary.TryGetValue(value, out var val);
                return val;
            }
            
        }

        public T1 GetValue( int key )
        {
            lock( _Lock )
            {
                _ValueKeyDictionary.TryGetValue( key, out T1 val );

                return val;
            }
        }

        public void Clear()
        {
            _KeyValueDictionary.Clear();
            _ValueKeyDictionary.Clear();
        }

        public override string ToString()
        {
            return $"Total de Keys {_KeyValueDictionary.Count}";

        }
    }
}
