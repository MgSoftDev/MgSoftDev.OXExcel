using System.ComponentModel;
using System.Dynamic;
using MgSoftDev.OXExcel.Attributes;

namespace MgSoftDev.OXExcel.Entities.Dynamic
{
    public class DynamicEntity : DynamicObject, INotifyPropertyChanged
    {
        public Dictionary<string, object> DynamicProperty = new Dictionary<string, object>();
        protected T Get<T>(string key)
        {
            return DynamicProperty[key] is DBNull ? default(T) : (T)DynamicProperty[key];
        }

        public void SetProperties(Dictionary<string, object> dictionary) { DynamicProperty = dictionary; }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            if (DynamicProperty.ContainsKey(binder.Name))
            {
                DynamicProperty[binder.Name] = value;

                return true;
            }

            return base.TrySetMember(binder, value);

        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (DynamicProperty.ContainsKey(binder.Name))
            {
                result = DynamicProperty[binder.Name];

                return true;
            }
            return base.TryGetMember(binder, out result);

        }

        public override IEnumerable<string> GetDynamicMemberNames() { return DynamicProperty.Keys.ToArray(); }

        #region PropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged( string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

    }

}
