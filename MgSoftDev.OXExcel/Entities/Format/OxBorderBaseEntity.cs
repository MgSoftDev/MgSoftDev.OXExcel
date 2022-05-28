using System.Drawing;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxBorderBaseEntity : IEquatable<OxBorderBaseEntity>
    {
        public Color Color { get; set; }
        public OxBorderStyles BorderStyle { get; set; }
        #region Equatable

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxBorderBaseEntity)obj);
        }
        public bool Equals(OxBorderBaseEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Color.Equals(other.Color) && BorderStyle == other.BorderStyle;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Color.GetHashCode()*397) ^ (int) BorderStyle;
            }
        }

        public static bool operator ==(OxBorderBaseEntity left, OxBorderBaseEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxBorderBaseEntity left, OxBorderBaseEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
