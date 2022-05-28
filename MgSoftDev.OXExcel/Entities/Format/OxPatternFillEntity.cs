using System.Drawing;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxPatternFillEntity : IEquatable<OxPatternFillEntity>
    {
       

        public Color Color { get; set; }
        public OxPatterns PatternType { get; set; }

        #region Equatable
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxPatternFillEntity)obj);
        }


        public bool Equals(OxPatternFillEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Color.Equals(other.Color) && PatternType == other.PatternType;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Color.GetHashCode()*397) ^ (int) PatternType;
            }
        }

        public static bool operator ==(OxPatternFillEntity left, OxPatternFillEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxPatternFillEntity left, OxPatternFillEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
