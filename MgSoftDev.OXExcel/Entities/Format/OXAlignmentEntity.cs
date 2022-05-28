using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxAlignmentEntity : IEquatable<OxAlignmentEntity>
    {
        public OxTextHorizontalAlignments HorizontalAlignment { get; set; }
        public OxTextVerticalAlignments VerticalAlignment { get; set; }

        public bool JustifyLastLine { get; set; }
        public bool ShrinkToFit { get;  set; }
        public uint Rotation { get; set; }

        #region Equatable  
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxAlignmentEntity)obj);
        }

        public bool Equals(OxAlignmentEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return HorizontalAlignment == other.HorizontalAlignment && VerticalAlignment == other.VerticalAlignment &&
                   JustifyLastLine == other.JustifyLastLine && ShrinkToFit == other.ShrinkToFit &&
                   Rotation == other.Rotation;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (int) HorizontalAlignment;
                hashCode = (hashCode*397) ^ (int) VerticalAlignment;
                hashCode = (hashCode*397) ^ JustifyLastLine.GetHashCode();
                hashCode = (hashCode*397) ^ ShrinkToFit.GetHashCode();
                hashCode = (hashCode*397) ^ (int) Rotation;
                return hashCode;
            }
        }

        public static bool operator ==(OxAlignmentEntity left, OxAlignmentEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxAlignmentEntity left, OxAlignmentEntity right)
        {
            return !Equals(left, right);
        }

        #endregion


    }
}
