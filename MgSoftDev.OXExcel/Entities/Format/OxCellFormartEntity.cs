namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxCellFormartEntity : IEquatable<OxCellFormartEntity>
    {

        public OxNumberFormatEntity NumberFormat { get; set; }
        public OxFontEntity Font { get; set; }
        public OxFillEntity Fill { get; set; }
        public  OxBorderEntity Borders { get; set; }
        public OxAlignmentEntity Alignment { get; set; }

        #region equatable
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxCellFormartEntity)obj);
        }
        public bool Equals(OxCellFormartEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(NumberFormat, other.NumberFormat) && Equals(Font, other.Font) && Equals(Fill, other.Fill) && Equals(Borders, other.Borders) && Equals(Alignment, other.Alignment);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (NumberFormat != null ? NumberFormat.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Font != null ? Font.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Fill != null ? Fill.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Borders != null ? Borders.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Alignment != null ? Alignment.GetHashCode() : 0);
                return hashCode;
            }
        }

        public static bool operator ==(OxCellFormartEntity left, OxCellFormartEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxCellFormartEntity left, OxCellFormartEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
