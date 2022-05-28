namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxBorderEntity : IEquatable<OxBorderEntity>
    {

        public OxBorderBaseEntity Bottom { get; set; }
        public OxBorderBaseEntity Top { get; set; }
        public OxBorderBaseEntity Right { get; set; }
        public OxBorderBaseEntity Left { get; set; }
        public OxBorderBaseEntity Diagonal { get; set; }
        public bool DiagonalDown { get; set; }
        public bool DiagonalUp { get; set; }
        public bool Outline { get; set; }

        #region equatable

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxBorderEntity)obj);
        }

        public bool Equals(OxBorderEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(Bottom, other.Bottom) && Equals(Top, other.Top) && Equals(Right, other.Right) &&
                   Equals(Left, other.Left) && Equals(Diagonal, other.Diagonal) && DiagonalDown == other.DiagonalDown &&
                   DiagonalUp == other.DiagonalUp && Outline == other.Outline;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Bottom != null ? Bottom.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Top != null ? Top.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Right != null ? Right.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Left != null ? Left.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ (Diagonal != null ? Diagonal.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ DiagonalDown.GetHashCode();
                hashCode = (hashCode*397) ^ DiagonalUp.GetHashCode();
                hashCode = (hashCode*397) ^ Outline.GetHashCode();
                return hashCode;
            }
        }

        public static bool operator ==(OxBorderEntity left, OxBorderEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxBorderEntity left, OxBorderEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
