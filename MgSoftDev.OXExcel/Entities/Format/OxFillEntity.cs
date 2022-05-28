namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxFillEntity : IEquatable<OxFillEntity>
    {

        public OxPatternFillEntity PatternFill { get; set; }
        public OxGradientFillEntity GradientFill { get; set; }

        #region Equatable

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxFillEntity)obj);
        }
        public bool Equals(OxFillEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(PatternFill, other.PatternFill) && Equals(GradientFill, other.GradientFill);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((PatternFill != null ? PatternFill.GetHashCode() : 0)*397) ^ (GradientFill != null ? GradientFill.GetHashCode() : 0);
            }
        }

        public static bool operator ==(OxFillEntity left, OxFillEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxFillEntity left, OxFillEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
