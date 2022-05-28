using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxGradientFillEntity : IEquatable<OxGradientFillEntity>
    {
        

        public double Degree { get; set; }
        public double Bottom { get; set; }
        public double Top { get; set; }
        public double Left { get; set; }
        public double Right { get; set; }

        public OxGradients Gradient { get; set; }

        public List<OxGradientStopEntity> GradientStops { get; set; }
        #region Equatable

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxGradientFillEntity)obj);
        }

        public bool Equals(OxGradientFillEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Degree.Equals(other.Degree) && Bottom.Equals(other.Bottom) && Top.Equals(other.Top) &&
                   Left.Equals(other.Left) && Right.Equals(other.Right) && Gradient == other.Gradient &&
                   Equals(GradientStops, other.GradientStops);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Degree.GetHashCode();
                hashCode = (hashCode*397) ^ Bottom.GetHashCode();
                hashCode = (hashCode*397) ^ Top.GetHashCode();
                hashCode = (hashCode*397) ^ Left.GetHashCode();
                hashCode = (hashCode*397) ^ Right.GetHashCode();
                hashCode = (hashCode*397) ^ (int) Gradient;
                hashCode = (hashCode*397) ^ (GradientStops != null ? GradientStops.GetHashCode() : 0);
                return hashCode;
            }
        }

        public static bool operator ==(OxGradientFillEntity left, OxGradientFillEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxGradientFillEntity left, OxGradientFillEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
