using System.Drawing;

namespace MgSoftDev.OXExcel.Entities.Format
{
    internal class OxGradientStopEntity : IEquatable<OxGradientStopEntity>
    {

        public Color Color { get; set; }
        public double Position { get; set; }
        #region Equatable
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxGradientStopEntity)obj);
        }
        public bool Equals(OxGradientStopEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Color.Equals(other.Color) && Position.Equals(other.Position);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Color.GetHashCode()*397) ^ Position.GetHashCode();
            }
        }

        public static bool operator ==(OxGradientStopEntity left, OxGradientStopEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxGradientStopEntity left, OxGradientStopEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
