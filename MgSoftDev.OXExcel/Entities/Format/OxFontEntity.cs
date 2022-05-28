using System.Drawing;
using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxFontEntity : IEquatable<OxFontEntity>
    {
        

        public bool Bold { get; set; }
        public Color Color { get; set; }
        public bool Condense { get; set; }
        public bool Extend { get; set; }
        public string FontName { get; set; }
        public double Size { get; set; }
        public bool Italic { get; set; }
        public bool Outline { get; set; }
        public bool Shadow { get; set; }
        public bool Strike { get; set; }
        public OxUnderlines Underline { get; set; }
        public OxVerticalAlignments VerticalAlignments { get; set; }
        public OxFontSchemes FontScheme { get; set; }

        #region equatable
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxFontEntity)obj);
        }

        public bool Equals(OxFontEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Bold == other.Bold && Color.Equals(other.Color) && Condense == other.Condense &&
                   Extend == other.Extend && string.Equals(FontName, other.FontName) && Size.Equals(other.Size) &&
                   Italic == other.Italic && Outline == other.Outline && Shadow == other.Shadow &&
                   Strike == other.Strike && Underline == other.Underline &&
                   VerticalAlignments == other.VerticalAlignments && FontScheme == other.FontScheme;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Bold.GetHashCode();
                hashCode = (hashCode*397) ^ Color.GetHashCode();
                hashCode = (hashCode*397) ^ Condense.GetHashCode();
                hashCode = (hashCode*397) ^ Extend.GetHashCode();
                hashCode = (hashCode*397) ^ (FontName != null ? FontName.GetHashCode() : 0);
                hashCode = (hashCode*397) ^ Size.GetHashCode();
                hashCode = (hashCode*397) ^ Italic.GetHashCode();
                hashCode = (hashCode*397) ^ Outline.GetHashCode();
                hashCode = (hashCode*397) ^ Shadow.GetHashCode();
                hashCode = (hashCode*397) ^ Strike.GetHashCode();
                hashCode = (hashCode*397) ^ (int) Underline;
                hashCode = (hashCode*397) ^ (int) VerticalAlignments;
                hashCode = (hashCode*397) ^ (int) FontScheme;
                return hashCode;
            }
        }

        public static bool operator ==(OxFontEntity left, OxFontEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxFontEntity left, OxFontEntity right)
        {
            return !Equals(left, right);
        }
        #endregion
    }
}
