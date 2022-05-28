namespace MgSoftDev.OXExcel.Entities.Format
{
    [Serializable]
    internal class OxNumberFormatEntity : IEquatable<OxNumberFormatEntity>
    {
        public string FormatCode { get; set; }

        #region Equatable
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((OxNumberFormatEntity) obj);
        }
        public bool Equals(OxNumberFormatEntity other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(FormatCode, other.FormatCode);
        }

        public override int GetHashCode()
        {
            return FormatCode?.GetHashCode() ?? 0;
        }

        public static bool operator ==(OxNumberFormatEntity left, OxNumberFormatEntity right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(OxNumberFormatEntity left, OxNumberFormatEntity right)
        {
            return !Equals(left, right);
        }

        #endregion
    }
}
