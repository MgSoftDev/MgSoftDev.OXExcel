using MgSoftDev.OXExcel.Entities.Document;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxPackagePropertiesFactory
    {
        internal readonly OxPackagePropertiesEntity PackageProperties;

        public OxPackagePropertiesFactory()
        {
            PackageProperties = new OxPackagePropertiesEntity();
        }

        public OxPackagePropertiesFactory Title(string value)
        {
            PackageProperties.Title = value;
            return this;
        }
        public OxPackagePropertiesFactory Company(string value)
        {
            PackageProperties.Company = value;
            return this;
        }
        public OxPackagePropertiesFactory Version(string value)
        {
            PackageProperties.Version = value;
            return this;
        }
        public OxPackagePropertiesFactory Creator(string value)
        {
            PackageProperties.Creator = value;
            return this;
        }
        public OxPackagePropertiesFactory LastModifiedBy(string value)
        {
            PackageProperties.LastModifiedBy = value;
            return this;
        }

        public OxPackagePropertiesFactory Created(DateTime? value)
        {
            PackageProperties.Created = value;
            return this;
        }

        public OxPackagePropertiesFactory Modified(DateTime? value)
        {
            PackageProperties.Modified = value;
            return this;
        }
    }
}
