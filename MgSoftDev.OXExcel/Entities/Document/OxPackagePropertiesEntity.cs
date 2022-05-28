namespace MgSoftDev.OXExcel.Entities.Document
{
    internal class OxPackagePropertiesEntity
    {
        public string Title { get; set; }
        public string Version { get; set; }
        public string Creator { get; set; }
        public string LastModifiedBy { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public string Company { get; set; }
    }
}
