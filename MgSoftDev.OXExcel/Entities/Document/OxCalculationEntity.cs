using MgSoftDev.OXExcel.Commons;

namespace MgSoftDev.OXExcel.Entities.Document
{
    internal class OxCalculationEntity
    {
        public OxCalculateModes CalculationMode { get; set; }
        public bool CalculationOnSave { get; set; }
        public bool ConcurrentCalc { get; set; }
        public bool ForceFullCalc { get; set; }
        public bool FullCalcOnLoad { get; set; }
        public bool FullPrecision { get; set; }
        public bool CalculationIteration { get; set; }
        public uint IterateCount { get; set; }
        public double IterateDelta { get; set; }




    }
}
