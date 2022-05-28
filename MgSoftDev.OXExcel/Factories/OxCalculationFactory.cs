using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Document;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxCalculationFactory
    {
        internal readonly OxCalculationEntity Calculation;

        public OxCalculationFactory()
        {
            Calculation = new OxCalculationEntity();
            CalculationMode(OxCalculateModes.Auto);
            NotCalculationOnSave(true);
            NotConcurrentCalc(true);
            FullCalcOnLoad(false);
            ForceFullCalc(false);
            FullPrecision(false);
            CalculationIteration(false);
            IterateCount(100);
            IterateDelta(0.001);

        }

        public OxCalculationFactory CalculationMode(OxCalculateModes mode)
        {
            Calculation.CalculationMode = mode;
            return this;
        }

        public OxCalculationFactory NotCalculationOnSave(bool value = false)
        {
            Calculation.CalculationOnSave = value;
            return this;
        }

        public OxCalculationFactory NotConcurrentCalc(bool value = false)
        {
            Calculation.ConcurrentCalc = value;
            return this;
        }

        public OxCalculationFactory ForceFullCalc(bool value = true)
        {
            Calculation.ForceFullCalc = value;
            return this;
        }

        public OxCalculationFactory FullCalcOnLoad(bool value = true)
        {
            Calculation.FullCalcOnLoad = value;
            return this;
        }

        public OxCalculationFactory FullPrecision(bool value = true)
        {
            Calculation.FullPrecision = value;
            return this;
        }

        public OxCalculationFactory CalculationIteration(bool value = true)
        {
            Calculation.CalculationIteration = value;
            return this;
        }

        public OxCalculationFactory IterateCount(uint num)
        {
            Calculation.IterateCount = num;
            return this;
        }

        public OxCalculationFactory IterateDelta(double delta)
        {
            Calculation.IterateDelta = delta;
            return this;
        }



    }
}
