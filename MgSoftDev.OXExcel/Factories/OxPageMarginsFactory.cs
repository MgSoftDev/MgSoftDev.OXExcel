using MgSoftDev.OXExcel.Entities.Sheet;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxPageMarginsFactory
    {
        internal readonly OxPageMarginsEntity PageMargins;

        public OxPageMarginsFactory()
        {
            PageMargins = new OxPageMarginsEntity();
            Left(1.15);
            Right(1.15);
            Top(1.15);
            Bottom(1.15);
            Header(0.3);
            Footer(0.3);
        }

        public OxPageMarginsFactory Left(double value)
        {
            PageMargins.Left = value;
            return this;
        }

        public OxPageMarginsFactory Right(double value)
        {
            PageMargins.Right = value;
            return this;
        }

        public OxPageMarginsFactory Top(double value)
        {
            PageMargins.Top = value;
            return this;
        }

        public OxPageMarginsFactory Bottom(double value)
        {
            PageMargins.Bottom = value;
            return this;
        }

        public OxPageMarginsFactory Header(double value)
        {
            PageMargins.Header = value;
            return this;
        }

        public OxPageMarginsFactory Footer(double value)
        {
            PageMargins.Footer = value;
            return this;
        }

    }
}
