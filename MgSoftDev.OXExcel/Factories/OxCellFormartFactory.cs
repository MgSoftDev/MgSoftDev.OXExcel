using System.Drawing;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxCellFormartFactory
    {
        internal readonly OxCellFormartEntity Format;

        internal OxCellFormartFactory(OxCellFormartEntity format)
        {
            Format = format;
        }
        public OxCellFormartFactory()
        {
            Format = new OxCellFormartEntity
            {
                NumberFormat = null,
                Font = null,
                Fill = null,
                Borders = null
            };
        }

        public OxFontFactory Font()
        {
            var f = Format.Font == null ? new OxFontFactory() : new OxFontFactory(Format.Font);       
            Format.Font = f.Font;
            return f;
        }
        public OxCellFormartFactory Font(OxFontFactory font)
        {
            Format.Font = font.Font;
            return this;
        }
        public OxCellFormartFactory Font(Action<OxFontFactory> fontAction)
        {
            var f = Format.Font== null? new OxFontFactory(): new OxFontFactory(Format.Font);
            fontAction(f);
            Format.Font = f.Font;
            return this;
        }

        public OxCellFormartFactory FillPattern(Color color, OxPatterns pattern)
        {
            Format.Fill = new OxFillEntity
            {
                PatternFill = new OxPatternFillEntity
                {
                    Color = color,
                    PatternType = pattern
                }, GradientFill = null
            };
            return this;
        }

        public OxCellFormartFactory FillGradient(double degree, Action<OxGradientStopsFactory> gradientsAction)
        {
            Format.Fill = new OxFillEntity
            {
                PatternFill = null,
                GradientFill = new OxGradientFillEntity
                {
                    Degree = degree,
                    GradientStops = new List<OxGradientStopEntity>(),
                    Gradient = OxGradients.Linear
                }
            };
            gradientsAction(new OxGradientStopsFactory(Format.Fill.GradientFill.GradientStops));
            return this;
        }

        public OxCellFormartFactory NumberFormat(string format)
        {
            Format.NumberFormat = new OxNumberFormatEntity {FormatCode = format};
            return this;
        }

        public OxBorderFactory Borders()
        {
            var b = Format.Borders== null? new OxBorderFactory(): new OxBorderFactory(Format.Borders);
            Format.Borders = b.Border;
            return b;
        }
        public OxCellFormartFactory Borders(OxBorderFactory border)
        {
            Format.Borders = border.Border;
            return this;
        }
        public OxCellFormartFactory Borders(Action<OxBorderFactory> borderAction)
        {
            var b = Format.Borders == null ? new OxBorderFactory() : new OxBorderFactory(Format.Borders);
            borderAction(b);
            Format.Borders = b.Border; 
            return this;
        }


        public OxAlignmentFactory Alignment()
        {
            var b = Format.Alignment== null? new OxAlignmentFactory(): new OxAlignmentFactory(Format.Alignment);
            Format.Alignment = b.Alignment;
            return b;
        }
        public OxCellFormartFactory Alignment(OxAlignmentFactory alignment)
        {
            Format.Alignment = alignment.Alignment;
            return this;
        }
        public OxCellFormartFactory Alignment(Action<OxAlignmentFactory> alignmentAction)
        {
            var b = Format.Alignment == null ? new OxAlignmentFactory() : new OxAlignmentFactory(Format.Alignment);
            alignmentAction(b);
            Format.Alignment = b.Alignment;
            return this;
        }

    }
}
