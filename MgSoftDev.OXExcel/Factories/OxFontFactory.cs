using System.Drawing;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxFontFactory
    {
        internal readonly OxFontEntity Font;

        internal OxFontFactory(OxFontEntity font)
        {
            Font = font;
        }

        public OxFontFactory()
        {
            Font = new OxFontEntity();
            FontName("Calibri");
            Color(System.Drawing.Color.Black);
            Size(12);
            Underline(OxUnderlines.None);
            VerticalAlignments();
            FontScheme(OxFontSchemes.None);
            Bold(false);
            Condense(false);
            Extend(false);
            Italic(false);
            Outline(false);
            Shadow(false);
            Strike(false);
        }

        public OxFontFactory FontName(string value)
        {
            Font.FontName = value;
            return this;
        }
        public OxFontFactory Color(Color value)
        {
            Font.Color = value;
            return this;
        }
        public OxFontFactory Size(double value)
        {
            Font.Size = value;
            return this;
        }
        public OxFontFactory Underline(OxUnderlines value = OxUnderlines.Single)
        {
            Font.Underline = value;
            return this;
        }
        public OxFontFactory VerticalAlignments(OxVerticalAlignments value = OxVerticalAlignments.Baseline)
        {
            Font.VerticalAlignments = value;
            return this;
        }
        public OxFontFactory FontScheme(OxFontSchemes value =OxFontSchemes.None)
        {
            Font.FontScheme = value;
            return this;
        }

        public OxFontFactory Bold(bool value = true)
        {
            Font.Bold = value;
            return this;
        }
        public OxFontFactory Condense(bool value = true)
        {
            Font.Condense = value;
            return this;
        }
        public OxFontFactory Extend(bool value = true)
        {
            Font.Extend = value;
            return this;
        }
        public OxFontFactory Italic(bool value = true)
        {
            Font.Italic = value;
            return this;
        }
        public OxFontFactory Outline(bool value = true)
        {
            Font.Outline = value;
            return this;
        }
        public OxFontFactory Shadow(bool value = true)
        {
            Font.Shadow = value;
            return this;
        }
        public OxFontFactory Strike(bool value = true)
        {
            Font.Strike = value;
            return this;
        }
    }
}
