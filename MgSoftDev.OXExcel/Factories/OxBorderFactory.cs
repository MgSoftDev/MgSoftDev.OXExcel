using System.Drawing;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;

namespace MgSoftDev.OXExcel.Factories
{
    public class OxBorderFactory
    {
        internal readonly OxBorderEntity Border;

        public OxBorderFactory()
        {
            Border = new OxBorderEntity
            {
                Left = null,
                Bottom = null,
                Diagonal = null,
                DiagonalUp = false,
                Right = null,
                Top = null,
                Outline = false,
                DiagonalDown = false
            };
        }

        internal OxBorderFactory(OxBorderEntity borders)
        {
            Border = borders;
        }

        public OxBorderFactory Left(Color color, OxBorderStyles style )
        {
            Border.Left = new OxBorderBaseEntity {Color = color,BorderStyle = style };
            return this;
        }
        public OxBorderFactory Bottom(Color color, OxBorderStyles style)
        {
            Border.Bottom = new OxBorderBaseEntity { Color = color, BorderStyle = style };
            return this;
        }
        public OxBorderFactory Diagonal(Color color, OxBorderStyles style)
        {
            Border.Diagonal = new OxBorderBaseEntity { Color = color, BorderStyle = style };
            return this;
        }
        public OxBorderFactory DiagonalDown(bool value = true)
        {
            Border.DiagonalDown = value;
            return this;
        }
        public OxBorderFactory DiagonalUp(bool value = true)
        {
            Border.DiagonalUp = value;
            return this;
        }
        
        public OxBorderFactory Outline(bool value = true)
        {
            Border.Outline = value;
            return this;
        }
        public OxBorderFactory Right(Color color, OxBorderStyles style)
        {
            Border.Right = new OxBorderBaseEntity { Color = color, BorderStyle = style };
            return this;
        }
        public OxBorderFactory Top(Color color, OxBorderStyles style)
        {
            Border.Top = new OxBorderBaseEntity { Color = color, BorderStyle = style };
            return this;
        }
        
    }
}
