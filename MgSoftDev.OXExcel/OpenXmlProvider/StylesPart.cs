using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Format;
using MgSoftDev.OXExcel.Helpers.Extensions;
using MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal partial class OpenXmlExcelProvider
    {

        private List<string> _DistinctColors = new List<string>();
        private void GenerateWorkbookStylesPart1Content(WorkbookPart workbookPart1)
        {
            var xw = OpenXmlWriter.Create(workbookPart1.AddNewPart<WorkbookStylesPart>("rId4"));
            uint iExcelIndex = 164;
            _DistinctColors = new List<string>();

            var stylesheet1 = new Stylesheet() {};
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            xw.WriteStartElement(stylesheet1);

            
            #region Create Default Font            
            var font1 = new Font
            {
                FontSize = new FontSize() {Val = 11D},
                Color = new Color() {Theme = 1U},
                FontName = new FontName() {Val = "Calibri"},
                FontFamilyNumbering = new FontFamilyNumbering() {Val = 2},
                FontScheme = new FontScheme() {Val = FontSchemeValues.Minor}
            };

            // font hyperlink
            var fontLink = new Font()
            {
                Underline = new Underline(),
                FontSize = new FontSize() { Val = 11D },
                Color = new Color() { Theme = 10U },
                FontName = new FontName() { Val = "Calibri" },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                FontScheme = new FontScheme() { Val = FontSchemeValues.Minor }
            };

            var fonts1 = new Fonts(font1, fontLink) { Count = (uint)(Const.Formats.Count + 2), KnownFonts = true };            
            #endregion
            #region create default fill            
            var fill1 = new Fill {PatternFill = new PatternFill() {PatternType = PatternValues.None}};
            var fill2 = new Fill {PatternFill = new PatternFill() {PatternType = PatternValues.Gray125}};
            var fills1 = new Fills(fill1,fill2) { Count = (uint)(Const.Formats.Count + 2) };
            #endregion
            #region Default Border            
            var border1 = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            var borders1 = new Borders(border1) { Count = (uint)(Const.Formats.Count + 1) };            
            #endregion
            #region Default Style
            var cellStyleFormats1 = new CellStyleFormats() { Count = 2U };
            cellStyleFormats1.Append(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U });
            cellStyleFormats1.Append(new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U , ApplyNumberFormat = false,ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false});
            
            var nfs = new NumberingFormats();
            var cellFormats1 = new CellFormats() { Count = (uint)(Const.Formats.Count + 2) };            
            cellFormats1.Append(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U });
            cellFormats1.Append(new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U, FormatId = 1U });
            #endregion

            #region Decine Cell FOrmats
            Const.Formats.ForEach(cf =>
            {

                #region fonts
                var font2 = (Font)font1.Clone();
                if (cf.Font != null)
                {
                    font2 = new Font
                    {
                        FontSize = new FontSize() {Val = cf.Font.Size},
                        Color = new Color() {Rgb = cf.Font.Color.ToHexFormat()},
                        FontName = new FontName()
                        {
                            Val = string.IsNullOrEmpty(cf.Font.FontName) ? "Calibri" : cf.Font.FontName
                        },
                        Bold = cf.Font.Bold ? new Bold() {Val = cf.Font.Bold} : null,
                        Italic = cf.Font.Italic ? new Italic() {Val = cf.Font.Italic} : null,
                        Underline = cf.Font.Underline != OxUnderlines.None
                            ? new Underline() {Val = cf.Font.Underline.ToUnderlineValues()}
                            : null,
                        Condense = cf.Font.Condense ? new Condense() {Val = cf.Font.Condense} : null,
                        Extend = cf.Font.Extend ? new Extend() {Val = cf.Font.Extend} : null,
                        Outline = cf.Font.Outline ? new Outline() {Val = cf.Font.Outline} : null,
                        Shadow = cf.Font.Shadow ? new Shadow() {Val = cf.Font.Shadow} : null,
                        Strike = cf.Font.Strike ? new Strike() {Val = cf.Font.Strike} : null,
                        VerticalTextAlignment = cf.Font.VerticalAlignments != OxVerticalAlignments.Baseline
                            ? new VerticalTextAlignment()
                            {
                                Val = cf.Font.VerticalAlignments.ToVerticalAlignmentRunValues()
                            }
                            : null,
                        FontScheme = cf.Font.FontScheme != OxFontSchemes.None
                            ? new FontScheme {Val = cf.Font.FontScheme.ToFontSchemeValues()}
                            : null
                    };
                    AddColor(cf.Font.Color);
                }

                #endregion

                var fillIndexFormat =  0U;
                #region Fill
                var fill3 = (Fill)fill1.Clone();
                if (cf.Fill != null)
                {
                    fill3 = new Fill();
                    if (cf.Fill.PatternFill != null)
                    {
                        fill3.PatternFill = new PatternFill
                        {
                            PatternType = cf.Fill.PatternFill.PatternType.ToPatternValues(),
                            ForegroundColor =
                                cf.Fill.PatternFill.Color != System.Drawing.Color.Transparent
                                    ? new ForegroundColor {Rgb = cf.Fill.PatternFill.Color.ToHexFormat()}
                                    : new ForegroundColor() {Theme = 1U, Tint = 0.59999389629810485D},
                            BackgroundColor = new BackgroundColor() {Indexed = 64U}
                        };
                        AddColor(cf.Fill.PatternFill.Color);
                        fillIndexFormat = (uint)(Const.Formats.IndexOf(cf) + 2);
                    }
                    else if (cf.Fill.GradientFill != null && cf.Fill.GradientFill.GradientStops.Count > 0)
                    {
                        var gradient = new GradientFill() {Degree = cf.Fill.GradientFill.Degree};
                        cf.Fill.GradientFill.GradientStops.ForEach(c =>
                        {
                            gradient.Append(new GradientStop()
                            {
                                Position = c.Position,
                                Color = new Color() {Rgb = c.Color.ToHexFormat()}
                            });
                            AddColor(c.Color);
                        });
                        fill3 = new Fill() {GradientFill = gradient};
                        fillIndexFormat = (uint)(Const.Formats.IndexOf(cf) + 2);
                    }
                }
                #endregion

                #region Border
                var border2 = (Border)border1.Clone();
                if (cf.Borders != null)
                {
                    border2 = new Border();
                    var leftBorder2 = new LeftBorder();
                    if (cf.Borders.Left != null)
                    {
                        leftBorder2.Style = cf.Borders.Left.BorderStyle.ToBorderStyleValues();
                        leftBorder2.Color = !cf.Borders.Left.Color.IsEmpty? new Color() { Rgb = cf.Borders.Left.Color.ToHexFormat() } : new Color() { Indexed = 64U };
                        AddColor(cf.Borders.Left.Color);
                    }

                    var rightBorder2 = new RightBorder();
                    if (cf.Borders.Right != null)
                    {
                        rightBorder2.Style = cf.Borders.Right.BorderStyle.ToBorderStyleValues();
                        rightBorder2.Color = !cf.Borders.Right.Color.IsEmpty? new Color() { Rgb = cf.Borders.Right.Color.ToHexFormat() } : new Color() { Indexed = 64U };
                        AddColor(cf.Borders.Right.Color);
                    }

                    var topBorder2 = new TopBorder();
                    if (cf.Borders.Top != null)
                    {
                        topBorder2.Style = cf.Borders.Top.BorderStyle.ToBorderStyleValues();
                        topBorder2.Color = !cf.Borders.Top.Color.IsEmpty? new Color() { Rgb = cf.Borders.Top.Color.ToHexFormat() } : new Color() { Indexed = 64U };
                        AddColor(cf.Borders.Top.Color);
                    }

                    var bottomBorder2 = new BottomBorder();
                    if (cf.Borders.Bottom != null)
                    {
                        bottomBorder2.Style = cf.Borders.Bottom.BorderStyle.ToBorderStyleValues();
                        bottomBorder2.Color = !cf.Borders.Bottom.Color.IsEmpty? new Color() { Rgb = cf.Borders.Bottom.Color.ToHexFormat() } : new Color() { Indexed = 64U };
                        AddColor(cf.Borders.Bottom.Color);
                    }

                    var diagonalBorder2 = new DiagonalBorder();
                    if (cf.Borders.Diagonal != null)
                    {
                        diagonalBorder2.Style = cf.Borders.Diagonal.BorderStyle.ToBorderStyleValues();
                        diagonalBorder2.Color = !cf.Borders.Diagonal.Color.IsEmpty? new Color() { Rgb = cf.Borders.Diagonal.Color.ToHexFormat() } : new Color() { Indexed = 64U };
                        AddColor(cf.Borders.Diagonal.Color);
                        border2.DiagonalUp = cf.Borders.DiagonalUp;
                        border2.DiagonalDown = cf.Borders.DiagonalDown;                        
                    }
                    border2.Outline = cf.Borders.Outline;
                    border2.Append(leftBorder2);
                    border2.Append(rightBorder2);
                    border2.Append(topBorder2);
                    border2.Append(bottomBorder2);
                    border2.Append(diagonalBorder2); 
                }
                #endregion

                #region NumberFormat
                var applyNf = false;
                var nf = new NumberingFormat();
                if (cf.NumberFormat != null)
                {
                     nf = new NumberingFormat
                    {
                        NumberFormatId = iExcelIndex++,
                        FormatCode = cf.NumberFormat.FormatCode
                    };
                    applyNf = true;
                    nfs.Append(nf);
                }
                #endregion

                fonts1.Append(font2);
                fills1.Append(fill3);
                borders1.Append(border2);

                var indexFormat = (uint)(Const.Formats.IndexOf(cf) + 1);
                var fontIndexFormat = (uint)(Const.Formats.IndexOf(cf) + 2);
                
                var cellFormat4 = new CellFormat()
                {
                    NumberFormatId = nf.NumberFormatId,
                    FontId = fontIndexFormat,
                    FillId = fillIndexFormat,
                    BorderId = indexFormat,
                    FormatId = 0U,
                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = true,
                    ApplyAlignment = true,
                    ApplyNumberFormat = applyNf
                };

                if (cf.Alignment != null)
                    cellFormat4.Append(new Alignment()
                    {
                        Horizontal = cf.Alignment.HorizontalAlignment.ToHorizontalAlignmentValues(),
                        Vertical = cf.Alignment.VerticalAlignment.ToVerticalAlignmentValues(),
                        TextRotation = cf.Alignment.Rotation,
                        ShrinkToFit = cf.Alignment.ShrinkToFit,
                        JustifyLastLine = cf.Alignment.JustifyLastLine
                    });
                cellFormats1.Append(cellFormat4);

            });
            #endregion


            xw.WriteElement(nfs);
            xw.WriteElement(fonts1);
            xw.WriteElement(fills1);
            xw.WriteElement(borders1);
            xw.WriteElement(cellStyleFormats1);
            xw.WriteElement(cellFormats1);

            #region CellStyles
            xw.WriteStartElement(new CellStyles() { Count = 2U });
            xw.WriteElement( new CellStyle() { Name = "Hipervínculo", FormatId = 1U, BuiltinId = 8U });
            xw.WriteElement( new CellStyle() { Name = "Normal", FormatId = 0U, BuiltinId = 0U });
            xw.WriteEndElement();
            #endregion

            #region DifferentialFormats
            xw.WriteStartElement( new DifferentialFormats() {  Count = 1U });
            xw.WriteStartElement( new DifferentialFormat());
            xw.WriteElement( new NumberingFormat() { NumberFormatId = 19U, FormatCode = "dd/mm/yyyy" });
            xw.WriteEndElement();
            xw.WriteEndElement();
            #endregion

            #region TableStyles
            xw.WriteElement(new TableStyles() { Count = 0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" });
            #endregion

            #region Colors
            if (_DistinctColors.Count > 0)
            {
                xw.WriteStartElement( new Colors());
                xw.WriteStartElement( new MruColors());
                _DistinctColors.ForEach(dc => xw.WriteElement(new Color() { Rgb = dc }));
                xw.WriteEndElement();
                xw.WriteEndElement();
            }
            #endregion



            xw.WriteEndElement();
            xw.Close();
        }

        private uint? GetFormartIndex(OxCellFormartEntity format)
        {
            if (format == null) return null;
            var index = Const.Formats.FindIndex(i => i.Equals(format));
            if (index < 0)
            {
                Const.Formats.Add(format);
                index = Const.Formats.Count-1;
            }
            return (uint)(index + 2);
        }

        public void AddColor(System.Drawing.Color color)
        {
            if (!_DistinctColors.Contains(color.ToHexFormat()))
                _DistinctColors.Add(color.ToHexFormat());
        }

    }
}
