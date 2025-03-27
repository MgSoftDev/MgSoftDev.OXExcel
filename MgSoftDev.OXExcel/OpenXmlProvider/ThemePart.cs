using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace MgSoftDev.OXExcel.OpenXmlProvider
{
    internal partial class OpenXmlExcelProvider
    {
        
        private void GenerateThemePart1Content(WorkbookPart workbookPart)
        {
            var themePart1 = workbookPart.AddNewPart<ThemePart>("rId3");
            var theme1     = new A.Theme() { Name = "Tema de Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var themeElements1 = new A.ThemeElements();

            var colorScheme1 = new A.ColorScheme() { Name = "Office" };

            var dark1Color1 = new A.Dark1Color
            {
                SystemColor = new A.SystemColor() {Val = A.SystemColorValues.WindowText, LastColor = "000000"}
            };

            var light1Color1 = new A.Light1Color
            {
                SystemColor = new A.SystemColor() {Val = A.SystemColorValues.Window, LastColor = "FFFFFF"}
            };

            var dark2Color1 = new A.Dark2Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "44546A"}};

            var light2Color1 = new A.Light2Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "E7E6E6"}};

            var accent1Color1 = new A.Accent1Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "5B9BD5"}};

            var accent2Color1 = new A.Accent2Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "ED7D31"}};

            var accent3Color1 = new A.Accent3Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "A5A5A5"}};

            var accent4Color1 = new A.Accent4Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "FFC000"}};

            var accent5Color1 = new A.Accent5Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "4472C4"}};

            var accent6Color1 = new A.Accent6Color {RgbColorModelHex = new A.RgbColorModelHex() {Val = "70AD47"}};

            var hyperlink1 = new A.Hyperlink {RgbColorModelHex = new A.RgbColorModelHex() {Val = "0563C1"}};

            var followedHyperlinkColor1 = new A.FollowedHyperlinkColor
            {
                RgbColorModelHex = new A.RgbColorModelHex() {Val = "954F72"}
            };

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            var fontScheme1 = new A.FontScheme() { Name = "Office" };
            var fontScheme3 = new A.FontScheme() { Name = "Office" };
            var majorFont1 = new A.MajorFont();
            var latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            var eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            var complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            var supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            var supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            var supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            var supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            var supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            var supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            var supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            var supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            var supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            var supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            var supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            var supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            var supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            var supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            var supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            var supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            var supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            var supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            var supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            var supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            var supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            var supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            var supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            var supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            var supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            var supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            var supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            var supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            var supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            var supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            var minorFont1 = new A.MinorFont();
            var latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            var eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            var complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            var supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            var supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            var supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            var supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            var supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            var supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            var supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            var supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            var supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            var supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            var supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            var supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            var supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            var supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            var supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            var supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            var supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            var supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            var supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            var supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            var supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            var supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            var supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            var supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            var supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            var supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            var supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            var supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            var supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            var supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            var formatScheme1 = new A.FormatScheme() { Name = "Office" };

            var fillStyleList1 = new A.FillStyleList();

            var solidFill1 = new A.SolidFill();
            var schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            var gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            var gradientStopList1 = new A.GradientStopList();

            var gradientStop1 = new A.GradientStop() { Position = 0 };

            var schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            var saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            var tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            var gradientStop2 = new A.GradientStop() { Position = 50000 };

            var schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            var saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            var tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            var gradientStop3 = new A.GradientStop() { Position = 100000 };

            var schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            var saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            var tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            var linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            var gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            var gradientStopList2 = new A.GradientStopList();

            var gradientStop4 = new A.GradientStop() { Position = 0 };

            var schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            var luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            var tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            var gradientStop5 = new A.GradientStop() { Position = 50000 };

            var schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            var luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            var shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            var gradientStop6 = new A.GradientStop() { Position = 100000 };

            var schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            var saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            var shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            var linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            var lineStyleList1 = new A.LineStyleList();

            var outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            var solidFill2 = new A.SolidFill();
            var schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            var presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            var miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            var outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            var solidFill3 = new A.SolidFill();
            var schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            var presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            var miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            var outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            var solidFill4 = new A.SolidFill();
            var schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            var presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            var miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            var effectStyleList1 = new A.EffectStyleList();

            var effectStyle1 = new A.EffectStyle();
            var effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            var effectStyle2 = new A.EffectStyle();
            var effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            var effectStyle3 = new A.EffectStyle();

            var effectList3 = new A.EffectList();

            var outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            var rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            var alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            var backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            var solidFill5 = new A.SolidFill();
            var schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            var solidFill6 = new A.SolidFill();

            var schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var tint5 = new A.Tint() { Val = 95000 };
            var saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            var gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            var gradientStopList3 = new A.GradientStopList();

            var gradientStop7 = new A.GradientStop() { Position = 0 };

            var schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var tint6 = new A.Tint() { Val = 93000 };
            var saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            var shade3 = new A.Shade() { Val = 98000 };
            var luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            var gradientStop8 = new A.GradientStop() { Position = 50000 };

            var schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var tint7 = new A.Tint() { Val = 98000 };
            var saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            var shade4 = new A.Shade() { Val = 90000 };
            var luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            var gradientStop9 = new A.GradientStop() { Position = 100000 };

            var schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            var shade5 = new A.Shade() { Val = 63000 };
            var saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            var linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            var objectDefaults1 = new A.ObjectDefaults();
            var extraColorSchemeList1 = new A.ExtraColorSchemeList();

            var extensionList1 = new A.ExtensionList();

            var extension1 = new A.Extension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            var openXmlUnknownElement3 = workbookPart.CreateUnknownElement("<thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\" />");

            extension1.Append(openXmlUnknownElement3);

            extensionList1.Append(extension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(extensionList1);

            themePart1.Theme = theme1;
        }
        
    }
}
