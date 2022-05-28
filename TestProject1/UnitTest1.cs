using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.Linq;
using MgSoftDev.OXExcel;
using MgSoftDev.OXExcel.Attributes;
using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.ColsRowsCells;
using MgSoftDev.OXExcel.Helpers.Extensions;
using NUnit.Framework;

namespace TestProject1
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

         [Test]
        public void EmpyDocument()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\EmpyDocument.xlsx";
            var doc = new OxExcelDocument();
            doc.AddSheet("Hoja 1").AddColumn(c => c.Add(1, 1).BestFit());
            doc.AddSheet("Hoja 2");
            doc.AddSheet("sheet 3");

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void DocumentProperty()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\DocumentProperty.xltx";
            var doc =
                new OxExcelDocument().Calculation(
                    c =>
                        c.CalculationIteration()
                            .CalculationMode(OxCalculateModes.AutoNoTable)
                            .ForceFullCalc()
                            .FullCalcOnLoad()
                            .FullPrecision()
                            .IterateCount(200)).DocumentType(OxDocumentTypes.Template)
                            .PackageProperties(p => p.Title("miguel").Company("MFGS").Created(DateTime.Now).Creator("yo").LastModifiedBy("mi").Modified(DateTime.Now).Version("1.0"));
            doc.AddSheet("Hoja 1").Cell(c => c.Add("A1").Value("hola"));

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void SheetProperty()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\SheetProperty.xlsx";
            var doc = new OxExcelDocument();
            doc.AddSheet("Hoja 1").AddColumn(c => c.Add(1, 3).Hidden()).BackGroundImage(new Uri(@"C:\Users\Migeru\Downloads\ife2.jpg"));
            doc.AddSheet("Hoja 2")
                .AddColumn(c => c.Add(1, 4).CollapsedOutlining().BestFit().OutlineLevel(1).Phonetic(false).Width(20)).SheetView(v => v.PaneFrozen("G3"))
                .Cell(c => c.Add("C3").Value("hello"))
                ;
            doc.AddSheet("h3")
                .AddColumn(c =>
                {
                    c.Add(12, 15).CollapsedOutlining().BestFit().OutlineLevel(1).Phonetic(false).Phonetic(false).Width(20);
                    c.Add(16, 16).CollapsedOutlining().BestFit().OutlineLevel(1).Phonetic(false).Phonetic(false).Width(20);
                    c.Add(17, 17).CollapsedOutlining().BestFit().OutlineLevel(1).Phonetic(false).Phonetic(false).Width(20);
                    c.Add(18, 18).CollapsedOutlining().BestFit().OutlineLevel(1).Phonetic(false).Phonetic(false).Width(20);
                })
                .PageMargins(m => m.Bottom(2).Left(1).Right(1).Top(1).Header(1).Footer(1))
                .PageSetup(
                    p =>
                        p.BlackAndWhite()
                            .Copies(5) //not found
                            .Draft(true)
                            .FirstPageNumber(1)
                            .Orientation(OxPageSetupOrientations.Landscape)
                            .PaperSize(OxPaperSizeDefault.A4));
            doc.AddSheet("h4")
                .SheetView(
                    s =>
                        s.HideGridLines()
                            .HideRowColHeaders()
                            .HideZeros()
                            .ShowFormulas()
                            .ShowRuler()
                            .TabSelected()
                            .ViewSheet(OxSheetViews.Normal)
                            .WindowProtection() //not found
                );

            doc.AddSheet("h5")
                .SheetProperties(p => p.EnableFormatConditionsCalculation().FilterMode().Published()
                    .OutlineProperties(o => o.OutlineSymbols().ApplyStyles().NotSummaryBelow().NotSummaryRight())
                    .TabColor(Color.Yellow))
                ;
            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void SimpleCells()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\SimpleCells.xlsx";
            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto));
            var sh1 = doc.AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("A1").Value("Reporte FInal").Hyperlink(new Uri("https://www.google.com.mx/", UriKind.Absolute), "este es mi link a google.com");
                        c.Add("B", 3).Value("Start Date:").Hyperlink("hoja2!A1", "goto shhet2");
                        c.Add("C", 3).Value(DateTime.Now);
                        c.Add("A", 4).Value(10);
                        c.Add("B", 4).Value(10).Phonetic();
                        c.Add("C", 4).Value(0).Formula("=A4+B4");
                    });
                sh.Add("hoja2");
            });
            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }
         [Test]
        public void RowCells()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\RowCells.xlsx";
            var doc = new OxExcelDocument();
            var sh1 = doc.AddSheet("Simple Cells");
            sh1.Cell("A1").Value("Reporte FInal");
            sh1.Cell(c =>
            {
                c.Add("B", 3).Value("Start Date:");
                c.Add("C", 3).Value(DateTime.Now);
            });
            sh1.AddRow(r =>
            {
                r.Add(1).Collapsed();
                r.Add(2).Collapsed().Height(35).OutlineLevel(1).ShowPhonetic().ThickBot().ThickTop();
                r.Add(3).Collapsed().OutlineLevel(1);
                r.Add(4).Collapsed().OutlineLevel(1);
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }


         [Test]
        public void MargenCell()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\MargenCell.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("A1").Value("Reporte FInal").Margen(4, 0);
                        c.Add("B", 3).Value("Start Date:");
                        c.Add("C", 3).Value(DateTime.Now).Margen(4, 4);
                    });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void FontFormatCell()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\FormatCell.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("A1").Value("Reporte FInal").Format(f =>
                        {
                            f.Font().Bold().Color(Color.Blue).Size(30).Condense().Extend();
                        });
                        c.Add("a2").Value(" FInal").Format(f =>
                        {
                            f.Font().Bold().FontName("Arial").FontScheme(OxFontSchemes.Minor).Italic().Outline().Shadow().Strike();
                        });

                        c.Add("a3").Value(" FInal").Format(f =>
                        {
                            f.Font().Underline(OxUnderlines.Double).VerticalAlignments(OxVerticalAlignments.Superscript);
                        });
                    });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void FillFormatCell()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\FillFormatCell.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("A1").Value("Reporte FInal").Format(f =>
                        {
                            f.FillPattern(Color.Blue, OxPatterns.Solid);
                        });
                        c.Add("a2").Value(" FInal").Format(f =>
                        {
                            f.FillPattern(Color.Blue, OxPatterns.Gray125);
                        });

                        c.Add("a3").Value(" FInal").Format(f =>
                        {
                            f.FillGradient(90, o =>
                            {
                                o.Add(Color.Blue, 0);
                                o.Add(Color.Yellow, 1);
                            });
                        });
                    });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void BorderFormatCell()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\BorderFormatCell.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("D3").Value("Reporte FInal").Format(f =>
                        {
                            f.Borders(b =>
                            {
                                b.Bottom(Color.Blue, OxBorderStyles.Double);
                                b.Top(Color.Yellow, OxBorderStyles.DashDotDot);
                                b.Left(Color.Gray, OxBorderStyles.Dotted);
                                b.Right(Color.Brown, OxBorderStyles.Thick);
                            });
                        });
                        c.Add("D6").Value("hola").Format(f =>
                        {
                            f.Borders()
                                .Diagonal(Color.Blue, OxBorderStyles.Double)
                                .Outline()
                                .DiagonalDown()
                                .DiagonalUp();
                        });
                    });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void NumberFormatCell()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\NumberFormatCell.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("D3").Value(DateTime.Now).Format(f =>
                        {
                            f.NumberFormat("dd/MM/yyyy hh:mm");
                        });
                        c.Add("a1").Value(DateTime.Now);
                        c.Add("b1").Value(DateTime.Now.TimeOfDay);
                        c.Add("c1").Value(123);
                        c.Add("d1").Value("123");
                    });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }


         [Test]
        public void AlignmentFormatCell()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\AlignmentFormatCell.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123).Margen(2, 2).Format(f =>
                        {
                            f.Alignment()
                                .Horizontal(OxTextHorizontalAlignments.Center)
                                .Vertical(OxTextVerticalAlignments.Center);
                        });
                        c.Add("a5").Value("123 asdsd asdad").Margen(2, 2).Format(f =>
                        {
                            f.Alignment().JustifyLastLine();
                        });
                        c.Add("f5").Value("123ssadasdd asd asda").Margen(2, 2).Format(f =>
                        {
                            f.Alignment().ShrinkToFit().Rotation(90);
                        });
                    });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void Image()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\Image.xlsx";
            var doc = new OxExcelDocument().AddSheet(sh =>
            {
                sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    }).AddImage(i =>
                    {
                        i.Add(new OxRangeEntity("B5:E24"))
                            .Name("img1")
                            .Url(@"C:\Users\Migeru\Downloads\ife2.jpg")
                            .Rectangle(new RectangleF(0, 0, 100, 50));

                        i.Add(new OxRangeEntity("B30:F40"))
                           .Name("img1")
                           .Url(@"C:\Users\Migeru\Downloads\ife2.jpg")
                           .Rectangle(new RectangleF(0, 0, 100, 100));

                        i.Add(new OxRangeEntity("J30:M40"))
                          .Name("img1")
                          .Url(@"C:\Users\Migeru\Downloads\ife2.jpg");
                    }).BackGroundImage(new Uri(@"C:\Users\Migeru\Downloads\ife2.jpg"));

            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void Table()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\Table.xlsx";
            var d = new List<data>()
            {
                new data {id = 1, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(5,5)}},
                new data {id = 2, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(10,10)}},
                new data {id = 3, fecha = DateTime.Now, nombre = "=5+100"}
            };
            object q = d.Select(s => new { s.fecha, s.id, num = 50 });
            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });
                sh1.AddTable(d, 1, 5, t =>
                {
                    t.Name("tbl1");
                    t.Columns(c =>
                    {
                        c.Add(p => p.id).Size(100).Filter(new[] { "1", "3" }).HyperlinkTemplate(e => new OxHyperlinkEntity(new Uri("http://www.bing.com/"), "Bing GO!!"));
                        c.Add(p => p.Data2.position.X).Header("pos x").DefaultVal("0.1");
                        c.Add(p => p.fecha)
                            .Type(typeof(string))
                            .TemplateValue(temp => "hola mundo" + temp.TableRowIndex);
                        c.Add(p => p.nombre)
                            .Header("Formula")
                            .IsFormula()
                            .CellType(OxCellTypeValues.Number)
                            .DefaultFormulaVal("0.1");
                        c.Add("Custom 1")
                            .Type(typeof(DateTime))
                            .TemplateValue(temp => temp.Row.GetPropertyVal("fecha"));
                    });
                    t.TableStyle(s => s.HideRowStripes().ShowColumnStripes().ShowFirstColumn().ShowLastColumn());
                });

                sh1.AddTable(q, 1, 20, t =>
                {
                    t.Columns(c =>
                    {
                        c.Add("id")
                            .Header("ID")
                            .Format(m => m.Font(mm => mm.Bold().Color(Color.Blue)))
                            .TemplateFormat(ff => ff.Format.Font(
                                mm =>
                                {
                                    mm.Bold(ff.CellValue.ToString() == "1");
                                }).FillPattern(Color.Black, OxPatterns.Solid));
                        c.Add("fecha").Type(typeof(DateTime)).Header("Fecha");
                        c.Add("Custom 1").DefaultVal("jajaj");
                    });
                    t.RowDefinition(r => r.Height(50).Format(f => f.FillPattern(Color.Yellow, OxPatterns.Solid)));

                });
                sh1.AddTable(q, 7, 20, t =>
                {
                    t.Columns(c =>
                    {
                        c.Add("id")
                            .Header("ID")
                            .Format(m => m.Font(mm => mm.Bold().Color(Color.Blue)))
                            .TemplateFormat(ff => ff.Format.Font(
                                mm =>
                                {
                                    mm.Bold(ff.CellValue.ToString() == "1");
                                }).FillPattern(Color.Black, OxPatterns.Solid));
                        c.Add("fecha").Type(typeof(DateTime)).Header("Fecha");
                        c.Add("Custom 1").DefaultVal("jajaj");
                    });
                    t.RowDefinition(r => r.Height(20).Format(f => f.FillPattern(Color.Yellow, OxPatterns.Solid)));
                    t.RowDefinitionTemplate(tt =>
                    {
                        if (tt.TableRowIndex % 2 != 0)
                            tt.RowDefinition.Height(40).Format().FillPattern(Color.Green, OxPatterns.Solid);
                        return tt.RowDefinition;
                    });
                });

            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }


         [Test]
        public void TableAttribute()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\TableAttribute.xlsx";
            var d = new List<data>()
            {
                new data {id = 1, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(5,5)}},
                new data {id = 2, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(10,10)}},
                new data {id = 3, fecha = DateTime.Now, nombre = "=5+100"}
            };
            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });
                sh1.AddTable(d, 1, 5, t =>
                {
                    t.Name("tbl1");
                    t.Columns(c =>
                    {
                        c.Add(p => p.id);
                        c.Add(p => p.Data2.position.X);
                        c.Add(p => p.fecha);
                        c.Add(p => p.nombre);
                        c.Add("Custom 1").Order(1);
                    });

                    t.AutoGenerateColumns(true);
                    t.TableStyle(s => s.HideRowStripes().ShowColumnStripes().ShowFirstColumn().ShowLastColumn());
                });




            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }


         [Test]
        public void TableMassibeData()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\TableMassibeData.xlsx";
            var d = new List<data>()
            {
                new data {id = 1, fecha = DateTime.Now, nombre = "100",Data2 = new data2() {position = new Point(5,5)}},
                new data {id = 2, fecha = DateTime.Now, nombre = "100",Data2 = new data2() {position = new Point(10,10)}},
                new data {id = 3, fecha = DateTime.Now, nombre = "100"}
            };

            for (int i = 0; i < 10000; i++)
            {
                d.Add(new data
                {
                    id = i,
                    fecha = DateTime.Now,
                    nombre = "5",
                    Data2 = new data2() { position = new Point(5, 5) }
                });
            }

            object q = d.Select(s => new { s.fecha, s.id, num = 50 });
            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });
                sh1.AddTable(d, 1, 5, t =>
                {
                    t.Name("tbl1");
                    t.Columns(c =>
                    {
                        c.Add(p => p.id).Size(100);//.Filter(new[] {"1", "3"});
                        c.Add(p => p.Data2.position.X).Header("pos x").DefaultVal("0.1");
                        c.Add(p => p.fecha)
                            .Type(typeof(string))
                            ;//.TemplateValue(temp => "hola mundo" + temp.TableRowIndex);
                        c.Add(p => p.nombre)
                            .Header("Formula")
                            // .IsFormula()
                            .CellType(OxCellTypeValues.Number)
                            .DefaultFormulaVal("0.1");
                        c.Add("Custom 1")
                            .Type(typeof(DateTime))
                            ;//.TemplateValue(temp => temp.Rows.GetPropertyVal("fecha"));
                    });

                });




            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }



         [Test]
        public void TableFromDataTable()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\TableFromDataTable.xlsx";
            var d = new List<data>()
            {
                new data {id = 1, fecha = DateTime.Now, nombre = "100"},
                new data {id = 2, fecha = null, nombre = "100"},
                new data {id = 3, fecha = DateTime.Now, nombre = "100"}
            };

            for (int i = 0; i < 10000; i++)
            {
                d.Add(new data
                {
                    id = i,
                    fecha = DateTime.Now,
                    nombre = "5",

                });
            }

            var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("fecha", typeof(DateTime)).AllowDBNull = true;
            table.Columns.Add("nombre", typeof(string));

            d.ForEach(f =>
            {
                table.Rows.Add(f.id, f.fecha, f.nombre);
            });
            var dynamictable = table.ToDynamicList();

            object q = d.Select(s => new { s.fecha, s.id, num = 50 });
            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });
                sh1.AddTable(dynamictable, 1, 5, t =>
                {
                    t.Name("tbl1");
                    t.Columns(c =>
                    {
                        c.Add("id").Size(100).CellType(OxCellTypeValues.Number);//.Filter(new[] {"1", "3"});

                        c.Add("fecha")
                            .Type(typeof(DateTime))
                            ;//.TemplateValue(temp => "hola mundo" + temp.TableRowIndex);
                        c.Add("nombre")
                            .Header("Formula")
                            // .IsFormula()
                            .CellType(OxCellTypeValues.Number)
                            .DefaultFormulaVal("0.1");
                        c.Add("Custom 1")
                            .Type(typeof(DateTime))
                            .TemplateValue(temp => temp.Row.GetPropertyVal("fecha"));
                    });

                });




            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }


         [Test]
        public void TableAutoGenerateColumns()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\Table Autogenerate columns.xlsx";
            var d = new List<dataFull>();



            for (int i = 0; i < 20000; i++)
            {
                d.Add(new dataFull());
            }




            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });
                sh1.AddTable(d, 1, 5, t =>
                {
                    t.Name("tbl1").AutoGenerateColumns();
                    // t.Columns(c=>
                    // c.Add("nombre").Type(typeof(int?)));

                });




            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }


         [Test]
        public void TableFilterCustom()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\TableFilterCustom.xlsx";
            var d = new List<data>()
            {
                new data {id = 1, fecha = DateTime.Now, nombre = "miguel",Data2 = new data2() {position = new Point()}},
                new data {id = 2, fecha = DateTime.Now, nombre = "miguel garcia",Data2 = new data2() {position = new Point()}},
                new data {id = 3, fecha = DateTime.Now, nombre = "lalo"},
                new data {id = 3, fecha = DateTime.Now, nombre = "lalo"},
                new data {id = 3, fecha = DateTime.Now, nombre = "lalo"},
                new data {id = 3, fecha = DateTime.Now, nombre = "miguel maria"},
                new data {id = 3, fecha = DateTime.Now, nombre = "lalo"},
                new data {id = 3, fecha = DateTime.Now, nombre = "lalo"},
                new data {id = 3, fecha = DateTime.Now, nombre = "lalo"},
            };
            var rIndex = 5U;
            var q = d.Select(s => new { s.nombre, col2 = "miguel", col3 = "fer", bol = true, num = 100 });
            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("Simple Cells")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });
                sh1.AddTable(d, 3, rIndex, t =>
                {
                    t.Name("Equal");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre).Size(40).Filter("miguel", OxFilterOperators.Equal);
                    });
                });
                sh1.AddTable(d, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("Not Equal");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre).Size(40).Filter("miguel", OxFilterOperators.NotEqual);
                    });
                });
                sh1.AddTable(d, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("Start w");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre).Size(40).Filter("mi", OxFilterOperators.StartWith);
                    });
                });
                sh1.AddTable(d, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("End w");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre).Size(40).Filter("lo", OxFilterOperators.EndWith);
                    });
                });
                sh1.AddTable(d, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("Contrains w");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre).Size(40).Filter("guel", OxFilterOperators.Contrains);
                    });
                });
                sh1.AddTable(d, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("not Contrains w");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre).Size(40).Filter("gel", OxFilterOperators.NotContrains);
                    });
                });
                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("Equal Num ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.num).Size(40).Filter(100, OxFilterOperators.Equal);
                    });
                });
                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("LessThan Num ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.num).Size(40).Filter(100, OxFilterOperators.LessThan);
                    });
                });
                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("LessThanOrEqual Num ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.num).Size(40).Filter(100, OxFilterOperators.LessThanOrEqual);
                    });
                });
                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("GreaterThanOrEqual Num ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.num).Size(40).Filter(100, OxFilterOperators.GreaterThanOrEqual);
                    });
                });
                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("GreaterThan Num ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.num).Size(40).Filter(100, OxFilterOperators.GreaterThan);
                    });
                });

                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("And Condition ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre)
                            .Size(40)
                            .Filter("guel", OxFilterOperators.Contrains, OxCustomFilterCondition.And, "m",
                                OxFilterOperators.StartWith);
                    });
                });

                sh1.AddTable(q, 3, (uint)(rIndex += 15), t =>
                {
                    t.Name("Or Condition ");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre)
                            .Size(40)
                            .Filter("guelf", OxFilterOperators.Contrains, OxCustomFilterCondition.Or, "l",
                                OxFilterOperators.StartWith);
                    });
                });

                sh1.AddTable(q, 3, rIndex += 15, t =>
                {
                    t.Name("bolean condition");
                    t.Columns(c =>
                    {
                        c.Add(p => p.bol)
                            .Size(40)
                            .Filter(true, OxFilterOperators.NotEqual);
                    });
                });

                sh1.AddTable(q, 3, rIndex += 15, t =>
                {
                    t.Name("Multi column condition");
                    t.Columns(c =>
                    {
                        c.Add(p => p.nombre)
                            .Size(40)
                            .Filter("l", OxFilterOperators.Contrains, OxCustomFilterCondition.And, "a", OxFilterOperators.EndWith);

                        c.Add(p => p.bol).Filter(true, OxFilterOperators.Equal);
                        c.Add(p => p.num).Filter(500, OxFilterOperators.GreaterThan);
                    });

                });
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

         [Test]
        public void TableTotalRows()
        {
            var path = @"C:\Users\miger\Downloads\Nueva carpeta (8)\TableTotalRows.xlsx";
            var d = new List<data>()
            {
                new data {id = 1, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(5,5)}},
                new data {id = 2, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(5,5)}},
                new data {id = 3, fecha = DateTime.Now, nombre = "=5+100",Data2 = new data2() {position = new Point(5,5)}}
            };

            var doc = new OxExcelDocument().Calculation(c => c.CalculationMode(OxCalculateModes.Auto)).AddSheet(sh =>
            {
                var sh1 = sh.Add("TableTotalRows")
                    .Cell(c =>
                    {
                        c.Add("a1").Value(123);
                    });

                sh1.AddTable(d, 1, 5, t =>
                {
                    t.Name("mytablita");
                    t.Columns(c =>
                    {
                        c.Add(p => p.id).Header("ID").TotalRow("total").HeaderFormat(f => f.FillPattern(Color.DarkOrange, OxPatterns.Solid));
                        c.Add(p => p.Data2.position.X).Size(30).Header("numero").TotalRow(TotalsRowFormulas.Sum, false).TotalRowFormat(f => f.Font(ff => ff.Size(20)).NumberFormat("0.00%")).HeaderFormat(f => f.Font(ff => ff.Size(20)));
                    }).TotalsRowShown();
                    t.RowDefinition(r => r.Height(20));
                });
                sh1.AddTable(d, 1, 40, t =>
                {
                    t.Columns(c =>
                    {
                        c.Add(p => p.id).Header("ID").TotalRow("total");
                        c.Add(p => p.Data2.position.X).Header("numero").TotalRow(TotalsRowFormulas.Sum, true);
                    }).TotalsRowShown();
                    t.RowDefinition(r => r.Height(20));
                });
                sh1.AddTable(d, 1, 20, t =>
                {
                    t.Columns(c =>
                    {
                        c.Add(p => p.id).Header("ID").TotalRow("total");
                        c.Add(p => p.Data2.position.X).Header("numero").TotalRow("=SUBTOTAL(1,[numero])", false);
                    }).TotalsRowShown();
                    t.RowDefinition(r => r.Height(20));
                }); //=SUBTOTALES(2,[numero])
            });

            doc.Save(path);
            System.Diagnostics.Process.Start(path);
        }

    }




    public class data
    {
        [DisplayName("My ID")]
        [OxColumn(Header = "ID", Order = 3, Size = 20, CellTypeValue = OxCellTypeValues.String, DefaultFormulaValue = "0", ShowPhonetic = true, DefaultValue = "0")]
        public int id { get; set; }
        [OxColumn(Header = "Nombre")]
        public string nombre { get; set; }
        [Display(Order = 2)]
        [OxColumn(Header = "FECHA", Size = 20, DefaultFormulaValue = "0", ShowPhonetic = true, DefaultValue = "0")]
        public DateTime? fecha { get; set; }
        public data2 Data2 { get; set; }
    }

    public class data2
    {

        public Point position { get; set; }
    }


    public class dataFull
    {
        public string Dato0 { get; set; } = "Dato 0";
        public string Dato1 { get; set; } = "Dato 1";
        public string Dato2 { get; set; } = "Dato 2";
        public string Dato3 { get; set; } = "Dato 3";
        public string Dato4 { get; set; } = "Dato 4";
        public string Dato5 { get; set; } = "Dato 5";
        public string Dato6 { get; set; } = "Dato 6";
        public string Dato7 { get; set; } = "Dato 7";
        public string Dato8 { get; set; } = "Dato 8";
        public string Dato9 { get; set; } = "Dato 9";
        public string Dato10 { get; set; } = "Dato 10";
        public string Dato11 { get; set; } = "Dato 11";
        public string Dato12 { get; set; } = "Dato 12";
        public string Dato13 { get; set; } = "Dato 13";
        public string Dato14 { get; set; } = "Dato 14";
        public string Dato15 { get; set; } = "Dato 15";
        public string Dato16 { get; set; } = "Dato 16";
        public string Dato17 { get; set; } = "Dato 17";
        public string Dato18 { get; set; } = "Dato 18";
        public string Dato19 { get; set; } = "Dato 19";
        public string Dato20 { get; set; } = "Dato 20";
        public string Dato21 { get; set; } = "Dato 21";
        public string Dato22 { get; set; } = "Dato 22";
        public string Dato23 { get; set; } = "Dato 23";
        public string Dato24 { get; set; } = "Dato 24";
        public string Dato25 { get; set; } = "Dato 25";
        public string Dato26 { get; set; } = "Dato 26";
        public string Dato27 { get; set; } = "Dato 27";
        public string Dato28 { get; set; } = "Dato 28";
        public string Dato29 { get; set; } = "Dato 29";
        public string Dato30 { get; set; } = "Dato 30";
        public string Dato31 { get; set; } = "Dato 31";
        public string Dato32 { get; set; } = "Dato 32";
        public string Dato33 { get; set; } = "Dato 33";
        public string Dato34 { get; set; } = "Dato 34";
        public string Dato35 { get; set; } = "Dato 35";
        public string Dato36 { get; set; } = "Dato 36";
        public string Dato37 { get; set; } = "Dato 37";
        public string Dato38 { get; set; } = "Dato 38";
        public string Dato39 { get; set; } = "Dato 39";
        public string Dato40 { get; set; } = "Dato 40";
        public string Dato41 { get; set; } = "Dato 41";
        public string Dato42 { get; set; } = "Dato 42";
        public string Dato43 { get; set; } = "Dato 43";
        public string Dato44 { get; set; } = "Dato 44";
        public string Dato45 { get; set; } = "Dato 45";
        public string Dato46 { get; set; } = "Dato 46";
        public string Dato47 { get; set; } = "Dato 47";
        public string Dato48 { get; set; } = "Dato 48";
        public string Dato49 { get; set; } = "Dato 49";
        public string Dato50 { get; set; } = "Dato 50";
        public string Dato51 { get; set; } = "Dato 51";
        public string Dato52 { get; set; } = "Dato 52";
        public string Dato53 { get; set; } = "Dato 53";
        public string Dato54 { get; set; } = "Dato 54";
        public string Dato55 { get; set; } = "Dato 55";
        public string Dato56 { get; set; } = "Dato 56";
        public string Dato57 { get; set; } = "Dato 57";
        public string Dato58 { get; set; } = "Dato 58";
        public string Dato59 { get; set; } = "Dato 59";
        public string Dato60 { get; set; } = "Dato 60";
        public string Dato61 { get; set; } = "Dato 61";
        public string Dato62 { get; set; } = "Dato 62";
        public string Dato63 { get; set; } = "Dato 63";
        public string Dato64 { get; set; } = "Dato 64";
        public string Dato65 { get; set; } = "Dato 65";
        public string Dato66 { get; set; } = "Dato 66";
        public string Dato67 { get; set; } = "Dato 67";
        public string Dato68 { get; set; } = "Dato 68";
        public string Dato69 { get; set; } = "Dato 69";
        public string Dato70 { get; set; } = "Dato 70";
        public string Dato71 { get; set; } = "Dato 71";
        public string Dato72 { get; set; } = "Dato 72";
        public string Dato73 { get; set; } = "Dato 73";
        public string Dato74 { get; set; } = "Dato 74";
        public string Dato75 { get; set; } = "Dato 75";
        public string Dato76 { get; set; } = "Dato 76";
        public string Dato77 { get; set; } = "Dato 77";
        public string Dato78 { get; set; } = "Dato 78";
        public string Dato79 { get; set; } = "Dato 79";
        public string Dato80 { get; set; } = "Dato 80";
        public string Dato81 { get; set; } = "Dato 81";
        public string Dato82 { get; set; } = "Dato 82";
        public string Dato83 { get; set; } = "Dato 83";
        public string Dato84 { get; set; } = "Dato 84";
        public string Dato85 { get; set; } = "Dato 85";
        public string Dato86 { get; set; } = "Dato 86";
        public string Dato87 { get; set; } = "Dato 87";
        public string Dato88 { get; set; } = "Dato 88";
        public string Dato89 { get; set; } = "Dato 89";
        public string Dato90 { get; set; } = "Dato 90";
        public string Dato91 { get; set; } = "Dato 91";
        public string Dato92 { get; set; } = "Dato 92";
        public string Dato93 { get; set; } = "Dato 93";
        public string Dato94 { get; set; } = "Dato 94";
        public string Dato95 { get; set; } = "Dato 95";
        public string Dato96 { get; set; } = "Dato 96";
        public string Dato97 { get; set; } = "Dato 97";
        public string Dato98 { get; set; } = "Dato 98";
        public string Dato99 { get; set; } = "Dato 99";


    }
}