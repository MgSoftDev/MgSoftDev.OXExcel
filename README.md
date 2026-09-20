# MgSoftDev.OXExcel

Una biblioteca .NET potente y fluida para la generación de documentos Excel (.xlsx) utilizando OpenXML. Diseñada para ofrecer una API intuitiva con patrón Builder/Factory que facilita la creación de hojas de cálculo complejas con formato avanzado.

## Tabla de Contenidos

- [Novedades](#novedades)
- [Instalacion](#instalacion)
- [Caracteristicas](#caracteristicas)
- [Inicio Rapido](#inicio-rapido)
- [API Reference](#api-reference)
  - [OxExcelDocument](#oxexceldocument)
  - [Hojas (Sheets)](#hojas-sheets)
  - [Celdas](#celdas)
  - [Filas y Columnas](#filas-y-columnas)
  - [Formato](#formato)
  - [Tablas](#tablas)
  - [Imagenes](#imagenes)
- [Ejemplos Completos](#ejemplos-completos)
- [Enumeraciones](#enumeraciones)
- [Atributos](#atributos)
- [Limites de Excel](#limites-de-excel)
- [Consideraciones de Rendimiento](#consideraciones-de-rendimiento)

---

## Novedades

### 1.0.7

- Las celdas combinadas, los hipervinculos y el rango declarado de la hoja (`dimension`) pasan a ser **de cada
  hoja**. Antes vivian en el documento, asi que en un libro de varias hojas la segunda repetia las combinaciones y
  los hipervinculos de la primera —con referencias a celdas que ahi son otra cosa y la URL escrita como valor— y su
  rango arrastraba el ancho y el alto de las anteriores.

### 1.0.6

- **Las filas de las tablas se escriben directo al archivo, una a la vez**, en vez de armar toda la hoja en memoria.
  Una tabla de 50 mil filas por 26 columnas pasa de 8,000 MB a 1,429 MB asignados y de 9.2 s a 4.3 s, con el mismo
  XML. No hay que cambiar nada: `AddTable` ya se comporta asi.
- Nuevo [`AddTableStream`](#tablas-grandes-addtablestream) para tablas que ni siquiera caben en memoria como lista.
- Nuevo [`TemplateCellType`](#tipo-de-celda-por-fila) para columnas donde conviven numeros y texto.
- Se avisa al pasar los [limites de Excel](#limites-de-excel) en vez de escribir un archivo que Excel no abre.
- Corregido: cuando un formato traia `Fill` o `Borders` y el otro no, `Combine` lanzaba una excepcion que se
  perdia en un `catch`, y la fila **se quedaba sin celdas en silencio**. Un reporte que use `RowDefinitionTemplate`
  con relleno junto a columnas con bordes puede salir ahora con celdas que antes faltaban.

---

## Instalacion

### NuGet Package

```bash
dotnet add package MgSoftDev.OXExcel
```

### Frameworks Soportados

- .NET 9.0
- .NET Standard 2.0
- .NET Framework 4.8

### Dependencias

- DocumentFormat.OpenXml (v3.3.0)
- Newtonsoft.Json (v13.0.3)
- System.Drawing.Common (v9.0.3)
- System.ComponentModel.Annotations (v5.0.0)
- System.Reflection.Emit (v4.7.0)

---

## Caracteristicas

- **API Fluida**: Construccion de documentos mediante encadenamiento de metodos
- **Multi-framework**: Compatible con .NET moderno y .NET Framework
- **Formato Completo**: Fuentes, bordes, rellenos, alineacion, degradados
- **Tablas Dinamicas**: Generacion automatica a partir de colecciones de datos
- **Soporte de Imagenes**: Insercion desde URL o bytes
- **Hipervinculos**: URLs externas y referencias internas
- **Filtros y Totales**: Filtros avanzados y filas de totales en tablas
- **Propiedades de Pagina**: Margenes, orientacion, tamaño de papel
- **Thread-safe**: Colecciones seguras para acceso concurrente
- **Gestion de Memoria**: Implementa IDisposable y escribe las filas de las tablas sin acumularlas en memoria

---

## Inicio Rapido

### Documento Vacio

```csharp
using MgSoftDev.OXExcel;

var doc = new OxExcelDocument();
doc.AddSheet("Hoja 1");
doc.AddSheet("Hoja 2");
doc.Save("documento.xlsx");
```

### Documento con Celdas Simples

```csharp
using MgSoftDev.OXExcel;
using MgSoftDev.OXExcel.Commons;

var doc = new OxExcelDocument()
    .Calculation(c => c.CalculationMode(OxCalculateModes.Auto));

doc.AddSheet("Datos")
    .Cell(c =>
    {
        c.Add("A1").Value("Titulo del Reporte");
        c.Add("A2").Value("Fecha:");
        c.Add("B2").Value(DateTime.Now);
        c.Add("A3").Value(100);
        c.Add("B3").Value(200);
        c.Add("C3").Value(0).Formula("=A3+B3");
    });

doc.Save("reporte.xlsx");
```

### Guardar en Stream

```csharp
using var ms = new MemoryStream();
doc.Save(ms);
File.WriteAllBytes("documento.xlsx", ms.ToArray());
```

---

## API Reference

### OxExcelDocument

Clase principal para la creacion de documentos Excel.

#### Constructor

```csharp
var doc = new OxExcelDocument();
```

#### Metodos Principales

| Metodo | Descripcion | Retorno |
|--------|-------------|---------|
| `DocumentType(OxDocumentTypes)` | Define el tipo de documento | `OxExcelDocument` |
| `DataCultureInfo(CultureInfo)` | Establece la cultura para datos | `OxExcelDocument` |
| `Calculation(Action<OxCalculationFactory>)` | Configura propiedades de calculo | `OxExcelDocument` |
| `PackageProperties(Action<OxPackagePropertiesFactory>)` | Propiedades del paquete | `OxExcelDocument` |
| `AddSheet(string)` | Agrega una hoja nueva | `OxSheetFactory` |
| `AddSheet(Action<OxSheetsFactory>)` | Agrega multiples hojas | `OxExcelDocument` |
| `Save(string)` | Guarda en archivo | `void` |
| `Save(Stream)` | Guarda en stream | `void` |
| `Dispose()` | Libera recursos | `void` |

#### Tipos de Documento

```csharp
doc.DocumentType(OxDocumentTypes.Workbook);      // .xlsx (default)
doc.DocumentType(OxDocumentTypes.Template);       // .xltx
doc.DocumentType(OxDocumentTypes.MacroEnabledWorkbook); // .xlsm
doc.DocumentType(OxDocumentTypes.MacroEnabledTemplate); // .xltm
doc.DocumentType(OxDocumentTypes.AddIn);          // .xlam
```

#### Propiedades del Paquete

```csharp
doc.PackageProperties(p => p
    .Title("Mi Reporte")
    .Creator("Usuario")
    .Company("Mi Empresa")
    .Created(DateTime.Now)
    .LastModifiedBy("Editor")
    .Modified(DateTime.Now)
    .Version("1.0"));
```

#### Configuracion de Calculo

```csharp
doc.Calculation(c => c
    .CalculationMode(OxCalculateModes.Auto)
    .CalculationIteration()
    .ForceFullCalc()
    .FullCalcOnLoad()
    .FullPrecision()
    .IterateCount(200));
```

---

### Hojas (Sheets)

#### Crear Hojas

```csharp
// Forma simple
var sheet = doc.AddSheet("Mi Hoja");

// Multiples hojas con Action
doc.AddSheet(sh =>
{
    sh.Add("Hoja 1");
    sh.Add("Hoja 2");
    sh.Add("Hoja 3");
});
```

#### Visibilidad de Hoja

```csharp
sheet.SheetVisibility(OxSheetVisibilities.Visible);   // Visible (default)
sheet.SheetVisibility(OxSheetVisibilities.Hidden);    // Oculta
sheet.SheetVisibility(OxSheetVisibilities.VeryHidden); // Muy oculta
```

#### Vista de Hoja (SheetView)

```csharp
sheet.SheetView(v => v
    .HideGridLines()           // Ocultar lineas de cuadricula
    .HideRowColHeaders()       // Ocultar encabezados de fila/columna
    .HideZeros()               // Ocultar ceros
    .ShowFormulas()            // Mostrar formulas
    .ShowRuler()               // Mostrar regla
    .TabSelected()             // Marcar como seleccionada
    .ViewSheet(OxSheetViews.Normal)
    .PaneFrozen("G3"));        // Congelar paneles en G3
```

#### Configuracion de Pagina

```csharp
sheet.PageSetup(p => p
    .Orientation(OxPageSetupOrientations.Landscape)
    .PaperSize(OxPaperSizeDefault.A4)
    .BlackAndWhite()
    .Draft(true)
    .Copies(5)
    .FirstPageNumber(1));

sheet.PageMargins(m => m
    .Top(1)
    .Bottom(2)
    .Left(1)
    .Right(1)
    .Header(0.5)
    .Footer(0.5));
```

#### Imagen de Fondo

```csharp
sheet.BackGroundImage(new Uri(@"C:\imagenes\fondo.jpg"));
```

---

### Celdas

#### Agregar Celdas

```csharp
sheet.Cell(c =>
{
    // Por referencia de celda
    c.Add("A1").Value("Texto");

    // Por columna (string) y fila
    c.Add("B", 3).Value("Otro texto");

    // Por columna (numero) y fila
    c.Add(3, 5).Value(123);
});
```

#### Tipos de Valores

```csharp
c.Add("A1").Value("Texto");              // String
c.Add("A2").Value(123);                  // Numero
c.Add("A3").Value(123.45);               // Decimal
c.Add("A4").Value(DateTime.Now);         // Fecha
c.Add("A5").Value(TimeSpan.FromHours(2)); // Hora
c.Add("A6").Value(true);                 // Booleano
```

#### Formulas

```csharp
c.Add("C1").Value(0).Formula("=A1+B1");
c.Add("D1").Value(0).Formula("=SUM(A1:C1)");
c.Add("E1").Value(0).Formula("=AVERAGE(A1:D1)");
```

#### Hipervinculos

```csharp
// URL externa
c.Add("A1").Value("Google")
    .Hyperlink(new Uri("https://www.google.com"), "Ir a Google");

// Referencia interna
c.Add("A2").Value("Ir a Hoja2")
    .Hyperlink("Hoja2!A1", "Navegar a Hoja2");
```

#### Combinar Celdas (Margen)

```csharp
// Combinar 2 columnas y 0 filas adicionales
c.Add("A1").Value("Titulo").Margen(2, 0);

// Combinar 4 columnas y 4 filas
c.Add("C3").Value(DateTime.Now).Margen(4, 4);
```

---

### Filas y Columnas

#### Definir Columnas

```csharp
sheet.AddColumn(c =>
{
    c.Add(1, 1).BestFit();                    // Columna A con ajuste automatico
    c.Add(2, 5).Width(20);                    // Columnas B-E con ancho 20
    c.Add(6, 6).Hidden();                     // Columna F oculta
    c.Add(7, 10).OutlineLevel(1).Collapsed(); // Columnas agrupadas
});
```

#### Metodos de Columna

| Metodo | Descripcion |
|--------|-------------|
| `BestFit()` | Ajuste automatico de ancho |
| `Width(double)` | Ancho especifico |
| `Hidden()` | Ocultar columna |
| `OutlineLevel(byte)` | Nivel de agrupacion |
| `CollapsedOutlining()` | Colapsar grupo |
| `Phonetic(bool)` | Mostrar fonetico |
| `Format(Action<OxCellFormartFactory>)` | Formato de columna |

#### Definir Filas

```csharp
sheet.AddRow(r =>
{
    r.Add(1).Height(25);                      // Fila 1 con altura 25
    r.Add(2).Collapsed().OutlineLevel(1);     // Fila 2 colapsada
    r.Add(3).ThickBot().ThickTop();           // Bordes gruesos
    r.Add(4).ShowPhonetic();                  // Mostrar fonetico
});
```

#### Metodos de Fila

| Metodo | Descripcion |
|--------|-------------|
| `Height(double)` | Altura de fila |
| `Hidden()` | Ocultar fila |
| `Collapsed()` | Colapsar |
| `OutlineLevel(byte)` | Nivel de agrupacion |
| `ThickBot()` | Borde inferior grueso |
| `ThickTop()` | Borde superior grueso |
| `ShowPhonetic()` | Mostrar fonetico |
| `Format(Action<OxCellFormartFactory>)` | Formato de fila |

---

### Formato

#### OxCellFormartFactory

Clase central para aplicar formato a celdas, filas, columnas y tablas.

#### Formato de Fuente

```csharp
c.Add("A1").Value("Texto").Format(f =>
{
    f.Font(font => font
        .Bold()                               // Negrita
        .Italic()                             // Cursiva
        .Underline(OxUnderlines.Single)       // Subrayado
        .Strike()                             // Tachado
        .Color(Color.Blue)                    // Color de texto
        .Size(14)                             // Tamaño
        .FontName("Arial")                    // Nombre de fuente
        .FontScheme(OxFontSchemes.Minor)      // Esquema
        .Shadow()                             // Sombra
        .Outline()                            // Contorno
        .Condense()                           // Condensado
        .Extend()                             // Extendido
        .VerticalAlignments(OxVerticalAlignments.Superscript)); // Superindice
});
```

#### Relleno Solido

```csharp
c.Add("A1").Value("Fondo").Format(f =>
{
    f.FillPattern(Color.Yellow, OxPatterns.Solid);
});
```

#### Patrones de Relleno

```csharp
f.FillPattern(Color.Blue, OxPatterns.Gray125);
f.FillPattern(Color.Green, OxPatterns.DarkHorizontal);
f.FillPattern(Color.Red, OxPatterns.LightGrid);
```

**Patrones disponibles:**
- `None`, `Solid`, `MediumGray`, `DarkGray`, `LightGray`
- `DarkHorizontal`, `DarkVertical`, `DarkDown`, `DarkUp`, `DarkGrid`, `DarkTrellis`
- `LightHorizontal`, `LightVertical`, `LightDown`, `LightUp`, `LightGrid`, `LightTrellis`
- `Gray125`, `Gray0625`

#### Relleno Degradado

```csharp
c.Add("A1").Value("Degradado").Format(f =>
{
    f.FillGradient(90, stops =>
    {
        stops.Add(Color.Blue, 0);    // Color inicial (posicion 0)
        stops.Add(Color.White, 0.5); // Color medio (posicion 0.5)
        stops.Add(Color.Yellow, 1);  // Color final (posicion 1)
    });
});
```

#### Bordes

```csharp
c.Add("D3").Value("Con Bordes").Format(f =>
{
    f.Borders(b =>
    {
        b.Top(Color.Black, OxBorderStyles.Thin);
        b.Bottom(Color.Black, OxBorderStyles.Medium);
        b.Left(Color.Gray, OxBorderStyles.Dotted);
        b.Right(Color.Brown, OxBorderStyles.Thick);
    });
});

// Bordes diagonales
c.Add("D6").Value("Diagonal").Format(f =>
{
    f.Borders()
        .Diagonal(Color.Red, OxBorderStyles.Double)
        .DiagonalDown()
        .DiagonalUp()
        .Outline();
});
```

**Estilos de borde disponibles:**
- `None`, `Thin`, `Medium`, `Thick`
- `Dashed`, `Dotted`, `Double`, `Hair`
- `MediumDashed`, `DashDot`, `MediumDashDot`
- `DashDotDot`, `MediumDashDotDot`, `SlantDashDot`

#### Formato Numerico

```csharp
c.Add("A1").Value(DateTime.Now).Format(f =>
{
    f.NumberFormat("dd/MM/yyyy hh:mm");
});

c.Add("A2").Value(1234.56).Format(f =>
{
    f.NumberFormat("#,##0.00");
});

c.Add("A3").Value(0.75).Format(f =>
{
    f.NumberFormat("0.00%");
});
```

#### Alineacion

```csharp
c.Add("A1").Value("Centrado").Margen(2, 2).Format(f =>
{
    f.Alignment(a => a
        .Horizontal(OxTextHorizontalAlignments.Center)
        .Vertical(OxTextVerticalAlignments.Center)
        .Rotation(45)        // Rotacion en grados
        .ShrinkToFit()       // Reducir para ajustar
        .JustifyLastLine()); // Justificar ultima linea
});
```

**Alineaciones horizontales:**
- `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`

**Alineaciones verticales:**
- `Top`, `Center`, `Bottom`, `Justify`, `Distributed`

#### Formato Combinado

```csharp
c.Add("A1").Value("Titulo").Margen(2, 0).Format(f =>
{
    f.Font(font => font.Bold().Size(14).Color(Color.White))
     .FillPattern(Color.DarkBlue, OxPatterns.Solid)
     .Alignment(a => a.Horizontal(OxTextHorizontalAlignments.Center))
     .Borders(b => b.Bottom(Color.Black, OxBorderStyles.Medium));
});
```

---

### Tablas

Las tablas permiten mostrar colecciones de datos con formato automatico, filtros y totales.

#### Tabla Basica con Tipo Generico

```csharp
public class Producto
{
    public int Id { get; set; }
    public string Nombre { get; set; }
    public decimal Precio { get; set; }
    public DateTime Fecha { get; set; }
}

var productos = new List<Producto>
{
    new Producto { Id = 1, Nombre = "Laptop", Precio = 999.99m, Fecha = DateTime.Now },
    new Producto { Id = 2, Nombre = "Mouse", Precio = 29.99m, Fecha = DateTime.Now },
    new Producto { Id = 3, Nombre = "Teclado", Precio = 79.99m, Fecha = DateTime.Now }
};

sheet.AddTable(productos, "A", 1, t =>
{
    t.Name("TblProductos");
    t.Columns(c =>
    {
        c.Add(p => p.Id).Header("ID");
        c.Add(p => p.Nombre).Header("Producto");
        c.Add(p => p.Precio).Header("Precio").Format(f => f.NumberFormat("$#,##0.00"));
        c.Add(p => p.Fecha).Header("Fecha");
    });
});
```

#### Posicionamiento de Tabla

```csharp
// Por columna (string) y fila
sheet.AddTable(datos, "A", 1, t => { ... });

// Por columna (numero) y fila
sheet.AddTable(datos, 1, 5, t => { ... });

// Por referencia de celda
sheet.AddTable(datos, "B5", t => { ... });
```

#### Auto-Generar Columnas

```csharp
sheet.AddTable(productos, 1, 1, t =>
{
    t.Name("Tabla1").AutoGenerateColumns();
});
```

#### Columnas Personalizadas

```csharp
t.Columns(c =>
{
    // Propiedad simple
    c.Add(p => p.Id).Header("ID");

    // Propiedad anidada
    c.Add(p => p.Direccion.Ciudad).Header("Ciudad");

    // Valor por defecto
    c.Add("Estado").DefaultVal("Activo");

    // Formula
    c.Add(p => p.Cantidad)
        .Header("Total")
        .IsFormula()
        .CellType(OxCellTypeValues.Number);

    // Valor calculado con Template
    c.Add("Descripcion")
        .Type(typeof(string))
        .TemplateValue(temp => $"{temp.Row.GetPropertyVal("Nombre")} - {temp.TableRowIndex}");

    // Tamaño de columna
    c.Add(p => p.Nombre).Size(30);

    // Orden de columna
    c.Add(p => p.Nombre).Order(1);
});
```

#### Formato de Columnas

```csharp
t.Columns(c =>
{
    c.Add(p => p.Id)
        .Header("ID")
        .HeaderFormat(f => f.FillPattern(Color.DarkOrange, OxPatterns.Solid))
        .Format(f => f.Font(font => font.Bold().Color(Color.Blue)));

    // Formato condicional con Template
    c.Add(p => p.Precio)
        .TemplateFormat(temp =>
        {
            var precio = Convert.ToDecimal(temp.CellValue);
            if (precio > 100)
                return temp.Format.FillPattern(Color.Green, OxPatterns.Solid);
            return temp.Format;
        });
});
```

#### Hipervinculos en Columnas

```csharp
c.Add(p => p.Id)
    .HyperlinkTemplate(e => new OxHyperlinkEntity(
        new Uri("http://ejemplo.com/producto/" + e.CellValue),
        "Ver producto"));
```

#### Filtros

```csharp
// Filtro simple (valores especificos)
c.Add(p => p.Estado).Filter(new[] { "Activo", "Pendiente" });

// Filtro de operador
c.Add(p => p.Precio).Filter(100, OxFilterOperators.GreaterThan);

// Filtros de texto
c.Add(p => p.Nombre).Filter("Pro", OxFilterOperators.StartWith);
c.Add(p => p.Nombre).Filter("max", OxFilterOperators.EndWith);
c.Add(p => p.Nombre).Filter("lap", OxFilterOperators.Contrains);
c.Add(p => p.Nombre).Filter("old", OxFilterOperators.NotContrains);

// Filtro con condicion compuesta
c.Add(p => p.Nombre).Filter(
    "a", OxFilterOperators.Contrains,
    OxCustomFilterCondition.And,
    "Pro", OxFilterOperators.StartWith);
```

**Operadores de filtro:**
- `Equal`, `NotEqual`
- `LessThan`, `LessThanOrEqual`
- `GreaterThan`, `GreaterThanOrEqual`
- `StartWith`, `EndWith`
- `Contrains`, `NotContrains`

#### Fila de Totales

```csharp
t.Columns(c =>
{
    c.Add(p => p.Categoria).TotalRow("Total:");

    c.Add(p => p.Cantidad)
        .TotalRow(TotalsRowFormulas.Sum, includeHidden: false)
        .TotalRowFormat(f => f.Font(ff => ff.Bold()).NumberFormat("#,##0"));

    c.Add(p => p.Precio)
        .TotalRow(TotalsRowFormulas.Average, includeHidden: true);

    // Formula personalizada
    c.Add(p => p.Total)
        .TotalRow("=SUBTOTAL(109,[Total])", includeHidden: false);
});

t.TotalsRowShown();
```

**Formulas de total disponibles:**
- `None`, `Sum`, `Minimum`, `Maximum`
- `Average`, `Count`, `CountNumbers`
- `StandardDeviation`, `Variance`, `Custom`

#### Definicion de Filas

```csharp
// Altura y formato para todas las filas
t.RowDefinition(r => r
    .Height(25)
    .Format(f => f.FillPattern(Color.LightGray, OxPatterns.Solid)));

// Template para filas alternadas
t.RowDefinitionTemplate(temp =>
{
    if (temp.TableRowIndex % 2 == 0)
        temp.RowDefinition.Format().FillPattern(Color.AliceBlue, OxPatterns.Solid);
    return temp.RowDefinition;
});
```

#### Estilo de Tabla

```csharp
t.TableStyle(s => s
    .HideRowStripes()
    .ShowColumnStripes()
    .ShowFirstColumn()
    .ShowLastColumn());

// Ocultar filtro automatico
t.HideAutoFilter();
```

#### Tablas Grandes: AddTableStream

`AddTable` recibe una coleccion y la guarda completa. Cuando ni eso cabe en memoria —los datos vienen de una
consulta por bloques, por ejemplo— use `AddTableStream`: las filas se leen de la secuencia **mientras se escribe el
archivo**, asi que la memoria no crece con el numero de filas.

```csharp
// El total se necesita antes que las filas: con el se declara el rango de la hoja.
var total = await db.Capturas.CountAsync(filtro);

sheet.AddTableStream(LeerCapturas(filtro), total, 1, 6, t =>
{
    t.TableType(OxTableType.Skeleton);
    t.Columns(c =>
    {
        c.Add("Fecha").Header("Fecha").Type(typeof(DateTime));
        c.Add("Valor").Header("Valor").Type(typeof(double));
    });
});
```

Reglas de la secuencia:

- `rowsCount` es el **total exacto, o un maximo**. Si la secuencia entrega mas filas se lanza una excepcion (el
  rango ya se escribio); si entrega menos, el rango declarado solo queda un poco mas grande. Si los datos pueden
  crecer mientras se escribe el archivo, cuente primero y corte con `.Take(total)`.
- Se recorre **una sola vez y en orden**. Para usar los mismos datos en otra hoja, entregue una secuencia nueva.
- No funciona `AutoGenerateColumns` (las columnas se declaran) y `MasterData` de las plantillas ya no trae la lista
  completa de filas.
- Con `TotalsRowShown()` el conteo tiene que ser exacto: la fila de totales se coloca segun ese numero.

Para tablas normales siga usando `AddTable`: es mas simple y desde 1.0.6 tambien escribe sus filas de una en una.

#### Tipo de Celda por Fila

`CellType` define el tipo de toda la columna. `TemplateCellType` lo decide fila por fila, que es lo que hace falta
cuando la misma columna trae numeros en unas filas y texto en otras (un valor no leido, una etiqueta):

```csharp
c.Add("Valor")
    .TemplateValue(t     => Lectura(t) is { Numero: { } n } ? n : (object)"ERR")
    .TemplateCellType(t  => Lectura(t)?.Numero is not null
                                ? OxCellTypeValues.Number
                                : OxCellTypeValues.String);
```

Sin esto hay que elegir entre mandar toda la columna como texto —y perder filtros y graficas en Excel— o generar
celdas numericas con texto adentro, que Excel repara al abrir.

#### Tipo de Tabla

```csharp
t.TableType(OxTableType.Excel);     // Tabla Excel completa
t.TableType(OxTableType.Skeleton);  // Solo filas y celdas
```

`Skeleton` escribe los datos sin crear el objeto Tabla de Excel: sin autofiltro, sin bandas de color y **sin la
exigencia de encabezados unicos**. Una tabla de Excel pide nombres de columna distintos, asi que si dos columnas
pueden llamarse igual —dos parametros con el mismo titulo en grupos distintos, por ejemplo— Excel repara el
archivo. Es tambien lo que conviene cuando la hoja ya trae su propio encabezado.

#### Tabla desde DataTable

```csharp
var dataTable = new DataTable();
dataTable.Columns.Add("Id", typeof(int));
dataTable.Columns.Add("Nombre", typeof(string));
dataTable.Columns.Add("Fecha", typeof(DateTime));

dataTable.Rows.Add(1, "Producto A", DateTime.Now);
dataTable.Rows.Add(2, "Producto B", DateTime.Now);

// Convertir a lista dinamica
var dynamicList = dataTable.ToDynamicList();

sheet.AddTable(dynamicList, 1, 1, t =>
{
    t.Name("TblDynamic");
    t.Columns(c =>
    {
        c.Add("Id").CellType(OxCellTypeValues.Number);
        c.Add("Nombre");
        c.Add("Fecha").Type(typeof(DateTime));
    });
});
```

#### Tabla con Objetos Anonimos

```csharp
var datos = productos.Select(p => new { p.Id, p.Nombre, Total = p.Precio * 1.16m });

sheet.AddTable(datos, 1, 1, t =>
{
    t.Columns(c =>
    {
        c.Add("Id");
        c.Add("Nombre");
        c.Add("Total").Type(typeof(decimal));
    });
});
```

---

### Imagenes

#### Agregar Imagenes

```csharp
sheet.AddImage(i =>
{
    // Desde archivo
    i.Add(new OxRangeEntity("B5:E24"))
        .Name("Logo")
        .Url(@"C:\imagenes\logo.png")
        .Rectangle(new RectangleF(0, 0, 100, 50));

    // Desde bytes
    var bytes = File.ReadAllBytes(@"C:\imagenes\foto.png");
    i.Add(new OxRangeEntity("F5:I15"))
        .Name("Foto")
        .ArrayBytes(bytes)
        .Rectangle(new RectangleF(0, 0, 100, 100));
});
```

#### OxRangeEntity

Define el rango donde se ubicara la imagen:

```csharp
// Desde string
var range = new OxRangeEntity("B5:E24");

// Propiedades
range.StartCol;  // Columna inicial
range.StartRow;  // Fila inicial
range.EndCol;    // Columna final
range.EndRow;    // Fila final
```

#### Metodos de Imagen

| Metodo | Descripcion |
|--------|-------------|
| `Url(string)` | Ruta del archivo de imagen |
| `ArrayBytes(byte[])` | Imagen desde array de bytes |
| `Name(string)` | Nombre de la imagen |
| `Rectangle(RectangleF)` | Dimensiones y posicion |

---

## Ejemplos Completos

### Reporte de Ventas

```csharp
public class Venta
{
    public int Id { get; set; }
    public string Producto { get; set; }
    public decimal Precio { get; set; }
    public int Cantidad { get; set; }
    public DateTime Fecha { get; set; }
}

var ventas = ObtenerVentas(); // Tu metodo de datos

var doc = new OxExcelDocument()
    .Calculation(c => c.CalculationMode(OxCalculateModes.Auto))
    .PackageProperties(p => p
        .Title("Reporte de Ventas")
        .Creator("Sistema")
        .Company("Mi Empresa"));

doc.AddSheet("Ventas")
    .SheetView(v => v.PaneFrozen("A2"))
    .Cell(c =>
    {
        c.Add("A1").Value("REPORTE DE VENTAS").Margen(5, 0)
            .Format(f => f
                .Font(font => font.Bold().Size(16).Color(Color.White))
                .FillPattern(Color.DarkBlue, OxPatterns.Solid)
                .Alignment(a => a.Horizontal(OxTextHorizontalAlignments.Center)));
    })
    .AddTable(ventas, "A", 3, t =>
    {
        t.Name("TblVentas");
        t.Columns(c =>
        {
            c.Add(p => p.Id).Header("ID").Size(10);
            c.Add(p => p.Producto).Header("Producto").Size(30);
            c.Add(p => p.Precio).Header("Precio")
                .Format(f => f.NumberFormat("$#,##0.00"))
                .TotalRow(TotalsRowFormulas.Sum, false);
            c.Add(p => p.Cantidad).Header("Cantidad")
                .TotalRow(TotalsRowFormulas.Sum, false);
            c.Add("Subtotal")
                .IsFormula()
                .DefaultFormulaVal("=[@Precio]*[@Cantidad]")
                .Format(f => f.NumberFormat("$#,##0.00"))
                .TotalRow(TotalsRowFormulas.Sum, false);
            c.Add(p => p.Fecha).Header("Fecha").Size(20);
        });
        t.TotalsRowShown();
        t.TableStyle(s => s.ShowFirstColumn());
    });

doc.Save("ReporteVentas.xlsx");
```

### Dashboard con Multiples Hojas

```csharp
var doc = new OxExcelDocument();

// Hoja de Resumen
doc.AddSheet("Resumen")
    .SheetView(v => v.HideGridLines())
    .Cell(c =>
    {
        c.Add("B2").Value("DASHBOARD EJECUTIVO").Margen(4, 0)
            .Format(f => f
                .Font(font => font.Bold().Size(18))
                .Alignment(a => a.Horizontal(OxTextHorizontalAlignments.Center)));

        // KPIs
        c.Add("B4").Value("Total Ventas:").Format(f => f.Font(font => font.Bold()));
        c.Add("C4").Value(125000).Format(f => f.NumberFormat("$#,##0"));

        c.Add("B5").Value("Promedio Mensual:").Format(f => f.Font(font => font.Bold()));
        c.Add("C5").Value(10416.67).Format(f => f.NumberFormat("$#,##0.00"));
    })
    .AddImage(i =>
    {
        i.Add(new OxRangeEntity("E2:H10"))
            .Url(@"logo.png");
    });

// Hoja de Datos Detallados
doc.AddSheet("Datos")
    .AddTable(datos, 1, 1, t =>
    {
        t.Name("TblDatos").AutoGenerateColumns();
    });

// Hoja oculta con configuracion
doc.AddSheet("Config")
    .SheetVisibility(OxSheetVisibilities.Hidden);

doc.Save("Dashboard.xlsx");
```

### Formato Condicional en Tablas

```csharp
sheet.AddTable(empleados, 1, 1, t =>
{
    t.Columns(c =>
    {
        c.Add(p => p.Nombre).Header("Empleado");

        c.Add(p => p.Ventas)
            .Header("Ventas")
            .Format(f => f.NumberFormat("$#,##0"))
            .TemplateFormat(temp =>
            {
                var ventas = Convert.ToDecimal(temp.CellValue);
                if (ventas >= 10000)
                    return temp.Format
                        .FillPattern(Color.LightGreen, OxPatterns.Solid)
                        .Font(f => f.Bold().Color(Color.DarkGreen));
                else if (ventas >= 5000)
                    return temp.Format
                        .FillPattern(Color.LightYellow, OxPatterns.Solid);
                else
                    return temp.Format
                        .FillPattern(Color.LightPink, OxPatterns.Solid)
                        .Font(f => f.Color(Color.DarkRed));
            });

        c.Add(p => p.Comision)
            .Header("Comision")
            .IsFormula()
            .DefaultFormulaVal("=[@Ventas]*0.05");
    });

    t.RowDefinitionTemplate(temp =>
    {
        // Filas alternadas
        if (temp.TableRowIndex % 2 == 0)
            temp.RowDefinition.Format().FillPattern(Color.WhiteSmoke, OxPatterns.Solid);
        return temp.RowDefinition;
    });
});
```

---

## Enumeraciones

### OxDocumentTypes
```csharp
Workbook              // Libro de trabajo (.xlsx)
Template              // Plantilla (.xltx)
MacroEnabledWorkbook  // Con macros (.xlsm)
MacroEnabledTemplate  // Plantilla con macros (.xltm)
AddIn                 // Complemento (.xlam)
```

### OxCalculateModes
```csharp
Auto          // Calculo automatico
Manual        // Calculo manual
AutoNoTable   // Automatico excepto tablas
```

### OxSheetVisibilities
```csharp
Visible       // Visible
Hidden        // Oculta
VeryHidden    // Muy oculta (solo por codigo)
```

### OxSheetViews
```csharp
Normal            // Vista normal
PageLayout        // Diseño de pagina
PageBreakPreview  // Vista previa de saltos
```

### OxCellTypeValues
```csharp
Default       // Automatico
String        // Texto
Number        // Numero
Date          // Fecha (Office 2010+)
Error         // Error
SharedString  // String compartido
InlineString  // String en linea
```

### OxTextHorizontalAlignments
```csharp
General           // General
Left              // Izquierda
Center            // Centro
Right             // Derecha
Fill              // Rellenar
Justify           // Justificar
CenterContinuous  // Centro continuo
Distributed       // Distribuido
```

### OxTextVerticalAlignments
```csharp
Top           // Superior
Center        // Centro
Bottom        // Inferior
Justify       // Justificar
Distributed   // Distribuido
```

### OxBorderStyles
```csharp
None, Thin, Medium, Thick, Double
Dashed, Dotted, Hair
MediumDashed, DashDot, MediumDashDot
DashDotDot, MediumDashDotDot, SlantDashDot
```

### OxPatterns
```csharp
None, Solid, MediumGray, DarkGray, LightGray
DarkHorizontal, DarkVertical, DarkDown, DarkUp
DarkGrid, DarkTrellis
LightHorizontal, LightVertical, LightDown, LightUp
LightGrid, LightTrellis
Gray125, Gray0625
```

### OxFilterOperators
```csharp
Equal, NotEqual
LessThan, LessThanOrEqual
GreaterThan, GreaterThanOrEqual
StartWith, EndWith
Contrains, NotContrains
```

### TotalsRowFormulas
```csharp
None, Sum, Minimum, Maximum
Average, Count, CountNumbers
StandardDeviation, Variance, Custom
```

### OxPaperSizeDefault
```csharp
Letter, Legal
A2, A3, A4, A5
B4, B5
// ... mas tamaños
```

---

## Atributos

### OxColumnAttribute

Permite definir metadatos de columna directamente en las propiedades del modelo:

```csharp
public class Empleado
{
    [OxColumn(Header = "ID", Order = 1, Size = 10)]
    public int Id { get; set; }

    [OxColumn(Header = "Nombre Completo", Size = 30)]
    [DisplayName("Empleado")]
    public string Nombre { get; set; }

    [OxColumn(Header = "Salario",
              CellTypeValue = OxCellTypeValues.Number,
              DefaultValue = 0)]
    public decimal Salario { get; set; }

    [OxColumn(Header = "Fecha Ingreso", Size = 20)]
    [Display(Order = 5)]
    public DateTime FechaIngreso { get; set; }

    [OxColumn(Header = "Bonificacion",
              IsFormula = true,
              DefaultFormulaValue = "=[@Salario]*0.1")]
    public decimal Bonificacion { get; set; }
}
```

#### Propiedades del Atributo

| Propiedad | Tipo | Descripcion |
|-----------|------|-------------|
| `Header` | `string` | Texto del encabezado |
| `Size` | `uint` | Ancho de columna |
| `Order` | `int` | Orden de la columna |
| `CellTypeValue` | `OxCellTypeValues` | Tipo de celda |
| `DefaultValue` | `object` | Valor por defecto |
| `DefaultFormulaValue` | `string` | Formula por defecto |
| `IsFormula` | `bool` | Es una formula |
| `ShowPhonetic` | `bool` | Mostrar fonetico |
| `CellFormart` | `OxCellFormartFactory` | Formato de celda |
| `HeaderCellFormart` | `OxCellFormartFactory` | Formato de encabezado |

### Compatibilidad con DataAnnotations

La libreria tambien reconoce atributos estandar de .NET:

```csharp
[DisplayName("Mi Nombre")]     // System.ComponentModel
[Display(Name = "Titulo", Order = 1)]  // System.ComponentModel.DataAnnotations
```

---

## Limites de Excel

Una hoja admite **1,048,576 filas** y **16,384 columnas**. Pasarse produce un archivo que Excel rechaza al abrir,
asi que desde 1.0.6 la libreria lanza una excepcion clara al escribir, diciendo la hoja y la fila o columna que se
paso:

```
La hoja "Detalle" llego a la fila 1048577 y Excel solo admite 1048576 filas.
Acorte el rango de datos o repartalos en varias hojas.
```

Si sus datos pueden crecer hasta ahi, cuente antes de escribir y reparta en varias hojas o acote el rango.

---

## Consideraciones de Rendimiento

### Datos Masivos

Las filas de una tabla se escriben directo al archivo, una a la vez, y se descartan enseguida: la memoria no crece
con el numero de filas, solo con la coleccion que usted entrega.

```csharp
var datos = new List<MiClase>();
for (int i = 0; i < 100000; i++)
{
    datos.Add(new MiClase { /* ... */ });
}

sheet.AddTable(datos, 1, 1, t =>
{
    t.Name("TblMasiva").AutoGenerateColumns();
});
```

Referencia con 50 mil filas por 26 columnas, con formato por columna y color por celda:

| | 1.0.5 | 1.0.6 |
|---|---|---|
| Memoria asignada | 8,000 MB | 1,429 MB |
| Tiempo de `Save` | 9.2 s | 4.3 s |
| Working set pico | 473 MB | 259 MB |

Para exprimirlo:

- **Comparta las instancias de formato**: una por estilo, guardada en un diccionario, en vez de crear una
  `OxCellFormartFactory` por celda. El archivo deduplica estilos iguales, pero crearlos cuesta.
- **No ponga un `Format(...)` base en las columnas cuyo `TemplateFormat` ya devuelve el formato completo**. La
  libreria combina lo que devuelve la plantilla con el formato de la columna y esa combinacion *modifica* la
  instancia devuelta: si la tiene en cache, le queda contaminada con lo que traia la columna. Sin formato base la
  plantilla manda tal cual y encima se ahorra una copia por celda.
- **Si ni la lista cabe en memoria**, use [`AddTableStream`](#tablas-grandes-addtablestream).

Si algo sale distinto a como salia antes, `OxExcelDocument.MaterializeTableRows = true` (estatico, antes de crear
el documento) vuelve al camino anterior —armar la hoja completa en memoria— para comparar. Es un respaldo, no una
opcion para produccion.

### Liberacion de Recursos

Siempre use `using` o llame `Dispose()` para liberar recursos:

```csharp
using (var doc = new OxExcelDocument())
{
    // ... crear documento
    doc.Save("archivo.xlsx");
} // Dispose automatico

// O manualmente
var doc = new OxExcelDocument();
try
{
    // ... crear documento
    doc.Save("archivo.xlsx");
}
finally
{
    doc.Dispose();
}
```

---

## Licencia

MIT License - Copyright (c) Miguel Fernando Garcias Salazar

## Repositorio

[https://github.com/MgSoftDev/MgSoftDev.OXExcel](https://github.com/MgSoftDev/MgSoftDev.OXExcel)
