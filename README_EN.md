## Special Notice

MudTools.OfficeInterop project has joined the [dotNET China](https://gitee.com/dotnetchina) organization.

![dotnetchina](https://gitee.com/dotnetchina/home/raw/master/assets/dotnetchina-raw.png "dotNET China LOGO")

# MudTools.OfficeInterop

A .NET wrapper library for Microsoft Office applications, designed to simplify the use of Office COM components.

This library provides developers with a modern, object-oriented API for manipulating Microsoft Office applications (Excel, Word, PowerPoint). By using this library, developers can avoid dealing with complex COM interactions directly and focus more on implementing business logic.

## Project Overview

MudTools.OfficeInterop is a set of .NET wrapper libraries for Microsoft Office applications (including Excel, Word, PowerPoint, and VBE). This project reduces the complexity of directly using Office COM components by providing concise, unified API interfaces, making it easier for developers to integrate and manipulate Office documents in .NET applications.

### Module Overview

| Module | Current Version | Download | Open Source License |
|---|---|---|---|
| [![OfficeInterop-Core](https://img.shields.io/badge/Office.Interop.Core-mudtools-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/nuget/v/MudTools.OfficeInterop.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop/) | [![Nuget](https://img.shields.io/nuget/dt/MudTools.OfficeInterop.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE) |
| [![OfficeInterop-Excel](https://img.shields.io/badge/Office.Interop.Excel-mudtools-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/nuget/v/MudTools.OfficeInterop.Excel.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.Excel/) | [![Nuget](https://img.shields.io/nuget/dt/MudTools.OfficeInterop.Excel.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.Excel/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE) |
| [![OfficeInterop-Word](https://img.shields.io/badge/Office.Interop.Word-mudtools-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/nuget/v/MudTools.OfficeInterop.Word.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.Word/) | [![Nuget](https://img.shields.io/nuget/dt/MudTools.OfficeInterop.Word.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.Word/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE) |
| [![OfficeInterop-PowerPoint](https://img.shields.io/badge/Office.Interop.PowerPoint-mudtools-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/nuget/v/MudTools.OfficeInterop.PowerPoint.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.PowerPoint/) | [![Nuget](https://img.shields.io/nuget/dt/MudTools.OfficeInterop.PowerPoint.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.PowerPoint/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE) |
| [![OfficeInterop-Vbe](https://img.shields.io/badge/Office.Interop.Vbe-mudtools-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/nuget/v/MudTools.OfficeInterop.Vbe.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.Vbe/) | [![Nuget](https://img.shields.io/nuget/dt/MudTools.OfficeInterop.Vbe.svg)](https://www.nuget.org/packages/MudTools.OfficeInterop.Vbe/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE) |

### Project Goals

The main objectives of this project are:

1. **Simplify Office Automation**: Provide a cleaner, more user-friendly .NET API by encapsulating complex COM interfaces
2. **Improve Development Efficiency**: Reduce the time and effort developers need for Office automation
3. **Enhance Code Maintainability**: Make code easier to understand and maintain through object-oriented design and clear interfaces
4. **Provide Complete Feature Coverage**: Support common features of Office applications, including document creation, editing, formatting, etc.
5. **Ensure Type Safety**: Leverage .NET's type system to reduce runtime errors

### Use Cases

MudTools.OfficeInterop is suitable for the following scenarios:

- Enterprise report generation and data processing
- Batch document processing and formatting
- Office plugin development
- Office application automation
- Data import/export functionality
- Document template processing

### Design Philosophy

This project follows these design principles:

1. **Simplicity**: Provide intuitive, easy-to-use APIs to reduce learning costs
2. **Consistency**: Maintain similar interface designs across different Office applications
3. **Extensibility**: Allow developers to access underlying COM objects when needed
4. **Resource Management**: Ensure proper release of COM resources through IDisposable interface implementation
5. **Compatibility**: Support multiple .NET Framework versions and different Office versions

## Feature Modules

### Core Module (MudTools.OfficeInterop)
- Provides basic interfaces and common functionality for Office applications
- Encapsulates common operations of Office core components
- Provides basic support for other Office application modules
- Provides Office UI related component encapsulation, including Ribbon and Custom Task Pane (CTP)

### Excel Module (MudTools.OfficeInterop.Excel)
- Complete Excel application manipulation interfaces
- Convenient operations for workbooks, worksheets, cells, and other objects
- Advanced functionality encapsulation for charts, pivot tables, etc.
- Formatting settings, style management, and other features

### Word Module (MudTools.OfficeInterop.Word)
- Word document manipulation interfaces
- Document content, style, formatting, and other management features
- Table, image, and other element operation encapsulation

### PowerPoint Module (MudTools.OfficeInterop.PowerPoint)
- PowerPoint presentation manipulation interfaces
- Management of slides, masters, animations, and other objects
- Presentation creation, editing, and formatting functionality

### VBE Module (MudTools.OfficeInterop.Vbe)
- Visual Basic Editor related functionality encapsulation
- Macro, code module, project, and other object operation interfaces

## Supported Frameworks

- .NET Framework 4.6.2
- .NET Framework 4.7
- .NET Framework 4.8
- .NET Standard 2.1
- .NET 6.0-windows
- .NET 7.0-windows
- .NET 8.0-windows
- .NET 9.0-windows

## Installation

This project depends on Microsoft Office COM components. Before using, ensure that the appropriate version of Microsoft Office is installed on the system.

```xml
<PackageReference Include="MudTools.OfficeInterop" Version="2.0.1" />
<PackageReference Include="MudTools.OfficeInterop.Excel" Version="2.0.1" />
<PackageReference Include="MudTools.OfficeInterop.Word" Version="2.0.1" />
<PackageReference Include="MudTools.OfficeInterop.PowerPoint" Version="2.0.1" />
<PackageReference Include="MudTools.OfficeInterop.Vbe" Version="2.0.1" />
```

## Factory Classes Usage

This project provides multiple factory classes for creating and manipulating Office application objects:

- [OfficeUIFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop/OfficeUIFactory.cs#L16-L51) - Used for creating Office UI related components, such as Ribbon and Custom Task Pane (CTP)
- [ExcelFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Excel/ExcelFactory.cs#L22-L152) - Used for creating and manipulating Excel application instances
- [WordFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/WordFactory.cs#L15-L97) - Used for creating and manipulating Word application instances
- [PowerPointFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.PowerPoint/PowerPointFactory.cs#L15-L74) - Used for creating and manipulating PowerPoint application instances

All factory classes provide multiple methods for creating application instances:
- `Connection` - Connect to a running application instance through an existing COM object
- `BlankWorkbook` - Create a new blank document/workbook/presentation
- `CreateFrom` - Create a new document/workbook/presentation based on a template (Word and Excel support)
- `Open` - Open an existing document/workbook/presentation
- `CreateInstance` - (ExcelFactory only) Create a specific version of the application instance through ProgID

## Usage Examples

### Excel Operation Examples

#### Basic Operations

```csharp
// Create Excel application instance
using var app = ExcelFactory.BlankWorkbook();
app.Visible = true;

// Get active worksheet
var worksheet = app.ActiveSheetWrap;

// Manipulate cells
worksheet.Cells[1, 1].Value = "Hello";
worksheet.Cells[1, 2].Value = "World";

// Save workbook
app.ActiveWorkbook.SaveAs(@"C:\temp\example.xlsx");
app.Quit();
```

#### Creating Excel Workbook from Template

```csharp
// Create workbook based on template
using var app = ExcelFactory.CreateFrom(@"C:\templates\ReportTemplate.xltx");
var worksheet = app.ActiveSheetWrap;

// Fill data
worksheet.Cells[1, 1].Value = "Sales Report";
worksheet.Cells[2, 1].Value = DateTime.Now.ToString("yyyy-MM-dd");

// Save and close
app.ActiveWorkbook.SaveAs(@"C:\reports\SalesReport.xlsx");
app.Quit();
```

#### Reading Excel Data

```csharp
// Open existing workbook
using var app = ExcelFactory.Open(@"C:\data\SalesData.xlsx");
var worksheet = app.Worksheets[1];

// Read data range
var dataRange = worksheet.Range("A1:D100");
var rowCount = dataRange.Rows.Count;
var columnCount = dataRange.Columns.Count;

// Process data
for (int row = 1; row <= rowCount; row++)
{
    for (int col = 1; col <= columnCount; col++)
    {
        var cellValue = dataRange.Cells[row, col].Value?.ToString();
        Console.WriteLine($"Row {row}, Column {col}: {cellValue}");
    }
}

app.Quit();
```

#### Excel Chart Operations

```csharp
using var app = ExcelFactory.BlankWorkbook();
var worksheet = app.ActiveSheetWrap;

// Add sample data
worksheet.Cells[1, 1].Value = "Month";
worksheet.Cells[1, 2].Value = "Sales";
worksheet.Cells[2, 1].Value = "January";
worksheet.Cells[2, 2].Value = 10000;
worksheet.Cells[3, 1].Value = "February";
worksheet.Cells[3, 2].Value = 15000;
worksheet.Cells[4, 1].Value = "March";
worksheet.Cells[4, 2].Value = 12000;

// Create chart
var chartObjects = worksheet.ChartObjects();
var chartObject = chartObjects.Add(100, 50, 300, 200);
var chart = chartObject.Chart;

// Set chart data source
chart.SetSourceData(worksheet.Range("A1:B4"));
chart.ChartType = XlChartType.xlColumnClustered;

app.ActiveWorkbook.SaveAs(@"C:\charts\SalesChart.xlsx");
app.Quit();
```

### Word Operation Examples

#### Basic Operations

```csharp
// Create Word application instance
using var app = WordFactory.BlankDocument();
app.Visible = true;

// Get active document
var document = app.ActiveDocument;

// Add content
var range = document.Range();
range.Text = "Hello World!";

// Save document
document.SaveAs2(@"C:\temp\example.docx");
app.Quit();
```

#### Creating Word Document from Template

```csharp
// Create document based on template
using var app = WordFactory.CreateFrom(@"C:\templates\ReportTemplate.dotx");
var document = app.ActiveDocument;

// Replace placeholders in template
document.FindAndReplace("{REPORT_TITLE}", "Quarterly Sales Report");

// Add table
var tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
var table = document.Tables.Add(tableRange, 3, 3);
table.Cell(1, 1).Range.Text = "Product";
table.Cell(1, 2).Range.Text = "Quantity";
table.Cell(1, 3).Range.Text = "Revenue";

app.Quit();
```

#### Word Document Formatting

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// Add title
var titleRange = document.Range(0, 0);
titleRange.Text = "Document Title\n";
titleRange.Font.Bold = 1;
titleRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

// Add paragraph
var paraRange = document.Range(document.Content.End - 1, document.Content.End - 1);
paraRange.Text = "This is the content paragraph of the document, containing some sample text.\n";
paraRange.Font.Bold = 0;
paraRange.Font.Size = 12;

// Add list
var listRange = document.Range(document.Content.End - 1, document.Content.End - 1);
listRange.Text = "Item 1\nItem 2\nItem 3\n";
listRange.ListFormat.ApplyBulletDefault();

document.SaveAs2(@"C:\documents\FormattedDocument.docx");
app.Quit();
```

### PowerPoint Operation Examples

#### Creating Presentation

```csharp
// Create PowerPoint application instance
using var app = PowerPointFactory.BlankWorkbook();
app.Visible = true;

// Get presentation
var presentation = app.ActivePresentation;

// Add slide
var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);

// Set title
slide.Shapes.Title.TextFrame.TextRange.Text = "Welcome to PowerPoint";

// Add content
slide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "This is the content section of the presentation";

// Save presentation
presentation.SaveAs(@"C:\presentations\example.pptx");
app.Quit();
```

#### Manipulating Existing Presentation

```csharp
// Open existing presentation
using var app = PowerPointFactory.Open(@"C:\presentations\ExistingPresentation.pptx");
var presentation = app.ActivePresentation;

// Iterate through all slides
foreach (var slide in presentation.Slides)
{
    Console.WriteLine($"Slide {slide.SlideIndex}: {slide.Name}");
    
    // Modify slide content
    if (slide.Shapes.HasTitle == MsoTriState.msoTrue)
    {
        slide.Shapes.Title.TextFrame.TextRange.Text += " - Updated";
    }
}

// Add new slide
var newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, 
                                      PowerPoint.PpSlideLayout.ppLayoutText);

newSlide.Shapes.Title.TextFrame.TextRange.Text = "New Slide";
newSlide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "This is the content of the newly added slide";

presentation.Save();
app.Quit();
```

#### PowerPoint Formatting and Animation

```csharp
using var app = PowerPointFactory.BlankWorkbook();
var presentation = app.ActivePresentation;

// Add slide
var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

// Add shape
var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 100, 100, 200, 100);
shape.TextFrame.TextRange.Text = "Sample Shape";

// Set shape format
shape.Fill.ForeColor.RGB = 0x00FF00; // Green fill
shape.Line.ForeColor.RGB = 0xFF0000; // Red border

// Add animation
var animation = shape.AnimationSettings;
animation.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFade;
animation.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnClick;

presentation.SaveAs(@"C:\presentations\AnimatedPresentation.pptx");
app.Quit();
```

### Office UI Operation Examples

#### Using Custom Task Pane

```csharp
// Create custom task pane
var ctpFactory = OfficeUIFactory.CreateCTPFactory(officeCTPFactory);
var ctp = ctpFactory.CreateCTP("MyAddin.UserControl", "My Task Pane");

// Set task pane properties
ctp.Visible = true;
ctp.Width = 200;

// Show task pane
ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
```

#### Using Ribbon Controls

```csharp
// Handle ribbon control events
public void OnRibbonButtonClicked(IRibbonControl control)
{
    switch (control.Id)
    {
        case "buttonNewDocument":
            // Create new document
            using var app = ExcelFactory.BlankWorkbook();
            break;
        case "buttonOpenDocument":
            // Open document
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                using var app = ExcelFactory.Open(openFileDialog.FileName);
            }
            break;
    }
}
```

## License

This project adopts the MIT License model:

- [MIT License](LICENSE-MIT)

## Disclaimer

The copyright, trademarks, patents, and other related rights of this project are protected by relevant laws and regulations. The use of this project shall comply with the requirements of relevant laws and regulations and licenses.

This project must not be used for activities prohibited by laws and regulations, such as endangering national security, disrupting social order, or infringing on the legitimate rights and interests of others! We assume no responsibility for any legal disputes and liabilities arising from secondary development based on this project.
