# MudTools.OfficeInterop.PowerPoint

PowerPoint 操作模块，提供完整的 PowerPoint 演示文稿操作接口。

## 项目概述

MudTools.OfficeInterop.PowerPoint 是专门用于操作 Microsoft PowerPoint 应用程序的 .NET 封装库。该模块提供了完整的 PowerPoint 演示文稿操作接口，包括幻灯片、母版、动画等对象的管理，以及演示文稿的创建、编辑和格式化功能。

通过使用本模块，开发者可以避免直接处理复杂的 PowerPoint COM 交互，从而更专注于业务逻辑的实现。

## 主要功能

- PowerPoint 演示文稿操作接口
- 幻灯片、母版、动画等对象的管理
- 演示文稿的创建、编辑和格式化功能

## 支持的框架

- .NET Framework 4.6.2
- .NET Framework 4.7
- .NET Framework 4.8
- .NET Standard 2.1

## 安装

```xml
<PackageReference Include="MudTools.OfficeInterop.PowerPoint" Version="1.0.1" />
```

## 核心组件

### PowerPointFactory

[PowerPointFactory](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.PowerPoint/PowerPointFactory.cs#L15-L74) 是用于创建和操作 PowerPoint 应用程序实例的工厂类，提供以下方法：

- `Connection` - 通过现有 COM 对象连接到已运行的 PowerPoint 应用程序实例
- `BlankWorkbook` - 创建新的空白 PowerPoint 演示文稿
- `Open` - 打开现有的 PowerPoint 演示文稿文件

## 使用示例

### 创建演示文稿

```csharp
// 创建 PowerPoint 应用程序实例
using var app = PowerPointFactory.CreateApplication();
app.Visible = true;

// 创建新演示文稿
var presentation = app.Presentations.Add();

// 添加幻灯片
var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);

// 设置标题
slide.Shapes.Title.TextFrame.TextRange.Text = "欢迎使用 PowerPoint";

// 添加内容
slide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "这是演示文稿的内容部分";

// 保存演示文稿
presentation.SaveAs(@"C:\presentations\example.pptx");
```

### 操作现有演示文稿

```csharp
// 打开现有演示文稿
using var app = PowerPointFactory.Open(@"C:\presentations\ExistingPresentation.pptx");
var presentation = app.ActivePresentation;

// 遍历所有幻灯片
foreach (var slide in presentation.Slides)
{
    Console.WriteLine($"幻灯片 {slide.SlideIndex}: {slide.Name}");
    
    // 修改幻灯片内容
    if (slide.Shapes.HasTitle == MsoTriState.msoTrue)
    {
        slide.Shapes.Title.TextFrame.TextRange.Text += " - 已更新";
    }
}

// 添加新幻灯片
var newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, 
                                      PowerPoint.PpSlideLayout.ppLayoutText);

newSlide.Shapes.Title.TextFrame.TextRange.Text = "新幻灯片";
newSlide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "这是新增的幻灯片内容";

presentation.Save();
app.Quit();
```

### PowerPoint 格式化和动画

```csharp
using var app = PowerPointFactory.BlankWorkbook();
var presentation = app.ActivePresentation;

// 添加幻灯片
var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

// 添加形状
var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 100, 100, 200, 100);
shape.TextFrame.TextRange.Text = "示例形状";

// 设置形状格式
shape.Fill.ForeColor.RGB = 0x00FF00; // 绿色填充
shape.Line.ForeColor.RGB = 0xFF0000; // 红色边框

// 添加动画
var animation = shape.AnimationSettings;
animation.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFade;
animation.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnClick;

presentation.SaveAs(@"C:\presentations\AnimatedPresentation.pptx");
app.Quit();
```

## 许可证

本项目采用双重许可证模式：

- [MIT 许可证](../../LICENSE-MIT)
- [Apache 许可证 2.0](../../LICENSE-APACHE)

## 免责声明

本项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。

不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任。