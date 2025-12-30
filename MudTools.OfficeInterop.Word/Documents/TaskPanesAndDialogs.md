# 第13章：任务窗格和对话框

任务窗格和对话框是Word用户界面的重要组成部分，它们提供了一种与用户交互的方式。MudTools.OfficeInterop.Word库提供了对自定义任务窗格和对话框的支持，允许开发者创建更加丰富的用户体验。本章将详细介绍如何使用这些功能创建交互式文档处理工具。

## 自定义任务窗格

自定义任务窗格是停靠在Word窗口侧面的面板，可以包含各种控件和内容。

```csharp
// 注意：自定义任务窗格通常需要在VSTO插件环境中实现
// 以下代码展示了核心概念和使用方法

using var app = WordFactory.BlankDocument();

// 访问任务窗格集合
var taskPanes = app.TaskPanes;
```

获取Word应用程序的任务窗格集合。

```csharp
// 获取任务窗格数量
int paneCount = taskPanes.Count;
Console.WriteLine($"当前任务窗格数量: {paneCount}");

// 访问特定任务窗格
if (paneCount > 0)
{
    var taskPane = taskPanes[1];
    Console.WriteLine($"第一个任务窗格可见性: {taskPane.Visible}");
    Console.WriteLine($"任务窗格宽度: {taskPane.Width}");
}
```

获取任务窗格信息。

```csharp
// 在VSTO插件中创建自定义任务窗格的示例：
/*
// 创建用户控件
public partial class CustomTaskPaneControl : UserControl
{
    public CustomTaskPaneControl()
    {
        InitializeComponent();
        // 初始化控件
    }
    
    private void btnProcess_Click(object sender, EventArgs e)
    {
        // 处理按钮点击事件
        ProcessDocument();
    }
    
    private void ProcessDocument()
    {
        // 文档处理逻辑
        var app = Globals.ThisAddIn.Application;
        var doc = app.ActiveDocument;
        // 执行文档处理操作
    }
}

// 在插件启动时添加任务窗格
private Microsoft.Office.Tools.CustomTaskPane customTaskPane;

private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    CustomTaskPaneControl taskPaneControl = new CustomTaskPaneControl();
    customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "文档工具");
    customTaskPane.Visible = true;
    customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
    customTaskPane.Width = 300;
}
*/
```

VSTO插件中创建自定义任务窗格的示例代码。

## 对话框操作

Word提供了多种内置对话框，可以通过代码打开和操作这些对话框。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 访问对话框集合
var dialogs = app.Dialogs;
```

获取Word应用程序的对话框集合。

```csharp
// 打开字体对话框
var fontDialog = dialogs[WdWordDialog.wdDialogFormatFont];
int fontResult = fontDialog.Show();
if (fontResult == 1) // 用户点击了确定
{
    Console.WriteLine("字体设置已应用");
}
```

打开字体对话框：
- WdWordDialog.wdDialogFormatFont：字体格式对话框
- Show()：显示对话框并返回结果（1表示确定，0表示取消）

```csharp
// 打开段落对话框
var paragraphDialog = dialogs[WdWordDialog.wdDialogFormatParagraph];
int paragraphResult = paragraphDialog.Show();
if (paragraphResult == 1)
{
    Console.WriteLine("段落格式已应用");
}

// 打开页面设置对话框
var pageSetupDialog = dialogs[WdWordDialog.wdDialogFilePageSetup];
int pageSetupResult = pageSetupDialog.Show();
if (pageSetupResult == 1)
{
    Console.WriteLine("页面设置已应用");
}

// 打开查找替换对话框
var findDialog = dialogs[WdWordDialog.wdDialogEditFind];
int findResult = findDialog.Show();
if (findResult == 1)
{
    Console.WriteLine("查找操作已完成");
}
```

打开其他常用对话框。

## 用户交互处理

可以通过代码处理用户在对话框中的输入和选择。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 创建自定义对话框交互示例
void ShowCustomFontDialog()
{
    var fontDialog = app.Dialogs[WdWordDialog.wdDialogFormatFont];
    
    // 设置默认值
    fontDialog.DefaultTab = WdFontDialogTab.wdFontTabFont;
```

设置字体对话框的默认选项卡。

```csharp
    // 显示对话框并获取结果
    int result = fontDialog.Show();
    
    if (result == 1) // 用户点击了确定
    {
        // 获取用户选择的设置
        string fontName = fontDialog.Font;
        int fontSize = fontDialog.Points;
        bool isBold = fontDialog.Bold != 0;
        bool isItalic = fontDialog.Italic != 0;
```

获取用户在对话框中的选择。

```csharp
        Console.WriteLine($"用户选择了字体: {fontName}, 大小: {fontSize}");
        Console.WriteLine($"粗体: {isBold}, 斜体: {isItalic}");
        
        // 应用到当前选择
        var selection = app.Selection;
        if (selection != null)
        {
            selection.Font.Name = fontName;
            selection.Font.Size = fontSize;
            selection.Font.Bold = isBold ? 1 : 0;
            selection.Font.Italic = isItalic ? 1 : 0;
        }
    }
}
```

将用户选择应用到当前文档选择区域。

```csharp
// 显示自定义页面设置对话框
void ShowCustomPageSetupDialog()
{
    var pageSetupDialog = app.Dialogs[WdWordDialog.wdDialogFilePageSetup];
    
    // 显示对话框
    int result = pageSetupDialog.Show();
    
    if (result == 1) // 用户点击了确定
    {
        // 应用页面设置到当前节
        var section = document.Sections[1];
        var pageSetup = section.PageSetup;
```

应用页面设置。

```csharp
        Console.WriteLine("页面设置已更新");
        Console.WriteLine($"页面宽度: {pageSetup.PageWidth}");
        Console.WriteLine($"页面高度: {pageSetup.PageHeight}");
        Console.WriteLine($"上边距: {pageSetup.TopMargin}");
        Console.WriteLine($"下边距: {pageSetup.BottomMargin}");
    }
}
```

## 实际应用示例

以下示例演示了如何创建一个包含自定义任务窗格和对话框交互的文档处理工具：

```csharp
using MudTools.OfficeInterop;
using System;

class TaskPaneAndDialogDemo
{
    public static void DemonstrateTaskPaneConcepts()
    {
        Console.WriteLine("=== 任务窗格和对话框概念演示 ===");
        Console.WriteLine();
        
        Console.WriteLine("1. 自定义任务窗格XML结构示例:");
        Console.WriteLine(@"<?xml version='1.0' encoding='utf-8' ?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabCustom' label='自定义工具'>
        <group id='grpTaskPane' label='任务窗格'>
          <button id='btnShowTaskPane' label='显示任务窗格' 
                  onAction='OnShowTaskPane' />
          <button id='btnHideTaskPane' label='隐藏任务窗格' 
                  onAction='OnHideTaskPane' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>");
```

展示自定义任务窗格的XML结构示例。

```csharp
        Console.WriteLine();
        Console.WriteLine("2. 任务窗格用户控件示例 (C# Windows Forms):");
        Console.WriteLine(@"
public partial class DocumentToolsPane : UserControl
{
    public DocumentToolsPane()
    {
        InitializeComponent();
        InitializeCustomComponents();
    }
    
    private void InitializeCustomComponents()
    {
        // 创建控件
        var groupBox = new GroupBox();
        groupBox.Text = ""文档格式化"";
        groupBox.Location = new Point(10, 10);
        groupBox.Size = new Size(280, 150);
        
        var lblFontSize = new Label();
        lblFontSize.Text = ""字体大小:"";
        lblFontSize.Location = new Point(10, 20);
        
        var cmbFontSize = new ComboBox();
        cmbFontSize.Items.AddRange(new[] { ""8"", ""10"", ""12"", ""14"", ""16"", ""18"", ""20"" });
        cmbFontSize.Location = new Point(80, 20);
        cmbFontSize.SelectedIndex = 2; // 默认选择12
```

展示任务窗格用户控件示例。

```csharp
        var btnBold = new Button();
        btnBold.Text = ""加粗"";
        btnBold.Location = new Point(10, 50);
        btnBold.Click += BtnBold_Click;
        
        var btnItalic = new Button();
        btnItalic.Text = ""斜体"";
        btnItalic.Location = new Point(90, 50);
        btnItalic.Click += BtnItalic_Click;
        
        var btnUnderline = new Button();
        btnUnderline.Text = ""下划线"";
        btnUnderline.Location = new Point(170, 50);
        btnUnderline.Click += BtnUnderline_Click;
        
        // 添加控件到GroupBox
        groupBox.Controls.AddRange(new Control[] { lblFontSize, cmbFontSize, btnBold, btnItalic, btnUnderline });
        
        // 添加到用户控件
        this.Controls.Add(groupBox);
    }
```

继续展示用户控件示例。

```csharp
    private void BtnBold_Click(object sender, EventArgs e)
    {
        ApplyFormatting(""Bold"");
    }
    
    private void BtnItalic_Click(object sender, EventArgs e)
    {
        ApplyFormatting(""Italic"");
    }
    
    private void BtnUnderline_Click(object sender, EventArgs e)
    {
        ApplyFormatting(""Underline"");
    }
    
    private void ApplyFormatting(string formatType)
    {
        try
        {
            var app = Globals.ThisAddIn.Application;
            var selection = app.Selection;
```

展示格式化应用方法。

```csharp
            if (selection != null)
            {
                switch (formatType)
                {
                    case ""Bold"":
                        selection.Font.Bold = (selection.Font.Bold == 1) ? 0 : 1;
                        break;
                    case ""Italic"":
                        selection.Font.Italic = (selection.Font.Italic == 1) ? 0 : 1;
                        break;
                    case ""Underline"":
                        selection.Font.Underline = (selection.Font.Underline == WdUnderline.wdUnderlineSingle) 
                            ? WdUnderline.wdUnderlineNone 
                            : WdUnderline.wdUnderlineSingle;
                        break;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($""格式化出错: {ex.Message}"");
        }
    }
}");
    }
```

## 对话框交互演示

```csharp
    public static void DemonstrateDialogInteractions()
    {
        Console.WriteLine("=== 对话框交互演示 ===");
        Console.WriteLine();
        
        Console.WriteLine("1. 常用对话框操作示例:");
        Console.WriteLine(@"
// 打开文件对话框
public void OpenFileDialog()
{
    var app = Globals.ThisAddIn.Application;
    var fileDialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
    
    fileDialog.Title = ""选择要打开的文档"";
    fileDialog.Filters.Add(""Word文档"", ""*.docx;*.doc"");
    fileDialog.Filters.Add(""所有文件"", ""*.*"");
```

展示打开文件对话框示例。

```csharp
    if (fileDialog.Show() == -1) // 用户点击了确定
    {
        string selectedFile = fileDialog.SelectedItems[1];
        app.Documents.Open(selectedFile);
    }
}

// 保存文件对话框
public void SaveFileDialog()
{
    var app = Globals.ThisAddIn.Application;
    var fileDialog = app.FileDialog[MsoFileDialogType.msoFileDialogSaveAs];
    
    fileDialog.Title = ""保存文档"";
    fileDialog.InitialFileName = ""新文档.docx"";
    
    if (fileDialog.Show() == -1)
    {
        string fileName = fileDialog.InitialFileName;
        var doc = app.ActiveDocument;
        doc.SaveAs2(fileName);
    }
}");
```

展示保存文件对话框示例。

```csharp
        Console.WriteLine();
        Console.WriteLine("2. 自定义对话框交互:");
        Console.WriteLine(@"
// 自定义文档属性对话框
public void ShowCustomDocumentProperties()
{
    var app = Globals.ThisAddIn.Application;
    var doc = app.ActiveDocument;
    
    // 创建自定义对话框
    var dialog = new DocumentPropertiesDialog();
    
    // 设置初始值
    dialog.Title = doc.Title ?? """";
    dialog.Author = doc.Author ?? """";
    dialog.Subject = doc.Subject ?? """";
    dialog.Keywords = doc.Keywords ?? """";
```

展示自定义对话框交互示例。

```csharp
    // 显示对话框
    if (dialog.ShowDialog() == DialogResult.OK)
    {
        // 应用更改
        doc.Title = dialog.Title;
        doc.Author = dialog.Author;
        doc.Subject = dialog.Subject;
        doc.Keywords = dialog.Keywords;
        
        MessageBox.Show(""文档属性已更新"");
    }
}");
    }
```

```csharp
    // 使用MudTools.OfficeInterop.Word创建支持任务窗格和对话框的文档
    public static void CreateDialogSupportingDocument()
    {
        using var app = WordFactory.BlankDocument();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            
            // 创建示例文档内容
            document.Range().Text = "任务窗格和对话框支持文档\n\n" +
                                  "此文档演示了如何为Word开发支持自定义任务窗格和对话框交互的插件。\n\n" +
                                  "主要特性包括：\n" +
                                  "1. 自定义任务窗格\n" +
                                  "2. 对话框操作\n" +
                                  "3. 用户交互处理\n" +
                                  "4. 实时格式化\n\n" +
                                  "完整实现需要在VSTO插件环境中进行。";
```

创建支持任务窗格和对话框的示例文档。

```csharp
            // 格式化标题
            var titleRange = document.Range(0, 15);
            titleRange.Font.Size = 16;
            titleRange.Font.Bold = 1;
            titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 格式化列表
            var listStart = document.Range().Text.IndexOf("主要特性包括：");
            var listEnd = document.Range().Text.IndexOf("完整实现需要在VSTO插件环境中进行。");
            if (listStart > 0 && listEnd > listStart)
            {
                var listRange = document.Range(listStart, listEnd);
                listRange.ListFormat.ApplyBulletDefault();
            }
```

格式化文档内容。

```csharp
            // 保存文档
            document.SaveAs2(@"C:\temp\DialogSupportingDocument.docx");
            
            Console.WriteLine("对话框支持文档已创建: DialogSupportingDocument.docx");
            Console.WriteLine("注意：完整的任务窗格和对话框功能需要在VSTO插件环境中实现");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建文档时出错: {ex.Message}");
        }
    }
}

// 任务窗格和对话框的最佳实践
class TaskPaneAndDialogBestPractices
{
    public static void ShowBestPractices()
    {
        Console.WriteLine("=== 任务窗格和对话框最佳实践 ===");
        Console.WriteLine();
        
        Console.WriteLine("1. 任务窗格设计:");
        Console.WriteLine("   - 保持界面简洁，避免过于复杂");
        Console.WriteLine("   - 提供清晰的标签和说明");
        Console.WriteLine("   - 合理组织控件布局");
        Console.WriteLine("   - 支持键盘导航");
```

展示任务窗格和对话框最佳实践。

```csharp
        Console.WriteLine();
        Console.WriteLine("2. 对话框设计:");
        Console.WriteLine("   - 使用标准对话框以保持一致性");
        Console.WriteLine("   - 提供明确的确定/取消按钮");
        Console.WriteLine("   - 验证用户输入");
        Console.WriteLine("   - 记住用户偏好设置");
        
        Console.WriteLine();
        Console.WriteLine("3. 用户体验:");
        Console.WriteLine("   - 提供有意义的默认值");
        Console.WriteLine("   - 给出清晰的操作反馈");
        Console.WriteLine("   - 处理异常情况");
        Console.WriteLine("   - 支持撤销操作");
        
        Console.WriteLine();
        Console.WriteLine("4. 性能考虑:");
        Console.WriteLine("   - 避免在UI线程执行耗时操作");
        Console.WriteLine("   - 及时释放资源");
        Console.WriteLine("   - 优化控件响应速度");
    }
}
```

## 应用场景

1. **文档编辑器**：创建专业的文档编辑工具
2. **格式化工具**：提供快速格式化选项
3. **数据导入导出**：实现数据交互功能
4. **模板管理**：管理文档模板和样式

## 要点总结

- 自定义任务窗格提供停靠在Word窗口侧面的交互面板
- 对话框操作允许打开和控制Word内置对话框
- 用户交互处理可以响应用户的操作和输入
- 完整实现需要在VSTO插件环境中进行
- 应遵循UI设计最佳实践，提供良好的用户体验

掌握任务窗格和对话框技能对于开发交互式Word插件和文档处理工具非常重要，这些功能使开发者能够创建更加直观和高效的用户界面。