# 工作簿与工作表操作基础

## 引言：Excel自动化的"心脏"与"骨架"

在前两篇文章中，我们已经成功打开了Excel自动化的大门，掌握了应用程序的创建与管理。现在，让我们深入Excel自动化的核心地带——工作簿与工作表操作！

如果把Excel自动化比作一个完整的人体，那么工作簿就是"心脏"，负责承载所有的数据和功能；而工作表则是"骨架"，为数据提供结构和支撑。没有心脏，人体无法存活；没有骨架，人体无法站立。同样，没有工作簿和工作表的熟练操作，Excel自动化就无从谈起。

想象一下这样的场景：你需要处理一个包含多个部门数据的复杂报表，每个部门的数据分布在不同的工作表中。你需要合并数据、重新组织结构、应用统一的格式。如果手动操作，这可能需要数小时甚至数天的时间。但通过自动化技术，这一切都可以在几分钟内完成！

本篇将带你探索工作簿和工作表的奥秘，从基础属性访问到高级操作技巧，从简单的数据管理到复杂的结构重组。准备好让你的Excel自动化技能更上一层楼了吗？

## 工作簿操作详解

### 工作簿属性与状态管理

工作簿是Excel文档的容器，包含了所有的工作表、图表、宏等元素。让我们先来了解工作簿的基本属性和状态管理。

1. **IExcelApplication（Excel应用程序）** - 代表整个Excel应用程序实例
2. **IExcelWorkbooks（工作簿集合）** - 包含所有打开的工作簿
3. **IExcelWorkbook（工作簿）** - 代表单个工作簿文件
4. **IExcelWorksheets、IExcelSheets、IExcelComSheets（工作表集合）** - 包含工作簿中的所有工作表
5. **IExcelWorksheet、IExcelComSheet（工作表）** - 代表单个工作表
6. **IExcelRange（单元格区域）** - 代表工作表中的单元格或单元格区域

#### 核心属性访问

```csharp
public class WorkbookInspector
{
    public void InspectWorkbook(IExcelWorkbook workbook)
    {
        // 基本信息
        Console.WriteLine($"工作簿名称: {workbook.Name}");
        Console.WriteLine($"完整路径: {workbook.FullName}");
        Console.WriteLine($"文件路径: {workbook.Path}");
        Console.WriteLine($"文件大小: {workbook.FileSize} 字节");
        
        // 时间信息
        Console.WriteLine($"创建时间: {workbook.CreatedTime}");
        Console.WriteLine($"修改时间: {workbook.ModifiedTime}");
        
        // 保护状态
        Console.WriteLine($"是否受密码保护: {workbook.HasPassword}");
        Console.WriteLine($"结构是否受保护: {workbook.ProtectStructure}");
        Console.WriteLine($"窗口是否受保护: {workbook.ProtectWindows}");
        Console.WriteLine($"是否只读: {workbook.ReadOnly}");
        
        // 多用户状态
        Console.WriteLine($"多用户编辑状态: {workbook.MultiUserEditing}");
        
        // 文档属性
        Console.WriteLine($"关键词: {workbook.Keywords}");
        Console.WriteLine($"主题: {workbook.Subject}");
        Console.WriteLine($"作者: {workbook.Author}");
        
        // 技术信息
        Console.WriteLine($"文件格式: {workbook.FileFormat}");
        Console.WriteLine($"是否包含VB工程: {workbook.HasVBProject}");
        Console.WriteLine($"是否为插件工作簿: {workbook.IsAddin}");
    }
    
    public void UpdateWorkbookProperties(IExcelWorkbook workbook)
    {
        // 更新文档属性
        workbook.Keywords = "销售数据,报表,分析";
        workbook.Subject = "2024年销售季度报告";
        workbook.Author = "销售分析部门";
        
        // 设置精确计算
        workbook.PrecisionAsDisplayed = true;
        
        // 设置图形显示方式
        workbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlDisplayShapes;
    }
}
```

#### 工作簿保护与安全

```csharp
public class WorkbookSecurityManager
{
    public void ProtectWorkbook(IExcelWorkbook workbook, string password)
    {
        // 保护工作簿结构
        workbook.Protect(password, true, false);
        
        // 验证保护状态
        if (workbook.ProtectStructure)
        {
            Console.WriteLine("工作簿结构已受保护");
        }
        
        if (workbook.ProtectWindows)
        {
            Console.WriteLine("工作簿窗口已受保护");
        }
    }
    
    public void UnprotectWorkbook(IExcelWorkbook workbook, string password)
    {
        try
        {
            workbook.Unprotect(password);
            Console.WriteLine("工作簿保护已解除");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"解除保护失败: {ex.Message}");
        }
    }
    
    public bool VerifyPassword(IExcelWorkbook workbook, string password)
    {
        if (!workbook.HasPassword)
        {
            return true; // 没有密码保护
        }
        
        try
        {
            // 尝试解除保护来验证密码
            workbook.Unprotect(password);
            // 如果成功，重新保护
            workbook.Protect(password, true, false);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
```

### 工作簿文件操作

#### 保存与另存为

```csharp
public class WorkbookFileManager
{
    public void SaveWorkbookWithOptions(IExcelWorkbook workbook)
    {
        // 检查是否需要保存
        if (!workbook.Saved)
        {
            Console.WriteLine("工作簿有未保存的更改");
        }
        
        // 保存工作簿
        workbook.Save();
        Console.WriteLine("工作簿已保存");
    }
    
    public void SaveAsWithFormat(IExcelWorkbook workbook, string filePath, XlFileFormat format)
    {
        // 保存为不同格式
        workbook.SaveAs(filePath, format);
        Console.WriteLine($"工作簿已保存为: {filePath}");
    }
    
    public void ExportToDifferentFormats(IExcelWorkbook workbook, string baseName)
    {
        // 导出为多种格式
        var formats = new Dictionary<string, XlFileFormat>
        {
            ["xlsx"] = XlFileFormat.xlOpenXMLWorkbook,
            ["xls"] = XlFileFormat.xlExcel8,
            ["pdf"] = XlFileFormat.xlPDF,
            ["csv"] = XlFileFormat.xlCSV
        };
        
        foreach (var format in formats)
        {
            string filePath = $"{baseName}.{format.Key}";
            try
            {
                workbook.SaveAs(filePath, format.Value);
                Console.WriteLine($"成功导出为: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出失败 {filePath}: {ex.Message}");
            }
        }
    }
    
    public void CreateBackup(IExcelWorkbook workbook)
    {
        string originalPath = workbook.FullName;
        string backupPath = Path.Combine(
            Path.GetDirectoryName(originalPath) ?? string.Empty,
            $"{Path.GetFileNameWithoutExtension(originalPath)}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(originalPath)}"
        );
        
        workbook.SaveCopyAs(backupPath);
        Console.WriteLine($"备份已创建: {backupPath}");
    }
}
```

#### 工作簿关闭与清理

```csharp
public class WorkbookCloser
{
    public bool CloseWorkbook(IExcelWorkbook workbook, bool saveChanges = true)
    {
        try
        {
            if (saveChanges && !workbook.Saved)
            {
                workbook.Save();
            }
            
            workbook.Close();
            Console.WriteLine("工作簿已关闭");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"关闭工作簿失败: {ex.Message}");
            return false;
        }
    }
    
    public void CloseAllWorkbooks(IExcelApplication excelApp)
    {
        var workbooks = excelApp.Workbooks;
        if (workbooks != null)
        {
            foreach (var workbook in workbooks)
            {
                try
                {
                    CloseWorkbook(workbook, false); // 不保存更改
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"关闭工作簿 {workbook.Name} 失败: {ex.Message}");
                }
            }
        }
    }
}
```

## 工作表操作详解

### 工作表集合管理

工作表集合（Worksheets）提供了对工作簿中所有工作表的管理功能。

```csharp
public class WorksheetsManager
{
    public void ManageWorksheets(IExcelWorksheets worksheets)
    {
        // 获取工作表数量
        int count = worksheets.Count;
        Console.WriteLine($"工作表总数: {count}");
        
        // 遍历所有工作表
        foreach (var worksheet in worksheets)
        {
            Console.WriteLine($"工作表: {worksheet.Name}");
        }
        
        // 按索引访问
        if (count > 0)
        {
            var firstSheet = worksheets[1];
            Console.WriteLine($"第一个工作表: {firstSheet?.Name}");
        }
        
        // 按名称访问
        var sheetByName = worksheets["Sheet1"];
        if (sheetByName != null)
        {
            Console.WriteLine("找到名为Sheet1的工作表");
        }
    }
    
    public void FilterWorksheets(IExcelWorksheets worksheets)
    {
        // 获取可见工作表
        var visibleSheets = worksheets.GetVisibleWorksheets();
        Console.WriteLine($"可见工作表数量: {visibleSheets.Length}");
        
        // 获取隐藏工作表
        var hiddenSheets = worksheets.GetHiddenWorksheets();
        Console.WriteLine($"隐藏工作表数量: {hiddenSheets.Length}");
        
        // 获取受保护的工作表
        var protectedSheets = worksheets.GetProtectedWorksheets();
        Console.WriteLine($"受保护工作表数量: {protectedSheets.Length}");
        
        // 获取未受保护的工作表
        var unprotectedSheets = worksheets.GetUnprotectedWorksheets();
        Console.WriteLine($"未受保护工作表数量: {unprotectedSheets.Length}");
    }
    
    public void BatchAddWorksheets(IExcelWorksheets worksheets, string[] sheetNames)
    {
        // 批量添加工作表
        int addedCount = worksheets.AddRange(sheetNames);
        Console.WriteLine($"成功添加 {addedCount} 个工作表");
        
        // 验证添加结果
        foreach (var name in sheetNames)
        {
            var sheet = worksheets[name];
            if (sheet != null)
            {
                Console.WriteLine($"工作表 {name} 添加成功");
            }
        }
    }
    
    public void ReorganizeWorksheets(IExcelWorksheets worksheets)
    {
        // 移动工作表
        var targetSheet = worksheets["目标位置"];
        var sheetToMove = worksheets["要移动的工作表"];
        
        if (targetSheet != null && sheetToMove != null)
        {
            worksheets.Move(sheetToMove, before: targetSheet);
            Console.WriteLine("工作表移动完成");
        }
        
        // 复制工作表
        var sheetToCopy = worksheets["要复制的工作表"];
        if (sheetToCopy != null)
        {
            var copiedSheet = worksheets.Copy(sheetToCopy, newName: $"{sheetToCopy.Name}_副本");
            if (copiedSheet != null)
            {
                Console.WriteLine($"工作表复制完成: {copiedSheet.Name}");
            }
        }
    }
}
```

### 单个工作表操作

#### 工作表属性与配置

```csharp
public class WorksheetConfigurator
{
    public void ConfigureWorksheet(IExcelWorksheet worksheet)
    {
        // 基本属性
        Console.WriteLine($"工作表名称: {worksheet.Name}");
        Console.WriteLine($"代码名称: {worksheet.CodeName}");
        Console.WriteLine($"标签颜色: {worksheet.TabColor}");
        
        // 显示设置
        worksheet.DisplayPageBreaks = true;
        worksheet.DisplayAutomaticPageBreaks = true;
        
        // 功能设置
        worksheet.EnableOutlining = true;
        worksheet.EnablePivotTable = true;
        worksheet.EnableSelection = XlEnableSelection.xlNoRestrictions;
        
        // 计算设置
        worksheet.TransitionExpEval = false;
        worksheet.StandardWidth = 8.5; // 标准列宽
        
        // 设置标签颜色
        worksheet.TabColor = Color.LightBlue;
    }
    
    public void InspectWorksheetState(IExcelWorksheet worksheet)
    {
        // 保护状态
        Console.WriteLine($"图形保护: {worksheet.ProtectDrawingObjects}");
        Console.WriteLine($"方案保护: {worksheet.ProtectScenarios}");
        Console.WriteLine($"筛选模式: {worksheet.FilterMode}");
        Console.WriteLine($"自动筛选模式: {worksheet.AutoFilterMode}");
        
        // 导航关系
        var nextSheet = worksheet.Next;
        var prevSheet = worksheet.Previous;
        
        Console.WriteLine($"下一个工作表: {nextSheet?.Name ?? "无"}");
        Console.WriteLine($"上一个工作表: {prevSheet?.Name ?? "无"}");
    }
}
```

#### 工作表保护

```csharp
public class WorksheetProtector
{
    public void ProtectWorksheet(IExcelWorksheet worksheet, string password)
    {
        // 保护工作表
        worksheet.Protect(
            password: password,
            drawingObjects: true,
            contents: true,
            scenarios: true,
            userInterfaceOnly: false
        );
        
        Console.WriteLine("工作表保护已启用");
    }
    
    public void ProtectWithOptions(IExcelWorksheet worksheet, string password)
    {
        // 详细的保护选项
        worksheet.Protect(
            password: password,
            drawingObjects: true,      // 保护图形对象
            contents: true,            // 保护内容
            scenarios: true,           // 保护方案
            userInterfaceOnly: false,  // 是否仅保护用户界面
            allowFormattingCells: true, // 允许格式化单元格
            allowFormattingColumns: true, // 允许格式化列
            allowFormattingRows: true,  // 允许格式化行
            allowInsertingColumns: false, // 禁止插入列
            allowInsertingRows: false,  // 禁止插入行
            allowInsertingHyperlinks: false, // 禁止插入超链接
            allowDeletingColumns: false, // 禁止删除列
            allowDeletingRows: false,   // 禁止删除行
            allowSorting: true,         // 允许排序
            allowFiltering: true,       // 允许筛选
            allowUsingPivotTables: true // 允许使用数据透视表
        );
        
        Console.WriteLine("工作表已使用详细选项保护");
    }
    
    public void UnprotectWorksheet(IExcelWorksheet worksheet, string password)
    {
        try
        {
            worksheet.Unprotect(password);
            Console.WriteLine("工作表保护已解除");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"解除保护失败: {ex.Message}");
        }
    }
}
```

### 工作表创建与删除

#### 创建工作表

```csharp
public class WorksheetCreator
{
    public IExcelWorksheet CreateWorksheetWithOptions(IExcelWorksheets worksheets, string name)
    {
        // 创建新工作表
        var newWorksheet = worksheets.Add(name);
        
        if (newWorksheet != null)
        {
            // 配置新工作表
            ConfigureNewWorksheet(newWorksheet);
            Console.WriteLine($"工作表 {name} 创建成功");
        }
        
        return newWorksheet;
    }
    
    private void ConfigureNewWorksheet(IExcelWorksheet worksheet)
    {
        // 设置默认格式
        worksheet.StandardWidth = 10.0;
        
        // 设置默认内容
        worksheet.Range("A1").Value = "创建时间:";
        worksheet.Range("B1").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        
        worksheet.Range("A2").Value = "创建者:";
        worksheet.Range("B2").Value = Environment.UserName;
        
        // 设置表头样式
        var headerRange = worksheet.Range("A1:B2");
        headerRange.Font.Bold = true;
        headerRange.Interior.Color = Color.LightGray;
    }
    
    public void CreateTemplateWorksheets(IExcelWorksheets worksheets, string[] templateNames)
    {
        foreach (var templateName in templateNames)
        {
            var worksheet = worksheets.Add(templateName);
            if (worksheet != null)
            {
                ApplyTemplateLayout(worksheet, templateName);
            }
        }
    }
    
    private void ApplyTemplateLayout(IExcelWorksheet worksheet, string templateType)
    {
        switch (templateType)
        {
            case "数据输入":
                SetupDataInputTemplate(worksheet);
                break;
            case "分析报告":
                SetupAnalysisTemplate(worksheet);
                break;
            case "汇总统计":
                SetupSummaryTemplate(worksheet);
                break;
        }
    }
    
    private void SetupDataInputTemplate(IExcelWorksheet worksheet)
    {
        // 设置数据输入模板
        string[] headers = { "序号", "日期", "产品", "数量", "单价", "金额" };
        
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
            worksheet.Cells[1, i + 1].Font.Bold = true;
        }
        
        // 设置数据验证
        SetupDataValidation(worksheet);
    }
    
    private void SetupDataValidation(IExcelWorksheet worksheet)
    {
        // 示例：设置数量列的数据验证
        var quantityColumn = worksheet.Columns[4]; // D列
        // 实际的数据验证设置需要更复杂的逻辑
    }
}
```

#### 删除工作表

```csharp
public class WorksheetRemover
{
    public void DeleteWorksheetSafely(IExcelWorksheets worksheets, string sheetName)
    {
        var worksheet = worksheets[sheetName];
        if (worksheet != null)
        {
            try
            {
                // 检查是否是最后一个工作表
                if (worksheets.Count <= 1)
                {
                    Console.WriteLine("不能删除最后一个工作表");
                    return;
                }
                
                worksheet.Delete();
                Console.WriteLine($"工作表 {sheetName} 已删除");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除工作表失败: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"未找到工作表: {sheetName}");
        }
    }
    
    public void DeleteMultipleWorksheets(IExcelWorksheets worksheets, string[] sheetNames)
    {
        foreach (var sheetName in sheetNames)
        {
            DeleteWorksheetSafely(worksheets, sheetName);
        }
    }
    
    public void CleanupEmptyWorksheets(IExcelWorksheets worksheets)
    {
        var sheetsToDelete = new List<string>();
        
        foreach (var worksheet in worksheets)
        {
            if (IsWorksheetEmpty(worksheet))
            {
                sheetsToDelete.Add(worksheet.Name);
            }
        }
        
        // 删除空工作表（保留至少一个）
        if (sheetsToDelete.Count < worksheets.Count)
        {
            DeleteMultipleWorksheets(worksheets, sheetsToDelete.ToArray());
        }
    }
    
    private bool IsWorksheetEmpty(IExcelWorksheet worksheet)
    {
        var usedRange = worksheet.UsedRange;
        if (usedRange == null) return true;
        
        // 检查是否有数据
        for (int row = 1; row <= usedRange.Rows.Count; row++)
        {
            for (int col = 1; col <= usedRange.Columns.Count; col++)
            {
                var cellValue = usedRange.Cells[row, col].Value?.ToString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    return false;
                }
            }
        }
        
        return true;
    }
}
```

### 工作表导航与选择

```csharp
public class WorksheetNavigator
{
    public void NavigateWorksheets(IExcelWorksheets worksheets)
    {
        // 激活第一个工作表
        var firstSheet = worksheets[1];
        if (firstSheet != null)
        {
            firstSheet.Activate();
            Console.WriteLine($"已激活工作表: {firstSheet.Name}");
        }
        
        // 按名称激活工作表
        var targetSheet = worksheets["数据汇总"];
        if (targetSheet != null)
        {
            targetSheet.Activate();
            Console.WriteLine($"已激活工作表: {targetSheet.Name}");
        }
        
        // 遍历所有工作表
        Console.WriteLine("所有工作表:");
        foreach (var worksheet in worksheets)
        {
            Console.WriteLine($"- {worksheet.Name}");
        }
    }
    
    public void CreateWorksheetNavigation(IExcelApplication excelApp)
    {
        var worksheets = excelApp.Worksheets;
        if (worksheets == null) return;
        
        // 创建导航工作表
        var navSheet = worksheets.Add("导航");
        if (navSheet != null)
        {
            SetupNavigationSheet(navSheet, worksheets);
        }
    }
    
    private void SetupNavigationSheet(IExcelWorksheet navSheet, IExcelWorksheets worksheets)
    {
        // 设置导航表头
        navSheet.Range("A1").Value = "工作表导航";
        navSheet.Range("A1").Font.Bold = true;
        navSheet.Range("A1").Font.Size = 14;
        
        // 创建工作表链接
        int row = 3;
        foreach (var worksheet in worksheets)
        {
            if (worksheet.Name != "导航") // 排除导航表本身
            {
                navSheet.Cells[row, 1].Value = worksheet.Name;
                
                // 实际中需要创建超链接来跳转到对应工作表
                // 这里只是示例
                row++;
            }
        }
        
        // 自动调整列宽
        navSheet.Columns.AutoFit();
    }
}
```

## 实际应用场景

### 场景1：多工作表报表系统

```csharp
public class MultiSheetReportSystem
{
    public void GenerateComprehensiveReport(IExcelApplication excelApp, ReportData data)
    {
        using var excelAppWrapper = excelApp; // 确保资源释放
        
        // 创建主报告工作表
        var summarySheet = excelApp.Worksheets?.Add("报告汇总");
        if (summarySheet != null)
        {
            SetupSummarySheet(summarySheet, data);
        }
        
        // 创建详细数据工作表
        var detailSheet = excelApp.Worksheets?.Add("详细数据");
        if (detailSheet != null)
        {
            SetupDetailSheet(detailSheet, data);
        }
        
        // 创建分析图表工作表
        var chartSheet = excelApp.Worksheets?.Add("分析图表");
        if (chartSheet != null)
        {
            SetupChartSheet(chartSheet, data);
        }
        
        // 创建数据验证工作表
        var validationSheet = excelApp.Worksheets?.Add("数据验证");
        if (validationSheet != null)
        {
            SetupValidationSheet(validationSheet, data);
        }
        
        // 设置工作表保护
        ProtectReportSheets(excelApp.Worksheets);
        
        // 激活汇总表
        summarySheet?.Activate();
    }
    
    private void SetupSummarySheet(IExcelWorksheet sheet, ReportData data)
    {
        // 设置汇总表格式和内容
        sheet.Range("A1").Value = "报告汇总";
        sheet.Range("A1").Font.Bold = true;
        sheet.Range("A1").Font.Size = 16;
        
        // 添加汇总数据
        // ...
    }
    
    private void ProtectReportSheets(IExcelWorksheets worksheets)
    {
        if (worksheets == null) return;
        
        foreach (var worksheet in worksheets)
        {
            if (worksheet.Name != "详细数据") // 详细数据表允许编辑
            {
                worksheet.Protect("report123", true, true);
            }
        }
    }
}

public class ReportData
{
    // 报告数据结构
    public DateTime ReportDate { get; set; }
    public List<SalesRecord> SalesData { get; set; } = new();
    public List<AnalysisResult> AnalysisResults { get; set; } = new();
}
```

### 场景2：模板化工作簿生成器

```csharp
public class TemplateWorkbookGenerator
{
    private readonly WorkbookTemplate _template;
    
    public TemplateWorkbookGenerator(WorkbookTemplate template)
    {
        _template = template;
    }
    
    public IExcelWorkbook GenerateWorkbook()
    {
        using var excelApp = ExcelFactory.BlankWorkbook();
        excelApp.Visible = false;
        
        var workbook = excelApp.ActiveWorkbook;
        
        // 删除默认工作表
        CleanDefaultSheets(excelApp.Worksheets);
        
        // 根据模板创建工作表
        foreach (var sheetTemplate in _template.SheetTemplates)
        {
            CreateSheetFromTemplate(excelApp.Worksheets, sheetTemplate);
        }
        
        // 设置工作簿属性
        SetupWorkbookProperties(workbook);
        
        return workbook;
    }
    
    private void CleanDefaultSheets(IExcelWorksheets worksheets)
    {
        if (worksheets == null) return;
        
        // 保留至少一个工作表
        while (worksheets.Count > 1)
        {
            var lastSheet = worksheets[worksheets.Count];
            if (lastSheet != null)
            {
                lastSheet.Delete();
            }
        }
        
        // 重命名剩余的工作表
        var remainingSheet = worksheets[1];
        if (remainingSheet != null)
        {
            remainingSheet.Name = _template.SheetTemplates.First().Name;
        }
    }
    
    private void CreateSheetFromTemplate(IExcelWorksheets worksheets, SheetTemplate template)
    {
        var worksheet = worksheets.Add(template.Name);
        if (worksheet != null)
        {
            // 应用模板设置
            ApplyTemplateToSheet(worksheet, template);
        }
    }
    
    private void ApplyTemplateToSheet(IExcelWorksheet worksheet, SheetTemplate template)
    {
        // 设置工作表属性
        worksheet.TabColor = template.TabColor;
        worksheet.StandardWidth = template.DefaultColumnWidth;
        
        // 设置表头和格式
        SetupTemplateHeaders(worksheet, template);
        
        // 设置保护
        if (template.IsProtected)
        {
            worksheet.Protect(template.ProtectionPassword);
        }
    }
    
    private void SetupTemplateHeaders(IExcelWorksheet worksheet, SheetTemplate template)
    {
        // 设置表头
        for (int i = 0; i < template.Headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = template.Headers[i];
            worksheet.Cells[1, i + 1].Font.Bold = true;
            worksheet.Cells[1, i + 1].Interior.Color = Color.LightBlue;
        }
    }
    
    private void SetupWorkbookProperties(IExcelWorkbook workbook)
    {
        workbook.Keywords = _template.Keywords;
        workbook.Subject = _template.Subject;
        workbook.Author = _template.Author;
    }
}

public class WorkbookTemplate
{
    public string Name { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Author { get; set; } = "";
    public string Keywords { get; set; } = "";
    public List<SheetTemplate> SheetTemplates { get; set; } = new();
}

public class SheetTemplate
{
    public string Name { get; set; } = "";
    public Color TabColor { get; set; } = Color.White;
    public double DefaultColumnWidth { get; set; } = 8.5;
    public string[] Headers { get; set; } = Array.Empty<string>();
    public bool IsProtected { get; set; } = false;
    public string ProtectionPassword { get; set; } = "";
}
```

### 场景3：动态工作表管理系统

```csharp
public class DynamicWorksheetManager
{
    public void ManageSheetsBasedOnData(IExcelApplication excelApp, DynamicData data)
    {
        var worksheets = excelApp.Worksheets;
        if (worksheets == null) return;
        
        // 根据数据类别创建对应的工作表
        foreach (var category in data.Categories)
        {
            var sheetName = $"数据_{category.Name}";
            
            // 检查是否已存在
            var existingSheet = worksheets[sheetName];
            if (existingSheet == null)
            {
                // 创建新工作表
                var newSheet = worksheets.Add(sheetName);
                if (newSheet != null)
                {
                    SetupCategorySheet(newSheet, category);
                }
            }
            else
            {
                // 更新现有工作表
                UpdateCategorySheet(existingSheet, category);
            }
        }
        
        // 清理不再需要的工作表
        CleanupObsoleteSheets(worksheets, data.Categories.Select(c => $"数据_{c.Name}").ToArray());
    }
    
    private void SetupCategorySheet(IExcelWorksheet worksheet, DataCategory category)
    {
        // 设置类别工作表
        worksheet.Range("A1").Value = $"{category.Name} 数据";
        worksheet.Range("A1").Font.Bold = true;
        
        // 设置数据表头
        for (int i = 0; i < category.Headers.Length; i++)
        {
            worksheet.Cells[3, i + 1].Value = category.Headers[i];
            worksheet.Cells[3, i + 1].Font.Bold = true;
        }
        
        // 设置标签颜色
        worksheet.TabColor = category.Color;
    }
    
    private void UpdateCategorySheet(IExcelWorksheet worksheet, DataCategory category)
    {
        // 更新现有工作表的数据
        // 这里可以添加数据追加或更新的逻辑
    }
    
    private void CleanupObsoleteSheets(IExcelWorksheets worksheets, string[] validSheetNames)
    {
        var sheetsToDelete = new List<string>();
        
        foreach (var worksheet in worksheets)
        {
            if (worksheet.Name.StartsWith("数据_") && 
                !validSheetNames.Contains(worksheet.Name))
            {
                sheetsToDelete.Add(worksheet.Name);
            }
        }
        
        // 删除过时的工作表
        foreach (var sheetName in sheetsToDelete)
        {
            var sheet = worksheets[sheetName];
            if (sheet != null && worksheets.Count > 1)
            {
                sheet.Delete();
            }
        }
    }
}

public class DynamicData
{
    public List<DataCategory> Categories { get; set; } = new();
}

public class DataCategory
{
    public string Name { get; set; } = "";
    public Color Color { get; set; } = Color.LightGray;
    public string[] Headers { get; set; } = Array.Empty<string>();
    public List<object[]> Data { get; set; } = new();
}
```

## 总结

通过本文的学习，我们深入掌握了工作簿和工作表的各种操作技巧，包括：

**工作簿操作要点：**
- 属性访问与状态管理
- 文件保存与格式转换
- 保护与安全设置
- 多用户协作配置

**工作表操作要点：**
- 工作表集合管理
- 单个工作表配置
- 保护与权限控制
- 创建、删除与重命名

**实际应用价值：**
- 多工作表报表系统实现复杂业务需求
- 模板化工作簿生成器提高开发效率
- 动态工作表管理系统适应变化的数据结构

**最佳实践：**
- 使用using语句确保资源释放
- 实现完善的错误处理机制
- 考虑性能优化和内存管理
- 提供用户友好的界面和导航

在下一篇文章中，我们将深入探讨单元格和区域操作，这是Excel数据处理的核心功能。

---

**下一篇预告：**《单元格和区域操作详解》将详细介绍单元格的各种操作技巧，包括数据读写、格式设置、批量操作等高级功能。