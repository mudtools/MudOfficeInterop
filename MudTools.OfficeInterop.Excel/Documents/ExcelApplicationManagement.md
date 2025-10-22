# Excel应用程序的创建与管理

## 引言：掌握Excel自动化的"钥匙"

在上一篇中，我们揭开了MudTools.OfficeInterop.Excel的神秘面纱，了解了它的强大功能和设计理念。现在，让我们正式开启Excel自动化的大门！如果说Excel自动化是一座宝库，那么Excel应用程序的创建与管理就是打开这座宝库的"金钥匙"。

想象一下，你是一名经验丰富的建筑师，ExcelFactory类就是你的工具箱。这个工具箱里装满了各种精良的工具，每一种工具都有其独特的用途和优势。学会使用这些工具，你就能轻松建造出功能强大、稳定可靠的Excel自动化应用。

本篇将带你深入探索ExcelFactory类的奥秘，从基础的实例创建到高级的配置管理，从简单的文件操作到复杂的模板应用。无论你是Excel自动化的新手还是经验丰富的老手，这里都有值得你学习的宝贵知识。准备好开启这段精彩的旅程了吗？

## ExcelFactory类：Excel自动化的"万能工厂"

如果说Excel自动化是一个庞大的工业体系，那么ExcelFactory类就是这个体系的"万能工厂"。它就像一个高度智能化的生产线，能够根据不同的需求生产出各种类型的Excel应用程序实例。

这个工厂的设计理念非常巧妙：它将复杂的COM对象创建过程封装在简洁的API后面，让开发者能够像使用普通.NET类一样轻松创建Excel实例。想象一下，你只需要告诉工厂你想要什么（"给我一个空白工作簿"或者"打开这个文件"），工厂就会自动完成所有的复杂工作，然后把成品交到你手中。

这种设计不仅大大简化了开发流程，更重要的是提高了代码的可维护性和可扩展性。无论你的需求如何变化，ExcelFactory都能灵活应对。

### 工厂模式设计

让我们先来看一下ExcelFactory类的完整结构：

```csharp
// ExcelFactory类的核心方法
public static class ExcelFactory
{
    // 连接到现有Excel实例
    public static IExcelApplication? Connection(object comObj)
    
    // 通过ProgID创建特定版本实例
    public static IExcelApplication CreateInstance(string? progId)
    
    // 创建空白工作簿
    public static IExcelApplication BlankWorkbook()
    
    // 基于模板创建工作簿
    public static IExcelApplication CreateFrom(string templatePath)
    
    // 打开现有工作簿文件
    public static IExcelApplication Open(string filePath)
}
```

### 核心方法详解

#### 1. Connection方法 - 连接到现有实例

```csharp
/// <summary>
/// 通过COM对象连接到现有的Excel应用程序实例
/// </summary>
/// <param name="comObj">COM对象，应为Excel应用程序实例</param>
/// <returns>如果comObj是有效的Excel应用程序实例，则返回封装的IExcelApplication对象；否则返回null</returns>
public static IExcelApplication? Connection(object comObj)
{
    MsExcel.Application? excelCom = comObj as MsExcel.Application;
    if (excelCom == null)
        return null;
    return new ExcelApplication(excelCom);
}
```

**应用场景：**
- 与现有Excel插件集成
- 监控用户打开的Excel文件
- 多应用程序协作

**示例代码：**

```csharp
public class ExcelMonitorService
{
    public void MonitorExistingExcelInstances()
    {
        // 获取当前运行的Excel进程
        var excelProcesses = Process.GetProcessesByName("EXCEL");
        
        foreach (var process in excelProcesses)
        {
            try
            {
                // 尝试连接到现有Excel实例
                var excelApp = ExcelFactory.Connection(GetExcelComObject(process));
                if (excelApp != null)
                {
                    Console.WriteLine($"已连接到Excel实例: {excelApp.Name}");
                    
                    // 监控工作簿变化
                    MonitorWorkbookChanges(excelApp);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"连接失败: {ex.Message}");
            }
        }
    }
    
    private object GetExcelComObject(Process excelProcess)
    {
        // 实际实现需要更复杂的COM对象获取逻辑
        // 这里仅为示例
        return null;
    }
}
```

#### 2. CreateInstance方法 - 创建特定版本实例

```csharp
/// <summary>
/// 根据ProgID创建Excel应用程序的新实例
/// </summary>
/// <param name="progId">Excel应用程序的ProgID</param>
/// <returns>返回新创建的Excel应用程序实例</returns>
public static IExcelApplication CreateInstance(string? progId)
{
    Type type = Type.GetTypeFromProgID(progId);
    if (type == null)
    {
        throw new InvalidOperationException($"无法从 ProgID '{progId}' 获取类型。");
    }

    MsExcel.Application instance = (MsExcel.Application)Activator.CreateInstance(type);
    instance.UserControl = true; // 允许用户控制该实例
    ExcelApplication excel = new(instance);
    return excel;
}
```

**支持的ProgID示例：**
- `Excel.Application` - 最新版本
- `Excel.Application.16` - Excel 2016
- `Excel.Application.15` - Excel 2013
- `Excel.Application.14` - Excel 2010

**应用场景：**
- 多版本Excel兼容性测试
- 特定版本功能需求
- 版本控制要求严格的环境

**示例代码：**

```csharp
public class MultiVersionExcelManager
{
    private readonly Dictionary<string, string> _excelVersions = new()
    {
        ["Latest"] = "Excel.Application",
        ["Excel2016"] = "Excel.Application.16",
        ["Excel2013"] = "Excel.Application.15",
        ["Excel2010"] = "Excel.Application.14"
    };
    
    public IExcelApplication CreateSpecificVersion(string versionKey)
    {
        if (!_excelVersions.ContainsKey(versionKey))
            throw new ArgumentException($"不支持的Excel版本: {versionKey}");
            
        string progId = _excelVersions[versionKey];
        return ExcelFactory.CreateInstance(progId);
    }
    
    public void TestCompatibility()
    {
        foreach (var version in _excelVersions.Keys)
        {
            try
            {
                using var excelApp = CreateSpecificVersion(version);
                Console.WriteLine($"{version} 版本测试成功");
                
                // 测试基本功能
                TestBasicFunctionality(excelApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{version} 版本测试失败: {ex.Message}");
            }
        }
    }
    
    private void TestBasicFunctionality(IExcelApplication excelApp)
    {
        // 测试基本功能
        excelApp.ActiveSheetWrap.Range("A1").Value = $"版本测试 - {DateTime.Now}";
    }
}
```

#### 3. BlankWorkbook方法 - 创建空白工作簿

```csharp
/// <summary>
/// 创建一个新的空白Excel工作簿
/// </summary>
/// <returns>返回实现了IExcelApplication接口的Excel应用程序实例</returns>
public static IExcelApplication BlankWorkbook()
{
    // 创建ExcelApplication实例，这会启动Excel应用程序进程
    ExcelApplication excel = new ExcelApplication();

    // 调用BlankWorkbook方法创建空白工作簿
    _ = excel.BlankWorkbook();

    // 返回已配置好空白工作簿的Excel应用程序实例
    return excel;
}
```

**应用场景：**
- 从零开始创建报表
- 动态数据生成
- 临时数据处理

**示例代码：**

```csharp
public class DynamicReportGenerator
{
    public void GenerateSalesReport(List<SalesData> salesData)
    {
        using var excelApp = ExcelFactory.BlankWorkbook();
        excelApp.Visible = false; // 后台运行
        
        var worksheet = excelApp.ActiveSheetWrap;
        
        // 设置报表标题
        worksheet.Range("A1").Value = "销售数据报告";
        worksheet.Range("A1").Font.Bold = true;
        worksheet.Range("A1").Font.Size = 16;
        
        // 设置表头
        string[] headers = { "日期", "产品", "数量", "金额", "销售人员" };
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[3, i + 1].Value = headers[i];
            worksheet.Cells[3, i + 1].Font.Bold = true;
        }
        
        // 填充数据
        int row = 4;
        foreach (var data in salesData)
        {
            worksheet.Cells[row, 1].Value = data.Date.ToString("yyyy-MM-dd");
            worksheet.Cells[row, 2].Value = data.ProductName;
            worksheet.Cells[row, 3].Value = data.Quantity;
            worksheet.Cells[row, 4].Value = data.Amount;
            worksheet.Cells[row, 5].Value = data.SalesPerson;
            row++;
        }
        
        // 自动调整列宽
        worksheet.Columns.AutoFit();
        
        // 保存文件
        string fileName = $"SalesReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        excelApp.ActiveWorkbook.SaveAs(fileName);
    }
}

public class SalesData
{
    public DateTime Date { get; set; }
    public string ProductName { get; set; }
    public int Quantity { get; set; }
    public decimal Amount { get; set; }
    public string SalesPerson { get; set; }
}
```

#### 4. CreateFrom方法 - 基于模板创建工作簿

```csharp
/// <summary>
/// 基于指定模板创建新的Excel工作簿
/// </summary>
/// <param name="templatePath">Excel模板文件的完整路径</param>
/// <returns>返回实现了IExcelApplication接口的Excel应用程序实例</returns>
public static IExcelApplication CreateFrom(string templatePath)
{
    // 创建ExcelApplication实例，初始化Excel应用程序环境
    var excel = new ExcelApplication();

    // 调用CreateFrom方法基于模板创建工作簿
    _ = excel.CreateFrom(templatePath);

    // 返回已加载模板内容的Excel应用程序实例
    return excel;
}
```

**应用场景：**
- 标准化报表生成
- 企业文档模板
- 批量文档处理

**示例代码：**

```csharp
public class TemplateBasedReportGenerator
{
    private readonly string _templateDirectory;
    
    public TemplateBasedReportGenerator(string templateDirectory)
    {
        _templateDirectory = templateDirectory;
    }
    
    public void GenerateReport(string templateName, Dictionary<string, object> reportData)
    {
        string templatePath = Path.Combine(_templateDirectory, templateName);
        
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"模板文件不存在: {templatePath}");
        
        using var excelApp = ExcelFactory.CreateFrom(templatePath);
        excelApp.Visible = false;
        
        var worksheet = excelApp.ActiveSheetWrap;
        
        // 填充模板数据
        FillTemplateData(worksheet, reportData);
        
        // 保存生成的报告
        string reportPath = $"C:\\Reports\\SalesReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        excelApp.ActiveWorkbook.SaveAs(reportPath);
    }
    
    private void FillTemplateData(IExcelWorksheet worksheet, Dictionary<string, object> data)
    {
        // 根据模板中的命名区域填充数据
        foreach (var item in data)
        {
            var namedRange = worksheet.Names?[item.Key];
            if (namedRange != null)
            {
                namedRange.Value = item.Value;
            }
        }
    }
}
```

#### 5. Open方法 - 打开现有工作簿文件

```csharp
/// <summary>
/// 打开现有的Excel工作簿文件
/// </summary>
/// <param name="filePath">要打开的Excel文件的完整路径</param>
/// <returns>返回实现了IExcelApplication接口的Excel应用程序实例</returns>
public static IExcelApplication Open(string filePath)
{
    // 创建ExcelApplication实例，准备Excel应用程序运行环境
    var excel = new ExcelApplication();

    // 调用Open方法打开指定路径的Excel文件
    _ = excel.Open(filePath);

    // 返回已加载指定文件的Excel应用程序实例
    return excel;
}
```

**应用场景：**
- 数据分析和处理
- 文件格式转换
- 批量数据更新

**示例代码：**

```csharp
public class ExcelFileProcessor
{
    public void ProcessMultipleFiles(string directoryPath)
    {
        var excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");
        
        foreach (var filePath in excelFiles)
        {
            try
            {
                ProcessSingleFile(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文件失败 {filePath}: {ex.Message}");
            }
        }
    }
    
    private void ProcessSingleFile(string filePath)
    {
        using var excelApp = ExcelFactory.Open(filePath);
        excelApp.Visible = false;
        
        var workbook = excelApp.ActiveWorkbook;
        
        // 处理每个工作表
        foreach (var worksheet in workbook.Worksheets)
        {
            ProcessWorksheet(worksheet);
        }
        
        // 保存处理后的文件
        string newFilePath = filePath.Replace(".xlsx", "_processed.xlsx");
        workbook.SaveAs(newFilePath);
        
        Console.WriteLine($"文件处理完成: {newFilePath}");
    }
    
    private void ProcessWorksheet(IExcelWorksheet worksheet)
    {
        // 示例处理逻辑：清理空行
        var usedRange = worksheet.UsedRange;
        if (usedRange != null)
        {
            for (int row = usedRange.Rows.Count; row >= 1; row--)
            {
                bool isEmpty = true;
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    var cellValue = usedRange.Cells[row, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        isEmpty = false;
                        break;
                    }
                }
                
                if (isEmpty)
                {
                    worksheet.Rows[row].Delete();
                }
            }
        }
    }
}
```

## 应用程序配置与管理

### 应用程序属性设置

IExcelApplication接口提供了丰富的属性来配置Excel应用程序的行为：

```csharp
public class ExcelAppConfigurator
{
    public void ConfigureApplication(IExcelApplication excelApp)
    {
        // 显示设置
        excelApp.Visible = true; // 是否显示Excel窗口
        excelApp.DisplayAlerts = false; // 是否显示警告对话框
        excelApp.ScreenUpdating = false; // 是否更新屏幕显示（提高性能）
        
        // 计算设置
        excelApp.Calculation = XlCalculation.xlCalculationManual; // 手动计算模式
        
        // 用户界面设置
        excelApp.ShowDevTools = true; // 显示开发工具
        excelApp.EnableLivePreview = true; // 启用实时预览
        excelApp.ShowSelectionFloaties = false; // 禁用浮动工具栏
        
        // 安全设置
        excelApp.IgnoreRemoteRequests = true; // 忽略远程请求
        excelApp.EnableCancelKey = XlEnableCancelKey.xlErrorHandler; // 取消键处理
    }
    
    public void OptimizeForBatchProcessing(IExcelApplication excelApp)
    {
        // 批量处理优化配置
        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        excelApp.ScreenUpdating = false;
        excelApp.EnableEvents = false; // 禁用事件以提高性能
        excelApp.Calculation = XlCalculation.xlCalculationManual;
    }
    
    public void RestoreDefaultSettings(IExcelApplication excelApp)
    {
        // 恢复默认设置
        excelApp.ScreenUpdating = true;
        excelApp.DisplayAlerts = true;
        excelApp.EnableEvents = true;
        excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
        excelApp.Calculate(); // 触发计算
    }
}
```

### 多实例管理

在实际应用中，经常需要同时管理多个Excel实例：

```csharp
public class MultiInstanceExcelManager : IDisposable
{
    private readonly List<IExcelApplication> _instances = new();
    
    public IExcelApplication CreateInstance(ExcelInstanceType type, string? parameter = null)
    {
        IExcelApplication instance = type switch
        {
            ExcelInstanceType.BlankWorkbook => ExcelFactory.BlankWorkbook(),
            ExcelInstanceType.FromTemplate => ExcelFactory.CreateFrom(parameter ?? throw new ArgumentNullException(nameof(parameter))),
            ExcelInstanceType.OpenFile => ExcelFactory.Open(parameter ?? throw new ArgumentNullException(nameof(parameter))),
            _ => throw new ArgumentException("不支持的实例类型")
        };
        
        _instances.Add(instance);
        return instance;
    }
    
    public void CloseAllInstances()
    {
        foreach (var instance in _instances)
        {
            try
            {
                instance.Quit();
                ((IDisposable)instance).Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"关闭实例失败: {ex.Message}");
            }
        }
        _instances.Clear();
    }
    
    public void Dispose()
    {
        CloseAllInstances();
    }
}

public enum ExcelInstanceType
{
    BlankWorkbook,
    FromTemplate,
    OpenFile
}
```

## 应用程序生命周期管理

### 资源管理最佳实践

正确的资源管理对于Excel自动化应用至关重要。以下是几种推荐的模式：

#### 模式1：使用using语句（推荐）

```csharp
public void ProcessWithUsing()
{
    using var excelApp = ExcelFactory.BlankWorkbook();
    
    // 执行操作
    excelApp.ActiveSheetWrap.Range("A1").Value = "Hello World";
    
    // using语句会自动调用Dispose()
}
```

#### 模式2：显式资源管理

```csharp
public void ProcessWithExplicitManagement()
{
    IExcelApplication excelApp = null;
    try
    {
        excelApp = ExcelFactory.BlankWorkbook();
        
        // 执行操作
        excelApp.ActiveSheetWrap.Range("A1").Value = "Hello World";
    }
    finally
    {
        if (excelApp != null)
        {
            try
            {
                excelApp.Quit();
                ((IDisposable)excelApp).Dispose();
            }
            catch (Exception ex)
            {
                // 记录日志但不抛出异常
                Console.WriteLine($"资源清理异常: {ex.Message}");
            }
        }
    }
}
```

#### 模式3：安全操作包装器

```csharp
public static class ExcelOperationWrapper
{
    public static T ExecuteSafe<T>(Func<IExcelApplication, T> operation)
    {
        IExcelApplication excelApp = null;
        try
        {
            excelApp = ExcelFactory.BlankWorkbook();
            return operation(excelApp);
        }
        finally
        {
            SafeCleanup(excelApp);
        }
    }
    
    private static void SafeCleanup(IExcelApplication excelApp)
    {
        if (excelApp == null) return;
        
        try
        {
            // 先尝试正常退出
            excelApp.Quit();
        }
        catch
        {
            // 忽略退出异常
        }
        finally
        {
            try
            {
                ((IDisposable)excelApp).Dispose();
            }
            catch
            {
                // 忽略释放异常
            }
        }
    }
}

// 使用示例
public void SafeOperationExample()
{
    ExcelOperationWrapper.ExecuteSafe(excelApp =>
    {
        // 安全执行操作
        excelApp.ActiveSheetWrap.Range("A1").Value = "安全操作示例";
        excelApp.ActiveWorkbook.SaveAs(@"C:\temp\safe_example.xlsx");
        return true;
    });
}
```

### 异常处理策略

完善的异常处理是保证应用稳定性的关键：

```csharp
public class RobustExcelProcessor
{
    public bool TryProcessWithRetry(Action<IExcelApplication> operation, int maxRetries = 3)
    {
        int retryCount = 0;
        
        while (retryCount < maxRetries)
        {
            IExcelApplication excelApp = null;
            try
            {
                excelApp = ExcelFactory.BlankWorkbook();
                operation(excelApp);
                return true;
            }
            catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010001))
            {
                // 调用被拒绝错误，可能是Excel忙
                retryCount++;
                Thread.Sleep(1000 * retryCount); // 指数退避
            }
            catch (FileNotFoundException)
            {
                // 文件不存在错误
                Console.WriteLine("指定的文件不存在");
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                // 权限不足错误
                Console.WriteLine("没有访问文件的权限");
                return false;
            }
            catch (Exception ex)
            {
                // 其他异常
                LogException(ex);
                return false;
            }
            finally
            {
                SafeCleanup(excelApp);
            }
        }
        
        Console.WriteLine($"操作失败，已达到最大重试次数: {maxRetries}");
        return false;
    }
    
    private void LogException(Exception ex)
    {
        // 记录异常信息
        Console.WriteLine($"异常类型: {ex.GetType().Name}");
        Console.WriteLine($"异常消息: {ex.Message}");
        Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
    }
    
    private void SafeCleanup(IExcelApplication excelApp)
    {
        if (excelApp == null) return;
        
        try
        {
            excelApp.Quit();
            ((IDisposable)excelApp).Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"清理过程中发生异常: {ex.Message}");
        }
    }
}
```

## 性能优化技巧

### 批量操作优化

```csharp
public class PerformanceOptimizer
{
    public void OptimizedDataInsertion(IExcelWorksheet worksheet, List<object[]> data)
    {
        // 禁用屏幕更新和自动计算
        var excelApp = worksheet.Application;
        excelApp.ScreenUpdating = false;
        excelApp.Calculation = XlCalculation.xlCalculationManual;
        
        try
        {
            // 批量写入数据
            int batchSize = 1000; // 每批处理1000行
            for (int batchStart = 0; batchStart < data.Count; batchStart += batchSize)
            {
                int endIndex = Math.Min(batchStart + batchSize, data.Count);
                var batchData = data.Skip(batchStart).Take(batchSize).ToArray();
                
                // 转换为二维数组进行批量写入
                var valueArray = ConvertTo2DArray(batchData);
                
                // 批量写入
                var targetRange = worksheet.Range($"A{batchStart + 1}:{GetColumnName(batchData[0].Length)}{endIndex}");
                targetRange.Value = valueArray;
            }
        }
        finally
        {
            // 恢复设置
            excelApp.ScreenUpdating = true;
            excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            excelApp.Calculate();
        }
    }
    
    private object[,] ConvertTo2DArray(object[][] data)
    {
        int rows = data.Length;
        int cols = data[0].Length;
        var result = new object[rows, cols];
        
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                result[i, j] = data[i][j];
            }
        }
        
        return result;
    }
    
    private string GetColumnName(int columnNumber)
    {
        // 将列号转换为列名（如1->A, 2->B, 27->AA等）
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }
}
```

### 内存管理优化

```csharp
public class MemoryOptimizedExcelProcessor
{
    public void ProcessLargeFile(string filePath)
    {
        using var excelApp = ExcelFactory.Open(filePath);
        excelApp.Visible = false;
        
        var workbook = excelApp.ActiveWorkbook;
        
        // 处理每个工作表
        foreach (var worksheet in workbook.Worksheets)
        {
            ProcessWorksheetWithMemoryOptimization(worksheet);
        }
        
        // 保存处理结果
        workbook.Save();
    }
    
    private void ProcessWorksheetWithMemoryOptimization(IExcelWorksheet worksheet)
    {
        var usedRange = worksheet.UsedRange;
        if (usedRange == null) return;
        
        int rowCount = usedRange.Rows.Count;
        int colCount = usedRange.Columns.Count;
        
        // 分块处理，避免一次性加载过多数据
        int chunkSize = 1000; // 每块处理1000行
        
        for (int startRow = 1; startRow <= rowCount; startRow += chunkSize)
        {
            int endRow = Math.Min(startRow + chunkSize - 1, rowCount);
            
            // 处理当前块
            ProcessDataChunk(worksheet, startRow, endRow, colCount);
            
            // 强制垃圾回收（谨慎使用）
            if (startRow % 5000 == 0)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
    
    private void ProcessDataChunk(IExcelWorksheet worksheet, int startRow, int endRow, int colCount)
    {
        // 读取当前块的数据
        var chunkRange = worksheet.Range(worksheet.Cells[startRow, 1], worksheet.Cells[endRow, colCount]);
        var chunkData = chunkRange.Value as object[,];
        
        if (chunkData != null)
        {
            // 处理数据
            for (int i = 1; i <= chunkData.GetLength(0); i++)
            {
                for (int j = 1; j <= chunkData.GetLength(1); j++)
                {
                    // 数据处理逻辑
                    ProcessCellData(chunkData[i, j]);
                }
            }
            
            // 写回处理后的数据
            chunkRange.Value = chunkData;
        }
    }
    
    private void ProcessCellData(object cellValue)
    {
        // 单元格数据处理逻辑
        // 例如：数据清洗、格式转换等
    }
}
```

## 实际应用场景

### 场景1：多实例报表生成系统

```csharp
public class MultiInstanceReportSystem
{
    public void GenerateDailyReports()
    {
        var reportDate = DateTime.Today;
        
        // 同时生成多个报表
        var tasks = new List<Task>
        {
            Task.Run(() => GenerateSalesReport(reportDate)),
            Task.Run(() => GenerateInventoryReport(reportDate)),
            Task.Run(() => GenerateFinancialReport(reportDate))
        };
        
        Task.WaitAll(tasks.ToArray());
    }
    
    private void GenerateSalesReport(DateTime reportDate)
    {
        using var excelApp = ExcelFactory.BlankWorkbook();
        excelApp.Visible = false;
        
        // 销售报表生成逻辑
        var worksheet = excelApp.ActiveSheetWrap;
        worksheet.Range("A1").Value = $"销售日报 - {reportDate:yyyy-MM-dd}";
        
        // ... 更多报表生成逻辑
        
        string fileName = $"SalesReport_{reportDate:yyyyMMdd}.xlsx";
        excelApp.ActiveWorkbook.SaveAs(fileName);
    }
    
    private void GenerateInventoryReport(DateTime reportDate)
    {
        using var excelApp = ExcelFactory.CreateFrom(@"C:\Templates\InventoryTemplate.xltx");
        excelApp.Visible = false;
        
        // 库存报表生成逻辑
        // ...
        
        string fileName = $"InventoryReport_{reportDate:yyyyMMdd}.xlsx";
        excelApp.ActiveWorkbook.SaveAs(fileName);
    }
    
    private void GenerateFinancialReport(DateTime reportDate)
    {
        using var excelApp = ExcelFactory.Open(@"C:\Templates\FinancialTemplate.xlsx");
        excelApp.Visible = false;
        
        // 财务报表生成逻辑
        // ...
        
        string fileName = $"FinancialReport_{reportDate:yyyyMMdd}.xlsx";
        excelApp.ActiveWorkbook.SaveAs(fileName);
    }
}
```

### 场景2：模板应用系统

```csharp
public class TemplateApplicationSystem
{
    private readonly Dictionary<string, string> _templates = new()
    {
        ["Sales"] = @"C:\Templates\SalesReport.xltx",
        ["Inventory"] = @"C:\Templates\InventoryReport.xltx",
        ["Financial"] = @"C:\Templates\FinancialStatement.xltx"
    };
    
    public void ApplyTemplate(string templateType, Dictionary<string, object> data)
    {
        if (!_templates.ContainsKey(templateType))
            throw new ArgumentException($"不支持的模板类型: {templateType}");
        
        string templatePath = _templates[templateType];
        
        using var excelApp = ExcelFactory.CreateFrom(templatePath);
        excelApp.Visible = false;
        
        var worksheet = excelApp.ActiveSheetWrap;
        
        // 应用模板数据
        ApplyTemplateData(worksheet, data);
        
        // 生成报告文件
        string outputPath = $"C:\Reports\\{templateType}Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        excelApp.ActiveWorkbook.SaveAs(outputPath);
    }
    
    private void ApplyTemplateData(IExcelWorksheet worksheet, Dictionary<string, object> data)
    {
        // 根据模板中的命名区域填充数据
        var names = worksheet.Names;
        if (names != null)
        {
            foreach (var name in names)
            {
                if (data.ContainsKey(name.Name))
                {
                    name.Value = data[name.Name];
                }
            }
        }
        
        // 处理动态数据区域
        ProcessDynamicDataRegions(worksheet, data);
    }
    
    private void ProcessDynamicDataRegions(IExcelWorksheet worksheet, Dictionary<string, object> data)
    {
        // 处理动态数据区域（如表格数据）
        if (data.ContainsKey("TableData") && data["TableData"] is object[,] tableData)
        {
            // 找到表格起始位置
            var tableStart = worksheet.Range("A5");
            var tableRange = worksheet.Range(tableStart, worksheet.Cells[tableStart.Row + tableData.GetLength(0) - 1, tableStart.Column + tableData.GetLength(1) - 1]);
            tableRange.Value = tableData;
        }
    }
}
```

## 总结

通过本文的学习，我们深入探讨了Excel应用程序的创建与管理，掌握了ExcelFactory类的各种用法和最佳实践。

**关键知识点回顾：**

1. **ExcelFactory方法**：Connection、CreateInstance、BlankWorkbook、CreateFrom、Open
2. **应用程序配置**：显示设置、计算模式、用户界面配置
3. **资源管理**：using语句、显式管理、安全包装器
4. **异常处理**：重试机制、特定异常处理、日志记录
5. **性能优化**：批量操作、内存管理、屏幕更新控制

**实际应用价值：**

- 多实例报表生成系统可以显著提高处理效率
- 模板应用系统确保文档格式的统一性
- 完善的异常处理机制保证应用的稳定性
- 性能优化技巧处理大规模数据时尤为重要

在下一篇文章中，我们将深入探讨工作簿与工作表操作基础，包括工作表的创建、删除、重命名、保护等高级功能。

---

**下一篇预告：**《工作簿与工作表操作基础》将详细介绍工作簿和工作表的各种操作技巧，包括多工作表管理、工作表保护、数据验证等实用功能。