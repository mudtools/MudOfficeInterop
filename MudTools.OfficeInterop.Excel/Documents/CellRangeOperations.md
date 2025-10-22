# 单元格和区域操作详解

## 引言：Excel自动化的"细胞"与"组织"

在前三篇文章中，我们已经搭建了Excel自动化的基础框架，掌握了应用程序、工作簿和工作表的操作。现在，让我们深入到Excel自动化的微观世界——单元格和区域操作！

如果把Excel自动化比作一个完整的生命体，那么单元格就是最基本的"细胞"，而区域则是细胞的"组织"。每一个细胞都承载着数据信息，而组织的协同工作则构成了复杂的功能系统。没有细胞的精确操作，就没有组织的协调运作；没有区域的灵活管理，就没有复杂应用的实现。

想象一下这样的场景：你需要处理一个包含数万行数据的销售报表，需要根据不同的条件筛选数据、计算汇总、应用格式。如果手动操作，这不仅耗时耗力，而且极易出错。但通过自动化技术，你可以像指挥一支训练有素的军队一样，精确控制每一个单元格，高效完成复杂的任务。

本篇将带你探索单元格和区域的奥秘，从基础引用到高级操作，从简单赋值到复杂计算。准备好让你的Excel自动化技能达到细胞级别的精确控制了吗？

## 单元格基础操作

### 单元格引用与访问

在MudTools.OfficeInterop.Excel中，单元格操作主要通过IExcelRange接口实现。让我们先了解基本的单元格引用方式。

```csharp
public class CellReferenceManager
{
    public void DemonstrateCellReferences(IExcelWorksheet worksheet)
    {
        // 方法1：使用Range方法
        var cellA1 = worksheet.Range("A1");
        if (cellA1 != null)
        {
            cellA1.Value = "使用Range方法引用";
        }
        
        // 方法2：使用索引器（字符串地址）
        var cellB1 = worksheet["B1"];
        if (cellB1 != null)
        {
            cellB1.Value = "使用字符串索引器引用";
        }
        
        // 方法3：使用索引器（行列号）
        var cellC1 = worksheet[1, 3]; // 第1行，第3列（C列）
        if (cellC1 != null)
        {
            cellC1.Value = "使用行列号索引器引用";
        }
        
        // 方法4：使用Cells属性
        var cells = worksheet.Cells;
        if (cells != null)
        {
            var cellD1 = cells[1, 4]; // 第1行，第4列（D列）
            if (cellD1 != null)
            {
                cellD1.Value = "使用Cells属性引用";
            }
        }
    }
    
    public void AdvancedCellReferences(IExcelWorksheet worksheet)
    {
        // 相对引用
        var activeCell = worksheet.Application?.ActiveCell;
        if (activeCell != null)
        {
            // 获取相对位置的单元格
            var offsetCell = activeCell.Offset(2, 3); // 向下2行，向右3列
            offsetCell.Value = "相对引用示例";
        }
        
        // 命名区域引用
        var namedRange = worksheet.Names?["DataRange"];
        if (namedRange != null)
        {
            namedRange.Value = "命名区域引用示例";
        }
        
        // 特殊单元格引用
        var usedRange = worksheet.UsedRange;
        if (usedRange != null)
        {
            // 获取已使用区域的最后一个单元格
            var lastCell = usedRange.Cells[usedRange.Rows.Count, usedRange.Columns.Count];
            if (lastCell != null)
            {
                lastCell.Value = "已使用区域最后一个单元格";
            }
        }
    }
}
```

### 单元格数据读写

单元格数据的读写是Excel自动化中最基本的操作。让我们看看各种数据类型的处理方法。

```csharp
public class CellDataManager
{
    public void BasicDataOperations(IExcelWorksheet worksheet)
    {
        // 写入不同类型的数据
        worksheet["A1"].Value = "文本数据";                    // 字符串
        worksheet["A2"].Value = 123.45;                        // 数字
        worksheet["A3"].Value = true;                          // 布尔值
        worksheet["A4"].Value = DateTime.Now;                  // 日期时间
        worksheet["A5"].Value = null;                         // 空值
        
        // 读取数据
        var textValue = worksheet["A1"]?.Value?.ToString();
        var numberValue = worksheet["A2"]?.Value as double?;
        var boolValue = worksheet["A3"]?.Value as bool?;
        var dateValue = worksheet["A4"]?.Value as DateTime?;
        
        Console.WriteLine($"文本: {textValue}");
        Console.WriteLine($"数字: {numberValue}");
        Console.WriteLine($"布尔: {boolValue}");
        Console.WriteLine($"日期: {dateValue}");
    }
    
    public void FormulaOperations(IExcelWorksheet worksheet)
    {
        // 设置公式
        worksheet["B1"].Value = 10;
        worksheet["B2"].Value = 20;
        worksheet["B3"].Formula = "=SUM(B1:B2)";
        
        // 读取公式
        var formula = worksheet["B3"]?.Formula?.ToString();
        Console.WriteLine($"公式: {formula}");
        
        // 设置数组公式
        var range = worksheet.Range("C1:C3");
        if (range != null)
        {
            range.FormulaArray = "={1,2,3}";
        }
        
        // 检查公式类型
        var hasFormula = worksheet["B3"]?.HasFormula ?? false;
        var hasArray = worksheet["C1"]?.HasArray ?? false;
        
        Console.WriteLine($"B3有公式: {hasFormula}");
        Console.WriteLine($"C1有数组公式: {hasArray}");
    }
    
    public void DataTypeConversion(IExcelWorksheet worksheet)
    {
        // 安全的数据类型转换
        var cell = worksheet["D1"];
        if (cell != null)
        {
            // 写入数据
            cell.Value = "123.45";
            
            // 尝试转换为数字
            if (double.TryParse(cell.Value?.ToString(), out double number))
            {
                Console.WriteLine($"成功转换为数字: {number}");
            }
            
            // 使用NumberValue属性
            var numberValue = cell.NumberValue;
            Console.WriteLine($"NumberValue: {numberValue}");
            
            // 处理空值
            cell.Value = null;
            var isNull = cell.Value == null || cell.Value == DBNull.Value;
            Console.WriteLine($"单元格为空: {isNull}");
        }
    }
}
```

## 区域操作详解

### 区域引用与创建

区域是单元格的集合，可以是连续的矩形区域，也可以是不连续的多块区域。

```csharp
public class RangeManager
{
    public void CreateRanges(IExcelWorksheet worksheet)
    {
        // 方法1：使用字符串地址
        var rangeA1B10 = worksheet.Range("A1:B10");
        if (rangeA1B10 != null)
        {
            rangeA1B10.Value = "A1到B10区域";
        }
        
        // 方法2：使用两个单元格引用
        var cellA1 = worksheet["A1"];
        var cellC5 = worksheet["C5"];
        if (cellA1 != null && cellC5 != null)
        {
            var rangeA1C5 = worksheet.Range(cellA1, cellC5);
            if (rangeA1C5 != null)
            {
                rangeA1C5.Value = "A1到C5区域";
            }
        }
        
        // 方法3：使用行列号
        var rangeByNumbers = worksheet.Range(worksheet.Cells[1, 1], worksheet.Cells[10, 3]);
        if (rangeByNumbers != null)
        {
            rangeByNumbers.Value = "使用行列号创建的区域";
        }
        
        // 方法4：整行或整列
        var entireRow = worksheet.Rows?[1]; // 第1行
        var entireColumn = worksheet.Columns?["A"]; // A列
        
        if (entireRow != null)
        {
            entireRow.Value = "整行数据";
        }
        
        if (entireColumn != null)
        {
            entireColumn.Value = "整列数据";
        }
    }
    
    public void AdvancedRangeOperations(IExcelWorksheet worksheet)
    {
        // 获取已使用区域
        var usedRange = worksheet.UsedRange;
        if (usedRange != null)
        {
            Console.WriteLine($"已使用区域: {usedRange.Address}");
            Console.WriteLine($"行数: {usedRange.Rows.Count}");
            Console.WriteLine($"列数: {usedRange.Columns.Count}");
        }
        
        // 获取当前区域（连续数据区域）
        var currentRegion = worksheet["A1"]?.CurrentRegion;
        if (currentRegion != null)
        {
            Console.WriteLine($"当前区域: {currentRegion.Address}");
        }
        
        // 获取特殊单元格
        var specialCells = worksheet.Range("A1:Z100")?.SpecialCells(
            XlCellType.xlCellTypeConstants, 
            XlSpecialCellsValue.xlTextValues
        );
        
        if (specialCells != null)
        {
            Console.WriteLine($"包含文本的特殊单元格数量: {specialCells.Count}");
        }
    }
    
    public void RangeNavigation(IExcelWorksheet worksheet)
    {
        var startCell = worksheet["C3"];
        if (startCell != null)
        {
            // 相对导航
            var rightCell = startCell.Offset(0, 1);  // 向右1列
            var downCell = startCell.Offset(1, 0);   // 向下1行
            var diagonalCell = startCell.Offset(1, 1); // 右下角
            
            // 绝对导航
            var firstCell = worksheet.Cells[1, 1]; // A1单元格
            var lastRow = worksheet.Cells[worksheet.Rows.Count, 1]; // 最后一行的A列
            
            // 设置导航标记
            rightCell.Value = "→";
            downCell.Value = "↓";
            diagonalCell.Value = "↘";
            firstCell.Value = "起点";
        }
    }
}
```

### 区域数据批量操作

批量操作是提高Excel自动化性能的关键技术。

```csharp
public class BatchDataManager
{
    public void BulkDataOperations(IExcelWorksheet worksheet)
    {
        // 批量写入数据（推荐方式）
        var dataRange = worksheet.Range("A1:E5");
        if (dataRange != null)
        {
            // 创建二维数组
            object[,] dataArray = new object[5, 5];
            
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    dataArray[row, col] = $"数据{row + 1}-{col + 1}";
                }
            }
            
            // 一次性写入
            dataRange.Value = dataArray;
            Console.WriteLine("批量数据写入完成");
        }
        
        // 批量读取数据
        var readRange = worksheet.Range("A1:E5");
        if (readRange != null)
        {
            var readData = readRange.Value as object[,];
            
            if (readData != null)
            {
                for (int row = 0; row < readData.GetLength(0); row++)
                {
                    for (int col = 0; col < readData.GetLength(1); col++)
                    {
                        Console.Write($"{readData[row, col]}\t");
                    }
                    Console.WriteLine();
                }
            }
        }
    }
    
    public void PerformanceOptimizedOperations(IExcelWorksheet worksheet)
    {
        var excelApp = worksheet.Application;
        
        // 优化设置
        excelApp.ScreenUpdating = false;
        excelApp.Calculation = XlCalculation.xlCalculationManual;
        excelApp.EnableEvents = false;
        
        try
        {
            // 执行批量操作
            ProcessLargeDataRange(worksheet, "A1:Z1000");
            
            // 手动触发计算
            excelApp.Calculate();
        }
        finally
        {
            // 恢复设置
            excelApp.ScreenUpdating = true;
            excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            excelApp.EnableEvents = true;
        }
    }
    
    private void ProcessLargeDataRange(IExcelWorksheet worksheet, string rangeAddress)
    {
        var largeRange = worksheet.Range(rangeAddress);
        if (largeRange == null) return;
        
        int totalRows = largeRange.Rows.Count;
        int batchSize = 100; // 每批处理100行
        
        for (int startRow = 1; startRow <= totalRows; startRow += batchSize)
        {
            int endRow = Math.Min(startRow + batchSize - 1, totalRows);
            
            // 处理当前批次
            ProcessBatch(worksheet, startRow, endRow, largeRange.Columns.Count);
            
            // 可选：定期垃圾回收
            if (startRow % 500 == 0)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
    
    private void ProcessBatch(IExcelWorksheet worksheet, int startRow, int endRow, int columnCount)
    {
        // 创建批次数据
        object[,] batchData = new object[endRow - startRow + 1, columnCount];
        
        for (int row = 0; row < batchData.GetLength(0); row++)
        {
            for (int col = 0; col < batchData.GetLength(1); col++)
            {
                batchData[row, col] = $"批次{startRow}-数据{row + 1}-{col + 1}";
            }
        }
        
        // 写入批次数据
        var batchRange = worksheet.Range(
            worksheet.Cells[startRow, 1], 
            worksheet.Cells[endRow, columnCount]
        );
        
        if (batchRange != null)
        {
            batchRange.Value = batchData;
        }
    }
}
```

### 区域格式设置

格式设置是Excel自动化中的重要部分，直接影响数据的可读性和专业性。

```csharp
public class RangeFormatter
{
    public void BasicFormatting(IExcelRange range)
    {
        if (range == null) return;
        
        // 字体格式
        range.Font.Bold = true;
        range.Font.Size = 12;
        range.Font.Color = Color.Blue;
        range.Font.Name = "微软雅黑";
        
        // 单元格填充
        range.Interior.Color = Color.LightYellow;
        range.Interior.Pattern = XlPattern.xlPatternSolid;
        
        // 边框设置
        range.Borders.LineStyle = XlLineStyle.xlContinuous;
        range.Borders.Weight = XlBorderWeight.xlThin;
        range.Borders.Color = Color.Black;
        
        // 对齐方式
        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        range.VerticalAlignment = XlVAlign.xlVAlignCenter;
        
        // 数字格式
        range.NumberFormat = "#,##0.00";
    }
    
    public void ConditionalFormatting(IExcelRange range)
    {
        if (range == null) return;
        
        var formatConditions = range.FormatConditions;
        if (formatConditions != null)
        {
            // 条件1：大于平均值
            var aboveAverage = formatConditions.AddAboveAverage();
            aboveAverage.Interior.Color = Color.LightGreen;
            
            // 条件2：包含特定文本
            var textCondition = formatConditions.Add(
                XlFormatConditionType.xlTextString, 
                XlFormatConditionOperator.xlContains, 
                "重要", 
                ""
            );
            textCondition.Font.Bold = true;
            textCondition.Font.Color = Color.Red;
            
            // 条件3：数据条
            var dataBar = formatConditions.AddDatabar();
            dataBar.BarColor.Color = Color.Blue;
            
            // 条件4：色阶
            var colorScale = formatConditions.AddColorScale(2); // 2色色阶
            colorScale.ColorScaleCriteria[1].Type = XlConditionValueTypes.xlConditionValueLowestValue;
            colorScale.ColorScaleCriteria[1].FormatColor.Color = Color.Green;
            colorScale.ColorScaleCriteria[2].Type = XlConditionValueTypes.xlConditionValueHighestValue;
            colorScale.ColorScaleCriteria[2].FormatColor.Color = Color.Red;
        }
    }
    
    public void AdvancedFormatting(IExcelWorksheet worksheet)
    {
        // 设置整个工作表的默认格式
        var allCells = worksheet.Cells;
        if (allCells != null)
        {
            allCells.Font.Name = "宋体";
            allCells.Font.Size = 11;
            allCells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }
        
        // 设置表头格式
        var headerRange = worksheet.Range("A1:Z1");
        if (headerRange != null)
        {
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Color.LightGray;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }
        
        // 设置数据区域格式
        var dataRange = worksheet.Range("A2:Z1000");
        if (dataRange != null)
        {
            // 交替行颜色
            ApplyAlternatingRowColors(dataRange);
            
            // 设置数字格式
            ApplyNumberFormats(dataRange);
        }
    }
    
    private void ApplyAlternatingRowColors(IExcelRange dataRange)
    {
        // 实际实现需要遍历每一行并设置颜色
        // 这里只是示例逻辑
        for (int row = 1; row <= dataRange.Rows.Count; row++)
        {
            var rowRange = dataRange.Rows[row];
            if (rowRange != null)
            {
                if (row % 2 == 0)
                {
                    rowRange.Interior.Color = Color.White;
                }
                else
                {
                    rowRange.Interior.Color = Color.FromArgb(240, 240, 240); // 浅灰色
                }
            }
        }
    }
    
    private void ApplyNumberFormats(IExcelRange dataRange)
    {
        // 根据列位置设置不同的数字格式
        for (int col = 1; col <= dataRange.Columns.Count; col++)
        {
            var columnRange = dataRange.Columns[col];
            if (columnRange != null)
            {
                switch (col)
                {
                    case 1: // 序号列
                        columnRange.NumberFormat = "0";
                        break;
                    case 2: // 金额列
                        columnRange.NumberFormat = "#,##0.00";
                        break;
                    case 3: // 百分比列
                        columnRange.NumberFormat = "0.00%";
                        break;
                    case 4: // 日期列
                        columnRange.NumberFormat = "yyyy-mm-dd";
                        break;
                    default:
                        columnRange.NumberFormat = "@"; // 文本格式
                        break;
                }
            }
        }
    }
}
```

## 数据验证与保护

### 数据验证规则

数据验证确保输入数据的正确性和一致性。

```csharp
public class DataValidationManager
{
    public void ApplyDataValidations(IExcelWorksheet worksheet)
    {
        // 整数范围验证
        var intRange = worksheet.Range("A1:A10");
        if (intRange?.Validation != null)
        {
            intRange.Validation.Add(
                XlDVType.xlValidateWholeNumber,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlBetween,
                "1", "100"
            );
            intRange.Validation.InputTitle = "输入整数";
            intRange.Validation.InputMessage = "请输入1到100之间的整数";
            intRange.Validation.ErrorMessage = "输入值必须在1到100之间";
        }
        
        // 列表验证
        var listRange = worksheet.Range("B1:B10");
        if (listRange?.Validation != null)
        {
            listRange.Validation.Add(
                XlDVType.xlValidateList,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlEqual,
                "苹果,香蕉,橙子,葡萄",
                ""
            );
            listRange.Validation.InCellDropdown = true;
        }
        
        // 日期验证
        var dateRange = worksheet.Range("C1:C10");
        if (dateRange?.Validation != null)
        {
            dateRange.Validation.Add(
                XlDVType.xlValidateDate,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlBetween,
                DateTime.Today.ToString("yyyy-mm-dd"),
                DateTime.Today.AddMonths(1).ToString("yyyy-mm-dd")
            );
        }
        
        // 自定义公式验证
        var formulaRange = worksheet.Range("D1:D10");
        if (formulaRange?.Validation != null)
        {
            formulaRange.Validation.Add(
                XlDVType.xlValidateCustom,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlEqual,
                "=AND(ISNUMBER(D1), D1>0)",
                ""
            );
            formulaRange.Validation.ErrorMessage = "必须输入大于0的数字";
        }
    }
    
    public void ValidateData(IExcelWorksheet worksheet)
    {
        var validationRange = worksheet.Range("A1:D10");
        if (validationRange?.Validation != null)
        {
            // 检查验证状态
            bool isValid = validationRange.Validation.Value;
            Console.WriteLine($"数据验证状态: {isValid}");
            
            // 获取无效数据
            var invalidData = validationRange.Validation.InvalidData;
            if (invalidData != null)
            {
                Console.WriteLine($"发现 {invalidData.Count} 个无效数据");
                
                foreach (var cell in invalidData)
                {
                    Console.WriteLine($"无效单元格 {cell.Address}: {cell.Value}");
                }
            }
            
            // 清除验证规则
            // validationRange.Validation.Delete();
        }
    }
}
```

### 单元格保护

```csharp
public class CellProtectionManager
{
    public void ApplyCellProtection(IExcelWorksheet worksheet)
    {
        // 先取消整个工作表的保护（如果已保护）
        worksheet.Unprotect();
        
        // 设置单元格的锁定状态
        var headerRange = worksheet.Range("A1:Z1");
        if (headerRange != null)
        {
            headerRange.Locked = true; // 锁定表头
        }
        
        var dataRange = worksheet.Range("A2:Z100");
        if (dataRange != null)
        {
            dataRange.Locked = false; // 解锁数据区域（允许编辑）
        }
        
        // 保护工作表（现在只有数据区域可编辑）
        worksheet.Protect("password123", 
            allowFormattingCells: true,
            allowFormattingColumns: true,
            allowFormattingRows: true,
            allowInsertingColumns: false,
            allowInsertingRows: false,
            allowDeletingColumns: false,
            allowDeletingRows: false,
            allowSorting: true,
            allowFiltering: true,
            allowUsingPivotTables: true
        );
    }
    
    public void AdvancedProtection(IExcelWorksheet worksheet)
    {
        // 隐藏公式
        var formulaRange = worksheet.Range("E1:E10");
        if (formulaRange != null)
        {
            formulaRange.Locked = true;
            formulaRange.FormulaHidden = true; // 隐藏公式
        }
        
        // 保护特定单元格
        var sensitiveRange = worksheet.Range("F1:F10");
        if (sensitiveRange != null)
        {
            sensitiveRange.Locked = true;
            // 设置额外的保护选项
        }
        
        // 重新保护工作表
        worksheet.Protect("securepassword", 
            userInterfaceOnly: true, // 仅保护用户界面
            contents: true,
            scenarios: true
        );
    }
}
```

## 实际应用场景

### 场景1：数据录入系统

```csharp
public class DataEntrySystem
{
    public void SetupDataEntrySheet(IExcelWorksheet worksheet)
    {
        // 设置表头
        string[] headers = { "序号", "日期", "产品名称", "数量", "单价", "金额", "备注" };
        
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
            worksheet.Cells[1, i + 1].Font.Bold = true;
            worksheet.Cells[1, i + 1].Interior.Color = Color.LightBlue;
        }
        
        // 设置数据验证
        SetupValidations(worksheet);
        
        // 设置公式
        SetupFormulas(worksheet);
        
        // 设置保护
        SetupProtection(worksheet);
        
        // 自动调整列宽
        worksheet.Columns.AutoFit();
    }
    
    private void SetupValidations(IExcelWorksheet worksheet)
    {
        // 日期列验证
        var dateColumn = worksheet.Range("B2:B100");
        if (dateColumn?.Validation != null)
        {
            dateColumn.Validation.Add(
                XlDVType.xlValidateDate,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlBetween,
                DateTime.Today.AddYears(-1).ToString("yyyy-mm-dd"),
                DateTime.Today.ToString("yyyy-mm-dd")
            );
        }
        
        // 数量列验证（正整数）
        var quantityColumn = worksheet.Range("D2:D100");
        if (quantityColumn?.Validation != null)
        {
            quantityColumn.Validation.Add(
                XlDVType.xlValidateWholeNumber,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlGreaterEqual,
                "1", ""
            );
        }
        
        // 单价列验证（正数）
        var priceColumn = worksheet.Range("E2:E100");
        if (priceColumn?.Validation != null)
        {
            priceColumn.Validation.Add(
                XlDVType.xlValidateDecimal,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlGreaterEqual,
                "0", ""
            );
        }
    }
    
    private void SetupFormulas(IExcelWorksheet worksheet)
    {
        // 金额公式（数量×单价）
        var amountColumn = worksheet.Range("F2:F100");
        if (amountColumn != null)
        {
            amountColumn.Formula = "=D2*E2";
            amountColumn.NumberFormat = "#,##0.00";
        }
        
        // 汇总行
        var summaryRow = worksheet.Range("A101:G101");
        if (summaryRow != null)
        {
            worksheet.Cells[101, 1].Value = "总计";
            worksheet.Cells[101, 4].Formula = "=SUM(D2:D100)";
            worksheet.Cells[101, 6].Formula = "=SUM(F2:F100)";
            
            summaryRow.Font.Bold = true;
            summaryRow.Interior.Color = Color.LightGreen;
        }
    }
    
    private void SetupProtection(IExcelWorksheet worksheet)
    {
        // 解锁可编辑区域
        var editableRange = worksheet.Range("B2:E100,G2:G100");
        if (editableRange != null)
        {
            editableRange.Locked = false;
        }
        
        // 保护工作表
        worksheet.Protect("entry123", 
            allowFormattingCells: true,
            allowSorting: true,
            allowFiltering: true
        );
    }
}
```

### 场景2：批量数据处理工具

```csharp
public class BatchDataProcessor
{
    public void ProcessDataInBatches(IExcelWorksheet worksheet, List<DataRecord> records)
    {
        int batchSize = 1000;
        int totalRecords = records.Count;
        
        for (int batchStart = 0; batchStart < totalRecords; batchStart += batchSize)
        {
            int batchEnd = Math.Min(batchStart + batchSize, totalRecords);
            var batchRecords = records.Skip(batchStart).Take(batchSize).ToList();
            
            ProcessBatch(worksheet, batchRecords, batchStart + 2); // +2 跳过表头
        }
        
        // 更新汇总信息
        UpdateSummary(worksheet, totalRecords);
    }
    
    private void ProcessBatch(IExcelWorksheet worksheet, List<DataRecord> batchRecords, int startRow)
    {
        object[,] batchData = new object[batchRecords.Count, 6]; // 6列数据
        
        for (int i = 0; i < batchRecords.Count; i++)
        {
            var record = batchRecords[i];
            batchData[i, 0] = startRow + i - 1; // 序号
            batchData[i, 1] = record.Date.ToString("yyyy-MM-dd");
            batchData[i, 2] = record.ProductName;
            batchData[i, 3] = record.Quantity;
            batchData[i, 4] = record.UnitPrice;
            batchData[i, 5] = record.Quantity * record.UnitPrice; // 金额
        }
        
        // 写入批次数据
        var batchRange = worksheet.Range(
            worksheet.Cells[startRow, 1],
            worksheet.Cells[startRow + batchRecords.Count - 1, 6]
        );
        
        if (batchRange != null)
        {
            batchRange.Value = batchData;
        }
    }
    
    private void UpdateSummary(IExcelWorksheet worksheet, int totalRecords)
    {
        int summaryRow = totalRecords + 3; // 汇总行位置
        
        worksheet.Cells[summaryRow, 1].Value = "数据统计";
        worksheet.Cells[summaryRow, 1].Font.Bold = true;
        
        worksheet.Cells[summaryRow, 2].Value = $"总记录数: {totalRecords}";
        worksheet.Cells[summaryRow, 3].Formula = $"=SUM(D2:D{totalRecords + 1})"; // 总数量
        worksheet.Cells[summaryRow, 4].Formula = $"=AVERAGE(E2:E{totalRecords + 1})"; // 平均单价
        worksheet.Cells[summaryRow, 5].Formula = $"=SUM(F2:F{totalRecords + 1})"; // 总金额
        
        var summaryRange = worksheet.Range("A" + summaryRow + ":F" + summaryRow);
        if (summaryRange != null)
        {
            summaryRange.Interior.Color = Color.LightYellow;
            summaryRange.Font.Bold = true;
        }
    }
}

public class DataRecord
{
    public DateTime Date { get; set; }
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
```

### 场景3：动态报表生成器

```csharp
public class DynamicReportGenerator
{
    public void GenerateReport(IExcelWorksheet worksheet, ReportTemplate template)
    {
        // 设置报告标题
        SetupReportHeader(worksheet, template);
        
        // 生成数据区域
        GenerateDataArea(worksheet, template);
        
        // 应用格式
        ApplyReportFormatting(worksheet, template);
        
        // 添加图表和分析
        AddChartsAndAnalysis(worksheet, template);
        
        // 保护报告
        ProtectReport(worksheet);
    }
    
    private void SetupReportHeader(IExcelWorksheet worksheet, ReportTemplate template)
    {
        worksheet.Range("A1").Value = template.Title;
        worksheet.Range("A1").Font.Bold = true;
        worksheet.Range("A1").Font.Size = 18;
        
        worksheet.Range("A2").Value = $"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        worksheet.Range("A3").Value = $"报告期间: {template.StartDate:yyyy-MM-dd} 至 {template.EndDate:yyyy-MM-dd}";
        
        // 合并标题单元格
        worksheet.Range("A1:F1").Merge();
        worksheet.Range("A2:F2").Merge();
        worksheet.Range("A3:F3").Merge();
    }
    
    private void GenerateDataArea(IExcelWorksheet worksheet, ReportTemplate template)
    {
        int startRow = 5;
        
        // 设置表头
        string[] headers = template.Columns;
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[startRow, i + 1].Value = headers[i];
            worksheet.Cells[startRow, i + 1].Font.Bold = true;
            worksheet.Cells[startRow, i + 1].Interior.Color = Color.LightBlue;
        }
        
        // 填充数据
        int dataRow = startRow + 1;
        foreach (var dataRow in template.Data)
        {
            for (int col = 0; col < dataRow.Length; col++)
            {
                worksheet.Cells[dataRow, col + 1].Value = dataRow[col];
            }
            dataRow++;
        }
        
        // 添加汇总行
        AddSummaryRow(worksheet, dataRow, headers.Length);
    }
    
    private void AddSummaryRow(IExcelWorksheet worksheet, int row, int columnCount)
    {
        worksheet.Cells[row, 1].Value = "汇总";
        worksheet.Cells[row, 1].Font.Bold = true;
        
        // 为数值列添加汇总公式
        for (int col = 3; col <= columnCount; col++) // 假设从第3列开始是数值列
        {
            worksheet.Cells[row, col].Formula = $"=SUM({GetColumnName(col)}6:{GetColumnName(col)}{row - 1})";
            worksheet.Cells[row, col].Font.Bold = true;
        }
        
        var summaryRange = worksheet.Range("A" + row + ":" + GetColumnName(columnCount) + row);
        if (summaryRange != null)
        {
            summaryRange.Interior.Color = Color.LightGreen;
        }
    }
    
    private string GetColumnName(int columnNumber)
    {
        // 将列号转换为列名（1->A, 2->B, 27->AA等）
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

public class ReportTemplate
{
    public string Title { get; set; } = "";
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public string[] Columns { get; set; } = Array.Empty<string>();
    public List<object[]> Data { get; set; } = new();
}
```

## 总结

通过本文的学习，我们深入掌握了单元格和区域的各种操作技巧，包括：

**单元格操作要点：**
- 多种引用方式（地址、行列号、相对引用）
- 数据类型处理（文本、数字、日期、公式）
- 安全的数据读写和类型转换

**区域操作要点：**
- 区域创建和引用方法
- 批量数据操作（性能优化关键）
- 区域格式设置（字体、边框、填充）
- 条件格式和数据验证

**高级功能：**
- 数据验证规则设置
- 单元格保护和权限控制
- 动态报表生成
- 批量数据处理

**实际应用价值：**
- 数据录入系统确保数据质量
- 批量处理工具提高处理效率
- 动态报表生成器适应业务变化
- 格式设置提升文档专业性

**最佳实践：**
- 使用批量操作提高性能
- 实现完善的数据验证
- 考虑用户体验和界面友好性
- 提供适当的保护和权限控制

在下一篇文章中，我们将深入探讨数据导入导出与转换，这是Excel自动化中数据交换的核心功能。

---

**下一篇预告：**《数据导入导出与转换》将详细介绍Excel数据的导入导出功能，包括数据库集成、文件格式转换、数据清洗等高级功能。