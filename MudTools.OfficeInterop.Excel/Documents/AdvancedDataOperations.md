# Excel自动化开发指南（十三）：高级数据操作详解

## 引言：Excel自动化的"数据整理大师"

在Excel自动化开发中，如果说数据是"原材料"，那么高级数据操作就是将这些原材料整理成精美成品的"工艺大师"！排序、筛选、分组、分类汇总——这些功能就像是给数据装上了"智能整理系统"，能够快速将杂乱的数据转化为有序的信息。

想象一下这样的场景：你有一个包含数万行销售数据的报表，需要按照产品类别、销售地区、时间顺序进行多重排序，然后筛选出特定条件下的数据，再进行分组汇总。如果手动操作，这不仅耗时耗力，而且极易出错。但通过高级数据操作技术，你可以一键完成所有这些复杂的整理任务！

MudTools.OfficeInterop.Excel项目就像是专业的"数据整理大师"，它提供了完整的高级数据操作接口。从简单的单列排序到复杂的多条件筛选，从基础的数据分组到高级的分类汇总，每一个功能都能让你的数据处理效率提升数倍。

本篇将带你探索高级数据操作的奥秘，学习如何通过代码创建智能、高效、精确的数据整理系统。准备好让你的数据从"杂乱无章"变成"井然有序"了吗？

## 1. 数据排序技术

### 1.1 排序基础概念

数据排序是Excel中最常用的数据操作之一，MudTools.OfficeInterop.Excel提供了完整的排序功能支持：

```csharp
// 核心排序接口概览
public interface IExcelSort : IDisposable
{
    IExcelApplication Application { get; }
    IExcelRange? Range { get; }
    XlYesNoGuess Header { get; set; }
    XlSortMethod SortMethod { get; set; }
    IExcelSortFields? SortFields { get; }
    object Parent { get; }
    
    void SetRange(IExcelRange range);
    void Apply();
    bool MatchCase { get; set; }
    XlSortOrientation Orientation { get; set; }
}
```

### 1.2 单列排序实现

```csharp
public class SingleColumnSortManager
{
    /// <summary>
    /// 单列升序排序
    /// </summary>
    public static void SortSingleColumnAscending(IExcelWorksheet worksheet, string rangeAddress, int sortColumn)
    {
        var sortRange = worksheet.Range(rangeAddress);
        var sort = worksheet.Sort;
        
        if (sort != null)
        {
            sort.SetRange(sortRange);
            sort.Header = XlYesNoGuess.xlYes; // 包含标题行
            sort.Orientation = XlSortOrientation.xlSortColumns; // 按列排序
            
            // 清除现有排序字段
            sort.SortFields?.Clear();
            
            // 添加排序字段
            sort.SortFields?.Add(sortRange.Columns[sortColumn], 
                XlSortOn.xlSortOnValues, XlSortOrder.xlAscending);
            
            sort.Apply();
        }
    }
    
    /// <summary>
    /// 单列降序排序
    /// </summary>
    public static void SortSingleColumnDescending(IExcelWorksheet worksheet, string rangeAddress, int sortColumn)
    {
        var sortRange = worksheet.Range(rangeAddress);
        var sort = worksheet.Sort;
        
        if (sort != null)
        {
            sort.SetRange(sortRange);
            sort.Header = XlYesNoGuess.xlYes;
            sort.Orientation = XlSortOrientation.xlSortColumns;
            
            sort.SortFields?.Clear();
            sort.SortFields?.Add(sortRange.Columns[sortColumn], 
                XlSortOn.xlSortOnValues, XlSortOrder.xlDescending);
            
            sort.Apply();
        }
    }
}
```

### 1.3 多列排序技术

```csharp
public class MultiColumnSortManager
{
    /// <summary>
    /// 多列排序 - 销售数据示例
    /// </summary>
    public static void SortSalesDataByRegionAndAmount(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:E100"); // 销售数据范围
        var sort = worksheet.Sort;
        
        if (sort != null)
        {
            sort.SetRange(dataRange);
            sort.Header = XlYesNoGuess.xlYes;
            sort.Orientation = XlSortOrientation.xlSortColumns;
            
            // 清除现有排序字段
            sort.SortFields?.Clear();
            
            // 第一排序字段：区域（升序）
            sort.SortFields?.Add(dataRange.Columns[2], // 区域列
                XlSortOn.xlSortOnValues, XlSortOrder.xlAscending);
            
            // 第二排序字段：销售额（降序）
            sort.SortFields?.Add(dataRange.Columns[4], // 销售额列
                XlSortOn.xlSortOnValues, XlSortOrder.xlDescending);
            
            sort.Apply();
        }
    }
    
    /// <summary>
    /// 自定义排序规则
    /// </summary>
    public static void SortWithCustomOrder(IExcelWorksheet worksheet, string[] customOrder)
    {
        var dataRange = worksheet.Range("A1:C50");
        var sort = worksheet.Sort;
        
        if (sort != null)
        {
            sort.SetRange(dataRange);
            sort.Header = XlYesNoGuess.xlYes;
            
            // 使用自定义排序列表
            sort.SortFields?.Clear();
            sort.SortFields?.Add(dataRange.Columns[1], // 状态列
                XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, 
                XlSortDataOption.xlSortNormal, customOrder);
            
            sort.Apply();
        }
    }
}
```

### 1.4 排序字段管理

```csharp
public class SortFieldManager
{
    /// <summary>
    /// 排序字段接口概览
    /// </summary>
    public interface IExcelSortField : IDisposable
    {
        object Key { get; set; }
        XlSortOn SortOn { get; set; }
        XlSortOrder Order { get; set; }
        XlSortDataOption DataOption { get; set; }
        object CustomOrder { get; set; }
        int Priority { get; set; }
        object Parent { get; }
        
        void Delete();
        void Modify(XlSortOn sortOn, XlSortOrder order, 
            XlSortDataOption dataOption, object customOrder);
    }
    
    /// <summary>
    /// 排序字段集合管理
    /// </summary>
    public static void ManageSortFields(IExcelSort sort)
    {
        var sortFields = sort.SortFields;
        
        if (sortFields != null)
        {
            // 添加多个排序字段
            for (int i = 0; i < 3; i++)
            {
                sortFields.Add(sort.Range?.Columns[i], 
                    XlSortOn.xlSortOnValues, XlSortOrder.xlAscending);
            }
            
            // 修改排序字段属性
            if (sortFields.Count > 0)
            {
                var firstField = sortFields[1]; // 索引从1开始
                firstField.Order = XlSortOrder.xlDescending;
                firstField.DataOption = XlSortDataOption.xlSortTextAsNumbers;
            }
            
            // 删除排序字段
            if (sortFields.Count > 2)
            {
                sortFields[3].Delete(); // 删除第三个字段
            }
            
            // 清除所有排序字段
            sortFields.Clear();
        }
    }
}
```

## 2. 数据筛选技术

### 2.1 自动筛选功能

```csharp
public class AutoFilterManager
{
    /// <summary>
    /// 应用自动筛选
    /// </summary>
    public static void ApplyAutoFilter(IExcelWorksheet worksheet, string rangeAddress)
    {
        var filterRange = worksheet.Range(rangeAddress);
        
        // 应用自动筛选
        filterRange.AutoFilter();
        
        // 或者使用工作表级别的筛选
        worksheet.AutoFilterMode = false; // 先关闭
        filterRange.AutoFilter(1, "*销售*", XlAutoFilterOperator.xlAnd); // 筛选包含"销售"的文本
    }
    
    /// <summary>
    /// 多条件筛选
    /// </summary>
    public static void ApplyMultiConditionFilter(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:E100");
        
        // 应用自动筛选
        dataRange.AutoFilter();
        
        // 第一列筛选：文本包含"北京"
        dataRange.AutoFilter(1, "*北京*");
        
        // 第二列筛选：数值大于10000
        dataRange.AutoFilter(2, "10000", XlAutoFilterOperator.xlGreater);
        
        // 第三列筛选：日期在指定范围内
        dataRange.AutoFilter(3, ">=" + DateTime.Today.AddDays(-30).ToString("yyyy-MM-dd"), 
            XlAutoFilterOperator.xlAnd, 
            "<=" + DateTime.Today.ToString("yyyy-MM-dd"));
    }
    
    /// <summary>
    /// 清除筛选
    /// </summary>
    public static void ClearAutoFilter(IExcelWorksheet worksheet)
    {
        if (worksheet.AutoFilterMode)
        {
            worksheet.AutoFilterMode = false;
        }
        
        // 或者使用范围清除
        var usedRange = worksheet.UsedRange;
        if (usedRange != null)
        {
            usedRange.AutoFilter(); // 不带参数调用清除筛选
        }
    }
}
```

### 2.2 高级筛选技术

```csharp
public class AdvancedFilterManager
{
    /// <summary>
    /// 应用高级筛选
    /// </summary>
    public static void ApplyAdvancedFilter(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:E100"); // 数据区域
        var criteriaRange = worksheet.Range("G1:H3"); // 条件区域
        var outputRange = worksheet.Range("J1:N1"); // 输出区域
        
        // 设置条件区域
        worksheet["G1"].Value = "区域";
        worksheet["G2"].Value = "北京";
        worksheet["G3"].Value = "上海";
        
        worksheet["H1"].Value = "销售额";
        worksheet["H2"].Value = ">10000";
        
        // 设置输出区域标题
        worksheet["J1"].Value = "区域";
        worksheet["K1"].Value = "销售员";
        worksheet["L1"].Value = "产品";
        worksheet["M1"].Value = "数量";
        worksheet["N1"].Value = "销售额";
        
        // 应用高级筛选
        dataRange.AdvancedFilter(XlFilterAction.xlFilterCopy, 
            criteriaRange, outputRange, false);
    }
    
    /// <summary>
    /// 复杂条件筛选
    /// </summary>
    public static void ApplyComplexCriteriaFilter(IExcelWorksheet worksheet)
    {
        // 设置复杂条件区域
        var criteriaRange = worksheet.Range("G1:K3");
        
        // 第一行：字段标题
        criteriaRange[1, 1].Value = "区域";
        criteriaRange[1, 2].Value = "产品类型";
        criteriaRange[1, 3].Value = "销售季度";
        criteriaRange[1, 4].Value = "销售额";
        criteriaRange[1, 5].Value = "客户等级";
        
        // 第二行：条件1 - 北京地区，电子产品，第一季度，销售额>50000，VIP客户
        criteriaRange[2, 1].Value = "北京";
        criteriaRange[2, 2].Value = "电子产品";
        criteriaRange[2, 3].Value = "Q1";
        criteriaRange[2, 4].Value = ">50000";
        criteriaRange[2, 5].Value = "VIP";
        
        // 第三行：条件2 - 上海地区，办公用品，第二季度，销售额>30000，普通客户
        criteriaRange[3, 1].Value = "上海";
        criteriaRange[3, 2].Value = "办公用品";
        criteriaRange[3, 3].Value = "Q2";
        criteriaRange[3, 4].Value = ">30000";
        criteriaRange[3, 5].Value = "普通";
        
        var dataRange = worksheet.Range("A1:F200");
        var outputRange = worksheet.Range("M1:R1");
        
        // 应用高级筛选
        dataRange.AdvancedFilter(XlFilterAction.xlFilterCopy, 
            criteriaRange, outputRange, false);
    }
}
```

### 2.3 筛选结果处理

```csharp
public class FilterResultManager
{
    /// <summary>
    /// 获取筛选结果
    /// </summary>
    public static List<object[]> GetFilteredData(IExcelWorksheet worksheet)
    {
        var filteredData = new List<object[]>();
        var usedRange = worksheet.UsedRange;
        
        if (usedRange != null && worksheet.AutoFilterMode)
        {
            // 获取可见行
            for (int row = 2; row <= usedRange.Rows.Count; row++) // 从第二行开始（跳过标题）
            {
                var rowRange = usedRange.Rows[row];
                
                // 检查行是否可见（未被筛选隐藏）
                if (rowRange.Hidden == false)
                {
                    var rowData = new List<object>();
                    
                    for (int col = 1; col <= usedRange.Columns.Count; col++)
                    {
                        var cell = usedRange[row, col];
                        rowData.Add(cell?.Value);
                    }
                    
                    filteredData.Add(rowData.ToArray());
                }
            }
        }
        
        return filteredData;
    }
    
    /// <summary>
    /// 统计筛选结果
    /// </summary>
    public static FilterStatistics GetFilterStatistics(IExcelWorksheet worksheet)
    {
        var stats = new FilterStatistics();
        var usedRange = worksheet.UsedRange;
        
        if (usedRange != null)
        {
            stats.TotalRows = usedRange.Rows.Count - 1; // 减去标题行
            stats.VisibleRows = 0;
            
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                var rowRange = usedRange.Rows[row];
                if (rowRange.Hidden == false)
                {
                    stats.VisibleRows++;
                }
            }
            
            stats.HiddenRows = stats.TotalRows - stats.VisibleRows;
        }
        
        return stats;
    }
    
    public class FilterStatistics
    {
        public int TotalRows { get; set; }
        public int VisibleRows { get; set; }
        public int HiddenRows { get; set; }
        public double VisiblePercentage => TotalRows > 0 ? (double)VisibleRows / TotalRows * 100 : 0;
    }
}
```

## 3. 数据分组技术

### 3.1 行分组操作

```csharp
public class RowGroupingManager
{
    /// <summary>
    /// 创建行分组
    /// </summary>
    public static void CreateRowGroups(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:F50");
        
        // 按区域分组销售数据
        var regions = new[] { "北京", "上海", "广州", "深圳" };
        
        foreach (var region in regions)
        {
            // 查找区域开始行
            int startRow = FindRegionStartRow(worksheet, region);
            int endRow = FindRegionEndRow(worksheet, region, startRow);
            
            if (startRow > 0 && endRow > startRow)
            {
                // 创建分组
                var groupRange = worksheet.Range($"A{startRow}:F{endRow}");
                groupRange.Rows.Group();
            }
        }
    }
    
    private static int FindRegionStartRow(IExcelWorksheet worksheet, string region)
    {
        var regionColumn = worksheet.Range("B2:B50"); // 区域列
        
        for (int row = 1; row <= regionColumn.Rows.Count; row++)
        {
            var cell = regionColumn[row, 1];
            if (cell?.Value?.ToString() == region)
            {
                return row + 1; // 转换为实际行号（从1开始）
            }
        }
        
        return -1;
    }
    
    private static int FindRegionEndRow(IExcelWorksheet worksheet, string region, int startRow)
    {
        var regionColumn = worksheet.Range($"B{startRow}:B50");
        
        for (int row = 1; row <= regionColumn.Rows.Count; row++)
        {
            var cell = regionColumn[row, 1];
            if (cell?.Value?.ToString() != region)
            {
                return startRow + row - 2; // 返回前一行
            }
        }
        
        return startRow + regionColumn.Rows.Count - 1;
    }
    
    /// <summary>
    /// 展开/折叠分组
    /// </summary>
    public static void ToggleRowGroups(IExcelWorksheet worksheet, bool expand)
    {
        var outline = worksheet.Outline;
        
        if (outline != null)
        {
            if (expand)
            {
                outline.ShowAllData(); // 展开所有分组
            }
            else
            {
                // 折叠到指定级别
                outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
                
                // 获取分组级别
                int maxLevel = outline.SummaryRowLevel;
                
                for (int level = maxLevel; level >= 1; level--)
                {
                    outline.ShowLevels(level, level); // 显示指定级别
                }
            }
        }
    }
}
```

### 3.2 列分组操作

```csharp
public class ColumnGroupingManager
{
    /// <summary>
    /// 创建列分组
    /// </summary>
    public static void CreateColumnGroups(IExcelWorksheet worksheet)
    {
        // 按季度分组销售数据
        var quarters = new[] 
        {
            ("Q1", 1, 3), // 第一季度：1-3月
            ("Q2", 4, 6), // 第二季度：4-6月
            ("Q3", 7, 9), // 第三季度：7-9月
            ("Q4", 10, 12) // 第四季度：10-12月
        };
        
        foreach (var quarter in quarters)
        {
            int startCol = quarter.Item2 + 1; // 月份从第2列开始（第1列是产品名称）
            int endCol = quarter.Item3 + 1;
            
            var groupRange = worksheet.Range(
                worksheet.Cells[1, startCol], 
                worksheet.Cells[20, endCol]);
            
            groupRange.Columns.Group();
            
            // 添加季度标题
            worksheet[1, startCol - 1].Value = quarter.Item1;
        }
    }
    
    /// <summary>
    /// 多级列分组
    /// </summary>
    public static void CreateMultiLevelColumnGroups(IExcelWorksheet worksheet)
    {
        // 第一级分组：年份
        var years = new[] { "2023", "2024" };
        
        foreach (var year in years)
        {
            int yearStartCol = years[0] == year ? 2 : 14; // 2023从第2列开始，2024从第14列开始
            int yearEndCol = yearStartCol + 11; // 每年12个月
            
            var yearGroup = worksheet.Range(
                worksheet.Cells[1, yearStartCol], 
                worksheet.Cells[20, yearEndCol]);
            
            yearGroup.Columns.Group();
            
            // 第二级分组：季度
            for (int quarter = 0; quarter < 4; quarter++)
            {
                int quarterStartCol = yearStartCol + quarter * 3;
                int quarterEndCol = quarterStartCol + 2;
                
                var quarterGroup = worksheet.Range(
                    worksheet.Cells[1, quarterStartCol], 
                    worksheet.Cells[20, quarterEndCol]);
                
                quarterGroup.Columns.Group();
            }
        }
    }
}
```

### 3.3 分组管理功能

```csharp
public class GroupManagementManager
{
    /// <summary>
    /// 获取分组信息
    /// </summary>
    public static GroupInfo GetGroupInfo(IExcelWorksheet worksheet)
    {
        var info = new GroupInfo();
        var outline = worksheet.Outline;
        
        if (outline != null)
        {
            info.HasRowGroups = outline.SummaryRow != XlSummaryRow.xlSummaryNone;
            info.HasColumnGroups = outline.SummaryColumn != XlSummaryColumn.xlSummaryNone;
            
            info.RowGroupLevels = outline.SummaryRowLevel;
            info.ColumnGroupLevels = outline.SummaryColumnLevel;
            
            info.TotalGroups = CountGroups(worksheet);
        }
        
        return info;
    }
    
    private static int CountGroups(IExcelWorksheet worksheet)
    {
        int count = 0;
        var usedRange = worksheet.UsedRange;
        
        if (usedRange != null)
        {
            // 检查行分组
            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                var rowRange = usedRange.Rows[row];
                if (rowRange?.OutlineLevel > 1) // 非顶级分组
                {
                    count++;
                }
            }
            
            // 检查列分组
            for (int col = 1; col <= usedRange.Columns.Count; col++)
            {
                var colRange = usedRange.Columns[col];
                if (colRange?.OutlineLevel > 1)
                {
                    count++;
                }
            }
        }
        
        return count;
    }
    
    /// <summary>
    /// 清除所有分组
    /// </summary>
    public static void ClearAllGroups(IExcelWorksheet worksheet)
    {
        var outline = worksheet.Outline;
        
        if (outline != null)
        {
            // 清除行分组
            if (outline.SummaryRow != XlSummaryRow.xlSummaryNone)
            {
                var usedRange = worksheet.UsedRange;
                if (usedRange != null)
                {
                    usedRange.Rows.Ungroup();
                }
            }
            
            // 清除列分组
            if (outline.SummaryColumn != XlSummaryColumn.xlSummaryNone)
            {
                var usedRange = worksheet.UsedRange;
                if (usedRange != null)
                {
                    usedRange.Columns.Ungroup();
                }
            }
        }
    }
    
    public class GroupInfo
    {
        public bool HasRowGroups { get; set; }
        public bool HasColumnGroups { get; set; }
        public int RowGroupLevels { get; set; }
        public int ColumnGroupLevels { get; set; }
        public int TotalGroups { get; set; }
    }
}
```

## 4. 分类汇总技术

### 4.1 基础分类汇总

```csharp
public class SubtotalManager
{
    /// <summary>
    /// 创建分类汇总
    /// </summary>
    public static void CreateSubtotals(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:F50");
        
        // 按区域分类汇总销售额
        dataRange.Subtotal(
            groupBy: 2, // 按区域列（第2列）分组
            function: XlConsolidationFunction.xlSum, // 求和函数
            totalList: new[] { 5 }, // 对销售额列（第5列）汇总
            replace: true, // 替换现有汇总
            pageBreaks: false, // 不添加分页符
            summaryBelowData: true // 汇总行在数据下方
        );
    }
    
    /// <summary>
    /// 多级分类汇总
    /// </summary>
    public static void CreateMultiLevelSubtotals(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:F100");
        
        // 第一级：按区域汇总
        dataRange.Subtotal(2, XlConsolidationFunction.xlSum, new[] { 5 }, true, false, true);
        
        // 第二级：按产品类型汇总（在区域分组内）
        dataRange.Subtotal(3, XlConsolidationFunction.xlSum, new[] { 5 }, false, false, true);
        
        // 第三级：按销售季度汇总（在产品类型分组内）
        dataRange.Subtotal(4, XlConsolidationFunction.xlSum, new[] { 5 }, false, false, true);
    }
    
    /// <summary>
    /// 多种汇总函数应用
    /// </summary>
    public static void ApplyMultipleSubtotalFunctions(IExcelWorksheet worksheet)
    {
        var dataRange = worksheet.Range("A1:F80");
        
        // 按区域分组，应用多种汇总函数
        
        // 1. 求和
        dataRange.Subtotal(2, XlConsolidationFunction.xlSum, new[] { 5 }, true, false, true);
        
        // 2. 平均值
        dataRange.Subtotal(2, XlConsolidationFunction.xlAverage, new[] { 5 }, false, false, true);
        
        // 3. 计数
        dataRange.Subtotal(2, XlConsolidationFunction.xlCount, new[] { 5 }, false, false, true);
        
        // 4. 最大值
        dataRange.Subtotal(2, XlConsolidationFunction.xlMax, new[] { 5 }, false, false, true);
        
        // 5. 最小值
        dataRange.Subtotal(2, XlConsolidationFunction.xlMin, new[] { 5 }, false, false, true);
    }
}
```

### 4.2 汇总结果处理

```csharp
public class SubtotalResultManager
{
    /// <summary>
    /// 获取汇总结果
    /// </summary>
    public static List<SubtotalResult> GetSubtotalResults(IExcelWorksheet worksheet)
    {
        var results = new List<SubtotalResult>();
        var usedRange = worksheet.UsedRange;
        
        if (usedRange != null)
        {
            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                var rowRange = usedRange.Rows[row];
                
                // 检查是否为汇总行（通常包含"汇总"或"小计"字样）
                var firstCell = rowRange[1, 1];
                if (firstCell?.Value?.ToString()?.Contains("汇总") == true ||
                    firstCell?.Value?.ToString()?.Contains("小计") == true)
                {
                    var result = new SubtotalResult
                    {
                        RowNumber = row,
                        GroupName = firstCell.Value?.ToString() ?? "",
                        SummaryValues = new Dictionary<int, object>()
                    };
                    
                    // 获取汇总值
                    for (int col = 2; col <= usedRange.Columns.Count; col++)
                    {
                        var cell = rowRange[1, col];
                        if (cell?.Value != null)
                        {
                            result.SummaryValues[col] = cell.Value;
                        }
                    }
                    
                    results.Add(result);
                }
            }
        }
        
        return results;
    }
    
    /// <summary>
    /// 清除分类汇总
    /// </summary>
    public static void RemoveSubtotals(IExcelWorksheet worksheet)
    {
        var usedRange = worksheet.UsedRange;
        
        if (usedRange != null)
        {
            usedRange.RemoveSubtotal();
        }
    }
    
    /// <summary>
    /// 导出汇总数据
    /// </summary>
    public static void ExportSubtotalData(IExcelWorksheet sourceWorksheet, IExcelWorksheet targetWorksheet)
    {
        var subtotalResults = GetSubtotalResults(sourceWorksheet);
        
        // 在目标工作表中创建汇总表
        int targetRow = 1;
        
        // 添加标题
        targetWorksheet["A1"].Value = "分组名称";
        targetWorksheet["B1"].Value = "汇总值1";
        targetWorksheet["C1"].Value = "汇总值2";
        targetWorksheet["D1"].Value = "汇总值3";
        
        targetRow++;
        
        // 添加汇总数据
        foreach (var result in subtotalResults)
        {
            targetWorksheet[$"A{targetRow}"].Value = result.GroupName;
            
            int col = 2;
            foreach (var summaryValue in result.SummaryValues.Values)
            {
                targetWorksheet[targetRow, col].Value = summaryValue;
                col++;
            }
            
            targetRow++;
        }
    }
    
    public class SubtotalResult
    {
        public int RowNumber { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<int, object> SummaryValues { get; set; } = new Dictionary<int, object>();
    }
}
```

## 5. 实际应用案例

### 5.1 销售数据分析系统

```csharp
public class SalesDataAnalysisSystem
{
    /// <summary>
    /// 完整的销售数据分析流程
    /// </summary>
    public static void AnalyzeSalesData(IExcelWorksheet worksheet)
    {
        // 1. 数据排序：按区域和销售额排序
        MultiColumnSortManager.SortSalesDataByRegionAndAmount(worksheet);
        
        // 2. 数据筛选：筛选高销售额记录
        AutoFilterManager.ApplyMultiConditionFilter(worksheet);
        
        // 3. 数据分组：按区域分组
        RowGroupingManager.CreateRowGroups(worksheet);
        
        // 4. 分类汇总：按区域汇总销售额
        SubtotalManager.CreateSubtotals(worksheet);
        
        // 5. 导出汇总结果
        var summaryWorksheet = worksheet.Application.Workbooks[1].Worksheets.Add();
        summaryWorksheet.Name = "销售汇总";
        
        SubtotalResultManager.ExportSubtotalData(worksheet, summaryWorksheet);
    }
    
    /// <summary>
    /// 生成销售分析报告
    /// </summary>
    public static void GenerateSalesReport(IExcelWorksheet dataWorksheet)
    {
        // 创建报告工作表
        var reportWorksheet = dataWorksheet.Application.Workbooks[1].Worksheets.Add();
        reportWorksheet.Name = "销售分析报告";
        
        // 添加报告标题
        reportWorksheet["A1"].Value = "销售数据分析报告";
        reportWorksheet["A1"].Font.Bold = true;
        reportWorksheet["A1"].Font.Size = 16;
        
        // 获取筛选统计数据
        var filterStats = FilterResultManager.GetFilterStatistics(dataWorksheet);
        
        // 添加统计信息
        reportWorksheet["A3"].Value = "数据统计:";
        reportWorksheet["A4"].Value = $"总记录数: {filterStats.TotalRows}";
        reportWorksheet["A5"].Value = $"可见记录数: {filterStats.VisibleRows}";
        reportWorksheet["A6"].Value = $"隐藏记录数: {filterStats.HiddenRows}";
        reportWorksheet["A7"].Value = $"可见比例: {filterStats.VisiblePercentage:F2}%";
        
        // 获取分组信息
        var groupInfo = GroupManagementManager.GetGroupInfo(dataWorksheet);
        
        reportWorksheet["C3"].Value = "分组信息:";
        reportWorksheet["C4"].Value = $"行分组级别: {groupInfo.RowGroupLevels}";
        reportWorksheet["C5"].Value = $"列分组级别: {groupInfo.ColumnGroupLevels}";
        reportWorksheet["C6"].Value = $"总分组数: {groupInfo.TotalGroups}";
        
        // 获取汇总结果
        var subtotalResults = SubtotalResultManager.GetSubtotalResults(dataWorksheet);
        
        reportWorksheet["A9"].Value = "汇总结果:";
        
        int reportRow = 10;
        foreach (var result in subtotalResults)
        {
            reportWorksheet[$"A{reportRow}"].Value = result.GroupName;
            reportRow++;
        }
    }
}
```

### 5.2 财务报表处理系统

```csharp
public class FinancialReportProcessor
{
    /// <summary>
    /// 财务报表数据处理
    /// </summary>
    public static void ProcessFinancialData(IExcelWorksheet worksheet)
    {
        // 1. 按科目分类排序
        SingleColumnSortManager.SortSingleColumnAscending(worksheet, "A2:F100", 1); // 按科目排序
        
        // 2. 应用高级筛选：筛选特定期间的数据
        var criteriaRange = worksheet.Range("H1:I2");
        criteriaRange[1, 1].Value = "期间";
        criteriaRange[2, 1].Value = ">=2023-01-01";
        criteriaRange[1, 2].Value = "金额";
        criteriaRange[2, 2].Value = ">1000";
        
        var dataRange = worksheet.Range("A2:F100");
        var outputRange = worksheet.Range("K2:P2");
        
        dataRange.AdvancedFilter(XlFilterAction.xlFilterCopy, criteriaRange, outputRange, false);
        
        // 3. 按科目分组
        RowGroupingManager.CreateRowGroups(worksheet);
        
        // 4. 创建多级分类汇总
        SubtotalManager.CreateMultiLevelSubtotals(worksheet);
    }
    
    /// <summary>
    /// 生成财务报表
    /// </summary>
    public static void GenerateFinancialStatement(IExcelWorksheet dataWorksheet)
    {
        // 创建财务报表
        var statementWorksheet = dataWorksheet.Application.Workbooks[1].Worksheets.Add();
        statementWorksheet.Name = "财务报表";
        
        // 添加报表结构
        statementWorksheet["A1"].Value = "损益表";
        statementWorksheet["A1"].Font.Bold = true;
        statementWorksheet["A1"].Font.Size = 14;
        
        // 添加报表项目
        string[] items = { "营业收入", "营业成本", "毛利润", "营业费用", "营业利润", "净利润" };
        
        for (int i = 0; i < items.Length; i++)
        {
            statementWorksheet[$"A{i + 3}"].Value = items[i];
        }
        
        // 从汇总结果获取数据
        var subtotalResults = SubtotalResultManager.GetSubtotalResults(dataWorksheet);
        
        // 填充报表数据
        int dataRow = 3;
        foreach (var result in subtotalResults.Take(items.Length))
        {
            if (result.SummaryValues.ContainsKey(5)) // 假设第5列是金额列
            {
                statementWorksheet[$"B{dataRow}"].Value = result.SummaryValues[5];
                statementWorksheet[$"B{dataRow}"].NumberFormat = "#,##0.00";
            }
            dataRow++;
        }
    }
}
```

## 6. 性能优化和最佳实践

### 6.1 批量操作优化

```csharp
public class DataOperationOptimizer
{
    /// <summary>
    /// 批量数据操作优化
    /// </summary>
    public static void OptimizeBatchOperations(IExcelWorksheet worksheet)
    {
        // 禁用屏幕更新
        worksheet.Application.ScreenUpdating = false;
        
        try
        {
            // 禁用自动计算
            worksheet.Application.Calculation = XlCalculation.xlCalculationManual;
            
            // 执行批量操作
            PerformBatchSorting(worksheet);
            PerformBatchFiltering(worksheet);
            PerformBatchGrouping(worksheet);
            
            // 重新启用自动计算
            worksheet.Application.Calculation = XlCalculation.xlCalculationAutomatic;
            
            // 强制重新计算
            worksheet.Application.Calculate();
        }
        finally
        {
            // 恢复屏幕更新
            worksheet.Application.ScreenUpdating = true;
        }
    }
    
    private static void PerformBatchSorting(IExcelWorksheet worksheet)
    {
        // 批量排序操作
        var sort = worksheet.Sort;
        if (sort != null)
        {
            sort.SetRange(worksheet.UsedRange);
            sort.SortFields?.Clear();
            
            // 添加多个排序字段
            for (int i = 1; i <= 3; i++)
            {
                sort.SortFields?.Add(worksheet.UsedRange.Columns[i], 
                    XlSortOn.xlSortOnValues, XlSortOrder.xlAscending);
            }
            
            sort.Apply();
        }
    }
    
    private static void PerformBatchFiltering(IExcelWorksheet worksheet)
    {
        // 批量筛选操作
        var usedRange = worksheet.UsedRange;
        if (usedRange != null)
        {
            usedRange.AutoFilter();
            
            // 应用多个筛选条件
            usedRange.AutoFilter(1, "*重要*");
            usedRange.AutoFilter(2, ">1000");
            usedRange.AutoFilter(3, ">=" + DateTime.Today.AddMonths(-1).ToString("yyyy-MM-dd"));
        }
    }
    
    private static void PerformBatchGrouping(IExcelWorksheet worksheet)
    {
        // 批量分组操作
        var usedRange = worksheet.UsedRange;
        if (usedRange != null)
        {
            // 创建行分组
            usedRange.Rows.Group();
            
            // 创建列分组
            usedRange.Columns.Group();
        }
    }
}
```

### 6.2 内存管理优化

```csharp
public class MemoryManagementOptimizer
{
    /// <summary>
    /// 优化内存使用
    /// </summary>
    public static void OptimizeMemoryUsage(IExcelWorksheet worksheet)
    {
        // 释放不需要的对象引用
        var usedRange = worksheet.UsedRange;
        
        if (usedRange != null)
        {
            // 处理完成后及时释放资源
            using (var sort = worksheet.Sort)
            using (var outline = worksheet.Outline)
            {
                // 执行数据操作
                PerformDataOperations(worksheet, usedRange);
                
                // 清理临时对象
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
    
    private static void PerformDataOperations(IExcelWorksheet worksheet, IExcelRange usedRange)
    {
        // 限制处理的数据量
        int maxRows = Math.Min(usedRange.Rows.Count, 1000);
        int maxCols = Math.Min(usedRange.Columns.Count, 50);
        
        var limitedRange = worksheet.Range(
            worksheet.Cells[1, 1], 
            worksheet.Cells[maxRows, maxCols]);
        
        // 在限制范围内执行操作
        limitedRange.Sort(limitedRange.Columns[1], XlSortOrder.xlAscending);
        limitedRange.AutoFilter();
    }
}
```

## 总结

本章详细介绍了MudTools.OfficeInterop.Excel项目中的高级数据操作技术，包括排序、筛选、分组和分类汇总等功能。通过实际的代码示例和完整的应用案例，展示了如何在实际项目中应用这些技术。

**关键技术要点：**

1. **排序技术**：支持单列、多列和自定义排序规则
2. **筛选技术**：提供自动筛选和高级筛选功能
3. **分组技术**：支持行分组和列分组的多级管理
4. **分类汇总**：实现多级汇总和多种汇总函数
5. **性能优化**：批量操作和内存管理优化技术

这些高级数据操作技术为Excel自动化开发提供了强大的数据处理能力，可以广泛应用于各种业务场景，如销售分析、财务报表、数据统计等。通过合理的性能优化和最佳实践，可以确保在大数据量情况下的高效运行。