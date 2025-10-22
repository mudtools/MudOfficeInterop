# 数据导入导出与转换

## 引言：Excel自动化的"数据高速公路"

在前四篇文章中，我们已经掌握了Excel自动化的基础操作，现在让我们进入一个更加激动人心的领域——数据导入导出与转换！这就像是给Excel自动化系统修建了一条"数据高速公路"，让数据能够在不同的系统之间自由流动。

想象一下这样的场景：你的企业有销售数据存储在SQL Server数据库中，有客户信息在CRM系统中，有产品数据在ERP系统里。传统的做法是手动从各个系统导出数据，然后在Excel中复制粘贴、整理格式。这个过程不仅效率低下，而且容易出错。

但是，通过数据导入导出技术，你可以建立一条自动化的数据管道！销售数据自动从数据库流入Excel，客户信息自动从CRM系统同步，产品数据自动从ERP系统获取。所有的数据就像在高速公路上飞驰的车辆，按照预设的路线自动到达目的地。

本篇将带你探索数据交换的奥秘，从基础的数据库连接到高级的Web服务集成，从简单的文件导入到复杂的数据转换。准备好让你的Excel自动化系统成为企业数据的"交通枢纽"了吗？

## 数据导入基础

### 从数据库导入数据

MudTools.OfficeInterop.Excel提供了强大的数据库连接功能，可以直接从各种数据源导入数据。

```csharp
public class DatabaseImporter
{
    public void ImportFromDatabase(IExcelWorkbook workbook, DatabaseConfig config)
    {
        var connections = workbook.Connections;
        if (connections == null) return;
        
        try
        {
            // 创建数据库连接
            var connection = connections.Add(
                name: config.ConnectionName,
                description: $"连接到 {config.DatabaseName}",
                connectionString: BuildConnectionString(config),
                commandText: config.SqlQuery,
                lCmdType: XlCmdType.xlCmdSql
            );
            
            if (connection != null)
            {
                // 创建查询表
                CreateQueryTable(workbook, connection, config);
                
                Console.WriteLine($"成功从数据库 {config.DatabaseName} 导入数据");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"数据库导入失败: {ex.Message}");
        }
    }
    
    private string BuildConnectionString(DatabaseConfig config)
    {
        return config.DatabaseType switch
        {
            DatabaseType.SqlServer => $"Provider=SQLOLEDB;Data Source={config.Server};Initial Catalog={config.DatabaseName};User ID={config.Username};Password={config.Password}",
            DatabaseType.Oracle => $"Provider=OraOLEDB.Oracle;Data Source={config.Server};User ID={config.Username};Password={config.Password}",
            DatabaseType.MySQL => $"Provider=MySQLProv;Data Source={config.Server};Database={config.DatabaseName};User ID={config.Username};Password={config.Password}",
            _ => throw new NotSupportedException($"不支持的数据库类型: {config.DatabaseType}")
        };
    }
    
    private void CreateQueryTable(IExcelWorkbook workbook, IExcelWorkbookConnection connection, DatabaseConfig config)
    {
        var worksheet = workbook.Worksheets?.Add($"{config.TableName}_数据");
        if (worksheet == null) return;
        
        // 创建查询表
        var queryTable = worksheet.QueryTables?.Add(
            connection: connection,
            destination: worksheet.Range("A1")
        );
        
        if (queryTable != null)
        {
            // 配置查询表
            queryTable.CommandText = config.SqlQuery;
            queryTable.RefreshOnFileOpen = true;
            queryTable.BackgroundQuery = false; // 同步执行
            
            // 刷新数据
            queryTable.Refresh();
            
            // 自动调整列宽
            worksheet.Columns.AutoFit();
        }
    }
}

public class DatabaseConfig
{
    public string ConnectionName { get; set; } = "";
    public DatabaseType DatabaseType { get; set; }
    public string Server { get; set; } = "";
    public string DatabaseName { get; set; } = "";
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string SqlQuery { get; set; } = "";
    public string TableName { get; set; } = "";
}

public enum DatabaseType
{
    SqlServer,
    Oracle,
    MySQL
}
```

### 从DataTable导入数据

DataTable是.NET中最常用的数据容器，Excel提供了直接的数据导入功能。

```csharp
public class DataTableImporter
{
    public void ImportFromDataTable(IExcelWorksheet worksheet, DataTable dataTable, string startCell = "A1")
    {
        var targetRange = worksheet.Range(startCell);
        if (targetRange == null) return;
        
        try
        {
            // 使用CopyFromDataTable方法导入数据
            bool success = targetRange.CopyFromDataTable(dataTable, startCell, true);
            
            if (success)
            {
                Console.WriteLine($"成功导入 {dataTable.Rows.Count} 行数据");
                
                // 应用格式
                ApplyDataTableFormat(worksheet, dataTable, startCell);
            }
            else
            {
                Console.WriteLine("数据导入失败");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DataTable导入失败: {ex.Message}");
        }
    }
    
    private void ApplyDataTableFormat(IExcelWorksheet worksheet, DataTable dataTable, string startCell)
    {
        // 设置表头格式
        var headerRange = worksheet.Range(startCell, worksheet.Cells[1, dataTable.Columns.Count]);
        if (headerRange != null)
        {
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }
        
        // 设置数据区域格式
        if (dataTable.Rows.Count > 0)
        {
            var dataRange = worksheet.Range(
                worksheet.Cells[2, 1],
                worksheet.Cells[dataTable.Rows.Count + 1, dataTable.Columns.Count]
            );
            
            if (dataRange != null)
            {
                // 交替行颜色
                ApplyAlternatingRowColors(dataRange);
                
                // 设置数字格式
                ApplyColumnFormats(worksheet, dataTable);
            }
        }
        
        // 自动调整列宽
        worksheet.Columns.AutoFit();
    }
    
    private void ApplyAlternatingRowColors(IExcelRange dataRange)
    {
        for (int row = 1; row <= dataRange.Rows.Count; row++)
        {
            var rowRange = dataRange.Rows[row];
            if (rowRange != null)
            {
                rowRange.Interior.Color = row % 2 == 0 ? Color.White : Color.FromArgb(248, 248, 248);
            }
        }
    }
    
    private void ApplyColumnFormats(IExcelWorksheet worksheet, DataTable dataTable)
    {
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            var column = dataTable.Columns[col];
            var columnRange = worksheet.Columns[col + 1];
            
            if (columnRange != null)
            {
                // 根据数据类型设置格式
                switch (column.DataType.Name)
                {
                    case "DateTime":
                        columnRange.NumberFormat = "yyyy-mm-dd";
                        break;
                    case "Decimal":
                    case "Double":
                    case "Single":
                        columnRange.NumberFormat = "#,##0.00";
                        break;
                    case "Int32":
                    case "Int64":
                        columnRange.NumberFormat = "#,##0";
                        break;
                    default:
                        columnRange.NumberFormat = "@"; // 文本格式
                        break;
                }
            }
        }
    }
    
    public DataTable CreateSampleDataTable()
    {
        DataTable table = new DataTable("销售数据");
        
        // 添加列
        table.Columns.Add("序号", typeof(int));
        table.Columns.Add("日期", typeof(DateTime));
        table.Columns.Add("产品名称", typeof(string));
        table.Columns.Add("数量", typeof(int));
        table.Columns.Add("单价", typeof(decimal));
        table.Columns.Add("金额", typeof(decimal));
        
        // 添加示例数据
        var random = new Random();
        for (int i = 1; i <= 100; i++)
        {
            int quantity = random.Next(1, 100);
            decimal price = (decimal)(random.NextDouble() * 1000);
            
            table.Rows.Add(
                i,
                DateTime.Today.AddDays(-random.Next(365)),
                $"产品{random.Next(1, 10)}",
                quantity,
                price,
                quantity * price
            );
        }
        
        return table;
    }
}
```

### 从文件导入数据

Excel支持从各种文件格式导入数据，包括CSV、TXT、XML等。

```csharp
public class FileImporter
{
    public void ImportFromCsvFile(IExcelWorksheet worksheet, string csvFilePath, string startCell = "A1")
    {
        if (!File.Exists(csvFilePath))
        {
            Console.WriteLine($"CSV文件不存在: {csvFilePath}");
            return;
        }
        
        try
        {
            // 读取CSV文件
            var csvData = ReadCsvFile(csvFilePath);
            
            // 导入到Excel
            ImportCsvData(worksheet, csvData, startCell);
            
            Console.WriteLine($"成功从CSV文件导入 {csvData.Count} 行数据");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CSV导入失败: {ex.Message}");
        }
    }
    
    private List<string[]> ReadCsvFile(string filePath)
    {
        var data = new List<string[]>();
        
        using var reader = new StreamReader(filePath, Encoding.UTF8);
        string? line;
        
        while ((line = reader.ReadLine()) != null)
        {
            // 简单的CSV解析（实际应用中可能需要更复杂的解析逻辑）
            var fields = line.Split(',');
            data.Add(fields);
        }
        
        return data;
    }
    
    private void ImportCsvData(IExcelWorksheet worksheet, List<string[]> csvData, string startCell)
    {
        var startRange = worksheet.Range(startCell);
        if (startRange == null) return;
        
        int startRow = startRange.Row;
        int startCol = startRange.Column;
        
        // 创建二维数组
        object[,] dataArray = new object[csvData.Count, csvData[0].Length];
        
        for (int row = 0; row < csvData.Count; row++)
        {
            for (int col = 0; col < csvData[row].Length; col++)
            {
                dataArray[row, col] = csvData[row][col];
            }
        }
        
        // 写入数据
        var targetRange = worksheet.Range(
            worksheet.Cells[startRow, startCol],
            worksheet.Cells[startRow + csvData.Count - 1, startCol + csvData[0].Length - 1]
        );
        
        if (targetRange != null)
        {
            targetRange.Value = dataArray;
        }
        
        // 应用格式
        ApplyCsvFormat(worksheet, csvData, startCell);
    }
    
    private void ApplyCsvFormat(IExcelWorksheet worksheet, List<string[]> csvData, string startCell)
    {
        var startRange = worksheet.Range(startCell);
        if (startRange == null) return;
        
        // 设置表头格式
        if (csvData.Count > 0)
        {
            var headerRange = worksheet.Range(
                startRange,
                worksheet.Cells[startRange.Row, startRange.Column + csvData[0].Length - 1]
            );
            
            if (headerRange != null)
            {
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = Color.LightGreen;
            }
        }
        
        // 自动调整列宽
        worksheet.Columns.AutoFit();
    }
    
    public void ImportFromJsonFile(IExcelWorksheet worksheet, string jsonFilePath, string startCell = "A1")
    {
        if (!File.Exists(jsonFilePath))
        {
            Console.WriteLine($"JSON文件不存在: {jsonFilePath}");
            return;
        }
        
        try
        {
            // 读取JSON文件
            var jsonData = File.ReadAllText(jsonFilePath);
            
            // 解析JSON（需要Newtonsoft.Json或System.Text.Json）
            var data = ParseJsonData(jsonData);
            
            // 导入到Excel
            ImportJsonData(worksheet, data, startCell);
            
            Console.WriteLine("成功从JSON文件导入数据");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"JSON导入失败: {ex.Message}");
        }
    }
    
    private List<Dictionary<string, object>> ParseJsonData(string jsonData)
    {
        // 简化的JSON解析（实际应用中需要使用JSON库）
        // 这里返回模拟数据
        return new List<Dictionary<string, object>>
        {
            new Dictionary<string, object>
            {
                ["ID"] = 1,
                ["Name"] = "产品A",
                ["Price"] = 100.50,
                ["Category"] = "电子产品"
            },
            new Dictionary<string, object>
            {
                ["ID"] = 2,
                ["Name"] = "产品B", 
                ["Price"] = 200.75,
                ["Category"] = "家居用品"
            }
        };
    }
    
    private void ImportJsonData(IExcelWorksheet worksheet, List<Dictionary<string, object>> data, string startCell)
    {
        if (data.Count == 0) return;
        
        var startRange = worksheet.Range(startCell);
        if (startRange == null) return;
        
        // 获取所有键（列名）
        var keys = data[0].Keys.ToList();
        
        // 写入表头
        for (int col = 0; col < keys.Count; col++)
        {
            worksheet.Cells[startRange.Row, startRange.Column + col].Value = keys[col];
        }
        
        // 写入数据
        for (int row = 0; row < data.Count; row++)
        {
            for (int col = 0; col < keys.Count; col++)
            {
                worksheet.Cells[startRange.Row + row + 1, startRange.Column + col].Value = data[row][keys[col]];
            }
        }
        
        // 应用格式
        ApplyJsonFormat(worksheet, data, startCell);
    }
    
    private void ApplyJsonFormat(IExcelWorksheet worksheet, List<Dictionary<string, object>> data, string startCell)
    {
        // 格式设置逻辑...
        worksheet.Columns.AutoFit();
    }
}
```

## 数据导出功能

### 导出到数据库

将Excel数据导出到数据库是常见的业务需求。

```csharp
public class DatabaseExporter
{
    public void ExportToDatabase(IExcelWorksheet worksheet, DatabaseConfig config, string rangeAddress = "A1")
    {
        try
        {
            // 读取Excel数据
            var excelData = ReadExcelData(worksheet, rangeAddress);
            
            // 连接到数据库
            using var connection = CreateDatabaseConnection(config);
            
            // 创建数据表（如果不存在）
            CreateDatabaseTable(connection, config, excelData);
            
            // 批量插入数据
            BulkInsertData(connection, config, excelData);
            
            Console.WriteLine($"成功导出 {excelData.Rows.Count} 行数据到数据库");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"数据库导出失败: {ex.Message}");
        }
    }
    
    private DataTable ReadExcelData(IExcelWorksheet worksheet, string rangeAddress)
    {
        var range = worksheet.Range(rangeAddress);
        if (range == null) throw new ArgumentException("无效的区域地址");
        
        var dataTable = new DataTable("ExcelData");
        
        // 获取数据
        var dataArray = range.Value as object[,];
        if (dataArray == null) return dataTable;
        
        int rowCount = dataArray.GetLength(0);
        int colCount = dataArray.GetLength(1);
        
        // 创建列（第一行作为列名）
        for (int col = 0; col < colCount; col++)
        {
            string columnName = dataArray[0, col]?.ToString() ?? $"Column{col + 1}";
            dataTable.Columns.Add(columnName, typeof(string));
        }
        
        // 添加数据行（从第二行开始）
        for (int row = 1; row < rowCount; row++)
        {
            var dataRow = dataTable.NewRow();
            for (int col = 0; col < colCount; col++)
            {
                dataRow[col] = dataArray[row, col]?.ToString() ?? "";
            }
            dataTable.Rows.Add(dataRow);
        }
        
        return dataTable;
    }
    
    private System.Data.Common.DbConnection CreateDatabaseConnection(DatabaseConfig config)
    {
        // 创建数据库连接（实际实现需要具体的数据库提供程序）
        // 这里返回模拟连接
        return null;
    }
    
    private void CreateDatabaseTable(System.Data.Common.DbConnection connection, DatabaseConfig config, DataTable dataTable)
    {
        // 创建数据库表的逻辑
        // 根据DataTable的结构创建对应的数据库表
    }
    
    private void BulkInsertData(System.Data.Common.DbConnection connection, DatabaseConfig config, DataTable dataTable)
    {
        // 批量插入数据的逻辑
        // 使用SqlBulkCopy或其他批量插入技术
    }
}
```

### 导出到文件

将Excel数据导出为各种文件格式。

```csharp
public class FileExporter
{
    public void ExportToCsvFile(IExcelWorksheet worksheet, string csvFilePath, string rangeAddress = "A1")
    {
        try
        {
            // 读取Excel数据
            var data = ReadExcelDataForExport(worksheet, rangeAddress);
            
            // 写入CSV文件
            WriteCsvFile(data, csvFilePath);
            
            Console.WriteLine($"成功导出数据到CSV文件: {csvFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CSV导出失败: {ex.Message}");
        }
    }
    
    private List<string[]> ReadExcelDataForExport(IExcelWorksheet worksheet, string rangeAddress)
    {
        var range = worksheet.Range(rangeAddress);
        if (range == null) return new List<string[]>();
        
        var dataArray = range.Value as object[,];
        if (dataArray == null) return new List<string[]>();
        
        var result = new List<string[]>();
        int rowCount = dataArray.GetLength(0);
        int colCount = dataArray.GetLength(1);
        
        for (int row = 0; row < rowCount; row++)
        {
            var rowData = new string[colCount];
            for (int col = 0; col < colCount; col++)
            {
                rowData[col] = dataArray[row, col]?.ToString() ?? "";
            }
            result.Add(rowData);
        }
        
        return result;
    }
    
    private void WriteCsvFile(List<string[]> data, string filePath)
    {
        using var writer = new StreamWriter(filePath, false, Encoding.UTF8);
        
        foreach (var row in data)
        {
            // 处理包含逗号或引号的字段
            var processedRow = row.Select(field => 
            {
                if (field.Contains(",") || field.Contains("\""))
                {
                    return $"\"{field.Replace("\"", "\"\"")}\"";
                }
                return field;
            });
            
            writer.WriteLine(string.Join(",", processedRow));
        }
    }
    
    public void ExportToJsonFile(IExcelWorksheet worksheet, string jsonFilePath, string rangeAddress = "A1")
    {
        try
        {
            // 读取Excel数据
            var data = ReadExcelDataForExport(worksheet, rangeAddress);
            
            // 转换为JSON格式
            var jsonData = ConvertToJson(data);
            
            // 写入JSON文件
            File.WriteAllText(jsonFilePath, jsonData, Encoding.UTF8);
            
            Console.WriteLine($"成功导出数据到JSON文件: {jsonFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"JSON导出失败: {ex.Message}");
        }
    }
    
    private string ConvertToJson(List<string[]> data)
    {
        if (data.Count == 0) return "[]";
        
        var jsonObjects = new List<Dictionary<string, object>>();
        var headers = data[0];
        
        // 从第二行开始（第一行是表头）
        for (int row = 1; row < data.Count; row++)
        {
            var jsonObject = new Dictionary<string, object>();
            for (int col = 0; col < headers.Length && col < data[row].Length; col++)
            {
                jsonObject[headers[col]] = data[row][col];
            }
            jsonObjects.Add(jsonObject);
        }
        
        // 简化的JSON序列化（实际应用中应使用JSON库）
        return $"[{string.Join(",", jsonObjects.Select(obj => $"{{{string.Join(",", obj.Select(kv => $"\\\"{kv.Key}\\\":\\\"{kv.Value}\\\""))}}}"))}]";
    }
    
    public void ExportToXmlFile(IExcelWorksheet worksheet, string xmlFilePath, string rangeAddress = "A1")
    {
        try
        {
            // 读取Excel数据
            var data = ReadExcelDataForExport(worksheet, rangeAddress);
            
            // 转换为XML格式
            var xmlData = ConvertToXml(data);
            
            // 写入XML文件
            File.WriteAllText(xmlFilePath, xmlData, Encoding.UTF8);
            
            Console.WriteLine($"成功导出数据到XML文件: {xmlFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"XML导出失败: {ex.Message}");
        }
    }
    
    private string ConvertToXml(List<string[]> data)
    {
        if (data.Count == 0) return "<Data></Data>";
        
        var xmlBuilder = new System.Text.StringBuilder();
        xmlBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        xmlBuilder.AppendLine("<Data>");
        
        var headers = data[0];
        
        // 从第二行开始
        for (int row = 1; row < data.Count; row++)
        {
            xmlBuilder.AppendLine("  <Row>");
            for (int col = 0; col < headers.Length && col < data[row].Length; col++)
            {
                string fieldName = SanitizeXmlName(headers[col]);
                string fieldValue = System.Security.SecurityElement.Escape(data[row][col]);
                xmlBuilder.AppendLine($"    <{fieldName}>{fieldValue}</{fieldName}>");
            }
            xmlBuilder.AppendLine("  </Row>");
        }
        
        xmlBuilder.AppendLine("</Data>");
        return xmlBuilder.ToString();
    }
    
    private string SanitizeXmlName(string name)
    {
        // 清理XML元素名称
        if (string.IsNullOrEmpty(name)) return "Field";
        
        // 移除无效字符
        var validChars = name.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray();
        var sanitized = new string(validChars);
        
        // 确保以字母或下划线开头
        if (string.IsNullOrEmpty(sanitized) || !char.IsLetter(sanitized[0]))
        {
            sanitized = "_" + sanitized;
        }
        
        return sanitized;
    }
}
```

## 数据转换与清洗

### 数据格式转换

在导入导出过程中，经常需要进行数据格式的转换。

```csharp
public class DataTransformer
{
    public void TransformData(IExcelWorksheet worksheet, DataTransformationConfig config)
    {
        var range = worksheet.Range(config.SourceRange);
        if (range == null) return;
        
        try
        {
            // 读取源数据
            var sourceData = range.Value as object[,];
            if (sourceData == null) return;
            
            // 应用转换
            var transformedData = ApplyTransformations(sourceData, config);
            
            // 写入目标区域
            var targetRange = worksheet.Range(config.TargetRange);
            if (targetRange != null)
            {
                targetRange.Value = transformedData;
            }
            
            Console.WriteLine("数据转换完成");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"数据转换失败: {ex.Message}");
        }
    }
    
    private object[,] ApplyTransformations(object[,] sourceData, DataTransformationConfig config)
    {
        int rowCount = sourceData.GetLength(0);
        int colCount = sourceData.GetLength(1);
        
        var result = new object[rowCount, colCount];
        
        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < colCount; col++)
            {
                var originalValue = sourceData[row, col];
                result[row, col] = TransformValue(originalValue, config.Transformations, col);
            }
        }
        
        return result;
    }
    
    private object TransformValue(object originalValue, List<TransformationRule> transformations, int columnIndex)
    {
        var transformedValue = originalValue;
        
        foreach (var rule in transformations.Where(r => r.ApplyToColumn == columnIndex || r.ApplyToColumn == -1))
        {
            transformedValue = ApplyTransformationRule(transformedValue, rule);
        }
        
        return transformedValue;
    }
    
    private object ApplyTransformationRule(object value, TransformationRule rule)
    {
        if (value == null) return rule.DefaultValue ?? "";
        
        string stringValue = value.ToString() ?? "";
        
        return rule.Type switch
        {
            TransformationType.Trim => stringValue.Trim(),
            TransformationType.ToUpper => stringValue.ToUpper(),
            TransformationType.ToLower => stringValue.ToLower(),
            TransformationType.RemoveSpaces => stringValue.Replace(" ", ""),
            TransformationType.ParseNumber => ParseNumber(stringValue),
            TransformationType.ParseDate => ParseDate(stringValue),
            TransformationType.Replace => stringValue.Replace(rule.FindValue, rule.ReplaceValue),
            TransformationType.CustomFormat => string.Format(rule.FormatString, value),
            _ => value
        };
    }
    
    private object ParseNumber(string value)
    {
        if (double.TryParse(value, out double result))
        {
            return result;
        }
        return value;
    }
    
    private object ParseDate(string value)
    {
        if (DateTime.TryParse(value, out DateTime result))
        {
            return result;
        }
        return value;
    }
}

public class DataTransformationConfig
{
    public string SourceRange { get; set; } = "A1";
    public string TargetRange { get; set; } = "A1";
    public List<TransformationRule> Transformations { get; set; } = new();
}

public class TransformationRule
{
    public TransformationType Type { get; set; }
    public int ApplyToColumn { get; set; } = -1; // -1表示应用到所有列
    public string FindValue { get; set; } = "";
    public string ReplaceValue { get; set; } = "";
    public string FormatString { get; set; } = "";
    public object DefaultValue { get; set; } = "";
}

public enum TransformationType
{
    Trim,           // 去除空格
    ToUpper,        // 转换为大写
    ToLower,        // 转换为小写
    RemoveSpaces,   // 移除所有空格
    ParseNumber,    // 解析为数字
    ParseDate,      // 解析为日期
    Replace,        // 替换文本
    CustomFormat    // 自定义格式
}
```

### 数据清洗与验证

```csharp
public class DataCleaner
{
    public DataCleaningResult CleanData(IExcelWorksheet worksheet, DataCleaningConfig config)
    {
        var result = new DataCleaningResult();
        
        try
        {
            var range = worksheet.Range(config.RangeAddress);
            if (range == null) return result;
            
            var data = range.Value as object[,];
            if (data == null) return result;
            
            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);
            
            // 执行数据清洗
            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < colCount; col++)
                {
                    var originalValue = data[row, col];
                    var cleanedValue = CleanCellValue(originalValue, config.CleaningRules);
                    
                    if (!Equals(originalValue, cleanedValue))
                    {
                        data[row, col] = cleanedValue;
                        result.ModifiedCells++;
                    }
                    
                    // 验证数据
                    var validationResult = ValidateCellValue(cleanedValue, config.ValidationRules, col);
                    if (!validationResult.IsValid)
                    {
                        result.InvalidCells++;
                        result.ValidationErrors.Add(new ValidationError
                        {
                            Row = row + 1,
                            Column = col + 1,
                            Value = cleanedValue,
                            ErrorMessage = validationResult.ErrorMessage
                        });
                    }
                }
            }
            
            // 写回清洗后的数据
            range.Value = data;
            
            result.Success = true;
            result.TotalCells = rowCount * colCount;
            
            Console.WriteLine($"数据清洗完成: 修改了 {result.ModifiedCells} 个单元格，发现 {result.InvalidCells} 个无效数据");
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
            Console.WriteLine($"数据清洗失败: {ex.Message}");
        }
        
        return result;
    }
    
    private object CleanCellValue(object value, List<CleaningRule> rules)
    {
        if (value == null) return "";
        
        var stringValue = value.ToString() ?? "";
        var cleanedValue = stringValue;
        
        foreach (var rule in rules)
        {
            cleanedValue = rule.Type switch
            {
                CleaningType.Trim => cleanedValue.Trim(),
                CleaningType.RemoveExtraSpaces => RemoveExtraSpaces(cleanedValue),
                CleaningType.RemoveSpecialCharacters => RemoveSpecialCharacters(cleanedValue, rule.AllowedCharacters),
                CleaningType.NormalizeWhitespace => NormalizeWhitespace(cleanedValue),
                CleaningType.StandardizeCase => StandardizeCase(cleanedValue, rule.CaseType),
                _ => cleanedValue
            };
        }
        
        return cleanedValue;
    }
    
    private string RemoveExtraSpaces(string value)
    {
        return System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " ");
    }
    
    private string RemoveSpecialCharacters(string value, string allowedCharacters)
    {
        if (string.IsNullOrEmpty(allowedCharacters))
        {
            allowedCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ";
        }
        
        return new string(value.Where(c => allowedCharacters.Contains(c)).ToArray());
    }
    
    private string NormalizeWhitespace(string value)
    {
        return System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " ").Trim();
    }
    
    private string StandardizeCase(string value, CaseType caseType)
    {
        return caseType switch
        {
            CaseType.Upper => value.ToUpper(),
            CaseType.Lower => value.ToLower(),
            CaseType.Proper => ToProperCase(value),
            _ => value
        };
    }
    
    private string ToProperCase(string value)
    {
        if (string.IsNullOrEmpty(value)) return value;
        
        var words = value.Split(' ');
        for (int i = 0; i < words.Length; i++)
        {
            if (!string.IsNullOrEmpty(words[i]))
            {
                words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
            }
        }
        
        return string.Join(" ", words);
    }
    
    private ValidationResult ValidateCellValue(object value, List<ValidationRule> rules, int columnIndex)
    {
        var result = new ValidationResult { IsValid = true };
        
        foreach (var rule in rules.Where(r => r.ApplyToColumn == columnIndex || r.ApplyToColumn == -1))
        {
            if (!ValidateAgainstRule(value, rule))
            {
                result.IsValid = false;
                result.ErrorMessage = rule.ErrorMessage;
                break;
            }
        }
        
        return result;
    }
    
    private bool ValidateAgainstRule(object value, ValidationRule rule)
    {
        if (value == null) return rule.AllowNull;
        
        var stringValue = value.ToString() ?? "";
        
        return rule.Type switch
        {
            ValidationType.Required => !string.IsNullOrWhiteSpace(stringValue),
            ValidationType.MinLength => stringValue.Length >= rule.MinLength,
            ValidationType.MaxLength => stringValue.Length <= rule.MaxLength,
            ValidationType.Regex => System.Text.RegularExpressions.Regex.IsMatch(stringValue, rule.Pattern),
            ValidationType.Numeric => double.TryParse(stringValue, out _),
            ValidationType.Date => DateTime.TryParse(stringValue, out _),
            ValidationType.Email => IsValidEmail(stringValue),
            _ => true
        };
    }
    
    private bool IsValidEmail(string email)
    {
        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }
}

public class DataCleaningConfig
{
    public string RangeAddress { get; set; } = "A1";
    public List<CleaningRule> CleaningRules { get; set; } = new();
    public List<ValidationRule> ValidationRules { get; set; } = new();
}

public class CleaningRule
{
    public CleaningType Type { get; set; }
    public string AllowedCharacters { get; set; } = "";
    public CaseType CaseType { get; set; }
}

public enum CleaningType
{
    Trim,
    RemoveExtraSpaces,
    RemoveSpecialCharacters,
    NormalizeWhitespace,
    StandardizeCase
}

public enum CaseType
{
    Upper,
    Lower,
    Proper
}

public class ValidationRule
{
    public ValidationType Type { get; set; }
    public int ApplyToColumn { get; set; } = -1;
    public bool AllowNull { get; set; } = true;
    public int MinLength { get; set; }
    public int MaxLength { get; set; }
    public string Pattern { get; set; } = "";
    public string ErrorMessage { get; set; } = "";
}

public enum ValidationType
{
    Required,
    MinLength,
    MaxLength,
    Regex,
    Numeric,
    Date,
    Email
}

public class DataCleaningResult
{
    public bool Success { get; set; }
    public string ErrorMessage { get; set; } = "";
    public int TotalCells { get; set; }
    public int ModifiedCells { get; set; }
    public int InvalidCells { get; set; }
    public List<ValidationError> ValidationErrors { get; set; } = new();
}

public class ValidationError
{
    public int Row { get; set; }
    public int Column { get; set; }
    public object Value { get; set; } = "";
    public string ErrorMessage { get; set; } = "";
}

public class ValidationResult
{
    public bool IsValid { get; set; }
    public string ErrorMessage { get; set; } = "";
}
```

## 实际应用场景

### 场景1：企业数据集成平台

```csharp
public class EnterpriseDataIntegrationPlatform
{
    private readonly DatabaseImporter _dbImporter;
    private readonly FileImporter _fileImporter;
    private readonly DatabaseExporter _dbExporter;
    private readonly FileExporter _fileExporter;
    private readonly DataCleaner _dataCleaner;
    
    public EnterpriseDataIntegrationPlatform()
    {
        _dbImporter = new DatabaseImporter();
        _fileImporter = new FileImporter();
        _dbExporter = new DatabaseExporter();
        _fileExporter = new FileExporter();
        _dataCleaner = new DataCleaner();
    }
    
    public void ProcessDataPipeline(IExcelWorkbook workbook, DataPipelineConfig config)
    {
        Console.WriteLine("开始数据处理管道...");
        
        // 步骤1：从数据源导入
        ImportData(workbook, config.ImportConfig);
        
        // 步骤2：数据清洗和转换
        CleanAndTransformData(workbook, config.CleaningConfig);
        
        // 步骤3：导出到目标系统
        ExportData(workbook, config.ExportConfig);
        
        Console.WriteLine("数据处理管道完成");
    }
    
    private void ImportData(IExcelWorkbook workbook, ImportConfig config)
    {
        var worksheet = workbook.Worksheets?.Add("原始数据");
        if (worksheet == null) return;
        
        switch (config.SourceType)
        {
            case DataSourceType.Database:
                _dbImporter.ImportFromDatabase(workbook, config.DatabaseConfig);
                break;
            case DataSourceType.CsvFile:
                _fileImporter.ImportFromCsvFile(worksheet, config.FilePath);
                break;
            case DataSourceType.JsonFile:
                _fileImporter.ImportFromJsonFile(worksheet, config.FilePath);
                break;
        }
    }
    
    private void CleanAndTransformData(IExcelWorkbook workbook, DataCleaningConfig config)
    {
        var rawWorksheet = workbook.Worksheets?["原始数据"];
        if (rawWorksheet == null) return;
        
        var cleanedWorksheet = workbook.Worksheets?.Add("清洗后数据");
        if (cleanedWorksheet == null) return;
        
        // 复制原始数据
        var usedRange = rawWorksheet.UsedRange;
        if (usedRange != null)
        {
            usedRange.Copy(cleanedWorksheet.Range("A1"));
        }
        
        // 执行数据清洗
        var cleaningResult = _dataCleaner.CleanData(cleanedWorksheet, config);
        
        // 记录清洗结果
        LogCleaningResult(cleaningResult);
    }
    
    private void ExportData(IExcelWorkbook workbook, ExportConfig config)
    {
        var worksheet = workbook.Worksheets?["清洗后数据"];
        if (worksheet == null) return;
        
        switch (config.TargetType)
        {
            case DataTargetType.Database:
                _dbExporter.ExportToDatabase(worksheet, config.DatabaseConfig);
                break;
            case DataTargetType.CsvFile:
                _fileExporter.ExportToCsvFile(worksheet, config.FilePath);
                break;
            case DataTargetType.JsonFile:
                _fileExporter.ExportToJsonFile(worksheet, config.FilePath);
                break;
        }
    }
    
    private void LogCleaningResult(DataCleaningResult result)
    {
        Console.WriteLine($"数据清洗统计:");
        Console.WriteLine($"- 总单元格数: {result.TotalCells}");
        Console.WriteLine($"- 修改单元格数: {result.ModifiedCells}");
        Console.WriteLine($"- 无效数据数: {result.InvalidCells}");
        
        if (result.ValidationErrors.Count > 0)
        {
            Console.WriteLine("验证错误详情:");
            foreach (var error in result.ValidationErrors.Take(10)) // 只显示前10个错误
            {
                Console.WriteLine($"- 位置({error.Row},{error.Column}): {error.Value} - {error.ErrorMessage}");
            }
        }
    }
}

public class DataPipelineConfig
{
    public ImportConfig ImportConfig { get; set; } = new();
    public DataCleaningConfig CleaningConfig { get; set; } = new();
    public ExportConfig ExportConfig { get; set; } = new();
}

public class ImportConfig
{
    public DataSourceType SourceType { get; set; }
    public DatabaseConfig DatabaseConfig { get; set; } = new();
    public string FilePath { get; set; } = "";
}

public class ExportConfig
{
    public DataTargetType TargetType { get; set; }
    public DatabaseConfig DatabaseConfig { get; set; } = new();
    public string FilePath { get; set; } = "";
}

public enum DataSourceType
{
    Database,
    CsvFile,
    JsonFile
}

public enum DataTargetType
{
    Database,
    CsvFile,
    JsonFile
}
```

### 场景2：批量文件处理系统

```csharp
public class BatchFileProcessor
{
    public void ProcessBatchFiles(string inputDirectory, string outputDirectory, FileProcessingConfig config)
    {
        if (!Directory.Exists(inputDirectory))
        {
            Console.WriteLine($"输入目录不存在: {inputDirectory}");
            return;
        }
        
        Directory.CreateDirectory(outputDirectory);
        
        var inputFiles = Directory.GetFiles(inputDirectory, $"*.{config.InputExtension}");
        
        Console.WriteLine($"发现 {inputFiles.Length} 个待处理文件");
        
        foreach (var inputFile in inputFiles)
        {
            try
            {
                ProcessSingleFile(inputFile, outputDirectory, config);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文件失败 {inputFile}: {ex.Message}");
            }
        }
        
        Console.WriteLine("批量文件处理完成");
    }
    
    private void ProcessSingleFile(string inputFile, string outputDirectory, FileProcessingConfig config)
    {
        var fileName = Path.GetFileNameWithoutExtension(inputFile);
        var outputFile = Path.Combine(outputDirectory, $"{fileName}.{config.OutputExtension}");
        
        using var excelApp = ExcelFactory.BlankWorkbook();
        excelApp.Visible = false;
        
        var worksheet = excelApp.ActiveSheetWrap;
        
        // 导入数据
        ImportDataToWorksheet(worksheet, inputFile, config.InputExtension);
        
        // 应用转换规则
        ApplyTransformations(worksheet, config.Transformations);
        
        // 导出数据
        ExportDataFromWorksheet(worksheet, outputFile, config.OutputExtension);
        
        Console.WriteLine($"成功处理: {inputFile} -> {outputFile}");
    }
    
    private void ImportDataToWorksheet(IExcelWorksheet worksheet, string inputFile, string extension)
    {
        var importer = new FileImporter();
        
        switch (extension.ToLower())
        {
            case "csv":
                importer.ImportFromCsvFile(worksheet, inputFile);
                break;
            case "txt":
                // 处理文本文件
                break;
            default:
                throw new NotSupportedException($"不支持的输入格式: {extension}");
        }
    }
    
    private void ApplyTransformations(IExcelWorksheet worksheet, List<DataTransformationConfig> transformations)
    {
        var transformer = new DataTransformer();
        
        foreach (var config in transformations)
        {
            transformer.TransformData(worksheet, config);
        }
    }
    
    private void ExportDataFromWorksheet(IExcelWorksheet worksheet, string outputFile, string extension)
    {
        var exporter = new FileExporter();
        
        switch (extension.ToLower())
        {
            case "csv":
                exporter.ExportToCsvFile(worksheet, outputFile);
                break;
            case "json":
                exporter.ExportToJsonFile(worksheet, outputFile);
                break;
            case "xml":
                exporter.ExportToXmlFile(worksheet, outputFile);
                break;
            default:
                throw new NotSupportedException($"不支持的输出格式: {extension}");
        }
    }
}

public class FileProcessingConfig
{
    public string InputExtension { get; set; } = "csv";
    public string OutputExtension { get; set; } = "json";
    public List<DataTransformationConfig> Transformations { get; set; } = new();
}
```

## 总结

通过本文的学习，我们深入掌握了Excel数据导入导出与转换的各种技术，包括：

**数据导入功能：**
- 从数据库导入数据（SQL Server、Oracle、MySQL等）
- 从DataTable导入数据（批量操作和格式设置）
- 从文件导入数据（CSV、JSON、XML等格式）

**数据导出功能：**
- 导出到数据库（批量插入和数据表创建）
- 导出到文件（CSV、JSON、XML等格式）
- 数据格式转换和编码处理

**数据转换技术：**
- 数据格式转换（文本处理、数字解析、日期格式化）
- 数据清洗（空格处理、特殊字符清理、大小写标准化）
- 数据验证（必填验证、格式验证、业务规则验证）

**实际应用价值：**
- 企业数据集成平台实现多系统数据交换
- 批量文件处理系统提高数据处理效率
- 数据质量管理系统确保数据准确性和一致性

**最佳实践：**
- 使用批量操作提高性能
- 实现完善的数据验证和错误处理
- 提供灵活的数据转换配置
- 考虑数据安全和权限控制

在下一篇文章中，我们将深入探讨公式与函数应用，这是Excel数据处理的核心功能。

---

**下一篇预告：**《公式与函数应用》将详细介绍Excel公式和函数的编程操作，包括常用公式设置、数组公式应用、自定义函数调用等高级功能。