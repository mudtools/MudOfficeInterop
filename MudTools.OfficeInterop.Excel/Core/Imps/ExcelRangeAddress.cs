//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！



namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelRangeAddress : IExcelRangeAddress, IEquatable<IExcelRangeAddress>, IDisposable
{
    private readonly MsExcel.Range _range;
    private bool _disposed;

    #region 属性

    /// <summary>起始行号 (1-based)</summary>
    public int StartRow { get; private set; }

    /// <summary>起始列号 (1-based)</summary>
    public int StartColumn { get; private set; }

    /// <summary>结束行号 (1-based)</summary>
    public int EndRow { get; private set; }

    /// <summary>结束列号 (1-based)</summary>
    public int EndColumn { get; private set; }

    /// <summary>工作表名称</summary>
    public string SheetName { get; private set; }

    /// <summary>工作簿名称</summary>
    public string WorkbookName { get; private set; }

    /// <summary>是否为绝对引用 ($A$1 格式)</summary>
    public bool IsAbsoluteReference { get; private set; }

    /// <summary>是否为单个单元格</summary>
    public bool IsSingleCell => StartRow == EndRow && StartColumn == EndColumn;

    /// <summary>单元格行数</summary>
    public int RowCount => EndRow - StartRow + 1;

    /// <summary>单元格列数</summary>
    public int ColumnCount => EndColumn - StartColumn + 1;

    #endregion

    #region 构造函数

    public ExcelRangeAddress(MsExcel.Range range)
    {
        _range = range ?? throw new ArgumentNullException(nameof(range));
        ParseAddress(range.Address);
    }

    public ExcelRangeAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("Address cannot be null or empty", nameof(address));

        ParseAddress(address);
    }

    #endregion

    #region 地址解析

    private void ParseAddress(string address)
    {
        // 处理外部引用格式 [Workbook]SheetName!Address
        if (address.StartsWith("["))
        {
            int closingBracket = address.IndexOf(']');
            if (closingBracket > 0)
            {
                WorkbookName = address.Substring(1, closingBracket - 1);
                address = address.Substring(closingBracket + 1);
            }
        }

        // 分离工作表名称和单元格地址
        int exclamationPos = address.IndexOf('!');
        if (exclamationPos > 0)
        {
            SheetName = address.Substring(0, exclamationPos);
            // 处理带单引号的工作表名
            if (SheetName.StartsWith("'") && SheetName.EndsWith("'"))
            {
                SheetName = SheetName.Substring(1, SheetName.Length - 2);
            }
            address = address.Substring(exclamationPos + 1);
        }

        // 检查是否为绝对引用
        IsAbsoluteReference = address.Contains("$");

        // 解析单元格地址
        ParseCellAddress(address);
    }

    private void ParseCellAddress(string address)
    {
        // 处理多区域地址（只取第一个区域）
        int commaPos = address.IndexOf(',');
        if (commaPos > 0)
        {
            address = address.Substring(0, commaPos);
        }

        // 处理 A1 引用样式
        if (address.Contains(":"))
        {
            // 范围地址 (A1:B2)
            string[] parts = address.Split(':');
            ParseCellReference(parts[0], out int startCol, out int startRow);
            ParseCellReference(parts[1], out int endCol, out int endRow);

            StartColumn = startCol;
            EndColumn = endCol;
            StartRow = startRow;
            EndRow = endRow;
        }
        else
        {
            // 单个单元格地址 (A1)
            ParseCellReference(address, out int startCol, out int startRow);
            StartColumn = startCol;
            StartRow = startRow;
            EndRow = StartRow;
            EndColumn = StartColumn;
        }
    }

    private void ParseCellReference(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;

        // 移除绝对引用符号
        reference = reference.Replace("$", "");

        // 提取行号和列号
        int rowStart = 0;
        for (int i = 0; i < reference.Length; i++)
        {
            if (char.IsDigit(reference[i]))
            {
                rowStart = i;
                break;
            }
        }

        if (rowStart == 0)
            throw new FormatException($"Invalid cell reference: {reference}");

        string columnRef = reference.Substring(0, rowStart);
        string rowRef = reference.Substring(rowStart);

        column = ColumnNameToNumber(columnRef);
        row = int.Parse(rowRef);
    }

    #endregion

    #region 公共方法

    /// <summary>获取A1样式的地址</summary>
    public string GetAddressA1(
        bool absoluteRow = true,
        bool absoluteCol = true,
        bool includeSheet = true,
        bool includeWorkbook = false)
    {
        string startAddr = FormatSingleCell(StartRow, StartColumn, absoluteRow, absoluteCol);

        if (IsSingleCell)
            return FormatFullAddress(startAddr, includeSheet, includeWorkbook);

        string endAddr = FormatSingleCell(EndRow, EndColumn, absoluteRow, absoluteCol);
        return FormatFullAddress($"{startAddr}:{endAddr}", includeSheet, includeWorkbook);
    }

    /// <summary>获取R1C1样式的地址</summary>
    public string GetAddressR1C1(
        bool absoluteRow = true,
        bool absoluteCol = true,
        bool includeSheet = true,
        bool includeWorkbook = false)
    {
        string startRowRef = absoluteRow ? StartRow.ToString() : $"R[{StartRow}]";
        string startColRef = absoluteCol ? StartColumn.ToString() : $"C[{StartColumn}]";
        string startAddr = $"{startRowRef}{startColRef}";

        if (IsSingleCell)
            return FormatFullAddress(startAddr, includeSheet, includeWorkbook);

        string endRowRef = absoluteRow ? EndRow.ToString() : $"R[{EndRow}]";
        string endColRef = absoluteCol ? EndColumn.ToString() : $"C[{EndColumn}]";
        return FormatFullAddress($"{startAddr}:{endRowRef}{endColRef}", includeSheet, includeWorkbook);
    }

    /// <summary>获取本地化地址（使用Excel的本地设置）</summary>
    public string GetLocalAddress()
    {
        try
        {
            return _range?.AddressLocal ?? GetAddressA1();
        }
        catch
        {
            return GetAddressA1();
        }
    }

    /// <summary>获取外部引用地址（包含工作簿和工作表）</summary>
    public string GetExternalAddress()
    {
        return GetAddressA1(true, true, true, true);
    }

    /// <summary>获取起始单元格地址 (A1)</summary>
    public string GetStartCellAddress()
    {
        return FormatSingleCell(StartRow, StartColumn);
    }

    /// <summary>获取结束单元格地址 (A1)</summary>
    public string GetEndCellAddress()
    {
        return FormatSingleCell(EndRow, EndColumn);
    }

    /// <summary>检查地址是否包含指定单元格</summary>
    public bool ContainsCell(int row, int column)
    {
        return row >= StartRow && row <= EndRow &&
               column >= StartColumn && column <= EndColumn;
    }

    /// <summary>检查地址是否包含指定单元格 (A1格式)</summary>
    public bool ContainsCell(string address)
    {
        var cellAddr = new ExcelRangeAddress(address);
        return ContainsCell(cellAddr.StartRow, cellAddr.StartColumn);
    }

    /// <summary>获取偏移后的新地址</summary>
    public IExcelRangeAddress Offset(int rowOffset, int colOffset)
    {
        return new ExcelRangeAddress(
            $"{FormatSingleCell(StartRow + rowOffset, StartColumn + colOffset)}:" +
            $"{FormatSingleCell(EndRow + rowOffset, EndColumn + colOffset)}")
        {
            SheetName = SheetName,
            WorkbookName = WorkbookName
        };
    }

    /// <summary>获取调整大小后的新地址</summary>
    public IExcelRangeAddress Resize(int rows, int columns)
    {
        return new ExcelRangeAddress(
            $"{FormatSingleCell(StartRow, StartColumn)}:" +
            $"{FormatSingleCell(StartRow + rows - 1, StartColumn + columns - 1)}")
        {
            SheetName = SheetName,
            WorkbookName = WorkbookName
        };
    }

    #endregion

    #region 辅助方法

    private string FormatFullAddress(string address, bool includeSheet, bool includeWorkbook)
    {
        string result = address;

        if (includeSheet && !string.IsNullOrEmpty(SheetName))
        {
            // 如果工作表名称包含空格，添加单引号
            string safeSheetName = SheetName.Contains(" ") ? $"'{SheetName}'" : SheetName;
            result = $"{safeSheetName}!{result}";
        }

        if (includeWorkbook && !string.IsNullOrEmpty(WorkbookName))
        {
            result = $"[{WorkbookName}]{result}";
        }

        return result;
    }

    private string FormatSingleCell(int row, int col,
        bool absoluteRow = true, bool absoluteCol = true)
    {
        string colName = ColumnNumberToName(col);
        string rowName = row.ToString();

        if (absoluteCol) colName = "$" + colName;
        if (absoluteRow) rowName = "$" + rowName;

        return $"{colName}{rowName}";
    }

    private string ColumnNumberToName(int columnNumber)
    {
        if (columnNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(columnNumber), "Column number must be positive");

        string columnName = "";
        while (columnNumber > 0)
        {
            int remainder = (columnNumber - 1) % 26;
            columnName = (char)('A' + remainder) + columnName;
            columnNumber = (columnNumber - 1) / 26;
        }
        return columnName;
    }

    private int ColumnNameToNumber(string columnName)
    {
        if (string.IsNullOrEmpty(columnName))
            throw new ArgumentNullException(nameof(columnName));

        columnName = columnName.ToUpper();
        int sum = 0;

        for (int i = 0; i < columnName.Length; i++)
        {
            if (!char.IsLetter(columnName[i]))
                throw new FormatException($"Invalid character in column name: {columnName}");

            sum *= 26;
            sum += (columnName[i] - 'A' + 1);
        }

        return sum;
    }

    #endregion

    #region 重写和接口实现

    public override string ToString() => GetAddressA1();

    public bool Equals(IExcelRangeAddress other)
    {
        if (other is null) return false;
        return StartRow == other.StartRow &&
               StartColumn == other.StartColumn &&
               EndRow == other.EndRow &&
               EndColumn == other.EndColumn &&
               SheetName == other.SheetName &&
               WorkbookName == other.WorkbookName;
    }

    public override bool Equals(object obj) => Equals(obj as ExcelRangeAddress);

    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + StartRow.GetHashCode();
            hash = hash * 23 + StartColumn.GetHashCode();
            hash = hash * 23 + EndRow.GetHashCode();
            hash = hash * 23 + EndColumn.GetHashCode();
            hash = hash * 23 + (SheetName?.GetHashCode() ?? 0);
            hash = hash * 23 + (WorkbookName?.GetHashCode() ?? 0);
            return hash;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            // 释放托管资源
        }

        // 释放COM对象
        if (_range != null)
        {
            Marshal.FinalReleaseComObject(_range);
        }

        _disposed = true;
    }

    ~ExcelRangeAddress()
    {
        Dispose(false);
    }

    #endregion
}