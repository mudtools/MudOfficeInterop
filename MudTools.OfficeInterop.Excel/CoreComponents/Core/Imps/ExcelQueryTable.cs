//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Excel;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// QueryTable COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelQueryTable : IExcelQueryTable
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.QueryTable _queryTable;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="queryTable">原始的 QueryTable COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 queryTable 为 null 时抛出。</exception>
    internal ExcelQueryTable(MsExcel.QueryTable queryTable)
    {
        _queryTable = queryTable ?? throw new ArgumentNullException(nameof(queryTable));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_queryTable != null)
            {
                Marshal.ReleaseComObject(_queryTable);
                _queryTable = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    public object? Parent => _queryTable?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication? Application =>
        _queryTable?.Application != null
            ? new ExcelApplication(_queryTable.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置查询表的名称。
    /// </summary>
    public string Name
    {
        get => _queryTable?.Name ?? string.Empty;
        set
        {
            if (_queryTable != null && !string.IsNullOrEmpty(value))
                _queryTable.Name = value;
        }
    }

    /// <summary>
    /// 获取或设置是否在刷新时保持列宽不变。
    /// </summary>
    public bool PreserveFormatting
    {
        get => _queryTable != null && _queryTable.PreserveFormatting;
        set
        {
            if (_queryTable != null)
                _queryTable.PreserveFormatting = value;
        }
    }

    public string TextFileDecimalSeparator
    {
        get => _queryTable != null ? _queryTable.TextFileDecimalSeparator : string.Empty;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileDecimalSeparator = value;
        }

    }

    public string TextFileThousandsSeparator
    {
        get => _queryTable != null ? _queryTable.TextFileThousandsSeparator : string.Empty;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileThousandsSeparator = value;
        }

    }

    public bool MaintainConnection
    {
        get => _queryTable != null && _queryTable.MaintainConnection;
        set
        {
            if (_queryTable != null)
                _queryTable.MaintainConnection = value;
        }

    }

    public bool EnableRefresh
    {
        get => _queryTable != null && _queryTable.EnableRefresh;
        set
        {
            if (_queryTable != null)
                _queryTable.EnableRefresh = value;
        }
    }

    public XlQueryType QueryType
    {
        get => _queryTable != null
            ? _queryTable.QueryType.EnumConvert(XlQueryType.xlOLEDBQuery)
            : XlQueryType.xlOLEDBQuery;
    }

    /// <summary>
    /// 获取或设置是否在刷新后调整列宽以适应内容。
    /// </summary>
    public bool AdjustColumnWidth
    {
        get => _queryTable != null && _queryTable.AdjustColumnWidth;
        set
        {
            if (_queryTable != null)
                _queryTable.AdjustColumnWidth = value;
        }
    }

    /// <summary>
    /// 获取或设置刷新样式（插入单元格的方式）。
    /// </summary>
    public XlCellInsertionMode RefreshStyle
    {
        get => _queryTable != null
            ? _queryTable.RefreshStyle.EnumConvert(XlCellInsertionMode.xlInsertDeleteCells)
            : XlCellInsertionMode.xlInsertDeleteCells;
        set
        {
            if (_queryTable != null)
                _queryTable.RefreshStyle = value.EnumConvert(MsExcel.XlCellInsertionMode.xlInsertDeleteCells);
        }
    }

    public XlTextParsingType TextFileParseType
    {
        get => _queryTable != null
            ? _queryTable.TextFileParseType.EnumConvert(XlTextParsingType.xlDelimited)
            : XlTextParsingType.xlDelimited;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileParseType = value.EnumConvert(MsExcel.XlTextParsingType.xlDelimited);
        }
    }

    public XlTextQualifier TextFileTextQualifier
    {
        get => _queryTable != null
            ? _queryTable.TextFileTextQualifier.EnumConvert(XlTextQualifier.xlTextQualifierNone)
            : XlTextQualifier.xlTextQualifierNone;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileTextQualifier = value.EnumConvert(MsExcel.XlTextQualifier.xlTextQualifierNone);
        }
    }

    public bool TextFileConsecutiveDelimiter
    {
        get => _queryTable != null && _queryTable.TextFileConsecutiveDelimiter;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileConsecutiveDelimiter = value;
        }
    }

    public bool TextFileTabDelimiter
    {
        get => _queryTable != null && _queryTable.TextFileTabDelimiter;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileTabDelimiter = value;
        }
    }

    public bool TextFileSemicolonDelimiter
    {
        get => _queryTable != null && _queryTable.TextFileSemicolonDelimiter;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileSemicolonDelimiter = value;
        }
    }

    public bool TextFileCommaDelimiter
    {
        get => _queryTable != null && _queryTable.TextFileCommaDelimiter;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileCommaDelimiter = value;
        }

    }

    public bool TextFileSpaceDelimiter
    {
        get => _queryTable != null && _queryTable.TextFileSpaceDelimiter;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileSpaceDelimiter = value;
        }
    }

    public string TextFileOtherDelimiter
    {
        get => _queryTable != null ? _queryTable.TextFileOtherDelimiter : string.Empty;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileOtherDelimiter = value;
        }
    }

    public bool PreserveColumnInfo
    {
        get => _queryTable != null && _queryTable.PreserveColumnInfo;
        set
        {
            if (_queryTable != null)
                _queryTable.PreserveColumnInfo = value;
        }

    }


    /// <summary>
    /// 获取或设置是否保存查询数据（即使连接断开）。
    /// </summary>
    public bool SaveData
    {
        get => _queryTable != null && _queryTable.SaveData;
        set
        {
            if (_queryTable != null)
                _queryTable.SaveData = value;
        }
    }

    /// <summary>
    /// 获取或设置是否在打开工作簿时自动刷新。
    /// </summary>
    public bool RefreshOnFileOpen
    {
        get => _queryTable != null && _queryTable.RefreshOnFileOpen;
        set
        {
            if (_queryTable != null)
                _queryTable.RefreshOnFileOpen = value;
        }
    }

    public bool FetchedRowOverflow
    {
        get => _queryTable != null && _queryTable.FetchedRowOverflow;
    }

    public bool Refreshing
    {
        get => _queryTable != null && _queryTable.Refreshing;
    }

    /// <summary>
    /// 获取或设置背景刷新（异步刷新）。
    /// </summary>
    public bool BackgroundQuery
    {
        get => _queryTable != null && _queryTable.BackgroundQuery;
        set
        {
            if (_queryTable != null)
                _queryTable.BackgroundQuery = value;
        }
    }

    /// <summary>
    /// 获取或设置连接字符串（如 ODBC、OLEDB、Web URL 等）。
    /// </summary>
    public string Connection
    {
        get => _queryTable?.Connection.ToString() ?? string.Empty;
        set
        {
            if (_queryTable != null && value != null)
                _queryTable.Connection = value;
        }
    }

    public int TextFilePlatform
    {
        get => _queryTable != null && _queryTable.TextFilePlatform > 0
            ? _queryTable.TextFilePlatform
            : 0;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFilePlatform = value;
        }
    }

    public int TextFileStartRow
    {
        get => _queryTable != null && _queryTable.TextFileStartRow > 0
            ? _queryTable.TextFileStartRow
            : 0;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileStartRow = value;
        }

    }

    public XlWebFormatting WebFormatting
    {
        get => _queryTable != null
            ? _queryTable.WebFormatting.EnumConvert(XlWebFormatting.xlWebFormattingNone)
            : XlWebFormatting.xlWebFormattingNone;
        set
        {
            if (_queryTable != null)
                _queryTable.WebFormatting = value.EnumConvert(MsExcel.XlWebFormatting.xlWebFormattingNone);
        }
    }

    public XlTextVisualLayoutType TextFileVisualLayout
    {
        get => _queryTable != null
            ? _queryTable.TextFileVisualLayout.EnumConvert(XlTextVisualLayoutType.xlTextVisualLTR)
            : XlTextVisualLayoutType.xlTextVisualLTR;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileVisualLayout = value.EnumConvert(MsExcel.XlTextVisualLayoutType.xlTextVisualLTR);
        }

    }

    public string WebTables
    {
        get => _queryTable?.WebTables?.ToString() ?? string.Empty;
        set
        {
            if (_queryTable != null && value != null)
                _queryTable.WebTables = value;
        }

    }

    public bool WebPreFormattedTextToColumns
    {
        get => _queryTable != null && _queryTable.WebPreFormattedTextToColumns;
        set
        {
            if (_queryTable != null)
                _queryTable.WebPreFormattedTextToColumns = value;
        }

    }

    public bool WebSingleBlockTextImport
    {
        get => _queryTable != null && _queryTable.WebSingleBlockTextImport;
        set
        {
            if (_queryTable != null)
                _queryTable.WebSingleBlockTextImport = value;
        }
    }

    public bool WebDisableDateRecognition
    {
        get => _queryTable != null && _queryTable.WebDisableDateRecognition;
        set
        {
            if (_queryTable != null)
                _queryTable.WebDisableDateRecognition = value;
        }
    }

    public bool WebConsecutiveDelimitersAsOne
    {
        get => _queryTable != null && _queryTable.WebConsecutiveDelimitersAsOne;
        set
        {
            if (_queryTable != null)
                _queryTable.WebConsecutiveDelimitersAsOne = value;
        }

    }

    public bool WebDisableRedirections
    {
        get => _queryTable != null && _queryTable.WebDisableRedirections;
        set
        {
            if (_queryTable != null)
                _queryTable.WebDisableRedirections = value;
        }

    }

    public string SourceConnectionFile
    {
        get => _queryTable?.SourceConnectionFile?.ToString() ?? string.Empty;
        set
        {
            if (_queryTable != null && value != null)
                _queryTable.SourceConnectionFile = value;
        }

    }

    public string SourceDataFile
    {
        get => _queryTable?.SourceDataFile?.ToString() ?? string.Empty;
        set
        {
            if (_queryTable != null && value != null)
                _queryTable.SourceDataFile = value;
        }
    }

    public bool TextFileTrailingMinusNumbers
    {
        get => _queryTable != null && _queryTable.TextFileTrailingMinusNumbers;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFileTrailingMinusNumbers = value;
        }
    }

    /// <summary>
    /// 获取或设置用于获取数据的 SQL 查询语句或命令文本。
    /// </summary>
    public string CommandText
    {
        get => _queryTable?.CommandText?.ToString() ?? string.Empty;
        set
        {
            if (_queryTable != null && value != null)
                _queryTable.CommandText = value;
        }
    }

    public string PostText
    {
        get => _queryTable?.PostText ?? string.Empty;
        set
        {
            if (_queryTable != null && value != null)
                _queryTable.PostText = value;
        }

    }

    /// <summary>
    /// 获取或设置命令类型（如 SQL、表、存储过程等）。
    /// </summary>
    public XlCmdType CommandType
    {
        get => _queryTable != null
            ? _queryTable.CommandType.EnumConvert(XlCmdType.xlCmdSql)
            : XlCmdType.xlCmdSql;

        set
        {
            if (_queryTable != null)
                _queryTable.CommandType = value.EnumConvert(MsExcel.XlCmdType.xlCmdSql);
        }
    }

    /// <summary>
    /// 获取或设置数据刷新时是否提示用户输入参数。
    /// </summary>
    public bool TextFilePromptOnRefresh
    {
        get => _queryTable != null && _queryTable.TextFilePromptOnRefresh;
        set
        {
            if (_queryTable != null)
                _queryTable.TextFilePromptOnRefresh = value;
        }
    }

    public bool SavePassword
    {
        get => _queryTable != null && _queryTable.SavePassword;
        set
        {
            if (_queryTable != null)
                _queryTable.SavePassword = value;
        }
    }

    /// <summary>
    /// 获取或设置第一行是否包含字段名（列标题）。
    /// </summary>
    public bool FieldNames
    {
        get => _queryTable != null && _queryTable.FieldNames;
        set
        {
            if (_queryTable != null)
                _queryTable.FieldNames = value;
        }
    }

    /// <summary>
    /// 获取或设置第一列是否包含行号。
    /// </summary>
    public bool RowNumbers
    {
        get => _queryTable != null && _queryTable.RowNumbers;
        set
        {
            if (_queryTable != null)
                _queryTable.RowNumbers = value;
        }
    }

    public bool HasAutoFormat
    {
        get => _queryTable != null && _queryTable.HasAutoFormat;
        set
        {
            if (_queryTable != null)
                _queryTable.HasAutoFormat = value;
        }

    }

    /// <summary>
    /// 获取或设置是否在数据前插入空行。
    /// </summary>
    public bool FillAdjacentFormulas
    {
        get => _queryTable != null && _queryTable.FillAdjacentFormulas;
        set
        {
            if (_queryTable != null)
                _queryTable.FillAdjacentFormulas = value;
        }
    }

    /// <summary>
    /// 获取或设置是否在刷新时覆盖原有数据。
    /// 语义封装：映射为 RefreshStyle = xlOverwriteCells
    /// </summary>
    public bool OverwriteCells
    {
        get => RefreshStyle == XlCellInsertionMode.xlOverwriteCells;
        set => RefreshStyle = value ? XlCellInsertionMode.xlOverwriteCells : XlCellInsertionMode.xlInsertEntireRows;
    }

    /// <summary>
    /// 获取查询表的数据范围（ResultRange）。
    /// </summary>
    public IExcelRange? ResultRange =>
        _queryTable?.ResultRange != null
            ? new ExcelRange(_queryTable.ResultRange)
            : null;

    /// <summary>
    /// 获取查询表的起始单元格位置（Destination）。
    /// </summary>
    public IExcelRange? Destination =>
        _queryTable?.Destination != null
            ? new ExcelRange(_queryTable.Destination)
            : null;

    public IExcelListObject? ListObject
    {
        get
        {
            var listObject = _queryTable?.ListObject;
            return listObject != null
                ? new ExcelListObject(listObject)
                : null;
        }
    }

    public IExcelSort? Sort
    {
        get
        {
            var sort = _queryTable?.Sort;
            return sort != null
                ? new ExcelSort(sort)
                : null;
        }

    }

    public IExcelWorkbookConnection? WorkbookConnection
    {
        get
        {
            var connection = _queryTable?.Connection as MsExcel.WorkbookConnection;
            return connection != null
                ? new ExcelWorkbookConnection(connection)
                : null;
        }

    }


    public void CancelRefresh()
    {
        _queryTable?.CancelRefresh();
    }

    /// <summary>
    /// 刷新查询表数据。
    /// </summary>
    /// <param name="backgroundQuery">是否后台异步刷新。</param>
    /// <returns>刷新是否成功（true=成功，false=失败或取消）。</returns>
    public bool Refresh(bool backgroundQuery = false)
    {
        if (_queryTable == null) return false;
        try
        {
            return _queryTable.Refresh(backgroundQuery);
        }
        catch
        {
            return false;
        }
    }

    public void ResetTimer()
    {
        _queryTable?.ResetTimer();
    }

    /// <summary>
    /// 删除此查询表（不会删除已导入的数据，仅解除查询绑定）。
    /// </summary>
    public void Delete()
    {
        _queryTable?.Delete();
    }
}