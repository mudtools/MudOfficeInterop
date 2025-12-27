//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 中一个查询表（QueryTable）的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.QueryTable
/// 用于管理从外部数据源（如数据库、Web、文本）导入的数据表。
/// </summary>
public interface IExcelQueryTable : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置查询表的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置是否在刷新时保持列宽不变。
    /// </summary>
    bool PreserveFormatting { get; set; }

    /// <summary>
    /// 获取或设置是否在刷新后调整列宽以适应内容。
    /// </summary>
    bool AdjustColumnWidth { get; set; }

    /// <summary>
    /// 获取与查询表关联的列表对象（ListObject）。
    /// </summary>
    IExcelListObject? ListObject { get; }

    /// <summary>
    /// 获取与查询表关联的排序对象（Sort）。
    /// </summary>
    IExcelSort? Sort { get; }

    /// <summary>
    /// 获取与查询表关联的工作簿连接对象（WorkbookConnection）。
    /// </summary>
    IExcelWorkbookConnection? WorkbookConnection { get; }

    /// <summary>
    /// 获取或设置文本文件的视觉布局类型。
    /// </summary>
    XlTextVisualLayoutType TextFileVisualLayout { get; set; }

    /// <summary>
    /// 获取或设置是否应用自动格式。
    /// </summary>
    bool HasAutoFormat { get; set; }

    /// <summary>
    /// 获取一个值，指示查询表当前是否正在刷新数据。
    /// </summary>
    bool Refreshing { get; }

    /// <summary>
    /// 获取一个值，指示查询是否获取的行数超过了可用行数。
    /// </summary>
    bool FetchedRowOverflow { get; }

    /// <summary>
    /// 获取或设置是否保存密码。
    /// </summary>
    bool SavePassword { get; set; }

    /// <summary>
    /// 获取或设置是否启用刷新功能。
    /// </summary>
    bool EnableRefresh { get; set; }

    /// <summary>
    /// 获取或设置发布到网页时附加的文本内容。
    /// </summary>
    string PostText { get; set; }

    /// <summary>
    /// 获取或设置文本文件的平台编码（如 MS-DOS、Windows 等）。
    /// </summary>
    int TextFilePlatform { get; set; }

    /// <summary>
    /// 获取或设置文本文件开始解析的行号（从1开始）。
    /// </summary>
    int TextFileStartRow { get; set; }

    /// <summary>
    /// 获取或设置文本文件的解析类型（分隔符或固定宽度）。
    /// </summary>
    XlTextParsingType TextFileParseType { get; set; }

    /// <summary>
    /// 获取或设置文本文件的文本识别符类型（如双引号、单引号等）。
    /// </summary>
    XlTextQualifier TextFileTextQualifier { get; set; }

    /// <summary>
    /// 获取或设置是否将连续的分隔符视为单个分隔符。
    /// </summary>
    bool TextFileConsecutiveDelimiter { get; set; }

    /// <summary>
    /// 获取或设置制表符是否作为文本分隔符使用。
    /// </summary>
    bool TextFileTabDelimiter { get; set; }

    /// <summary>
    /// 获取或设置分号是否作为文本分隔符使用。
    /// </summary>
    bool TextFileSemicolonDelimiter { get; set; }

    /// <summary>
    /// 获取或设置逗号是否作为文本分隔符使用。
    /// </summary>
    bool TextFileCommaDelimiter { get; set; }

    /// <summary>
    /// 获取或设置空格是否作为文本分隔符使用。
    /// </summary>
    bool TextFileSpaceDelimiter { get; set; }

    /// <summary>
    /// 获取或设置其他自定义文本分隔符。
    /// </summary>
    string TextFileOtherDelimiter { get; set; }

    /// <summary>
    /// 获取或设置是否保留列信息（如列宽等）。
    /// </summary>
    bool PreserveColumnInfo { get; set; }

    /// <summary>
    /// 获取或设置文本文件的小数分隔符。
    /// </summary>
    string TextFileDecimalSeparator { get; set; }

    /// <summary>
    /// 获取或设置文本文件的千位分隔符。
    /// </summary>
    string TextFileThousandsSeparator { get; set; }

    /// <summary>
    /// 获取或设置是否在查询完成后保持连接状态。
    /// </summary>
    bool MaintainConnection { get; set; }

    /// <summary>
    /// 获取查询的类型（如 ODBC、Web 查询、文本导入等）。
    /// </summary>
    XlQueryType QueryType { get; }

    /// <summary>
    /// 获取或设置 Web 查询的格式化方式（如全部格式、仅 RTF 等）。
    /// </summary>
    XlWebFormatting WebFormatting { get; set; }

    /// <summary>
    /// 获取或设置要从网页中导入的表格的标识符（ID 或索引）列表，多个标识符用逗号分隔。
    /// </summary>
    string WebTables { get; set; }

    /// <summary>
    /// 获取或设置是否将预格式化的文本按列拆分导入到单元格中。
    /// </summary>
    bool WebPreFormattedTextToColumns { get; set; }

    /// <summary>
    /// 获取或设置是否将网页中的所有数据作为单个文本块导入，而不是按列拆分。
    /// </summary>
    bool WebSingleBlockTextImport { get; set; }

    /// <summary>
    /// 获取或设置是否禁用日期识别功能（将日期数据视为普通文本）。
    /// </summary>
    bool WebDisableDateRecognition { get; set; }

    /// <summary>
    /// 获取或设置是否将连续的分隔符视为单个分隔符（Web 查询使用）。
    /// </summary>
    bool WebConsecutiveDelimitersAsOne { get; set; }

    /// <summary>
    /// 获取或设置是否禁用网页重定向（即是否跟随 HTTP 重定向）。
    /// </summary>
    bool WebDisableRedirections { get; set; }

    /// <summary>
    /// 获取或设置源连接文件的路径（如 .odc 文件）。
    /// </summary>
    string SourceConnectionFile { get; set; }

    /// <summary>
    /// 获取或设置源数据文件的路径（实际包含数据的文件）。
    /// </summary>
    string SourceDataFile { get; set; }

    /// <summary>
    /// 获取或设置是否将尾随负号（如 123- 表示 -123）识别为负数。
    /// </summary>
    bool TextFileTrailingMinusNumbers { get; set; }

    /// <summary>
    /// 获取或设置是否在刷新时覆盖单元格格式。
    /// </summary>
    XlCellInsertionMode RefreshStyle { get; set; }

    /// <summary>
    /// 获取或设置是否保存查询数据（即使连接断开）。
    /// </summary>
    bool SaveData { get; set; }

    /// <summary>
    /// 获取或设置是否在打开工作簿时自动刷新。
    /// </summary>
    bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// 获取或设置背景刷新（异步刷新）。
    /// true = 异步刷新，不阻塞 UI。
    /// </summary>
    bool BackgroundQuery { get; set; }

    /// <summary>
    /// 获取或设置连接字符串（如 ODBC、OLEDB、Web URL 等）。
    /// </summary>
    string Connection { get; set; }

    /// <summary>
    /// 获取或设置用于获取数据的 SQL 查询语句或命令文本。
    /// </summary>
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置命令类型（如 SQL、表、存储过程等）。
    /// 使用 <see cref="XlCmdType"/> 枚举。
    /// </summary>
    XlCmdType CommandType { get; set; }

    /// <summary>
    /// 获取或设置数据刷新时是否提示用户输入参数。
    /// </summary>
    bool TextFilePromptOnRefresh { get; set; }


    /// <summary>
    /// 获取或设置第一行是否包含字段名（列标题）。
    /// </summary>
    bool FieldNames { get; set; }

    /// <summary>
    /// 获取或设置第一列是否包含行号。
    /// </summary>
    bool RowNumbers { get; set; }

    /// <summary>
    /// 获取或设置是否在数据前插入空行。
    /// </summary>
    bool FillAdjacentFormulas { get; set; }

    /// <summary>
    /// 获取或设置是否在刷新时覆盖原有数据。
    /// </summary>
    bool OverwriteCells { get; set; }

    /// <summary>
    /// 获取查询表的数据范围（ResultRange）。
    /// 返回封装后的 <see cref="IExcelRange"/>。
    /// </summary>
    IExcelRange? ResultRange { get; }

    /// <summary>
    /// 获取查询表的起始单元格位置（Destination）。
    /// 返回封装后的 <see cref="IExcelRange"/>。
    /// </summary>
    IExcelRange? Destination { get; }

    /// <summary>
    /// 刷新查询表数据。
    /// </summary>
    /// <param name="backgroundQuery">是否后台异步刷新。</param>
    /// <returns>刷新是否成功。</returns>
    bool Refresh(bool backgroundQuery = false);

    /// <summary>
    /// 删除此查询表（不会删除已导入的数据，仅解除查询绑定）。
    /// </summary>
    void Delete();

    /// <summary>
    /// 取消正在进行的查询表刷新操作。
    /// </summary>
    void CancelRefresh();

    /// <summary>
    /// 重置查询表的自动刷新计时器。
    /// </summary>
    void ResetTimer();
}