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
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelQueryTable : IOfficeObject<IExcelQueryTable, MsExcel.QueryTable>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置查询表的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示数据源的字段名称是否作为返回数据的列标题出现。默认值为True。
    /// </summary>
    bool FieldNames { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否将行号添加为指定查询表的第一列。
    /// </summary>
    bool RowNumbers { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示每当刷新查询表时，查询表右侧的公式是否自动更新。
    /// </summary>
    bool FillAdjacentFormulas { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示每次打开工作簿时是否自动更新数据透视表缓存或查询表。默认值为False。
    /// </summary>
    bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// 获取一个布尔值，表示指定的查询表是否有后台查询正在进行。
    /// </summary>
    bool Refreshing { get; }

    /// <summary>
    /// 获取一个布尔值，表示最后一次使用Refresh方法返回的行数是否大于工作表中可用的行数。
    /// </summary>
    bool FetchedRowOverflow { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否异步执行数据透视表报告或查询表的查询。
    /// </summary>
    bool BackgroundQuery { get; set; }

    /// <summary>
    /// 取消指定查询表的所有后台查询。
    /// </summary>
    void CancelRefresh();

    /// <summary>
    /// 获取或设置工作表上添加或删除行的方式，以容纳查询返回的记录集中的行数。
    /// </summary>
    XlCellInsertionMode RefreshStyle { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示用户是否可以刷新数据透视表缓存或查询表。默认值为True。
    /// </summary>
    bool EnableRefresh { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否将ODBC连接字符串中的密码信息与指定的查询一起保存。False表示删除密码。
    /// </summary>
    bool SavePassword { get; set; }

    /// <summary>
    /// 获取查询表目标范围左上角的单元格。
    /// </summary>
    IExcelRange? Destination { get; }

    /// <summary>
    /// 获取或设置包含查询表连接信息的字符串。
    /// </summary>
    object Connection { get; set; }

    /// <summary>
    /// 获取或设置用于将数据输入到Web服务器以从Web查询返回数据的post方法字符串。
    /// </summary>
    string PostText { get; set; }

    /// <summary>
    /// 获取表示工作表被指定查询表占用的区域的Range对象。
    /// </summary>
    IExcelRange? ResultRange { get; }

    /// <summary>
    /// 删除查询表。
    /// </summary>
    void Delete();

    /// <summary>
    /// 更新外部数据范围。
    /// </summary>
    /// <param name="backgroundQuery">可选。仅适用于基于SQL查询结果的QueryTable。True表示在建立数据库连接并提交查询后立即将控制权返回给过程，在后台更新查询表。False表示仅在所有数据提取到工作表后才将控制权返回给过程。</param>
    /// <returns>刷新操作的结果。</returns>
    bool? Refresh(bool? backgroundQuery = null);

    /// <summary>
    /// 获取表示查询表参数的Parameters集合。
    /// </summary>
    IExcelParameters? Parameters { get; }

    /// <summary>
    /// 获取或设置用作指定查询表或数据透视表缓存数据源的Recordset对象。
    /// </summary>
    object Recordset { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否将数据透视表报告的数据与工作簿一起保存。False表示仅保存报告定义。
    /// </summary>
    bool SaveData { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示用户是否可以编辑指定的查询表。False表示用户只能刷新查询表。
    /// </summary>
    bool EnableEditing { get; set; }

    /// <summary>
    /// 获取或设置导入到查询表中的文本文件的来源。此属性确定在数据导入期间使用哪个代码页。
    /// </summary>
    int TextFilePlatform { get; set; }

    /// <summary>
    /// 获取或设置将文本文件导入查询表时开始文本解析的行号。有效值为1到32767的整数。默认值为1。
    /// </summary>
    int TextFileStartRow { get; set; }

    /// <summary>
    /// 获取或设置导入到查询表中的文本文件的数据列格式。
    /// </summary>
    XlTextParsingType TextFileParseType { get; set; }

    /// <summary>
    /// 获取或设置将文本文件导入查询表时的文本限定符。
    /// </summary>
    XlTextQualifier TextFileTextQualifier { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将文本文件导入查询表时是否将连续分隔符视为单个分隔符。默认值为False。
    /// </summary>
    bool TextFileConsecutiveDelimiter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示制表符是否是导入文本文件时的分隔符。默认值为False。
    /// </summary>
    bool TextFileTabDelimiter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示分号是否是导入文本文件时的分隔符。默认值为False。
    /// </summary>
    bool TextFileSemicolonDelimiter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示逗号是否是导入文本文件时的分隔符。False表示使用其他字符作为分隔符。默认值为False。
    /// </summary>
    bool TextFileCommaDelimiter { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示空格字符是否是导入文本文件时的分隔符。默认值为False。
    /// </summary>
    bool TextFileSpaceDelimiter { get; set; }

    /// <summary>
    /// 获取或设置将文本文件导入查询表时用作分隔符的字符。默认值为null。
    /// </summary>
    string TextFileOtherDelimiter { get; set; }

    /// <summary>
    /// 获取或设置指定应用于导入到查询表中的文本文件对应列的数据类型的有序数组。每列的默认常量为xlGeneral。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColumnDataType TextFileColumnDataTypes { get; set; }

    /// <summary>
    /// 获取或设置对应于导入到查询表中的文本文件列宽度的整数数组。有效宽度为1到32767个字符。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int TextFileFixedColumnWidths { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示每当刷新查询表时是否保留列排序、筛选和布局信息。默认值为False。
    /// </summary>
    bool PreserveColumnInfo { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，确定是否将常见于前五行数据的任何格式应用于查询表中的新数据行。未使用的单元格不会被格式化。
    /// </summary>
    bool PreserveFormatting { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示每次刷新指定查询表或XML映射时是否自动调整列宽以获得最佳拟合。默认值为True。
    /// </summary>
    bool AdjustColumnWidth { get; set; }

    /// <summary>
    /// 获取或设置指定数据源的命令字符串。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置描述与CommandText属性关联的命令类型的XlCmdType常量。默认值为xlCmdSQL。
    /// </summary>
    XlCmdType CommandType { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示每次刷新查询表时是否要指定导入的文本文件的名称。导入文本文件对话框允许指定路径和文件名。默认值为False。
    /// </summary>
    bool TextFilePromptOnRefresh { get; set; }

    /// <summary>
    /// 获取Excel用于填充查询表或数据透视表缓存的查询类型。
    /// </summary>
    XlQueryType QueryType { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示刷新后是否保持与指定数据源的连接直到工作簿关闭。默认值为True。
    /// </summary>
    bool MaintainConnection { get; set; }

    /// <summary>
    /// 获取或设置将文本文件导入查询表时Excel使用的小数分隔符字符。默认值为系统小数分隔符字符。
    /// </summary>
    string TextFileDecimalSeparator { get; set; }

    /// <summary>
    /// 获取或设置将文本文件导入查询表时Excel使用的千位分隔符字符。默认值为系统千位分隔符字符。
    /// </summary>
    string TextFileThousandsSeparator { get; set; }

    /// <summary>
    /// 获取或设置刷新之间的分钟数。
    /// </summary>
    int RefreshPeriod { get; set; }

    /// <summary>
    /// 将指定查询表或数据透视表报告的刷新计时器重置为使用RefreshPeriod属性设置的最后一个间隔。
    /// </summary>
    void ResetTimer();

    /// <summary>
    /// 获取或设置确定整个网页、网页上的所有表还是仅网页上的特定表导入查询表的值。
    /// </summary>
    XlWebSelectionType WebSelectionType { get; set; }

    /// <summary>
    /// 获取或设置确定将网页导入查询表时应用多少格式的值。
    /// </summary>
    XlWebFormatting WebFormatting { get; set; }

    /// <summary>
    /// 获取或设置将网页导入查询表时的表名或表索引号的逗号分隔列表。
    /// </summary>
    string WebTables { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将网页导入查询表时，网页中HTML &lt;PRE&gt;标签内的数据是否被解析为列。默认值为True。
    /// </summary>
    bool WebPreFormattedTextToColumns { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将网页导入查询表时，HTML &lt;PRE&gt;标签中的数据是否一次性处理。False表示以连续行的块导入数据，以便识别标题行。默认值为False。
    /// </summary>
    bool WebSingleBlockTextImport { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示将网页导入查询表时，类似于日期的数据是否作为文本解析。False表示使用日期识别。默认值为False。
    /// </summary>
    bool WebDisableDateRecognition { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示从网页的HTML &lt;PRE&gt;标签导入数据到查询表时，如果数据要解析为列，是否将连续分隔符视为单个分隔符。False表示将连续分隔符视为多个分隔符。默认值为True。
    /// </summary>
    bool WebConsecutiveDelimitersAsOne { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否为QueryTable对象禁用Web查询重定向。默认值为False。
    /// </summary>
    bool WebDisableRedirections { get; set; }

    /// <summary>
    /// 获取或设置Web查询的网页统一资源定位符。
    /// </summary>
    object EditWebPage { get; set; }

    /// <summary>
    /// 获取或设置用于创建数据透视表的Microsoft Office数据连接文件或类似文件的字符串。
    /// </summary>
    string SourceConnectionFile { get; set; }

    /// <summary>
    /// 获取或设置表示查询表源数据文件的字符串。
    /// </summary>
    string SourceDataFile { get; set; }

    /// <summary>
    /// 获取或设置数据透视表缓存如何连接到其数据源。
    /// </summary>
    XlRobustConnect RobustConnect { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示Excel是否将以"-"符号开头作为文本导入的数字视为负数。False表示Excel将作为文本导入的数字视为文本。
    /// </summary>
    bool TextFileTrailingMinusNumbers { get; set; }

    /// <summary>
    /// 将数据透视表缓存源保存为Microsoft Office数据连接文件。
    /// </summary>
    /// <param name="odcFileName">必需。文件保存的位置。</param>
    /// <param name="description">可选。将保存在文件中的描述。</param>
    /// <param name="keywords">可选。可用于搜索此文件的关键字。</param>
    void SaveAsODC(string odcFileName, string? description = null, string? keywords = null);

    /// <summary>
    /// 获取Range对象或QueryTable对象的ListObject对象。
    /// </summary>
    IExcelListObject? ListObject { get; }

    /// <summary>
    /// 获取或设置一个XlTextVisualLayoutType常量，指示正在导入的文本的可视布局是从左到右还是从右到左。
    /// </summary>
    XlTextVisualLayoutType TextFileVisualLayout { get; set; }

    /// <summary>
    /// 获取查询表使用的工作簿连接对象。
    /// </summary>
    IExcelWorkbookConnection? WorkbookConnection { get; }

    /// <summary>
    /// 获取查询表范围的排序条件。
    /// </summary>
    IExcelSort? Sort { get; }
}