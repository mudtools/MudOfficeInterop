//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Workbooks 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Workbooks 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelWorkbooks : IEnumerable<IExcelWorkbook?>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取集合中工作簿的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的工作簿对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">工作簿索引（从1开始）</param>
    /// <returns>工作簿对象</returns>
    IExcelWorkbook? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的工作簿对象
    /// </summary>
    /// <param name="name">工作簿名称</param>
    /// <returns>工作簿对象</returns>
    IExcelWorkbook? this[string name] { get; }

    /// <summary>
    /// 获取工作簿集合所在的父对象（通常是Application）
    /// 对应 Workbooks.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取工作簿集合所在的Application对象
    /// 对应 Workbooks.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    #endregion

    /// <summary>
    /// 创建一个新工作簿。新工作簿成为活动工作簿。
    /// </summary>
    /// <param name="template">可选。确定如何创建新工作簿。如果此参数是指定现有Excel文件名的字符串，则以指定文件为模板创建新工作簿。如果此参数是常量，则新工作簿包含指定类型的单个工作表。</param>
    /// <returns>新创建的Workbook对象。</returns>
    IExcelWorkbook? Add(string? template = null);

    /// <summary>
    /// 创建一个新工作簿。新工作簿成为活动工作簿。
    /// </summary>
    /// <param name="template">可选。确定如何创建新工作簿。可以是以下XlWBATemplate常量之一：xlWBATChart、xlWBATExcel4IntlMacroSheet、xlWBATExcel4MacroSheet或xlWBATWorksheet。如果省略此参数，则Excel创建一个包含多个空白工作表的新工作簿（工作表数量由SheetsInNewWorkbook属性设置）。</param>
    /// <returns>新创建的Workbook对象。</returns>
    IExcelWorkbook? Add(XlWBATemplate template);

    /// <summary>
    /// 关闭Workbooks集合中的所有工作簿。
    /// </summary>
    void Close();


    /// <summary>
    /// 打开一个工作簿。
    /// </summary>
    /// <param name="filename">必需。要打开的工作簿的文件名。</param>
    /// <param name="updateLinks">可选。指定文件中链接的更新方式。如果省略此参数，则会提示用户指定链接更新方式。否则，此参数为下表中列出的值之一。如果Excel正在打开WKS、WK1或WK3格式的文件且UpdateLinks参数为2，则Excel会从文件附加的图形生成图表。如果参数为0，则不创建图表。</param>
    /// <param name="readOnly">可选。True表示以只读模式打开工作簿。</param>
    /// <param name="format">可选。如果Excel正在打开文本文件，此参数指定分隔符字符，如下表所示。如果省略此参数，则使用当前分隔符。</param>
    /// <param name="password">可选。包含打开受保护工作簿所需密码的字符串。如果省略此参数且工作簿需要密码，则会提示用户输入密码。</param>
    /// <param name="writeResPassword">可选。包含写入写保护工作簿所需密码的字符串。如果省略此参数且工作簿需要密码，则会提示用户输入密码。</param>
    /// <param name="ignoreReadOnlyRecommended">可选。True表示让Excel不显示"建议只读"消息（如果工作簿保存时启用了"建议只读"选项）。</param>
    /// <param name="origin">可选。如果文件是文本文件，此参数指示其来源（以便正确映射代码页和回车/换行符）。可以是以下XlPlatform常量之一：xlMacintosh、xlWindows或xlMSDOS。如果省略此参数，则使用当前操作系统。</param>
    /// <param name="delimiter">可选。如果文件是文本文件且Format参数为6，则此参数是指定用作分隔符的字符的字符串。例如，使用Chr(9)表示制表符，使用","表示逗号，使用";"表示分号，或使用自定义字符。仅使用字符串的第一个字符。</param>
    /// <param name="editable">可选。如果文件是Excel 4.0加载项，此参数为True时将以可见窗口打开加载项。如果此参数为False或省略，则加载项将作为隐藏状态打开，且无法取消隐藏。此选项不适用于Excel 5.0或更高版本创建的加载项。如果文件是Excel模板，使用True可打开指定模板进行编辑，使用False可基于指定模板打开新工作簿。默认值为False。</param>
    /// <param name="notify">可选。如果文件无法以读/写模式打开，此参数为True会将文件添加到文件通知列表。Excel将以只读方式打开文件，轮询文件通知列表，然后在文件可用时通知用户。如果此参数为False或省略，则不请求通知，任何打开不可用文件的尝试都将失败。</param>
    /// <param name="converter">可选。打开文件时尝试使用的第一个文件转换器的索引。指定的文件转换器首先尝试；如果此转换器无法识别文件，则尝试所有其他转换器。转换器索引由Application.FileConverters属性返回的转换器的行号组成。</param>
    /// <param name="addToMru">可选。True表示将此工作簿添加到最近使用的文件列表。默认值为False。</param>
    /// <param name="local">可选。True表示根据Excel的语言（包括控制面板设置）保存文件。False（默认）表示根据VBA的语言（通常是美式英语，除非运行Workbooks.Open的VBA项目是旧的国际化XL5/95 VBA项目）保存文件。</param>
    /// <param name="corruptLoad">可选。可以是以下常量之一：xlNormalLoad、xlRepairFile和xlExtractData。如果未指定值，默认行为通常是正常加载，但如果Excel已尝试打开文件，则可能是安全加载或数据恢复。第一次尝试是正常加载。如果Excel在打开文件时停止运行，第二次尝试是安全加载。如果Excel再次停止运行，则下一次尝试是数据恢复。</param>
    /// <returns>打开的Workbook对象。</returns>
    IExcelWorkbook? Open(string filename, object? updateLinks = null,
                        bool? readOnly = null, string? format = null,
                        string? password = null, string? writeResPassword = null,
                        bool? ignoreReadOnlyRecommended = null, XlPlatform? origin = null,
                        string? delimiter = null, bool? editable = null,
                        bool? notify = null, object? converter = null,
                        bool? addToMru = null, bool? local = null,
                        object? corruptLoad = null);

    /// <summary>
    /// 加载并解析文本文件作为新工作簿，其中包含单个工作表，该工作表包含解析后的文本文件数据。
    /// </summary>
    /// <param name="filename">必需。指定要打开和解析的文本文件的文件名。</param>
    /// <param name="origin">可选。指定文本文件的来源。可以是以下XlPlatform常量之一：xlMacintosh、xlWindows或xlMSDOS。此外，这可以是一个表示所需代码页编号的整数。例如，"1256"指定源文本文件的编码为阿拉伯语（Windows）。如果省略此参数，则方法使用文本导入向导中"文件来源"选项的当前设置。</param>
    /// <param name="startRow">可选。开始解析文本的行号。默认值为1。</param>
    /// <param name="dataType">可选。指定文件中数据的列格式。可以是以下XlTextParsingType常量之一：xlDelimited或xlFixedWidth。如果未指定此参数，Excel将在打开文件时尝试确定列格式。</param>
    /// <param name="textQualifier">可选。指定文本限定符。可以是以下XlTextQualifier常量之一：xlTextQualifierDoubleQuote（默认）、xlTextQualifierNone、xlTextQualifierSingleQuote。</param>
    /// <param name="consecutiveDelimiter">可选。True表示将连续分隔符视为一个分隔符。默认为False。</param>
    /// <param name="tab">可选。True表示将制表符作为分隔符（DataType必须为xlDelimited）。默认值为False。</param>
    /// <param name="semicolon">可选。True表示将分号字符作为分隔符（DataType必须为xlDelimited）。默认值为False。</param>
    /// <param name="comma">可选。True表示将逗号字符作为分隔符（DataType必须为xlDelimited）。默认值为False。</param>
    /// <param name="space">可选。True表示将空格字符作为分隔符（DataType必须为xlDelimited）。默认值为False。</param>
    /// <param name="other">可选。True表示将OtherChar参数指定的字符作为分隔符（DataType必须为xlDelimited）。默认值为False。</param>
    /// <param name="otherChar">可选（如果Other为True则必需）。当Other为True时指定分隔符字符。如果指定多个字符，则仅使用字符串的第一个字符；其余字符将被忽略。</param>
    /// <param name="fieldInfo">可选。包含单个数据列解析信息的数组。解释取决于DataType的值。当数据为分隔格式时，此参数是二维数组的数组，每个二维数组指定特定列的转换选项。第一个元素是列号（从1开始），第二个元素是XlColumnDataType常量之一，指定如何解析列。列说明符可以按任意顺序排列。如果输入数据中特定列没有列说明符，则该列使用常规设置进行解析。</param>
    /// <param name="textVisualLayout">可选。文本的可视布局。</param>
    /// <param name="decimalSeparator">可选。Excel识别数字时使用的小数分隔符。默认设置为系统设置。</param>
    /// <param name="thousandsSeparator">可选。Excel识别数字时使用的千位分隔符。默认设置为系统设置。</param>
    /// <param name="trailingMinusNumbers">可选。处理尾部负号数字。</param>
    /// <param name="local">可选。本地化设置。</param>
    void OpenText(string filename, int? origin = null, int? startRow = null,
                XlTextParsingType? dataType = null, XlTextQualifier textQualifier = XlTextQualifier.xlTextQualifierDoubleQuote,
                bool? consecutiveDelimiter = null, bool? tab = null,
                bool? semicolon = null, bool? comma = null,
                bool? space = null, bool? other = null, string? otherChar = null,
                object? fieldInfo = null, object? textVisualLayout = null,
                string? decimalSeparator = null, string? thousandsSeparator = null,
                string? trailingMinusNumbers = null, object? local = null);

    /// <summary>
    /// 返回表示数据库的Workbook对象。
    /// </summary>
    /// <param name="filename">必需。连接字符串。</param>
    /// <param name="commandText">可选。查询的命令文本。</param>
    /// <param name="commandType">可选。查询的命令类型。可用的命令类型有：Default、SQL和Table。</param>
    /// <param name="backgroundQuery">可选。查询的背景。</param>
    /// <param name="importDataAs">可选。确定查询的格式。</param>
    /// <returns>数据库工作簿对象。</returns>
    IExcelWorkbook? OpenDatabase(string filename, string? commandText = null,
                                 string? commandType = null, string? backgroundQuery = null, string? importDataAs = null);

    /// <summary>
    /// 将指定工作簿从服务器复制到本地计算机进行编辑。
    /// </summary>
    /// <param name="filename">必需。要签出的文件名。</param>
    void CheckOut(string filename);

    /// <summary>
    /// 确定Excel是否可以从服务器签出指定工作簿。
    /// </summary>
    /// <param name="filename">必需。要检查的文件名。</param>
    /// <returns>如果可以签出则为True，否则为False。</returns>
    bool? CanCheckOut(string filename);

    /// <summary>
    /// 打开XML数据文件。
    /// </summary>
    /// <param name="filename">必需。要打开的文件名。</param>
    /// <param name="stylesheets">可选。指定要应用的XSL转换（XSLT）样式表处理指令的单个值或值数组。</param>
    /// <param name="loadOption">可选。指定Excel如何打开XML数据文件。可以是以下XlXmlLoadOption常量之一：xlXmlLoadImportToList（将XML数据文件内容放入XML列表）、xlXmlLoadMapXml（在XML结构任务窗格中显示XML数据文件的架构）、xlXmlLoadOpenXml（打开XML数据文件，文件内容将被展平）、xlXmlLoadPromptUser（提示用户选择如何打开文件）。</param>
    /// <returns>打开的XML工作簿对象。</returns>
    IExcelWorkbook? OpenXML(string filename, object? stylesheets = null, XlXmlLoadOption? loadOption = null);
}