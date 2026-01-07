//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Worksheet 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Worksheet 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel", NoneConstructor = true, NoneDisposed = true)]
public interface IExcelWorksheet : IOfficeObject<IExcelWorksheet, MsExcel.Worksheet>, IExcelComSheet, IDisposable
{

    /// <summary>
    /// 获取对象的代码名称。只读字符串。
    /// </summary>
    string CodeName { get; }

    /// <summary>
    /// 获取或设置对象的代码名称。此属性保留供内部使用。
    /// </summary>
    string _CodeName { get; set; }


    /// <summary>
    /// 获取表示下一个工作表或单元格的 Chart、Range 或 Worksheet 对象。只读。
    /// </summary>
    [ComPropertyWrap(PropertyName = "Next", NeedConvert = true)]
    IExcelChart? NextChart { get; }

    /// <summary>
    /// 获取表示下一个工作表或单元格的 Chart、Range 或 Worksheet 对象。只读。
    /// </summary>
    [ComPropertyWrap(PropertyName = "Next", NeedConvert = true)]
    IExcelRange? NextRange { get; }

    /// <summary>
    /// 获取表示下一个工作表或单元格的 Chart、Range 或 Worksheet 对象。只读。
    /// </summary>
    [ComPropertyWrap(PropertyName = "Next", NeedConvert = true)]
    IExcelWorksheet? NextWorksheet { get; }

    /// <summary>
    /// 获取表示前一个工作表或单元格的 Chart、Range 或 Worksheet 对象。只读。
    /// </summary>
    [ComPropertyWrap(PropertyName = "Previous", NeedConvert = true)]
    IExcelChart? PreviousChart { get; }

    /// <summary>
    /// 获取表示前一个工作表或单元格的 Chart、Range 或 Worksheet 对象。只读。
    /// </summary>
    [ComPropertyWrap(PropertyName = "Previous", NeedConvert = true)]
    IExcelRange? PreviousRange { get; }

    /// <summary>
    /// 获取表示前一个工作表或单元格的 Chart、Range 或 Worksheet 对象。只读。
    /// </summary>
    [ComPropertyWrap(PropertyName = "Previous", NeedConvert = true)]
    IExcelWorksheet? PreviousWorksheet { get; }

    /// <summary>
    /// 显示对象的打印预览。
    /// </summary>
    /// <param name="enableChanges">可选项。True 表示允许对指定工作表进行更改。</param>
    void PrintPreview(bool? enableChanges = null);

    /// <summary>
    /// 获取一个值，该值指示形状是否受保护。只读布尔值。
    /// </summary>
    bool ProtectDrawingObjects { get; }

    /// <summary>
    /// 获取一个值，该值指示工作表方案是否受保护。只读布尔值。
    /// </summary>
    bool ProtectScenarios { get; }

    /// <summary>
    /// 选择对象。
    /// </summary>
    /// <param name="replace">可选项。要替换的对象。</param>
    void Select(object? replace = null);

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Excel 是否为工作表使用 Lotus 1-2-3 表达式计算规则。可读写布尔值。
    /// </summary>
    bool TransitionExpEval { get; set; }

    /// <summary>
    /// 获取一个值，该值指示工作表上当前是否显示自动筛选下拉箭头。此属性独立于 FilterMode 属性。可读写布尔值。
    /// </summary>
    bool AutoFilterMode { get; set; }

    /// <summary>
    /// 为工作表或图表设置背景图形。
    /// </summary>
    /// <param name="filename">必需。图形文件的名称。</param>
    void SetBackgroundPicture(string filename);

    /// <summary>
    /// 计算所有打开的工作簿、工作簿中的特定工作表或工作表上的指定单元格区域。
    /// </summary>
    void Calculate();

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Excel 是否在必要时自动重新计算工作表。可读写布尔值。
    /// </summary>
    bool EnableCalculation { get; set; }

    /// <summary>
    /// 获取表示工作表中所有单元格的 Range 对象（不仅仅是当前使用的单元格）。只读。
    /// </summary>
    IExcelRange? Cells { get; }

    /// <summary>
    /// 获取表示单个嵌入图表（ChartObject 对象）或工作表上所有嵌入图表（ChartObjects 对象）的对象。
    /// </summary>
    /// <returns>ChartObject 或 ChartObjects 对象。</returns>
    [ValueConvert]
    IExcelChartObjects? ChartObjects();

    /// <summary>
    /// 获取表示单个嵌入图表（ChartObject 对象）或工作表上所有嵌入图表（ChartObjects 对象）的对象。
    /// </summary>
    /// <param name="name">可选项。图表的名称或编号。此参数可以是数组，用于指定多个图表。</param>
    /// <returns>ChartObject 或 ChartObjects 对象。</returns>
    [ValueConvert]
    IExcelChartObjects? ChartObjects(string[] name);

    /// <summary>
    /// 获取表示单个嵌入图表（ChartObject 对象）或工作表上所有嵌入图表（ChartObjects 对象）的对象。
    /// </summary>
    /// <param name="index">可选项。图表的名称或编号。此参数可以是数组，用于指定多个图表。</param>
    /// <returns>ChartObject 或 ChartObjects 对象。</returns>
    [ValueConvert]
    IExcelChartObjects? ChartObjects(int[] index);

    /// <summary>
    /// 获取表示单个嵌入图表（ChartObject 对象）或工作表上所有嵌入图表（ChartObjects 对象）的对象。
    /// </summary>
    /// <param name="name">可选项。图表的名称或编号。此参数可以是数组，用于指定多个图表。</param>
    /// <returns>ChartObject 或 ChartObjects 对象。</returns>
    [ValueConvert]
    IExcelChartObject? ChartObjects(string name);

    /// <summary>
    /// 获取表示单个嵌入图表（ChartObject 对象）或工作表上所有嵌入图表（ChartObjects 对象）的对象。
    /// </summary>
    /// <param name="index">可选项。图表的名称或编号。此参数可以是数组，用于指定多个图表。</param>
    /// <returns>ChartObject 或 ChartObjects 对象。</returns>
    [ValueConvert]
    IExcelChartObject? ChartObjects(int index);

    /// <summary>
    /// 检查对象的拼写。此形式没有返回值；Microsoft Excel 显示“拼写检查”对话框。
    /// </summary>
    /// <param name="customDictionary">可选项。字符串，指示如果在主词典中找不到单词，则要检查的自定义词典的文件名。如果省略此参数，则使用当前指定的词典。</param>
    /// <param name="ignoreUppercase">可选项。True 表示 Microsoft Excel 忽略全部大写的单词。False 表示 Microsoft Excel 检查全部大写的单词。如果省略此参数，则使用当前设置。</param>
    /// <param name="alwaysSuggest">可选项。True 表示当发现错误拼写时，Microsoft Excel 显示建议的替代拼写列表。False 表示 Microsoft Excel 等待您输入正确的拼写。如果省略此参数，则使用当前设置。</param>
    /// <param name="spellLang">可选项。所用词典的语言。可以是 MsoLanguageID 值之一。</param>
    void CheckSpelling(string? customDictionary = null, bool? ignoreUppercase = null, bool? alwaysSuggest = null, [ComNamespace("MsCore")] MsoLanguageID? spellLang = null);

    /// <summary>
    /// 获取表示工作表中包含第一个循环引用的区域的 Range 对象；如果工作表中没有循环引用，则返回 Nothing。
    /// 必须先删除循环引用，然后才能继续计算。只读。
    /// </summary>
    IExcelRange? CircularReference { get; }

    /// <summary>
    /// 从工作表中清除追踪箭头。追踪箭头是通过审核功能添加的。
    /// </summary>
    void ClearArrows();

    /// <summary>
    /// 获取表示指定工作表中所有列的 Range 对象。只读。
    /// </summary>
    IExcelRange? Columns { get; }

    /// <summary>
    /// 获取用于当前合并的函数代码。可以是 XlConsolidationFunction 常量之一。只读。
    /// </summary>
    XlConsolidationFunction ConsolidationFunction { get; }

    /// <summary>
    /// 获取合并选项的三元素数组。如果元素为 True，则表示设置了该选项。只读对象。
    /// </summary>
    object ConsolidationOptions { get; }

    /// <summary>
    /// 获取字符串值数组，这些字符串值为工作表当前合并的源工作表名称。如果工作表上没有合并，则返回 Empty。只读对象。
    /// </summary>
    object ConsolidationSources { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示仅启用用户界面保护时是否启用自动筛选箭头。可读写布尔值。
    /// </summary>
    bool EnableAutoFilter { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示可以在工作表上选择的内容。可读写 XlEnableSelection。
    /// </summary>
    XlEnableSelection EnableSelection { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示仅启用用户界面保护时是否启用分级显示符号。可读写布尔值。
    /// </summary>
    bool EnableOutlining { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示仅启用用户界面保护时是否启用数据透视表控件和操作。可读写布尔值。
    /// </summary>
    bool EnablePivotTable { get; set; }

    /// <summary>
    /// 将 Microsoft Excel 名称转换为对象或值。
    /// </summary>
    /// <param name="name">必需。对象的名称，使用 Microsoft Excel 的命名约定。</param>
    /// <returns>转换后的对象或值。</returns>
    object? Evaluate(string? name);

    /// <summary>
    /// 获取一个值，该值指示工作表是否处于筛选模式。只读布尔值。
    /// </summary>
    bool FilterMode { get; }

    /// <summary>
    /// 重置指定工作表上的所有分页符。
    /// </summary>
    void ResetAllPageBreaks();

    /// <summary>
    /// 获取表示所有特定于工作表的名称的 Names 集合（使用 "WorksheetName!" 前缀定义的名称）。只读 Names 对象。
    /// </summary>
    IExcelNames? Names { get; }

    /// <summary>
    /// 获取表示图表或工作表上的单个 OLE 对象（OLEObject）或所有 OLE 对象（OLEObjects 集合）的对象。只读。
    /// </summary>
    /// <returns>OLEObject 或 OLEObjects 对象。</returns>
    [ValueConvert]
    IExcelOLEObjects? OLEObjects();

    /// <summary>
    /// 获取表示图表或工作表上的单个 OLE 对象（OLEObject）或所有 OLE 对象（OLEObjects 集合）的对象。只读。
    /// </summary>
    /// <param name="index">可选项。OLE 对象的名称或编号。</param>
    /// <returns>OLEObject 或 OLEObjects 对象。</returns>
    [ValueConvert]
    IExcelOLEObject? OLEObjects(string index);

    /// <summary>
    /// 获取表示图表或工作表上的单个 OLE 对象（OLEObject）或所有 OLE 对象（OLEObjects 集合）的对象。只读。
    /// </summary>
    /// <param name="index">可选项。OLE 对象的名称或编号。</param>
    /// <returns>OLEObject 或 OLEObjects 对象。</returns>
    [ValueConvert]
    IExcelOLEObject? OLEObjects(int index);

    /// <summary>
    /// 获取表示工作表大纲的 Outline 对象。只读。
    /// </summary>
    IExcelOutline? Outline { get; }

    /// <summary>
    /// 将剪贴板的内容粘贴到工作表上。
    /// </summary>
    /// <param name="destination">可选项。Range 对象，指定剪贴板内容应粘贴的位置。如果省略此参数，则使用当前选定区域。仅当剪贴板内容可以粘贴到区域中时才能指定此参数。如果指定了此参数，则不能使用 link 参数。</param>
    /// <param name="link">可选项。True 表示建立与粘贴数据源的链接。如果指定了此参数，则不能使用 destination 参数。默认值为 False。</param>
    void Paste(IExcelRange? destination = null, bool? link = null);

    /// <summary>
    /// 获取表示单个数据透视表报告（PivotTable 对象）或工作表上所有数据透视表报告（PivotTables 对象）的对象。只读。
    /// </summary>
    /// <param name="index">可选项。报告的名称或编号。</param>
    /// <returns>PivotTable 或 PivotTables 对象。</returns>
    [ValueConvert]
    IExcelPivotTable? PivotTables(string index);

    /// <summary>
    /// 获取表示单个数据透视表报告（PivotTable 对象）或工作表上所有数据透视表报告（PivotTables 对象）的对象。只读。
    /// </summary>
    /// <param name="index">可选项。报告的名称或编号。</param>
    /// <returns>PivotTable 或 PivotTables 对象。</returns>
    [ValueConvert]
    IExcelPivotTable? PivotTables(int index);

    /// <summary>
    /// 获取表示单个数据透视表报告（PivotTable 对象）或工作表上所有数据透视表报告（PivotTables 对象）的对象。只读。
    /// </summary>
    /// <returns>PivotTable 或 PivotTables 对象。</returns>
    [ValueConvert]
    IExcelPivotTables? PivotTables();

    /// <summary>
    /// 创建 PivotTable 对象。此方法不显示数据透视表向导。此方法不适用于 OLE DB 数据源。
    /// </summary>
    /// <param name="sourceType">可选项。报告数据的来源。可以是 XlPivotTableSourceType 常量之一。</param>
    /// <param name="sourceData">可选项。新报告的数据。可以是 Range 对象、区域数组或表示另一个报告名称的文本常量。</param>
    /// <param name="tableDestination">可选项。Range 对象，指定报告在工作表上的放置位置。如果省略此参数，报告将放置在活动单元格处。</param>
    /// <param name="tableName">可选项。字符串，指定新报告的名称。</param>
    /// <param name="rowGrand">可选项。True 表示在报告中显示行的总计。</param>
    /// <param name="columnGrand">可选项。True 表示在报告中显示列的总计。</param>
    /// <param name="saveData">可选项。True 表示将数据与报告一起保存。False 表示仅保存报告定义。</param>
    /// <param name="hasAutoFormat">可选项。True 表示在刷新或移动字段时让 Microsoft Excel 自动设置报告格式。</param>
    /// <param name="autoPage">可选项。仅当 sourceType 为 xlConsolidation 时有效。True 表示让 Microsoft Excel 为合并创建页字段。</param>
    /// <param name="reserved">可选项。Microsoft Excel 未使用。</param>
    /// <param name="backgroundQuery">可选项。True 表示让 Excel 异步执行报告的查询（在后台）。默认值为 False。</param>
    /// <param name="optimizeCache">可选项。True 表示在构造数据透视表缓存时对其进行优化。默认值为 False。</param>
    /// <param name="pageFieldOrder">可选项。页字段添加到数据透视表报告布局的顺序。可以是 XlOrder 常量之一：xlDownThenOver 或 xlOverThenDown。默认值为 xlDownThenOver。</param>
    /// <param name="pageFieldWrapCount">可选项。数据透视表报告中每列或每行的页字段数。默认值为 0（零）。</param>
    /// <param name="readData">可选项。True 表示创建包含外部数据库中所有记录的数据透视表缓存；此缓存可能非常大。</param>
    /// <param name="connection">可选项。字符串，包含允许 Excel 连接到 ODBC 数据源的 ODBC 设置。</param>
    /// <returns>创建的 PivotTable 对象。</returns>
    IExcelPivotTable? PivotTableWizard(XlPivotTableSourceType? sourceType = null, IExcelRange? sourceData = null,
                                    IExcelRange? tableDestination = null, string? tableName = null,
                                    bool? rowGrand = null, bool? columnGrand = null,
                                    bool? saveData = null, bool? hasAutoFormat = null,
                                    bool? autoPage = null, object? reserved = null,
                                    bool? backgroundQuery = null, bool? optimizeCache = null,
                                    XlOrder? pageFieldOrder = null, int? pageFieldWrapCount = null,
                                    bool? readData = null, string? connection = null);

    /// <summary>
    /// 获取工作表中指定范围的区域对象
    /// </summary>
    /// <param name="cell1">起始单元格</param>
    /// <param name="cell2">结束单元格（可选）</param>
    /// <returns>区域对象</returns>
    [IgnoreGenerator]
    IExcelRange? Range(object? cell1, object? cell2 = null);

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    [IgnoreGenerator]
    IExcelRange? this[string address] { get; }
    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <param name="column">列号</param>
    /// <returns>单元格对象</returns>
    [IgnoreGenerator]
    IExcelRange? this[int row, int column] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    [IgnoreGenerator]
    IExcelRange? this[string begin, string end] { get; }

    /// <summary>
    /// 获取表示指定工作表中所有行的 Range 对象。只读 Range 对象。
    /// </summary>
    IExcelRange? Rows { get; }

    /// <summary>
    /// 获取表示单个方案（Scenario 对象）或工作表上所有方案（Scenarios 对象）的对象。
    /// </summary>
    /// <returns>Scenario 或 Scenarios 对象。</returns>
    [ValueConvert]
    IExcelScenarios? Scenarios();

    /// <summary>
    /// 获取表示单个方案（Scenario 对象）或工作表上所有方案（Scenarios 对象）的对象。
    /// </summary>
    /// <param name="index">可选项。方案的名称或编号。可以使用数组指定多个方案。</param>
    /// <returns>Scenario 或 Scenarios 对象。</returns>
    [ValueConvert]
    IExcelScenario? Scenarios(string index);

    /// <summary>
    /// 获取表示单个方案（Scenario 对象）或工作表上所有方案（Scenarios 对象）的对象。
    /// </summary>
    /// <param name="index">可选项。方案的名称或编号。可以使用数组指定多个方案。</param>
    /// <returns>Scenario 或 Scenarios 对象。</returns>
    [ValueConvert]
    IExcelScenario? Scenarios(int index);

    /// <summary>
    /// 获取或设置允许滚动的区域作为 A1 样式的区域引用。滚动区域外的单元格无法被选中。可读写字符串。
    /// </summary>
    string ScrollArea { get; set; }

    /// <summary>
    /// 使当前筛选列表的所有行可见。如果正在使用自动筛选，此方法会将箭头更改为“全部”。
    /// </summary>
    void ShowAllData();

    /// <summary>
    /// 显示与工作表关联的数据表单。
    /// </summary>
    void ShowDataForm();

    /// <summary>
    /// 获取所有行在工作表中的标准（默认）高度（以磅为单位）。只读 Double。
    /// </summary>
    double StandardHeight { get; }

    /// <summary>
    /// 获取或设置工作表中所有列的标准（默认）宽度。可读写 Double。
    /// </summary>
    double StandardWidth { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Excel 是否为工作表使用 Lotus 1-2-3 公式输入规则。可读写布尔值。
    /// </summary>
    bool TransitionFormEntry { get; set; }

    /// <summary>
    /// 获取表示指定工作表上已使用区域的 Range 对象。只读。
    /// </summary>
    IExcelRange? UsedRange { get; }

    /// <summary>
    /// 获取表示工作表上水平分页符的 HPageBreaks 集合。只读。
    /// </summary>
    IExcelHPageBreaks? HPageBreaks { get; }

    /// <summary>
    /// 获取表示工作表上垂直分页符的 VPageBreaks 集合。只读。
    /// </summary>
    IExcelVPageBreaks? VPageBreaks { get; }

    /// <summary>
    /// 获取 QueryTables 集合，该集合表示指定工作表上的所有查询表。只读。
    /// </summary>
    IExcelQueryTables? QueryTables { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示指定工作表上的分页符（包括自动和手动）。可读写布尔值。
    /// </summary>
    bool DisplayPageBreaks { get; set; }

    /// <summary>
    /// 获取 Comments 集合，该集合表示指定工作表的所有批注。只读。
    /// </summary>
    IExcelComments? Comments { get; }

    /// <summary>
    /// 清除工作表中无效条目上的圆圈。
    /// </summary>
    void ClearCircles();

    /// <summary>
    /// 圈释工作表中的无效条目。
    /// </summary>
    void CircleInvalid();

    /// <summary>
    /// 获取 AutoFilter 对象（如果筛选已打开）。如果筛选已关闭，则返回 Nothing。只读。
    /// </summary>
    IExcelAutoFilter? AutoFilter { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示指定窗口、工作表或 ListObject 是否从右到左显示，而不是从左到右。可读写布尔值。
    /// </summary>
    bool DisplayRightToLeft { get; set; }

    /// <summary>
    /// 打印对象。
    /// </summary>
    /// <param name="from">可选项。开始打印的页码。如果省略此参数，则从开头开始打印。</param>
    /// <param name="to">可选项。要打印的最后一页的页码。如果省略此参数，则打印到最后一页。</param>
    /// <param name="copies">可选项。要打印的份数。如果省略此参数，则打印一份。</param>
    /// <param name="preview">可选项。True 表示让 Microsoft Excel 在打印对象之前调用打印预览。False（或省略）表示立即打印对象。</param>
    /// <param name="activePrinter">可选项。设置活动打印机的名称。</param>
    /// <param name="printToFile">可选项。True 表示打印到文件。如果未指定 prToFileName，Microsoft Excel 会提示用户输入输出文件的名称。</param>
    /// <param name="collate">可选项。True 表示对多份副本进行排序。</param>
    /// <param name="prToFileName">可选项。如果将 printToFile 设置为 True，此参数指定要打印到的文件的名称。</param>
    void PrintOut(int? from = null, int? to = null,
                int? copies = null, bool? preview = null,
                string? activePrinter = null, bool? printToFile = null,
                bool? collate = null, string? prToFileName = null);

    /// <summary>
    /// 获取图表或工作表的 Tab 对象。
    /// </summary>
    IExcelTab? Tab { get; }

    /// <summary>
    /// 获取表示文档电子邮件标头的 MsoEnvelope 对象。
    /// </summary>
    IOfficeMsoEnvelope? MailEnvelope { get; }

    /// <summary>
    /// 将图表或工作表的更改保存到不同的文件中。
    /// </summary>
    /// <param name="filename">可选项。字符串，指示要保存的文件的名称。可以包含完整路径；如果不包含，Microsoft Excel 会将文件保存在当前文件夹中。</param>
    /// <param name="fileFormat">可选项。保存文件时使用的文件格式。</param>
    /// <param name="password">可选项。区分大小写的字符串（不超过 15 个字符），指示要提供给文件的保护密码。</param>
    /// <param name="writeResPassword">可选项。字符串，指示此文件的写保护密码。</param>
    /// <param name="readOnlyRecommended">可选项。True 表示在打开文件时显示建议以只读方式打开文件的消息。</param>
    /// <param name="createBackup">可选项。True 表示创建备份文件。</param>
    /// <param name="addToMru">可选项。True 表示将此工作簿添加到最近使用的文件列表中。默认值为 False。</param>
    /// <param name="textCodepage">可选项。在美国英语版的 Microsoft Excel 中不使用。</param>
    /// <param name="textVisualLayout">可选项。在美国英语版的 Microsoft Excel 中不使用。</param>
    /// <param name="local">可选项。True 表示根据 Microsoft Excel 的语言（包括控制面板设置）保存文件。False（默认值）表示根据 Visual Basic for Applications (VBA) 的语言保存文件。</param>
    void SaveAs(string filename, XlFileFormat? fileFormat = null,
                string? password = null, string? writeResPassword = null,
                bool? readOnlyRecommended = null, bool? createBackup = null,
                bool? addToMru = null, object? textCodepage = null,
                object? textVisualLayout = null, bool? local = null);

    /// <summary>
    /// 获取 CustomProperties 对象，该对象表示与工作表关联的标识符信息。
    /// </summary>
    IExcelCustomProperties? CustomProperties { get; }

    /// <summary>
    /// 获取 Protection 对象，该对象表示工作表的保护选项。
    /// </summary>
    IExcelProtection? Protection { get; }

    /// <summary>
    /// 使用指定格式将剪贴板的内容粘贴到工作表上。
    /// </summary>
    /// <param name="format">可选项。字符串，指定数据的剪贴板格式。</param>
    /// <param name="link">可选项。True 表示建立与粘贴数据源的链接。如果源数据不适合链接或源应用程序不支持链接，则忽略此参数。默认值为 False。</param>
    /// <param name="displayAsIcon">可选项。True 表示将粘贴的数据显示为图标。默认值为 False。</param>
    /// <param name="iconFileName">可选项。如果 displayAsIcon 为 True，则包含要使用的图标的文件名。</param>
    /// <param name="iconIndex">可选项。图标在图标文件中的索引号。</param>
    /// <param name="iconLabel">可选项。图标的文本标签。</param>
    /// <param name="noHtmlFormatting">可选项。True 表示从 HTML 中删除所有格式、超链接和图像。False 表示按原样粘贴 HTML。默认值为 False。</param>
    void PasteSpecial(string? format = null, bool? link = null,
                    bool? displayAsIcon = null, string? iconFileName = null,
                    int? iconIndex = null, string? iconLabel = null,
                    bool? noHtmlFormatting = null);

    /// <summary>
    /// 保护工作表，使其无法被修改。
    /// </summary>
    /// <param name="password">可选项。字符串，指定工作表或工作簿的区分大小写的密码。</param>
    /// <param name="drawingObjects">可选项。True 表示保护形状。默认值为 False。</param>
    /// <param name="contents">可选项。True 表示保护内容。对于图表，此参数保护整个图表；对于工作表，此参数保护锁定的单元格。默认值为 True。</param>
    /// <param name="scenarios">可选项。True 表示保护方案。此参数仅对工作表有效。默认值为 True。</param>
    /// <param name="userInterfaceOnly">可选项。True 表示仅保护用户界面而不保护宏。如果省略此参数，则保护既应用于宏也应用于用户界面。</param>
    /// <param name="allowFormattingCells">可选项。True 允许用户在受保护的工作表上设置任何单元格的格式。默认值为 False。</param>
    /// <param name="allowFormattingColumns">可选项。True 允许用户在受保护的工作表上设置任何列的格式。默认值为 False。</param>
    /// <param name="allowFormattingRows">可选项。True 允许用户在受保护的工作表上设置任何行的格式。默认值为 False。</param>
    /// <param name="allowInsertingColumns">可选项。True 允许用户在受保护的工作表上插入列。默认值为 False。</param>
    /// <param name="allowInsertingRows">可选项。True 允许用户在受保护的工作表上插入行。默认值为 False。</param>
    /// <param name="allowInsertingHyperlinks">可选项。True 允许用户在工作表上插入超链接。默认值为 False。</param>
    /// <param name="allowDeletingColumns">可选项。True 允许用户在受保护的工作表上删除列，其中要删除的列中的每个单元格都已解锁。默认值为 False。</param>
    /// <param name="allowDeletingRows">可选项。True 允许用户在受保护的工作表上删除行，其中要删除的行中的每个单元格都已解锁。默认值为 False。</param>
    /// <param name="allowSorting">可选项。True 允许用户在受保护的工作表上进行排序。排序区域中的每个单元格都必须已解锁或未受保护。默认值为 False。</param>
    /// <param name="allowFiltering">可选项。True 允许用户在受保护的工作表上设置筛选器。</param>
    /// <param name="allowUsingPivotTables">可选项。True 允许用户在受保护的工作表上使用数据透视表报告。默认值为 False。</param>
    void Protect(string? password = null, bool? drawingObjects = null,
                bool? contents = null, bool? scenarios = null,
                bool? userInterfaceOnly = null, bool? allowFormattingCells = null,
                bool? allowFormattingColumns = null, bool? allowFormattingRows = null,
                bool? allowInsertingColumns = null, bool? allowInsertingRows = null,
                bool? allowInsertingHyperlinks = null, bool? allowDeletingColumns = null,
                bool? allowDeletingRows = null, bool? allowSorting = null,
                bool? allowFiltering = null, bool? allowUsingPivotTables = null);

    /// <summary>
    /// 获取工作表中的 ListObject 对象集合。只读 ListObjects 集合。
    /// </summary>
    IExcelListObjects? ListObjects { get; }

    /// <summary>
    /// 获取映射到特定 XPath 的单元格的 Range 对象。如果指定的 XPath 尚未映射到工作表或映射的区域为空，则返回 Nothing。
    /// </summary>
    /// <param name="xPath">必需。要查询的 XPath。</param>
    /// <param name="selectionNamespaces">可选项。空格分隔的字符串，包含 XPath 参数中引用的命名空间。</param>
    /// <param name="map">可选项。XmlMap。如果要在特定映射中查询 XPath，请指定 XML 映射。</param>
    /// <returns>映射的 Range 对象。</returns>
    IExcelRange? XmlDataQuery(string xPath, string? selectionNamespaces = null, IExcelXmlMap? map = null);

    /// <summary>
    /// 获取映射到特定 XPath 的单元格的 Range 对象。如果指定的 XPath 尚未映射到工作表，则返回 Nothing。
    /// </summary>
    /// <param name="xPath">必需。要查询的 XPath。</param>
    /// <param name="selectionNamespaces">可选项。空格分隔的字符串，包含 XPath 参数中引用的命名空间。</param>
    /// <param name="map">可选项。XmlMap。如果要在特定映射中查询 XPath，请指定 XML 映射。</param>
    /// <returns>映射的 Range 对象。</returns>
    IExcelRange? XmlMapQuery(string xPath, string? selectionNamespaces = null, IExcelXmlMap? map = null);

    /// <summary>
    /// 获取或设置一个值，该值指示是否根据需要自动应用条件格式。可读写布尔值。
    /// </summary>
    bool EnableFormatConditionsCalculation { get; set; }

    /// <summary>
    /// 获取当前工作表中的排序值。只读 Sort 对象。
    /// </summary>
    IExcelSort? Sort { get; }

    /// <summary>
    /// 导出为指定格式的文件。
    /// </summary>
    /// <param name="type">要导出到的文件格式类型。</param>
    /// <param name="filename">要保存的文件的名称。可以包含完整路径，或者 Excel 会将文件保存在当前文件夹中。</param>
    /// <param name="quality">可选项。可选项。XlFixedFormatQuality。指定发布文件的质量。。指定发布文件的质量。</param>
    /// <param name="includeDocProperties">True 表示包含文档属性；否则为 False。</param>
    /// <param name="ignorePrintAreas">True 表示发布时忽略任何已设置的打印区域；否则为 False。</param>
    /// <param name="from">开始发布的页码。如果省略此参数，则从开头开始发布。</param>
    /// <param name="to">要发布的最后一页的页码。如果省略此参数，则发布到最后一页。</param>
    /// <param name="openAfterPublish">True 表示发布后在查看器中显示文件；否则为 False。</param>
    /// <param name="fixedFormatExtClassPtr">指向 FixedFormatExt 类的指针。</param>
    void ExportAsFixedFormat(XlFixedFormatType type, string? filename = null,
                            XlFixedFormatQuality? quality = null, bool? includeDocProperties = null,
                            bool? ignorePrintAreas = null, int? from = null,
                            int? to = null, bool? openAfterPublish = null,
                            object? fixedFormatExtClassPtr = null);

    /// <summary>
    /// 获取将为当前工作表打印的批注页数。
    /// </summary>
    int PrintedCommentPages { get; }

    #region 事件

    /// <summary>
    /// 当工作表内容发生改变时触发
    /// </summary>
    event ChangeEventHandler Change;

    /// <summary>
    /// 当工作表选择区域发生改变时触发
    /// </summary>
    event SelectionChangeEventHandler SelectionChange;

    /// <summary>
    /// 当工作表被激活时触发
    /// </summary>
    event ActivateEventHandler SheetActivate;

    /// <summary>
    /// 当工作表被取消激活时触发
    /// </summary>
    event DeactivateEventHandler SheetDeactivate;

    /// <summary>
    /// 当工作表被双击时触发
    /// </summary>
    event BeforeDoubleClickEventHandler BeforeDoubleClick;

    /// <summary>
    /// 当工作表被右键单击时触发
    /// </summary>
    event BeforeRightClickEventHandler BeforeRightClick;

    /// <summary>
    /// 当工作表计算完成后触发
    /// </summary>
    event CalculateEventHandler SheetCalculate;

    /// <summary>
    /// 在工作表被删除之前触发
    /// </summary>
    event BeforeDeleteEventHandler BeforeDelete;

    /// <summary>
    /// 当数据透视表发生更改时同步触发
    /// </summary>
    event PivotTableChangeSyncEventHandler PivotTableChangeSync;
    #endregion
}