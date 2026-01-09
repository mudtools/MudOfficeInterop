//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 工作表中一个单元格区域的包装器接口，提供对单元格区域的各种操作和属性访问功能。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelRange : IEnumerable<IExcelRange?>, IOfficeObject<IExcelRange, MsExcel.Range>, IDisposable
{

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置指定区域的值（重载索引器）。
    /// </summary>
    /// <param name="rowIndex">可选。行索引。</param>
    /// <param name="columnIndex">可选。列索引。</param>
    /// <returns>指定索引处的单元格值。</returns>
    [IgnoreGenerator]
    IExcelRange? this[int? rowIndex, int? columnIndex] { get; set; }

    /// <summary>
    /// 获取或设置指定区域的值（重载索引器）。
    /// </summary>
    /// <param name="rowAddress">行地址</param>
    /// <param name="columnAddress">列地址</param>
    /// <returns>指定索引处的单元格值。</returns>
    [IgnoreGenerator]
    IExcelRange? this[string? rowAddress, string? columnAddress] { get; set; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="address">地址</param>
    [IgnoreGenerator]
    IExcelRange? this[string address] { get; }

    /// <summary>
    /// 获取工作表中指定行列的单元格对象
    /// </summary>
    /// <param name="row">行号</param>
    /// <returns>单元格对象</returns>
    [IgnoreGenerator]
    IExcelRange? this[int row] { get; }

    /// <summary>
    /// 激活当前选定区域内的单个单元格。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Activate();

    /// <summary>
    /// 获取或设置一个值，该值指示当单元格文本对齐方式设置为水平或垂直均匀分布时，文本是否自动缩进。
    /// </summary>
    object AddIndent { get; set; }

    /// <summary>
    /// 获取区域的引用地址字符串。
    /// </summary>
    [MethodIndex]
    string? Address(bool? rowAbsolute = true,
        bool? columnAbsolute = true,
        XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
        bool? external = false,
        object? relativeTo = null);

    /// <summary>
    /// 获取区域的引用地址字符串（使用用户语言）。
    /// </summary>
    [MethodIndex]
    string? AddressLocal(bool? rowAbsolute = true,
        bool? columnAbsolute = true,
        XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
        bool? external = false,
        object? relativeTo = null);

    /// <summary>
    /// 基于条件区域筛选或复制列表中的数据。
    /// </summary>
    /// <param name="action">要执行的筛选操作，可以是 xlFilterCopy 或 xlFilterInPlace。</param>
    /// <param name="criteriaRange">可选。条件区域。如果省略，则没有条件。</param>
    /// <param name="copyToRange">可选。如果操作是 xlFilterCopy，则为复制行的目标区域；否则忽略此参数。</param>
    /// <param name="unique">可选。True 表示仅筛选唯一记录；False 表示筛选所有符合条件的记录。默认为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? AdvancedFilter(XlFilterAction action, IExcelRange? criteriaRange,
                           IExcelRange? copyToRange, bool? unique);

    /// <summary>
    /// 将名称应用到指定区域中的单元格。
    /// </summary>
    /// <param name="names">可选。要应用的名称数组。如果省略，则将工作表上的所有名称应用到区域。</param>
    /// <param name="ignoreRelativeAbsolute">可选。True 表示用名称替换引用，而不考虑名称或引用的引用类型；False 表示仅将绝对引用替换为绝对名称，相对引用替换为相对名称，混合引用替换为混合名称。默认为 True。</param>
    /// <param name="useRowColumnNames">可选。True 表示如果找不到区域的名称，则使用包含指定区域的列范围和行范围的名称；False 表示忽略 omitColumn 和 omitRow 参数。默认为 True。</param>
    /// <param name="omitColumn">可选。True 表示用面向行的名称替换整个引用。仅当引用的单元格与公式位于同一列且在面向行的命名区域内时，才能省略面向列的名称。默认为 True。</param>
    /// <param name="omitRow">可选。True 表示用面向列的名称替换整个引用。仅当引用的单元格与公式位于同一行且在面向列的命名区域内时，才能省略面向行的名称。默认为 True。</param>
    /// <param name="order">可选。确定当单元格引用被面向行和面向列的区域名称替换时，哪个区域名称列在前面。</param>
    /// <param name="appendLast">可选。True 表示替换 names 中名称的定义，并替换最后定义的名称的定义；False 表示仅替换 names 中名称的定义。默认为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? ApplyNames(string[]? names = null, bool? ignoreRelativeAbsolute = null, bool? useRowColumnNames = null,
                    bool? omitColumn = null, bool? omitRow = null, XlApplyNamesOrder order = XlApplyNamesOrder.xlRowThenColumn,
                    bool? appendLast = null);

    /// <summary>
    /// 对指定区域应用分级显示样式。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ApplyOutlineStyles();

    /// <summary>
    /// 获取一个 Areas 集合，该集合表示多区域选定中的所有区域。
    /// </summary>
    IExcelAreas? Areas { get; }

    /// <summary>
    /// 从列表中返回一个自动完成匹配项。
    /// </summary>
    /// <param name="text">要完成的字符串。</param>
    /// <returns>自动完成匹配的字符串。</returns>
    string? AutoComplete(string text);

    /// <summary>
    /// 对指定区域中的单元格执行自动填充。
    /// </summary>
    /// <param name="destination">要填充的单元格。目标区域必须包含源区域。</param>
    /// <param name="type">可选。指定填充类型。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? AutoFill(IExcelRange destination, XlAutoFillType type = XlAutoFillType.xlFillDefault);

    /// <summary>
    /// 使用自动筛选功能筛选列表。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    [ValueConvert]
    IExcelAutoFilter? AutoFilter();

    /// <summary>
    /// 使用自动筛选功能筛选列表。
    /// </summary>
    /// <param name="field">可选。要作为筛选依据的字段的整数偏移量（从列表左侧开始；最左侧字段为字段一）。</param>
    /// <param name="criteria1">可选。条件（字符串；例如“101”）。使用“=”查找空白字段，或使用“&lt;&gt;”查找非空白字段。如果省略，条件为“全部”。如果 operator 是 xlTop10Items，则 criteria1 指定项目数（例如“10”）。</param>
    /// <param name="filterOperator">可选。筛选运算符。可与 criteria1 和 criteria2 一起构造复合条件。</param>
    /// <param name="criteria2">可选。第二个条件（字符串）。与 criteria1 和 operator 一起使用以构造复合条件。</param>
    /// <param name="visibleDropDown">可选。True 表示显示筛选字段的自动筛选下拉箭头；False 表示隐藏筛选字段的自动筛选下拉箭头。默认为 True。</param>
    /// <returns>表示操作结果的对象。</returns>
    [ValueConvert]
    IExcelAutoFilter? AutoFilter(int? field, string? criteria1, XlAutoFilterOperator filterOperator = XlAutoFilterOperator.xlAnd,
                       string? criteria2 = null, bool? visibleDropDown = null);

    /// <summary>
    /// 更改区域中列的宽度或行的高度以实现最佳匹配。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? AutoFit();

    /// <summary>
    /// 使用预定义格式自动格式化指定区域。
    /// </summary>
    /// <param name="format">可选。指定的自动格式。</param>
    /// <param name="number">可选。True 表示在自动格式中包含数字格式。默认为 True。</param>
    /// <param name="font">可选。True 表示在自动格式中包含字体格式。默认为 True。</param>
    /// <param name="alignment">可选。True 表示在自动格式中包含对齐方式。默认为 True。</param>
    /// <param name="border">可选。True 表示在自动格式中包含边框格式。默认为 True。</param>
    /// <param name="pattern">可选。True 表示在自动格式中包含图案格式。默认为 True。</param>
    /// <param name="width">可选。True 表示在自动格式中包含列宽和行高。默认为 True。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? AutoFormat(XlRangeAutoFormat format = XlRangeAutoFormat.xlRangeAutoFormatClassic1,
                     bool? number = null, bool? font = null, bool? alignment = null,
                     bool? border = null, bool? pattern = null, bool? width = null);

    /// <summary>
    /// 为指定区域自动创建分级显示。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? AutoOutline();

    /// <summary>
    /// 为区域添加边框，并为新边框设置颜色、线型和粗细属性。
    /// </summary>
    /// <param name="lineStyle">可选。边框的线型。</param>
    /// <param name="weight">可选。边框的粗细。</param>
    /// <param name="colorIndex">可选。边框颜色，作为当前调色板中的索引或 XlColorIndex 常量。</param>
    /// <param name="color">可选。边框颜色，作为 RGB 值。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? BorderAround(XlLineStyle? lineStyle = null, XlBorderWeight weight = XlBorderWeight.xlThin,
                        XlColorIndex colorIndex = XlColorIndex.xlColorIndexAutomatic, [ConvertInt] Color? color = null);

    /// <summary>
    /// 获取一个 Borders 集合，该集合表示样式或单元格区域（包括定义为条件格式一部分的区域）的边框。
    /// </summary>
    IExcelBorders? Borders { get; }

    /// <summary>
    /// 计算工作表上指定的单元格区域。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Calculate();

    /// <summary>
    /// 获取一个 Range 对象，该对象表示指定区域中的单元格。
    /// </summary>
    IExcelRange? Cells { get; }

    /// <summary>
    /// 获取一个 Characters 对象，该对象表示对象文本中的一系列字符。
    /// </summary>
    /// <param name="start">可选。要返回的第一个字符。如果此参数为 1 或省略，则此属性返回从第一个字符开始的字符范围。</param>
    /// <param name="length">可选。要返回的字符数。如果省略，则此属性返回字符串的其余部分（起始字符之后的所有内容）。</param>
    [MethodIndex]
    IExcelCharacters? Characters(int? start = null, int? length = null);

    /// <summary>
    /// 检查对象的拼写。
    /// </summary>
    /// <param name="customDictionary">可选。一个字符串，指示如果在主词典中找不到单词时要检查的自定义词典的文件名。如果省略，则使用当前指定的词典。</param>
    /// <param name="ignoreUppercase">可选。True 表示让 Microsoft Excel 忽略全部大写的单词；False 表示让 Microsoft Excel 检查全部大写的单词。如果省略，则使用当前设置。</param>
    /// <param name="alwaysSuggest">可选。True 表示当发现拼写错误时，让 Microsoft Excel 显示建议的替代拼写列表；False 表示让 Microsoft Excel 在您输入正确拼写时暂停。如果省略，则使用当前设置。</param>
    /// <param name="spellLang">可选。所用词典的语言。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? CheckSpelling(string? customDictionary = null, bool? ignoreUppercase = null,
                         bool? alwaysSuggest = null, object? spellLang = null);

    /// <summary>
    /// 清除整个对象。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Clear();

    /// <summary>
    /// 清除区域中的公式。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ClearContents();

    /// <summary>
    /// 清除对象的格式。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ClearFormats();

    /// <summary>
    /// 清除指定区域中所有单元格的批注和声音批注。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ClearNotes();

    /// <summary>
    /// 清除指定区域的分级显示。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ClearOutline();

    /// <summary>
    /// 获取指定区域中第一个区域的第一列的编号。
    /// </summary>
    int Column { get; }

    /// <summary>
    /// 返回一个 Range 对象，该对象表示每一列中内容与比较单元格不同的所有单元格。
    /// </summary>
    /// <param name="comparison">要与指定区域进行比较的单个单元格。</param>
    /// <returns>表示列差异的 Range 对象。</returns>
    IExcelRange? ColumnDifferences(IExcelRange? comparison);

    /// <summary>
    /// 获取一个 Range 对象，该对象表示指定区域中的列。
    /// </summary>
    IExcelRange? Columns { get; }

    /// <summary>
    /// 获取或设置指定区域中所有列的宽度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float ColumnWidth { get; set; }

    /// <summary>
    /// 将多个工作表上多个区域的数据合并到单个工作表上的单个区域中。
    /// </summary>
    /// <param name="sources">可选。合并的源，作为 R1C1 样式表示法中的文本引用字符串数组。引用必须包含要合并的工作表的完整路径。</param>
    /// <param name="function">可选。合并函数。</param>
    /// <param name="topRow">可选。True 表示基于合并区域顶行的列标题合并数据；False 表示按位置合并数据。默认为 False。</param>
    /// <param name="leftColumn">可选。True 表示基于合并区域左列的行标题合并数据；False 表示按位置合并数据。默认为 False。</param>
    /// <param name="createLinks">可选。True 表示让合并使用工作表链接；False 表示让合并复制数据。默认为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Consolidate(string[]? sources = null, XlConsolidationFunction? function = null, bool? topRow = null, bool? leftColumn = null, bool? createLinks = null);

    /// <summary>
    /// 将区域复制到指定区域或剪贴板。
    /// </summary>
    /// <param name="destination">可选。指定区域将复制到的新区域。如果省略，Microsoft Excel 将区域复制到剪贴板。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Copy(IExcelRange? destination = null);

    /// <summary>
    /// 将 ADO 或 DAO Recordset 对象的内容复制到工作表上，从指定区域的左上角开始。
    /// </summary>
    /// <param name="data">要复制到区域中的 Recordset 对象。</param>
    /// <param name="maxRows">可选。要复制到工作表中的最大记录数。如果省略，则复制 Recordset 对象中的所有记录。</param>
    /// <param name="maxColumns">可选。要复制到工作表中的最大字段数。如果省略，则复制 Recordset 对象中的所有字段。</param>
    /// <returns>复制的行数。</returns>
    int? CopyFromRecordset(object data, int? maxRows = null, int? maxColumns = null);

    /// <summary>
    /// 将所选对象作为图片复制到剪贴板。
    /// </summary>
    /// <param name="appearance">可选。指定应如何复制图片。</param>
    /// <param name="format">可选。图片的格式。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? CopyPicture(XlPictureAppearance appearance = XlPictureAppearance.xlScreen,
                        XlCopyPictureFormat format = XlCopyPictureFormat.xlPicture);

    /// <summary>
    /// 获取集合中的对象数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 基于工作表中的文本标签在指定区域中创建名称。
    /// </summary>
    /// <param name="top">可选。True 表示使用顶行中的标签创建名称。默认为 False。</param>
    /// <param name="left">可选。True 表示使用左列中的标签创建名称。默认为 False。</param>
    /// <param name="bottom">可选。True 表示使用底行中的标签创建名称。默认为 False。</param>
    /// <param name="right">可选。True 表示使用右列中的标签创建名称。默认为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? CreateNames(bool? top = null, bool? left = null, bool? bottom = null, bool? right = null);

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含指定单元格的整个数组（如果该单元格是数组的一部分）。
    /// </summary>
    IExcelRange? CurrentArray { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示当前区域。
    /// </summary>
    IExcelRange? CurrentRegion { get; }

    /// <summary>
    /// 将对象剪切到剪贴板或将其粘贴到指定目标。
    /// </summary>
    /// <param name="destination">可选。应粘贴对象的位置。如果省略此参数，则将对象剪切到剪贴板。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Cut(IExcelRange? destination = null);

    /// <summary>
    /// 在指定区域中创建数据系列。
    /// </summary>
    /// <param name="rowcol">可选。可以是 xlRows 或 xlColumns 常量，以指定数据系列是按行还是按列输入。如果省略此参数，则使用区域的大小和形状。</param>
    /// <param name="type">可选。数据系列类型。</param>
    /// <param name="date">可选。如果 type 参数是 xlChronological，则 date 参数指示步进日期单位。</param>
    /// <param name="step">可选。系列的步长值。默认值为 1。</param>
    /// <param name="stop">可选。系列的终止值。如果省略此参数，Microsoft Excel 将填充到区域的末尾。</param>
    /// <param name="trend">可选。True 表示创建线性趋势或增长趋势；False 表示创建标准数据系列。默认值为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? DataSeries(XlRowCol? rowcol = null, XlDataSeriesType type = XlDataSeriesType.xlDataSeriesLinear,
                      XlDataSeriesDate date = XlDataSeriesDate.xlDay,
                      int? step = 1, int? stop = null, bool? trend = null);

    /// <summary>
    /// 删除对象。
    /// </summary>
    /// <param name="shift">可选。指定如何移动单元格以替换删除的单元格。如果省略，Microsoft Excel 将根据区域的形状决定。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Delete(object shift);

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含单元格所有从属项的单元格区域。
    /// </summary>
    IExcelRange? Dependents { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含单元格所有直接从属项的单元格区域。
    /// </summary>
    IExcelRange? DirectDependents { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含单元格所有直接引用项的单元格区域。
    /// </summary>
    IExcelRange? DirectPrecedents { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含源区域的区域末尾的单元格。
    /// </summary>
    /// <param name="direction">移动的方向。</param>
    /// <returns>表示区域末尾的 Range 对象。</returns>
    [MethodIndex]
    IExcelRange? End(XlDirection direction);

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含指定区域的整列（或列）。
    /// </summary>
    IExcelRange? EntireColumn { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含指定区域的整行（或行）。
    /// </summary>
    IExcelRange? EntireRow { get; }

    /// <summary>
    /// 从指定区域的顶部单元格（或单元格）向下填充到区域的底部。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? FillDown();

    /// <summary>
    /// 从指定区域中最右侧的单元格（或单元格）向左填充。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? FillLeft();

    /// <summary>
    /// 从指定区域中最左侧的单元格（或单元格）向右填充。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? FillRight();

    /// <summary>
    /// 从指定区域的底部单元格（或单元格）向上填充到区域的顶部。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? FillUp();

    /// <summary>
    /// 在区域中查找特定信息，并返回表示找到该信息的第一个单元格的 Range 对象。
    /// </summary>
    /// <param name="what">要搜索的数据。可以是字符串或任何 Microsoft Excel 数据类型。</param>
    /// <param name="after">可选。搜索开始之后的单元格。这对应于从用户界面进行搜索时活动单元格的位置。注意：After 必须是区域中的单个单元格。请记住，搜索从该单元格之后开始；直到该方法绕回到该单元格时才会搜索指定的单元格。如果未指定此参数，则搜索从区域左上角单元格之后开始。</param>
    /// <param name="lookIn">可选。信息类型。</param>
    /// <param name="lookAt">可选。可以是 xlWhole 或 xlPart。</param>
    /// <param name="searchOrder">可选。可以是 xlByRows 或 xlByColumns。</param>
    /// <param name="searchDirection">可选。搜索方向。</param>
    /// <param name="matchCase">可选。True 表示使搜索区分大小写。默认值为 False。</param>
    /// <param name="matchByte">可选。仅当选择或安装了双字节语言支持时使用。True 表示双字节字符仅匹配双字节字符；False 表示双字节字符匹配其单字节等效项。</param>
    /// <param name="searchFormat">可选。搜索格式。</param>
    /// <returns>表示找到的第一个匹配单元格的 Range 对象。</returns>
    IExcelRange? Find(object what, IExcelRange? after, XlFindLookIn? lookIn = null,
        XlLookAt? lookAt = null, XlSearchOrder? searchOrder = null,
        XlSearchDirection searchDirection = XlSearchDirection.xlNext,
        bool? matchCase = null, bool? matchByte = null,
        object? searchFormat = null);

    /// <summary>
    /// 继续从 Find 方法开始的搜索。
    /// </summary>
    /// <param name="after">可选。搜索开始之后的单元格。这对应于从用户界面进行搜索时活动单元格的位置。注意：After 必须是区域中的单个单元格。请记住，搜索从该单元格之后开始；直到该方法绕回到该单元格时才会搜索指定的单元格。如果未指定此参数，则搜索从区域左上角单元格之后开始。</param>
    /// <returns>表示下一个匹配单元格的 Range 对象。</returns>
    IExcelRange? FindNext(object after);

    /// <summary>
    /// 继续从 Find 方法开始的搜索（向前搜索）。
    /// </summary>
    /// <param name="after">可选。搜索开始之前的单元格。这对应于从用户界面进行搜索时活动单元格的位置。注意：After 必须是区域中的单个单元格。请记住，搜索从该单元格之前开始；直到该方法绕回到该单元格时才会搜索指定的单元格。如果未指定此参数，则搜索从区域左上角单元格之前开始。</param>
    /// <returns>表示上一个匹配单元格的 Range 对象。</returns>
    IExcelRange? FindPrevious(object after);

    /// <summary>
    /// 获取一个 Font 对象，该对象表示指定对象的字体。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置对象的公式（使用 A1 样式表示法和宏语言）。
    /// </summary>
    object Formula { get; set; }

    /// <summary>
    /// 获取或设置区域的数组公式。
    /// </summary>
    object FormulaArray { get; set; }

    /// <summary>
    /// 获取或设置指定区域的公式标签类型。
    /// </summary>
    XlFormulaLabel FormulaLabel { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示在工作表受保护时公式是否隐藏。
    /// </summary>
    object FormulaHidden { get; set; }

    /// <summary>
    /// 获取或设置对象的公式（使用 A1 样式引用和用户语言）。
    /// </summary>
    object FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置对象的公式（使用 R1C1 样式表示法和宏语言）。
    /// </summary>
    object FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置对象的公式（使用 R1C1 样式表示法和用户语言）。
    /// </summary>
    object FormulaR1C1Local { get; set; }

    /// <summary>
    /// 为区域的左上角单元格启动函数向导。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? FunctionWizard();

    /// <summary>
    /// 计算实现特定目标所需的值。
    /// </summary>
    /// <param name="goal">希望在此单元格中返回的值。</param>
    /// <param name="changingCell">指定为实现目标值应更改的单元格。</param>
    /// <returns>如果成功找到目标，则为 True；否则为 False。</returns>
    bool? GoalSeek(object goal, IExcelRange changingCell);

    /// <summary>
    /// 当 Range 对象表示数据透视表字段数据区域中的单个单元格时，Group 方法在该字段中执行数字或基于日期的分组。
    /// </summary>
    /// <param name="start">可选。要分组的第一个值。如果省略或为 True，则使用字段中的第一个值。</param>
    /// <param name="end">可选。要分组的最后一个值。如果省略或为 True，则使用字段中的最后一个值。</param>
    /// <param name="by">可选。如果字段是数字，则此参数指定每个组的大小。如果字段是日期，则当 periods 数组中的元素 4 为 True 且所有其他元素为 False 时，此参数指定每个组的天数。否则，将忽略此参数。如果省略，Microsoft Excel 会自动选择默认的组大小。</param>
    /// <param name="periods">可选。一个布尔值数组，用于指定分组的周期。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Group(bool? start = null, bool? end = null, object? by = null, bool[]? periods = null);

    /// <summary>
    /// 获取一个值，该值指示指定单元格是否为数组公式的一部分。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasArray { get; }

    /// <summary>
    /// 获取一个值，该值指示区域中的所有单元格是否都包含公式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasFormula { get; }

    /// <summary>
    /// 获取区域的高度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float Height { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示行或列是否隐藏。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Hidden { get; set; }

    /// <summary>
    /// 获取或设置指定对象的水平对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置单元格或区域的缩进级别。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int IndentLevel { get; set; }

    /// <summary>
    /// 向指定区域添加缩进。
    /// </summary>
    /// <param name="insertAmount">要添加到当前缩进中的量。</param>
    void InsertIndent(int insertAmount);

    /// <summary>
    /// 将单元格或单元格区域插入工作表或宏表，并移动其他单元格以腾出空间。
    /// </summary>
    /// <param name="shift">可选。指定单元格移动的方向。如果省略此参数，Microsoft Excel 将根据区域的形状决定。</param>
    /// <param name="copyOrigin">可选。复制来源。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Insert(XlInsertShiftDirection? shift = null,
                   XlInsertFormatOrigin copyOrigin = XlInsertFormatOrigin.xlFormatFromRightOrBelow);

    /// <summary>
    /// 获取一个 Interior 对象，该对象表示指定对象的内部。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 将区域中的文本重新排列，使其均匀填充区域。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Justify();

    /// <summary>
    /// 获取从 A 列左边缘到区域左边缘的距离。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float Left { get; }

    /// <summary>
    /// 获取指定区域的标题行数。
    /// </summary>
    int ListHeaderRows { get; }

    /// <summary>
    /// 将所有显示的名称粘贴到工作表上，从区域中的第一个单元格开始。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ListNames();

    /// <summary>
    /// 获取一个常量，该常量描述包含指定区域左上角的数据透视表报告的部分。
    /// </summary>
    XlLocationInTable LocationInTable { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示对象是否已锁定。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Locked { get; set; }

    /// <summary>
    /// 从指定的 Range 对象创建合并单元格。
    /// </summary>
    /// <param name="across">可选。True 表示将指定区域中每一行的单元格合并为单独的合并单元格。默认值为 False。</param>
    void Merge(bool? across);

    /// <summary>
    /// 将合并区域拆分为单独的单元格。
    /// </summary>
    void UnMerge();

    /// <summary>
    /// 获取一个 Range 对象，该对象表示包含指定单元格的合并区域。
    /// </summary>
    IExcelRange? MergeArea { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示区域或样式是否包含合并单元格。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool MergeCells { get; set; }

    /// <summary>
    /// 获取或设置对象的名称。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Name { get; set; }

    /// <summary>
    /// 将指定区域的追踪箭头导航到引用单元格、从属单元格或导致错误的单元格。
    /// </summary>
    /// <param name="towardPrecedent">可选。指定导航方向：True 表示向引用单元格导航；False 表示向从属单元格导航。</param>
    /// <param name="arrowNumber">可选。指定要导航的箭头编号；对应于单元格公式中的编号引用。</param>
    /// <param name="linkNumber">可选。如果箭头是外部引用箭头，则此参数指示要遵循哪个外部引用。如果省略此参数，则遵循第一个外部引用。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? NavigateArrow(bool? towardPrecedent = null, string? arrowNumber = null, string? linkNumber = null);

    /// <summary>
    /// 获取一个 Range 对象，该对象表示下一个单元格。
    /// </summary>
    IExcelRange? Next { get; }

    /// <summary>
    /// 获取或设置与区域左上角单元格关联的单元格批注。
    /// </summary>
    /// <param name="text">可选。要添加到批注的文本（最多 255 个字符）。文本从 start 位置开始插入，替换现有批注的 length 个字符。如果省略此参数，则此方法返回从 start 位置开始的 length 个字符的当前批注文本。</param>
    /// <param name="start">可选。要设置或返回的文本的起始位置。如果省略此参数，则此方法从第一个字符开始。要将文本追加到批注，请指定一个大于现有批注字符数的数字。</param>
    /// <param name="length">可选。要设置或返回的字符数。如果省略此参数，则 Microsoft Excel 设置或返回从起始位置到批注末尾的字符（最多 255 个字符）。如果从 start 到批注末尾的字符超过 255 个，则此方法仅返回 255 个字符。</param>
    /// <returns>批注文本。</returns>
    string? NoteText(string? text = null, int? start = null, int? length = null);

    /// <summary>
    /// 获取或设置对象的格式代码。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置对象的格式代码（作为用户语言中的字符串）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示与指定区域有一定偏移量的区域。
    /// </summary>
    /// <param name="rowOffset">可选。区域要偏移的行数（正数、负数或 0）。正值向下偏移，负值向上偏移。默认值为 0。</param>
    /// <param name="columnOffset">可选。区域要偏移的列数（正数、负数或 0）。正值向右偏移，负值向左偏移。默认值为 0。</param>
    /// <returns>表示偏移区域的 Range 对象。</returns>
    [MethodIndex]
    IExcelRange? Offset(int? rowOffset = null, int? columnOffset = null);

    /// <summary>
    /// 获取或设置文本方向。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置指定行或列的当前大纲级别。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    int OutlineLevel { get; set; }

    /// <summary>
    /// 获取或设置分页符的位置。
    /// </summary>
    int PageBreak { get; set; }

    /// <summary>
    /// 解析数据区域并将其拆分为多个单元格。
    /// </summary>
    /// <param name="parseLine">可选。一个字符串，包含左右括号，指示应在何处拆分单元格。例如，“[xxx][xxx]”会将前三个字符插入目标区域的第一列，并将接下来的三个字符插入第二列。如果省略此参数，Microsoft Excel 将根据区域左上角单元格的间距猜测拆分列的位置。如果要使用不同的区域来猜测解析行，请使用 Range 对象作为 parseLine 参数。该区域必须是正在解析的单元格之一。parseLine 参数的长度不能超过 255 个字符（包括括号和空格）。</param>
    /// <param name="destination">可选。一个 Range 对象，表示已解析数据的目标区域的左上角。如果省略此参数，Microsoft Excel 将就地解析。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Parse(string? parseLine = null, IExcelRange? destination = null);

    /// <summary>
    /// 获取一个 PivotField 对象，该对象表示包含指定区域左上角的数据透视表字段。
    /// </summary>
    IExcelPivotField? PivotField { get; }

    /// <summary>
    /// 获取一个 PivotItem 对象，该对象表示包含指定区域左上角的数据透视表项。
    /// </summary>
    IExcelPivotItem? PivotItem { get; }

    /// <summary>
    /// 获取一个 PivotTable 对象，该对象表示包含指定区域左上角的数据透视表报告，或与数据透视图报告关联的数据透视表报告。
    /// </summary>
    IExcelPivotTable? PivotTable { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示单元格的所有引用单元格。
    /// </summary>
    IExcelRange? Precedents { get; }

    /// <summary>
    /// 获取单元格的前缀字符。
    /// </summary>
    object PrefixCharacter { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示上一个单元格。
    /// </summary>
    IExcelRange? Previous { get; }

    /// <summary>
    /// 显示对象的打印预览。
    /// </summary>
    /// <param name="enableChanges">True 表示启用更改。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? PrintPreview(object enableChanges);

    /// <summary>
    /// 获取一个 QueryTable 对象，该对象表示与指定的 Range 对象相交的查询表。
    /// </summary>
    IExcelQueryTable? QueryTable { get; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示一个单元格或一系列单元格。
    /// </summary>
    /// <param name="cell1">区域的名称。这必须是宏语言中的 A1 样式引用。它可以包括区域运算符（冒号）、交集运算符（空格）或并集运算符（逗号）。它还可以包括美元符号，但会被忽略。您可以在区域的任何部分使用本地定义的名称。如果使用名称，则假定该名称使用宏语言。</param>
    /// <param name="cell2">可选。区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，也可以是宏语言中命名单个单元格的字符串。</param>
    [MethodIndex]
    IExcelRange? Range(string? cell1, string? cell2);

    /// <summary>
    /// 获取一个 Range 对象，该对象表示一个单元格或一系列单元格。
    /// </summary>
    /// <param name="cell1">区域的名称。这必须是宏语言中的 A1 样式引用。它可以包括区域运算符（冒号）、交集运算符（空格）或并集运算符（逗号）。它还可以包括美元符号，但会被忽略。您可以在区域的任何部分使用本地定义的名称。如果使用名称，则假定该名称使用宏语言。</param>
    /// <param name="cell2">可选。区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，也可以是宏语言中命名单个单元格的字符串。</param>
    [MethodIndex]
    IExcelRange? Range(string? cell1, IExcelRange? cell2 = null);

    /// <summary>
    /// 从列表中删除分类汇总。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? RemoveSubtotal();

    /// <summary>
    /// 替换指定区域单元格中的字符。
    /// </summary>
    /// <param name="what">Microsoft Excel 要搜索的字符串。</param>
    /// <param name="replacement">替换字符串。</param>
    /// <param name="lookAt">可选。可以是 xlWhole 或 xlPart。</param>
    /// <param name="searchOrder">可选。可以是 xlByRows 或 xlByColumns。</param>
    /// <param name="matchCase">可选。True 表示使搜索区分大小写。</param>
    /// <param name="matchByte">可选。仅当选择或安装了双字节语言支持时使用。True 表示双字节字符仅匹配双字节字符；False 表示双字节字符匹配其单字节等效项。</param>
    /// <param name="searchFormat">可选。方法的搜索格式。</param>
    /// <param name="replaceFormat">可选。方法的替换格式。</param>
    /// <returns>如果进行了任何替换，则为 True；否则为 False。</returns>
    bool? Replace(string what, string replacement, XlLookAt? lookAt = null,
                    XlSearchOrder? searchOrder = null, bool? matchCase = null, bool? matchByte = null,
                    object? searchFormat = null, object? replaceFormat = null);

    /// <summary>
    /// 调整指定区域的大小。
    /// </summary>
    /// <param name="rowSize">可选。新区域中的行数。如果省略此参数，则区域中的行数保持不变。</param>
    /// <param name="columnSize">可选。新区域中的列数。如果省略此参数，则区域中的列数保持不变。</param>
    /// <returns>调整大小后的新 Range 对象。</returns>
    [MethodIndex]
    IExcelRange? Resize(int? rowSize, int? columnSize);

    /// <summary>
    /// 获取区域中第一个区域的第一行的编号。
    /// </summary>
    int Row { get; }

    /// <summary>
    /// 返回一个 Range 对象，该对象表示每一行中内容与比较单元格不同的所有单元格。
    /// </summary>
    /// <param name="comparison">要与指定区域进行比较的单个单元格。</param>
    /// <returns>表示行差异的 Range 对象。</returns>
    IExcelRange? RowDifferences(IExcelRange comparison);

    /// <summary>
    /// 返回一个 Range 对象，该对象表示每一行中内容与比较单元格不同的所有单元格。
    /// </summary>
    /// <param name="comparison">要与指定区域进行比较的单个单元格。</param>
    /// <returns>表示行差异的 Range 对象。</returns>
    IExcelRange? RowDifferences(string comparison);

    /// <summary>
    /// 获取或设置指定区域中所有行的高度（以磅为单位）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float RowHeight { get; set; }

    /// <summary>
    /// 获取一个 Range 对象，该对象表示指定区域中的行。
    /// </summary>
    IExcelRange? Rows { get; }

    /// <summary>
    /// 运行位于此位置的 Microsoft Excel 宏。
    /// </summary>
    /// <param name="arg1">可选。应传递给函数的参数。</param>
    /// <param name="arg2">可选。应传递给函数的参数。</param>
    /// <param name="arg3">可选。应传递给函数的参数。</param>
    /// <param name="arg4">可选。应传递给函数的参数。</param>
    /// <param name="arg5">可选。应传递给函数的参数。</param>
    /// <param name="arg6">可选。应传递给函数的参数。</param>
    /// <param name="arg7">可选。应传递给函数的参数。</param>
    /// <param name="arg8">可选。应传递给函数的参数。</param>
    /// <param name="arg9">可选。应传递给函数的参数。</param>
    /// <param name="arg10">可选。应传递给函数的参数。</param>
    /// <param name="arg11">可选。应传递给函数的参数。</param>
    /// <param name="arg12">可选。应传递给函数的参数。</param>
    /// <param name="arg13">可选。应传递给函数的参数。</param>
    /// <param name="arg14">可选。应传递给函数的参数。</param>
    /// <param name="arg15">可选。应传递给函数的参数。</param>
    /// <param name="arg16">可选。应传递给函数的参数。</param>
    /// <param name="arg17">可选。应传递给函数的参数。</param>
    /// <param name="arg18">可选。应传递给函数的参数。</param>
    /// <param name="arg19">可选。应传递给函数的参数。</param>
    /// <param name="arg20">可选。应传递给函数的参数。</param>
    /// <param name="arg21">可选。应传递给函数的参数。</param>
    /// <param name="arg22">可选。应传递给函数的参数。</param>
    /// <param name="arg23">可选。应传递给函数的参数。</param>
    /// <param name="arg24">可选。应传递给函数的参数。</param>
    /// <param name="arg25">可选。应传递给函数的参数。</param>
    /// <param name="arg26">可选。应传递给函数的参数。</param>
    /// <param name="arg27">可选。应传递给函数的参数。</param>
    /// <param name="arg28">可选。应传递给函数的参数。</param>
    /// <param name="arg29">可选。应传递给函数的参数。</param>
    /// <param name="arg30">可选。应传递给函数的参数。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Run(object? arg1 = null, object? arg2 = null, object? arg3 = null, object? arg4 = null,
                object? arg5 = null, object? arg6 = null, object? arg7 = null, object? arg8 = null, object? arg9 = null,
                object? arg10 = null, object? arg11 = null, object? arg12 = null, object? arg13 = null, object? arg14 = null,
                object? arg15 = null, object? arg16 = null, object? arg17 = null, object? arg18 = null, object? arg19 = null,
                object? arg20 = null, object? arg21 = null, object? arg22 = null, object? arg23 = null, object? arg24 = null,
                object? arg25 = null, object? arg26 = null, object? arg27 = null, object? arg28 = null, object? arg29 = null, object? arg30 = null);

    /// <summary>
    /// 选择对象。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Select();

    /// <summary>
    /// 滚动活动窗口的内容以使区域进入视图。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Show();

    /// <summary>
    /// 绘制追踪箭头指向区域的直接从属项。
    /// </summary>
    /// <param name="remove">可选。True 表示删除指向直接从属项的一级追踪箭头；False 表示扩展一级追踪箭头。默认值为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? ShowDependents(bool? remove = true);

    /// <summary>
    /// 获取或设置一个值，该值指示是否为指定区域展开分级显示（以便列或行的详细信息可见）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowDetail { get; set; }

    /// <summary>
    /// 绘制追踪箭头穿过引用树，指向作为错误源的单元格，并返回包含该单元格的区域。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? ShowErrors();

    /// <summary>
    /// 绘制追踪箭头指向区域的直接引用项。
    /// </summary>
    /// <param name="remove">可选。True 表示删除指向直接引用项的一级追踪箭头；False 表示扩展一级追踪箭头。默认值为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? ShowPrecedents(bool? remove = true);

    /// <summary>
    /// 获取或设置一个值，该值指示文本是否自动收缩以适合可用列宽。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShrinkToFit { get; set; }

    /// <summary>
    /// 对数据透视表报告、区域或活动区域（如果指定区域仅包含一个单元格）进行排序。
    /// </summary>
    /// <param name="key1">可选。第一个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象（例如“Dept”或 Cells(1, 1)）。</param>
    /// <param name="order1">可选。为 key1 指定的字段或区域的排序顺序。</param>
    /// <param name="key2">可选。第二个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象。如果省略此参数，则没有第二个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="type">可选。指定要排序的元素。仅当对数据透视表报告排序时使用此参数。</param>
    /// <param name="order2">可选。为 key2 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="key3">可选。第三个排序字段，可以是文本（区域名称）或 Range 对象。如果省略此参数，则没有第三个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="order3">可选。为 key3 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="header">可选。指定第一行是否包含标题。对数据透视表报告排序时不能使用。</param>
    /// <param name="orderCustom">可选。此参数是自定义排序列表的从 1 开始的整数偏移量。如果省略 orderCustom，则使用正常排序。</param>
    /// <param name="matchCase">可选。True 表示执行区分大小写的排序；False 表示执行不区分大小写的排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="orientation">可选。排序方向。</param>
    /// <param name="sortMethod">可选。排序类型。</param>
    /// <param name="dataOption1">可选。指定如何对 key1 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption2">可选。指定如何对 key2 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption3">可选。指定如何对 key3 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Sort(string? key1, XlSortOrder order1 = XlSortOrder.xlAscending, object? key2 = null,
                XlSortType? type = null, XlSortOrder order2 = XlSortOrder.xlAscending, object? key3 = null,
                XlSortOrder order3 = XlSortOrder.xlAscending, XlYesNoGuess header = XlYesNoGuess.xlNo,
                int? orderCustom = null, bool? matchCase = null, XlSortOrientation orientation = XlSortOrientation.xlSortRows,
                XlSortMethod sortMethod = XlSortMethod.xlPinYin, XlSortDataOption dataOption1 = XlSortDataOption.xlSortNormal,
                XlSortDataOption dataOption2 = XlSortDataOption.xlSortNormal, XlSortDataOption dataOption3 = XlSortDataOption.xlSortNormal);

    /// <summary>
    /// 对数据透视表报告、区域或活动区域（如果指定区域仅包含一个单元格）进行排序。
    /// </summary>
    /// <param name="key1">可选。第一个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象（例如“Dept”或 Cells(1, 1)）。</param>
    /// <param name="order1">可选。为 key1 指定的字段或区域的排序顺序。</param>
    /// <param name="key2">可选。第二个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象。如果省略此参数，则没有第二个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="type">可选。指定要排序的元素。仅当对数据透视表报告排序时使用此参数。</param>
    /// <param name="order2">可选。为 key2 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="key3">可选。第三个排序字段，可以是文本（区域名称）或 Range 对象。如果省略此参数，则没有第三个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="order3">可选。为 key3 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="header">可选。指定第一行是否包含标题。对数据透视表报告排序时不能使用。</param>
    /// <param name="orderCustom">可选。此参数是自定义排序列表的从 1 开始的整数偏移量。如果省略 orderCustom，则使用正常排序。</param>
    /// <param name="matchCase">可选。True 表示执行区分大小写的排序；False 表示执行不区分大小写的排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="orientation">可选。排序方向。</param>
    /// <param name="sortMethod">可选。排序类型。</param>
    /// <param name="dataOption1">可选。指定如何对 key1 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption2">可选。指定如何对 key2 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption3">可选。指定如何对 key3 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Sort(IExcelRange? key1, XlSortOrder order1 = XlSortOrder.xlAscending, object? key2 = null,
                XlSortType? type = null, XlSortOrder order2 = XlSortOrder.xlAscending, object? key3 = null,
                XlSortOrder order3 = XlSortOrder.xlAscending, XlYesNoGuess header = XlYesNoGuess.xlNo,
                int? orderCustom = null, bool? matchCase = null, XlSortOrientation orientation = XlSortOrientation.xlSortRows,
                XlSortMethod sortMethod = XlSortMethod.xlPinYin, XlSortDataOption dataOption1 = XlSortDataOption.xlSortNormal,
                XlSortDataOption dataOption2 = XlSortDataOption.xlSortNormal, XlSortDataOption dataOption3 = XlSortDataOption.xlSortNormal);


    /// <summary>
    /// 使用东亚排序方法对区域或数据透视表报告进行排序，或者对活动区域进行排序（如果区域仅包含一个单元格）。
    /// </summary>
    /// <param name="sortMethod">可选。排序类型。</param>
    /// <param name="key1">可选。第一个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象（例如“Dept”或 Cells(1, 1)）。</param>
    /// <param name="order1">可选。为 key1 指定的字段或区域的排序顺序。</param>
    /// <param name="type">可选。指定要排序的元素。仅当对数据透视表报告排序时使用此参数。</param>
    /// <param name="key2">可选。第二个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象。如果省略此参数，则没有第二个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="order2">可选。为 key2 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="key3">可选。第三个排序字段，可以是文本（区域名称）或 Range 对象。如果省略此参数，则没有第三个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="order3">可选。为 key3 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="header">可选。指定第一行是否包含标题。对数据透视表报告排序时不能使用。</param>
    /// <param name="orderCustom">可选。此参数是自定义排序列表的从 1 开始的整数偏移量。如果省略 orderCustom，则使用正常排序。</param>
    /// <param name="matchCase">可选。True 表示执行区分大小写的排序；False 表示执行不区分大小写的排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="orientation">可选。排序方向。</param>
    /// <param name="dataOption1">可选。指定如何对 key1 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption2">可选。指定如何对 key2 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption3">可选。指定如何对 key3 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? SortSpecial(XlSortMethod sortMethod = XlSortMethod.xlPinYin, string? key1 = null,
                XlSortOrder order1 = XlSortOrder.xlAscending, XlSortType? type = null, object? key2 = null,
                XlSortOrder order2 = XlSortOrder.xlAscending, object? key3 = null,
                XlSortOrder order3 = XlSortOrder.xlAscending, XlYesNoGuess header = XlYesNoGuess.xlNo,
                int? orderCustom = null, bool? matchCase = null, XlSortOrientation orientation = XlSortOrientation.xlSortRows,
                XlSortDataOption dataOption1 = XlSortDataOption.xlSortNormal,
                XlSortDataOption dataOption2 = XlSortDataOption.xlSortNormal,
                XlSortDataOption dataOption3 = XlSortDataOption.xlSortNormal);

    /// <summary>
    /// 使用东亚排序方法对区域或数据透视表报告进行排序，或者对活动区域进行排序（如果区域仅包含一个单元格）。
    /// </summary>
    /// <param name="sortMethod">可选。排序类型。</param>
    /// <param name="key1">可选。第一个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象（例如“Dept”或 Cells(1, 1)）。</param>
    /// <param name="order1">可选。为 key1 指定的字段或区域的排序顺序。</param>
    /// <param name="type">可选。指定要排序的元素。仅当对数据透视表报告排序时使用此参数。</param>
    /// <param name="key2">可选。第二个排序字段，可以是文本（数据透视表字段或区域名称）或 Range 对象。如果省略此参数，则没有第二个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="order2">可选。为 key2 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="key3">可选。第三个排序字段，可以是文本（区域名称）或 Range 对象。如果省略此参数，则没有第三个排序字段。对数据透视表报告排序时不能使用。</param>
    /// <param name="order3">可选。为 key3 指定的字段或区域的排序顺序。对数据透视表报告排序时不能使用。</param>
    /// <param name="header">可选。指定第一行是否包含标题。对数据透视表报告排序时不能使用。</param>
    /// <param name="orderCustom">可选。此参数是自定义排序列表的从 1 开始的整数偏移量。如果省略 orderCustom，则使用正常排序。</param>
    /// <param name="matchCase">可选。True 表示执行区分大小写的排序；False 表示执行不区分大小写的排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="orientation">可选。排序方向。</param>
    /// <param name="dataOption1">可选。指定如何对 key1 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption2">可选。指定如何对 key2 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <param name="dataOption3">可选。指定如何对 key3 中的文本进行排序。对数据透视表报告排序时不能使用。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? SortSpecial(XlSortMethod sortMethod = XlSortMethod.xlPinYin, IExcelRange? key1 = null,
                XlSortOrder order1 = XlSortOrder.xlAscending, XlSortType? type = null, object? key2 = null,
                XlSortOrder order2 = XlSortOrder.xlAscending, object? key3 = null,
                XlSortOrder order3 = XlSortOrder.xlAscending, XlYesNoGuess header = XlYesNoGuess.xlNo,
                int? orderCustom = null, bool? matchCase = null, XlSortOrientation orientation = XlSortOrientation.xlSortRows,
                XlSortDataOption dataOption1 = XlSortDataOption.xlSortNormal,
                XlSortDataOption dataOption2 = XlSortDataOption.xlSortNormal,
                XlSortDataOption dataOption3 = XlSortDataOption.xlSortNormal);

    /// <summary>
    /// 返回一个 Range 对象，该对象表示匹配指定类型和值的所有单元格。
    /// </summary>
    /// <param name="type">要包含的单元格。可以是以下 XlCellType 常量之一：xlCellTypeAllFormatConditions（任何格式的单元格）、xlCellTypeAllValidation（具有验证条件的单元格）、xlCellTypeBlanks（空单元格）、xlCellTypeComments（包含批注的单元格）、xlCellTypeConstants（包含常量的单元格）、xlCellTypeFormulas（包含公式的单元格）、xlCellTypeLastCell（已用区域中的最后一个单元格）、xlCellTypeSameFormatConditions（具有相同格式的单元格）、xlCellTypeSameValidation（具有相同验证条件的单元格）、xlCellTypeVisible（所有可见单元格）。</param>
    /// <param name="value">可选。如果 type 是 xlCellTypeConstants 或 xlCellTypeFormulas，则使用此参数确定结果中包含哪些类型的单元格。这些值可以相加以返回多种类型。默认情况下，选择所有常量或公式，无论其类型如何。可以是以下 XlSpecialCellsValue 常量之一：xlErrors、xlLogical、xlNumbers、xlTextValues。</param>
    /// <returns>表示特殊单元格的 Range 对象。</returns>
    IExcelRange? SpecialCells(XlCellType type, XlSpecialCellsValue? value = null);

    /// <summary>
    /// 获取或设置一个 Style 对象，该对象表示指定区域的样式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelStyle? Style { get; set; }

    /// <summary>
    /// 为区域（或当前区域，如果区域是单个单元格）创建分类汇总。
    /// </summary>
    /// <param name="groupBy">要分组的字段，作为从 1 开始的整数偏移量。</param>
    /// <param name="function">分类汇总函数。</param>
    /// <param name="totalList">一个从 1 开始的字段偏移量数组，指示要添加分类汇总的字段。</param>
    /// <param name="replace">可选。True 表示替换现有分类汇总。默认值为 False。</param>
    /// <param name="pageBreaks">可选。True 表示在每个组后添加分页符。默认值为 False。</param>
    /// <param name="summaryBelowData">可选。将汇总数据相对于分类汇总放置。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Subtotal(int groupBy, XlConsolidationFunction function, int[] totalList,
                    bool? replace = null, bool? pageBreaks = null,
                    XlSummaryRow summaryBelowData = XlSummaryRow.xlSummaryBelow);

    /// <summary>
    /// 获取一个值，该值指示该区域是否为分级显示的汇总行或列。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Summary { get; }

    /// <summary>
    /// 基于工作表中定义的输入值和公式创建数据表。
    /// </summary>
    /// <param name="rowInput">可选。用作表的行输入的单个单元格。</param>
    /// <param name="columnInput">可选。用作表的列输入的单个单元格。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? Table(IExcelRange? rowInput = null, IExcelRange? columnInput = null);

    /// <summary>
    /// 获取指定对象的文本。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Text { get; }

    /// <summary>
    /// 将包含文本的单元格列解析为若干列。
    /// </summary>
    /// <param name="destination">可选。一个 Range 对象，指定 Microsoft Excel 将放置结果的位置。如果该区域大于单个单元格，则使用左上角单元格。</param>
    /// <param name="dataType">可选。要拆分为列的文本的格式。</param>
    /// <param name="textQualifier">可选。文本限定符。</param>
    /// <param name="consecutiveDelimiter">可选。True 表示让 Microsoft Excel 将连续分隔符视为一个分隔符。默认值为 False。</param>
    /// <param name="tab">可选。True 表示使 dataType 为 xlDelimited 并将制表符作为分隔符。默认值为 False。</param>
    /// <param name="semicolon">可选。True 表示使 dataType 为 xlDelimited 并将分号作为分隔符。默认值为 False。</param>
    /// <param name="comma">可选。True 表示使 dataType 为 xlDelimited 并将逗号作为分隔符。默认值为 False。</param>
    /// <param name="space">可选。True 表示使 dataType 为 xlDelimited 并将空格字符作为分隔符。默认值为 False。</param>
    /// <param name="other">可选。True 表示使 dataType 为 xlDelimited 并将 otherChar 参数指定的字符作为分隔符。默认值为 False。</param>
    /// <param name="otherChar">可选（如果 other 为 True，则为必需）。当 other 为 True 时的分隔符。如果指定了多个字符，则仅使用字符串的第一个字符；其余字符将被忽略。</param>
    /// <param name="fieldInfo">可选。包含各个数据列解析信息的数组。</param>
    /// <param name="decimalSeparator">可选。Microsoft Excel 在识别数字时使用的小数分隔符。默认设置为系统设置。</param>
    /// <param name="thousandsSeparator">可选。Excel 在识别数字时使用的千位分隔符。默认设置为系统设置。</param>
    /// <param name="trailingMinusNumbers">可选。以减号字符开头的数字。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? TextToColumns(IExcelRange? destination, XlTextParsingType dataType = XlTextParsingType.xlDelimited,
                        XlTextQualifier textQualifier = XlTextQualifier.xlTextQualifierDoubleQuote,
                        bool? consecutiveDelimiter = null, bool? tab = null, bool? semicolon = null, bool? comma = null,
                        bool? space = null, bool? other = null, string? otherChar = null, object? fieldInfo = null,
                        string? decimalSeparator = null, string? thousandsSeparator = null, string? trailingMinusNumbers = null);

    /// <summary>
    /// 获取从第 1 行顶部到区域顶部的距离（以磅为单位）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float Top { get; }

    /// <summary>
    /// 提升分级显示中的区域（即降低其大纲级别）。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? Ungroup();

    /// <summary>
    /// 获取或设置一个值，该值指示 Range 对象的行高是否等于工作表的标准高度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseStandardHeight { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Range 对象的列宽是否等于工作表的标准宽度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool UseStandardWidth { get; set; }

    /// <summary>
    /// 获取一个 Validation 对象，该对象表示指定区域的数据验证。
    /// </summary>
    IExcelValidation? Validation { get; }

    /// <summary>
    /// 获取或设置指定区域的值。
    /// </summary>
    /// <param name="rangeValueDataType">可选。区域值数据类型。</param>
    [IgnoreGenerator]
    object GetValue(XlRangeValueDataType rangeValueDataType);

    /// <summary>
    /// 获取或设置指定区域的值。
    /// </summary>
    /// <param name="rangeValueDataType">可选。区域值数据类型。</param>
    /// <param name="value">设置指定区域的值。</param>
    [IgnoreGenerator]
    void SetValue(XlRangeValueDataType rangeValueDataType, object? value);

    /// <summary>
    /// 获取或设置单元格值。
    /// </summary>
    object Value { get; set; }

    /// <summary>
    /// 获取或设置单元格数组值。
    /// </summary>
    [IgnoreGenerator]
    object[,]? ArrayValue { get; set; }

    /// <summary>
    /// 获取或设置单元格值。
    /// </summary>
    object Value2 { get; set; }

    /// <summary>
    /// 获取或设置指定对象的垂直对齐方式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取区域的宽度（以磅为单位）。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float Width { get; }

    /// <summary>
    /// 获取一个 Worksheet 对象，该对象表示包含指定区域的工作表。
    /// </summary>
    IExcelWorksheet? Worksheet { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Excel 是否在对象中自动换行。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool WrapText { get; set; }

    /// <summary>
    /// 向区域添加批注。
    /// </summary>
    /// <param name="text">可选。批注文本。</param>
    /// <returns>表示新批注的 Comment 对象。</returns>
    IExcelComment? AddComment(string? text = null);

    /// <summary>
    /// 获取一个 Comment 对象，该对象表示与区域左上角单元格关联的批注。
    /// </summary>
    IExcelComment? Comment { get; }

    /// <summary>
    /// 清除指定区域中的所有单元格批注。
    /// </summary>
    void ClearComments();

    /// <summary>
    /// 获取一个 Phonetic 对象，该对象包含有关单元格中特定拼音文本字符串的信息。
    /// </summary>
    IExcelPhonetic? Phonetic { get; }

    /// <summary>
    /// 获取一个 FormatConditions 集合，该集合表示指定区域的所有条件格式。
    /// </summary>
    IExcelFormatConditions? FormatConditions { get; }

    /// <summary>
    /// 获取或设置指定对象的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取一个 Hyperlinks 集合，该集合表示区域的超链接。
    /// </summary>
    IExcelHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 获取区域的 Phonetics 集合。
    /// </summary>
    IExcelPhonetics? Phonetics { get; }

    /// <summary>
    /// 为区域中的所有单元格创建 Phonetic 对象。
    /// </summary>
    void SetPhonetic();

    /// <summary>
    /// 获取或设置当页面另存为网页时指定单元格的标识标签。
    /// </summary>
    string ID { get; set; }

    /// <summary>
    /// 获取一个 PivotCell 对象，该对象表示数据透视表报告中的单元格。
    /// </summary>
    IExcelPivotCell? PivotCell { get; }

    /// <summary>
    /// 指定在下一次重新计算时要重新计算的区域。
    /// </summary>
    void Dirty();

    /// <summary>
    /// 允许用户访问错误检查选项。
    /// </summary>
    IExcelErrors? Errors { get; }

    /// <summary>
    /// 使区域中的单元格按行顺序或列顺序朗读。
    /// </summary>
    /// <param name="speakDirection">可选。朗读方向，按行或按列。</param>
    /// <param name="speakFormulas">可选。True 将导致将公式发送到文本到语音 (TTS) 引擎（对于具有公式的单元格）。如果单元格没有公式，则发送值；False（默认）将导致始终将值发送到 TTS 引擎。</param>
    void Speak(XlSpeakDirection? speakDirection = null, bool? speakFormulas = null);

    /// <summary>
    /// 将 Range 从剪贴板粘贴到指定区域。
    /// </summary>
    /// <param name="paste">可选。要粘贴的区域部分。</param>
    /// <param name="operation">可选。粘贴操作。</param>
    /// <param name="skipBlanks">可选。True 表示不将剪贴板上区域中的空白单元格粘贴到目标区域。默认值为 False。</param>
    /// <param name="transpose">可选。True 表示在粘贴区域时转置行和列。默认值为 False。</param>
    /// <returns>表示操作结果的对象。</returns>
    object? PasteSpecial(XlPasteType paste = XlPasteType.xlPasteAll,
                        XlPasteSpecialOperation operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                        bool? skipBlanks = null, bool? transpose = null);

    /// <summary>
    /// 获取一个值，该值指示是否可以在受保护的工作表上编辑该区域。
    /// </summary>
    bool AllowEdit { get; }

    /// <summary>
    /// 为 Range 对象或 QueryTable 对象返回一个 ListObject 对象。
    /// </summary>
    IExcelListObject? ListObject { get; }

    /// <summary>
    /// 获取一个 XPath 对象，该对象表示映射到指定 Range 对象的元素的 XPath。
    /// </summary>
    IExcelXPath? XPath { get; }

    /// <summary>
    /// 指定可以对 SharePoint 服务器上的 Range 对象执行的操作。
    /// </summary>
    IExcelActions? ServerActions { get; }

    /// <summary>
    /// 从值区域中删除重复值。
    /// </summary>
    /// <param name="columns">可选。包含重复信息的列的索引数组。如果未传递任何内容，则假定所有列都包含重复信息。</param>
    /// <param name="header">可选。指定第一行是否包含标题信息。Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo 是默认值；如果希望 Excel 尝试确定标题，请指定 Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess。</param>
    void RemoveDuplicates(object? columns = null, XlYesNoGuess header = XlYesNoGuess.xlNo);

    /// <summary>
    /// 获取指定 Range 对象的 MDX 名称。
    /// </summary>
    string MDX { get; }

    /// <summary>
    /// 导出为指定格式的文件。
    /// </summary>
    /// <param name="type">要导出到的文件格式类型。</param>
    /// <param name="filename">可选。要保存的文件名。可以包含完整路径，或者 short_Excel2007 将文件保存在当前文件夹中。</param>
    /// <param name="quality">可选。通常格式化为 Microsoft.Office.Interop.Excel.XlFixedFormatQuality。指定发布文件的质量。</param>
    /// <param name="includeDocProperties">可选。设置为 True 以包含文档属性；否则为 False。</param>
    /// <param name="ignorePrintAreas">可选。设置为 True 以在发布时忽略任何设置的打印区域；否则为 False。</param>
    /// <param name="from">可选。开始发布的页码。如果省略此参数，则从开头开始发布。</param>
    /// <param name="to">可选。要发布的最后一页的页码。如果省略此参数，则发布到最后一页。</param>
    /// <param name="openAfterPublish">可选。设置为 True 以在发布后在查看器中显示文件；否则为 False。</param>
    /// <param name="fixedFormatExtClassPtr">可选。指向 FixedFormatExt 类的指针。</param>
    void ExportAsFixedFormat(XlFixedFormatType type, string? filename = null, XlFixedFormatQuality? quality = null,
                            bool? includeDocProperties = null, bool? ignorePrintAreas = null, int? from = null, int? to = null,
                            bool? openAfterPublish = null, object? fixedFormatExtClassPtr = null);

    /// <summary>
    /// 计算给定值区域中的最大值。只读。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    long CountLarge { get; }

    /// <summary>
    /// 计算指定的单元格区域。
    /// </summary>
    /// <returns>表示操作结果的对象。</returns>
    object? CalculateRowMajorOrder();

    /// <summary>
    /// 获取一个 SparklineGroups 对象，该对象表示指定区域中现有的一组迷你图。
    /// </summary>
    IExcelSparklineGroups? SparklineGroups { get; }

    /// <summary>
    /// 从指定区域中删除所有超链接。
    /// </summary>
    void ClearHyperlinks();

    /// <summary>
    /// 获取一个 DisplayFormat 对象，该对象表示指定区域的显示设置。
    /// </summary>
    IExcelDisplayFormat? DisplayFormat { get; }

    /// <summary>
    /// 为基于 OLAP 数据源的范围中所有已编辑的单元格执行写回操作。
    /// </summary>
    void AllocateChanges();

    /// <summary>
    /// 放弃区域中已编辑单元格的所有更改。
    /// </summary>
    void DiscardChanges();

    /// <summary>
    /// 执行快速填充操作。
    /// </summary>
    void FlashFill();

}