//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 工作表中一个单元格区域的包装器基础接口，提供对单元格区域的各种操作和属性访问功能。
/// </summary>
/// <typeparam name="T"></typeparam>
public interface ICoreRange<T> : IEnumerable<T>, IDisposable
    where T : ICoreRange<T>
{
    /// <summary>
    /// 获取图表集合所在的Application对象
    /// 对应 Rang.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置单元格的值。可以是字符串、数字、布尔值、错误值或空值。
    /// </summary>
    object Value { get; set; }

    /// <summary>
    /// 获取或设置单元格的数组值。。
    /// </summary>
    object[,] ArrayValue { get; set; }

    int PageBreak { get; set; }

    /// <summary>
    /// 获取或设置单元格的数字或空值。
    /// </summary>
    double? NumberValue { get; set; }

    /// <summary>
    /// 获取或设置单元格的数字或空值。
    /// </summary>
    double[]? NumberValues { get; }

    /// <summary>
    /// 获取或设置单元格的公式。设置时应包含等号（如 "=SUM(A1:A10)"）。
    /// </summary>
    object Formula { get; set; }

    XlFormulaLabel FormulaLabel { get; set; }

    /// <summary>
    /// 获取单元格的前缀字符。前缀字符可以是用于标识文本标签的字符，
    /// 如单引号(')表示左对齐，双引号(")表示右对齐，插入符(^)表示居中对齐，
    /// 反斜杠(\)表示重复标签，或者为空。
    /// </summary>
    object PrefixCharacter { get; }

    /// <summary>
    /// 获取单元格区域的条件格式规则集合，用于管理和操作单元格区域的条件格式
    /// </summary>
    IExcelFormatConditions? FormatConditions { get; }

    /// <summary>
    /// 获取单元格区域中字符的集合，用于对单元格中文本的字符级操作
    /// </summary>
    IExcelCharacters? Characters { get; }

    /// <summary>
    /// 获取一个值，该值指示此区域是否是单元格的公式。
    /// </summary>
    bool HasFormula { get; }

    /// <summary>
    /// 获取一个值，该值指示此区域是否是数组公式的一部分。
    /// </summary>
    bool HasArray { get; }

    /// <summary>
    /// 获取或设置区域的数组公式
    /// 对应 Range.FormulaArray 属性
    /// </summary>
    string FormulaArray { get; set; }

    /// <summary>
    /// 获取或设置区域的 R1C1 格式公式
    /// 对应 Range.FormulaR1C1 属性
    /// </summary>
    object FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置是否自动换行
    /// 对应 Range.WrapText 属性
    /// </summary>
    bool WrapText { get; set; }

    /// <summary>
    /// 获取单元格区域的内部属性对象，用于设置单元格的背景色、图案等样式。
    /// 对应 Range.Interior 属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取或设置缩进级别（0-15）
    /// 对应 Range.IndentLevel 属性
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    /// 获取或设置阅读顺序
    /// 对应 Range.ReadingOrder 属性
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取当前区域中第一个单元格所在的行号（从1开始计数）。
    /// </summary>
    int Row { get; }

    /// <summary>
    /// 获取或设置图案类型
    /// -4142=无图案, 1=纯色, 2=75%灰色等
    /// 对应 Range.Interior.Pattern 属性
    /// </summary>
    XlPattern Pattern { get; set; }

    /// <summary>
    /// 获取或设置图案颜色（RGB值）
    /// 对应 Range.Interior.PatternColor 属性
    /// </summary>
    Color PatternColor { get; set; }

    /// <summary>
    /// 获取或设置样式名称
    /// 对应 Range.Style 属性
    /// </summary>
    IExcelStyle? Style { get; set; }

    /// <summary>
    /// 获取或设置分级显示展开状态
    /// </summary>
    bool ShowDetail { get; }

    /// <summary>
    /// 获取包含指定单元格的合并区域
    /// </summary>
    /// <remarks>
    /// 若单元格不属于合并区域，则返回单元格自身
    /// </remarks>
    T MergeArea { get; }

    /// <summary>
    /// 获取当前区域包含的所有行对象集合。
    /// </summary>
    IExcelRows Rows { get; }

    /// <summary>
    /// 获取当前区域中第一个单元格所在的列号（从1开始计数）。
    /// </summary>
    int Column { get; }

    /// <summary>
    /// 获取区域所在的工作表名称
    /// 对应 Range.Worksheet.Name 属性
    /// </summary>
    string WorksheetName { get; }

    /// <summary>
    /// 获取区域的总行数
    /// 对应 Range.Rows.Count 属性
    /// </summary>
    int RowsCount { get; }

    /// <summary>
    /// 获取区域的总列数
    /// 对应 Range.Columns.Count 属性
    /// </summary>
    int ColumnsCount { get; }

    /// <summary>
    /// 获取区域中的单元格对象。
    /// </summary>
    IExcelCells Cells { get; }

    /// <summary>
    /// 获取当前区域包含的所有列对象集合。
    /// </summary>
    IExcelColumns Columns { get; }

    /// <summary>
    /// 获取区域中的子区域集合
    /// 对应 Range.Areas 属性
    /// </summary>
    IExcelAreas Areas { get; }

    /// <summary>
    /// 获取当前区域的字体格式设置对象，用于设置字体样式、大小、颜色等属性。
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取区域中的第一个单元格
    /// </summary>
    T FirstCell { get; }

    /// <summary>
    /// 获取区域中的最后一个单元格
    /// </summary>
    T LastCell { get; }

    /// <summary>
    /// 获取下一个相邻区域
    /// 对应 Range.Next 属性
    /// </summary>
    T Next { get; }

    /// <summary>
    /// 获取上一个相邻区域
    /// 对应 Range.Previous 属性
    /// </summary>
    T Previous { get; }

    /// <summary>
    /// 获取或设置当前区域所有行的高度（以磅为单位）。
    /// </summary>
    double RowHeight { get; set; }

    /// <summary>
    /// 获取或设置当前区域所有列的宽度（以字符宽度为单位）。
    /// </summary>
    double ColumnWidth { get; set; }

    /// <summary>
    /// 获取或设置当前区域是否被隐藏。
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// 获取当前区域的地址字符串表示（如 "A1:B2" 或 "$A$1:$B$2"）。
    /// </summary>
    string Address { get; }

    /// <summary>
    /// 获取当前区域的地址包装对象。
    /// </summary>
    IExcelRangeAddress RangeAddress { get; }

    /// <summary>
    /// 获取或设置区域的名称。可以是内置名称或自定义名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取当前区域中包含的单元格总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取当前区域中包含的单元格总数。
    /// </summary>
    long CountLarge { get; }

    /// <summary>
    /// 获取当前区域左边缘到工作表左边缘的距离（以磅为单位）。
    /// </summary>
    double Left { get; }

    /// <summary>
    /// 获取从工作表第1行的上边缘到当前区域上边缘的距离（以磅为单位）。
    /// </summary>
    double Top { get; }

    /// <summary>
    /// 获取当前区域的总宽度（以磅为单位）。
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取当前区域的总高度（以磅为单位）。
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取或设置是否打印网格线
    /// 对应 Worksheet.PageSetup.PrintGridlines 属性
    /// </summary>
    bool PrintGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否打印行列标题
    /// 对应 Worksheet.PageSetup.PrintHeadings 属性
    /// </summary>
    bool PrintHeadings { get; set; }

    /// <summary>
    /// 获取工作表中指定范围的区域对象
    /// 对应 Worksheet.Range 属性
    /// </summary>
    /// <param name="cell1">起始单元格</param>
    /// <param name="cell2">结束单元格（可选）</param>
    /// <returns>区域对象</returns>
    T? Range(object? cell1, object? cell2 = null);


    /// <summary>
    /// 获取当前区域所属的工作表对象。
    /// </summary>
    IExcelWorksheet Worksheet { get; }

    /// <summary>
    /// 获取区域的智能标记集合
    /// 对应 Range.SmartTags 属性
    /// </summary>
    IExcelSmartTags SmartTags { get; }

    /// <summary>
    /// 获取区域的超链接集合
    /// 对应 Range.Hyperlinks 属性
    /// </summary>
    IExcelHyperlinks Hyperlinks { get; }

    IExcelErrors Errors { get; }

    /// <summary>
    /// 为区域添加超链接
    /// </summary>
    /// <param name="address">链接地址</param>
    /// <param name="subAddress">子地址（如工作表名称）</param>
    /// <param name="screenTip">鼠标悬停提示文本</param>
    /// <param name="textToDisplay">显示文本</param>
    /// <returns>创建的超链接对象</returns>
    IExcelHyperlink AddHyperlink(string address, string subAddress, string screenTip, string textToDisplay);

    /// <summary>
    /// 清除区域中的所有超链接
    /// 对应 Range.Hyperlinks.Delete() 方法
    /// </summary>
    void ClearHyperlinks();

    #region 注释

    /// <summary>
    /// 获取区域的注释对象
    /// 对应 Range.Comment 属性
    /// </summary>
    IExcelComment Comment { get; }

    /// <summary>
    /// 获取或设置注释的文本内容
    /// </summary>
    string? CommentText { get; set; }

    /// <summary>
    /// 为区域添加注释
    /// </summary>
    /// <param name="text">注释文本</param>
    /// <returns>创建的注释对象</returns>
    IExcelComment? AddComment(string? text);

    /// <summary>
    /// 删除区域的注释
    /// 对应 Comment.Delete() 方法
    /// </summary>
    void DeleteComment();

    /// <summary>
    /// 清除区域中的所有注释
    /// 对应 Range.ClearComments() 方法
    /// </summary>
    void ClearComments();

    #endregion

    /// <summary>
    /// 激活当前区域
    /// </summary>
    void Activate();

    /// <summary>
    /// 计算公式结果
    /// </summary>
    /// <remarks>
    /// 强制重新计算区域内的所有公式
    /// </remarks>
    void Calculate();

    /// <summary>
    /// 自动填充数据到目标区域
    /// </summary>
    /// <param name="destination">填充目标区域</param>
    /// <param name="type">填充类型</param>
    void AutoFill(T destination, AutoFillType type = AutoFillType.xlFillDefault);

    /// <summary>
    /// 获取区域内的直接引用单元格
    /// </summary>
    /// <returns>直接引用单元格区域</returns>
    T? GetDirectPrecedents();

    /// <summary>
    /// 获取区域内的直接依赖单元格
    /// </summary>
    /// <returns>直接前驱单元格区域</returns>
    T? GetDirectDependents();

    /// <summary>
    /// 获取区域的本地化地址。
    /// </summary>
    /// <param name="rowAbsolute">指定行号是否为绝对引用 (例如，$1)。默认为 true。</param>
    /// <param name="columnAbsolute">指定列号是否为绝对引用 (例如，$A)。默认为 true。</param>
    /// <param name="referenceStyle">指定地址样式。默认为 xlA1。</param>
    /// <param name="external">如果为 true，则返回包含工作簿和工作表名称的外部引用。默认为 false。</param>
    /// <param name="relativeTo">如果 ReferenceStyle 为 xlR1C1，则指定相对引用的起始点。</param>
    /// <returns>表示区域地址的字符串。</returns>
    /// <exception cref="System.Runtime.InteropServices.COMException">
    /// 如果与 Excel 的交互失败，可能会抛出 COM 异常。
    /// </exception>
    /// <exception cref="ArgumentNullException">
    /// 如果内部的 _range 对象为 null。
    /// </exception>
    string GetAddressLocal(
       bool? rowAbsolute = true,
       bool? columnAbsolute = true,
       XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
       bool? external = false,
       object? relativeTo = null);

    /// <summary>
    /// 获取区域的本地化地址，使用默认参数 (A1, 绝对引用, 非外部)。
    /// </summary>
    string AddressLocal { get; }

    /// <summary>
    /// 完全兼容原始调用的静态方法
    /// </summary>
    /// <param name="rowAbsolute">行是否绝对引用</param>
    /// <param name="columnAbsolute">列是否绝对引用</param>
    /// <param name="referenceStyle">引用样式</param>
    /// <param name="external">是否外部引用</param>
    /// <param name="relativeTo">相对引用基准</param>
    /// <returns>地址字符串</returns>
    string? GetAddress(
       bool? rowAbsolute = true,
       bool? columnAbsolute = true,
       XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
       bool? external = false,
        object? relativeTo = null);

    /// <summary>
    /// 直接替换原始调用：range.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing)
    /// </summary>
    /// <returns>地址字符串</returns>
    string? GetDefaultA1Address();

    /// <summary>
    /// 获取或设置区域是否锁定（用于保护工作表）
    /// 对应 Range.Locked 属性
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置公式是否隐藏（用于保护工作表）
    /// 对应 Range.FormulaHidden 属性
    /// </summary>
    bool FormulaHidden { get; set; }

    /// <summary>
    /// 获取区域的页面设置对象
    /// 对应 Worksheet.PageSetup 属性
    /// </summary>
    IExcelPageSetup PageSetup { get; }

    /// <summary>
    /// 将当前区域的内容复制到剪贴板中，以便进行粘贴操作。
    /// </summary>
    /// <returns>如果复制操作成功则返回 true，否则返回 false</returns>
    bool Copy();

    /// <summary>
    /// 复制区域到指定目标区域
    /// </summary>
    /// <param name="destination">目标区域</param>
    void Copy(T destination);

    void CopyPicture(XlPictureAppearance appearance = XlPictureAppearance.xlScreen, XlCopyPictureFormat format = XlCopyPictureFormat.xlPicture);

    /// <summary>
    /// 复制Range区域并粘贴到指定位置
    /// </summary>
    /// <param name="targetAddress">目标地址</param>
    /// <param name="pasteType">粘贴类型</param>
    /// <returns>是否操作成功</returns>
    bool CopyAndPaste(string targetAddress, XlPasteType pasteType = XlPasteType.xlPasteAll);


    /// <summary>
    /// 从指定源区域复制内容并粘贴到当前区域。
    /// </summary>
    /// <param name="from">要复制内容的源区域</param>
    /// <param name="type">粘贴的类型（如全部、值、格式等），默认为全部</param>
    /// <param name="operation">粘贴时的运算方式（如加、减、乘、除等），默认为无运算</param>
    /// <param name="skipBlanks">是否跳过源区域中的空白单元格，默认为 false</param>
    /// <param name="transpose">是否转置粘贴（行列互换），默认为 false</param>
    /// <returns>如果粘贴操作成功则返回 true，否则返回 false</returns>
    bool CopyAndPaste(T from,
        PasteType type = PasteType.All,
        PasteOperation operation = PasteOperation.None,
        bool skipBlanks = false,
        bool transpose = false);

    /// <summary>
    /// 特殊粘贴操作
    /// </summary>
    /// <param name="paste">粘贴内容类型</param>
    /// <param name="operation">粘贴操作类型</param>
    /// <param name="skipBlanks">是否跳过空单元格</param>
    /// <param name="transpose">是否转置</param>
    /// <returns>粘贴后的区域对象</returns>
    T? PasteSpecial(
        XlPasteType paste = XlPasteType.xlPasteAll,
        XlPasteSpecialOperation operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone,
        bool? skipBlanks = false,
        bool? transpose = false);


    /// <summary>
    /// 在当前区域插入单元格、行或列，并根据指定方向移动现有内容。
    /// </summary>
    /// <param name="direction">插入后原区域内容的移动方向（向下或向右），默认为向下</param>
    /// <param name="origin">新插入区域的格式来源方向（从右侧/下方或左侧/上方），默认为从右侧或下方</param>
    /// <returns>如果插入操作成功则返回 true，否则返回 false</returns>
    bool Insert(
        XlDirection direction = XlDirection.xlDown,
        XlInsertFormatOrigin origin = XlInsertFormatOrigin.FromRightOrBelow);

    /// <summary>
    /// 删除当前区域的单元格，并根据指定方向移动相邻内容来填补空白。
    /// </summary>
    /// <param name="direction">删除后相邻单元格的移动方向（向左或向上），默认为向左</param>
    /// <returns>如果删除操作成功则返回 true，否则返回 false</returns>
    bool Delete(XlDirection direction = XlDirection.xlToLeft);

    /// <summary>
    /// 获取指定方向上最后一个非空单元格（或工作表边界）所在的单元格。
    /// </summary>
    /// <param name="direction">搜索方向（上、下、左、右），默认为向下</param>
    /// <returns>返回指定方向上的最边缘单元格区域</returns>
    T? End(XlDirection direction = XlDirection.xlDown);

    /// <summary>
    /// 选中并激活当前区域。
    /// </summary>
    /// <returns>如果选择操作成功则返回 true，否则返回 false</returns>
    bool Select();

    /// <summary>
    /// 清除当前区域中的所有内容，但保留格式设置。
    /// </summary>
    void ClearContents();

    /// <summary>
    /// 清除当前区域中的所有内容和格式设置。
    /// </summary>
    void Clear();

    /// <summary>
    /// 清除当前区域中的所有格式设置，但保留内容。
    /// </summary>
    void ClearFormats();

    /// <summary>
    /// 分列区域内的数据并将这些数据分散放置于若干单元格中。
    /// </summary>
    /// <param name="parseLine">包含左括号和右括号的字符串，用于指示单元格应拆分的位置。例如，“[xxx][xxx]”会将前三个字符插入目标范围的第一列中，并将接下来的三个字符插入到第二列中。如果省略此参数，Microsoft Excel 会根据区域中左上单元格的间距猜测拆分列的位置。 如果要使用不同的区域来猜测分析行，请使用 Range 对象作为 ParseLine 参数。 该区域必须为进行分列处理的单元格之一。 参数 ParseLine 长度不能超过 255 个字符，包括括号和空格。</param>
    /// <param name="range">表示已分析数据的目标范围的左上角。 如果省略该参数，Excel 将在原处进行分列。</param>
    /// <returns></returns>
    object Parse(string? parseLine = null, T? range = default);

    /// <summary>
    /// 合并当前区域中的所有单元格为一个单元格。
    /// </summary>
    /// <param name="merge">是否执行合并操作，默认为 true。如果为 false 则取消合并</param>
    void Merge(bool merge = true);

    /// <summary>
    /// 取消合并当前区域中的单元格。
    /// </summary>
    void UnMerge();

    /// <summary>
    /// 为当前区域添加边框
    /// </summary>
    /// <param name="LineStyle">边框线条样式</param>
    /// <param name="Weight">边框粗细程度，默认为细线</param>
    /// <param name="ColorIndex">边框颜色索引，默认为自动颜色</param>
    /// <param name="Color">边框颜色值，默认为null</param>
    /// <returns>返回边框设置结果</returns>
    object? BorderAround(XlLineStyle? LineStyle = null,
         XlBorderWeight Weight = XlBorderWeight.xlThin,
         XlColorIndex ColorIndex = XlColorIndex.xlColorIndexAutomatic,
         Color? Color = null);

    /// <summary>
    /// 调整区域大小
    /// </summary>
    /// <param name="rowSize">新行数（若为-1则保持原行数）</param>
    /// <param name="columnSize">新列数（若为-1则保持原列数）</param>
    /// <returns>调整后的新区域</returns>
    T? Resize(int? rowSize = -1, int? columnSize = -1);

    /// <summary>
    /// 获取相对于当前区域指定偏移量的新单元格区域。
    /// </summary>
    /// <param name="rowOffset">行偏移量（正数向下，负数向上），默认为 0</param>
    /// <param name="columnOffset">列偏移量（正数向右，负数向左），默认为 0</param>
    /// <returns>返回偏移后的新单元格区域</returns>
    T? Offset(int? rowOffset = 0, int? columnOffset = 0);

    /// <summary>
    /// 获取相对于当前区域指定偏移量的新单元格区域。
    /// </summary>
    /// <param name="rowOffset">行偏移量（正数向下，负数向上），默认为 0</param>
    /// <param name="columnOffset">列偏移量（正数向右，负数向左），默认为 0</param>
    /// <returns>返回偏移后的新单元格区域</returns>
    T? Offset(long? rowOffset = 0, long? columnOffset = 0);

    /// <summary>
    /// 获取当前区域与另一个区域的交集部分。
    /// </summary>
    /// <param name="other">要计算交集的另一个区域</param>
    /// <returns>返回两个区域的交集区域，如果没有交集则可能返回 null 或空区域</returns>
    T? Intersect(T other);

    /// <summary>
    /// 获取当前区域与另一个区域的并集（组合区域）。
    /// </summary>
    /// <param name="other">要计算并集的另一个区域</param>
    /// <returns>返回包含两个区域所有单元格的新区域</returns>
    T? Union(T other);

    /// <summary>
    /// 获取或设置当前区域的背景填充颜色（RGB 值格式，如 0xFF0000 表示红色）。
    /// </summary>
    int InteriorColor { get; set; }

    /// <summary>
    /// 获取或设置当前区域内容的水平对齐方式（如左对齐、居中、右对齐等）。
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置当前区域内容的垂直对齐方式（如顶端对齐、居中、底端对齐等）。
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置当前区域中文本的旋转角度（以度为单位，范围 -90 到 90 度）。
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置当前区域的超链接地址。设置后该区域将变为可点击的超链接。
    /// </summary>
    string? Hyperlink { get; set; }

    /// <summary>
    /// 获取当前区域是否包含合并单元格。
    /// </summary>
    bool MergeCells { get; }

    /// <summary>
    /// 获取或设置当前区域的数字格式字符串（如 "0.00"、"#,##0.00"、"yyyy/mm/dd" 等）。
    /// </summary>
    string NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置当前区域的数字格式字符串（如 "0.00"、"#,##0.00"、"yyyy/mm/dd" 等）。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取当前区域在工作表中实际显示的文本内容（经过格式化后的结果）。
    /// </summary>
    string Text { get; }

    /// <summary>
    /// 获取当前区域的边框设置对象，用于设置边框样式、颜色、粗细等属性。
    /// </summary>
    IExcelBorders Borders { get; }

    /// <summary>
    /// 获取当前区域的注音符号集合对象，用于处理日文等语言的注音符号。
    /// </summary>
    IExcelPhonetics? Phonetics { get; }

    /// <summary>
    /// 获取当前区域的矩形尺寸。
    /// </summary>
    ExcelRectange RangeRect { get; }

    /// <summary>
    /// 自动调整当前区域所有列的宽度以适应其内容。
    /// </summary>
    void AutoFit();

    object? AutoFormat(XlRangeAutoFormat format = XlRangeAutoFormat.xlRangeAutoFormatClassic1,
        bool? number = true, bool? font = true, bool? alignment = true,
        bool? border = true, bool? pattern = true, bool? width = true);

    object? AutoOutline();

    /// <summary>
    /// 获取当前数据区域（自动识别连续数据区域）
    /// 对应 Range.CurrentRegion 属性
    /// </summary>
    T CurrentRegion { get; }

    /// <summary>
    /// 获取整个行区域
    /// 对应 Range.EntireRow 属性
    /// </summary>
    T EntireRow { get; }

    /// <summary>
    /// 获取整个列区域
    /// 对应 Range.EntireColumn 属性
    /// </summary>
    T EntireColumn { get; }

    /// <summary>
    /// 获取工作表的已使用区域
    /// 对应 Worksheet.UsedRange 属性
    /// </summary>
    T UsedRange { get; }

    /// <summary>
    /// 获取父级区域对象
    /// 对应 Range.Parent 属性
    /// </summary>
    T ParentRange { get; }

    /// <summary>
    /// 获取父级对象
    /// 对应 Range.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取父级工作表对象
    /// </summary>
    IExcelWorksheet? ParentSheet { get; }

    #region 高级操作

    /// <summary>
    /// 对区域应用自动筛选
    /// 对应 Range.AutoFilter() 方法
    /// </summary>
    void AutoFilter();

    /// <summary>
    /// 移除区域的自动筛选
    /// 对应 Range.AutoFilter() 方法
    /// </summary>
    void RemoveAutoFilter();

    /// <summary>
    /// 对区域进行排序
    /// </summary>
    /// <param name="key1">主要排序关键字</param>
    /// <param name="order1">主要排序顺序</param>
    /// <param name="type"></param>
    /// <param name="key2">次要排序关键字</param>
    /// <param name="order2">次要排序顺序</param>
    /// <param name="key3">第三排序关键字</param>
    /// <param name="order3">第三排序顺序</param>
    /// <param name="header">是否包含标题行</param>
    /// <param name="orderCustom">自定义排序顺序</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="orientation">排序方向</param>
    /// <param name="sortMethod">排序方法</param>
    /// <param name="dataOption1">第一数据选项</param>
    /// <param name="dataOption2">第二数据选项</param>
    /// <param name="dataOption3">第三数据选项</param>
    void Sort(
         object key1,
         object key2,
         object type,
         object key3,
         object orderCustom,
         object matchCase,
         XlSortOrder order1 = XlSortOrder.xlAscending,
         XlSortOrder order2 = XlSortOrder.xlAscending,
         XlSortOrder order3 = XlSortOrder.xlAscending,
         XlYesNoGuess header = XlYesNoGuess.xlNo,
         XlSortOrientation orientation = XlSortOrientation.xlSortRows,
         XlSortMethod sortMethod = XlSortMethod.xlPinYin,
         XlSortDataOption dataOption1 = XlSortDataOption.xlSortNormal,
         XlSortDataOption dataOption2 = XlSortDataOption.xlSortNormal,
         XlSortDataOption dataOption3 = XlSortDataOption.xlSortNormal);

    /// <summary>
    /// 在当前区域中查找下一个匹配的单元格区域
    /// </summary>
    /// <param name="after">查找的起始位置，搜索将从该位置之后开始。为null时表示从区域起始位置开始查找</param>
    /// <returns>返回找到的下一个匹配区域，如果未找到匹配项则返回null</returns>
    T? FindNext(T? after = default);

    /// <summary>
    /// 在当前区域中查找上一个匹配的单元格区域
    /// </summary>
    /// <param name="after">查找的起始位置，搜索将从该位置之前开始。为null时表示从区域末尾位置开始查找</param>
    /// <returns>返回找到的上一个匹配区域，如果未找到匹配项则返回null</returns>
    T? FindPrevious(T? after = default);

    /// <summary>
    /// 在区域中查找指定内容
    /// </summary>
    /// <param name="what">要查找的内容</param>
    /// <param name="after">开始查找的位置</param>
    /// <param name="lookIn">查找范围</param>
    /// <param name="lookAt">匹配方式</param>
    /// <param name="searchOrder">查找顺序</param>
    /// <param name="searchDirection">查找方向</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchByte">是否匹配字节</param>
    /// <param name="searchFormat">查找格式</param>
    /// <returns>找到的区域，未找到则返回null</returns>
    T? Find(object what,
        T? after = default,
        XlFindLookIn? lookIn = null,
        XlLookAt? lookAt = null,
        XlSearchOrder? searchOrder = null,
        XlSearchDirection searchDirection = XlSearchDirection.xlNext,
        bool? matchCase = null,
        bool? matchByte = null,
        object? searchFormat = null);

    /// <summary>
    /// 在单元格区域中查找并替换指定的内容
    /// </summary>
    /// <param name="what">要查找的内容</param>
    /// <param name="replacement">用于替换的内容</param>
    /// <param name="lookAt">指定匹配方式，是完全匹配还是部分匹配</param>
    /// <param name="searchOrder">指定搜索顺序，按行或按列</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <param name="matchByte">是否匹配字节</param>
    /// <param name="searchFormat">搜索的单元格格式条件</param>
    /// <param name="replaceFormat">替换的单元格格式条件</param>
    /// <returns>如果成功执行替换操作则返回true，否则返回false</returns>
    bool Replace(object what, object replacement,
       XlLookAt? lookAt = null, XlSearchOrder? searchOrder = null,
       bool? matchCase = null, bool? matchByte = null,
       object? searchFormat = null, object? replaceFormat = null);

    /// <summary>
    /// 获取区域中的特殊单元格
    /// </summary>
    /// <param name="type">特殊单元格类型</param>
    /// <param name="value">附加参数值</param>
    /// <returns>特殊单元格区域</returns>
    T? SpecialCells(XlCellType type, object? value = null);
    #endregion
}
