//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中表格的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTable : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 返回表示包含在指定对象中的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取指定表格中的所有行是否具有相同数量的列。
    /// </summary>
    bool Uniform { get; }

    /// <summary>
    /// 返回已应用于指定表格的自动格式类型。
    /// </summary>
    int AutoFormatType { get; }

    /// <summary>
    /// 获取指定表格的嵌套级别。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置一个值，允许 Microsoft Word 在表格中跨页断点。
    /// </summary>
    bool AllowPageBreaks { get; set; }

    /// <summary>
    /// 获取或设置一个值，允许 Microsoft Word 自动调整表格中单元格的大小以适应其内容。
    /// </summary>
    bool AllowAutoFit { get; set; }

    /// <summary>
    /// 获取或设置指定单元格、单元格组、列或表格的首选宽度（以磅为单位或窗口宽度的百分比）。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置用于指定表格宽度的首选测量单位。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 获取或设置要添加到单个单元格或表格中所有单元格内容上方的空间量（以磅为单位）。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取或设置要添加到单个单元格或表格中所有单元格内容下方的空间量（以磅为单位）。
    /// </summary>
    float BottomPadding { get; set; }

    /// <summary>
    /// 获取或设置要添加到单个单元格或表格中所有单元格内容左侧的空间量（以磅为单位）。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取或设置要添加到单个单元格或表格中所有单元格内容右侧的空间量（以磅为单位）。
    /// </summary>
    float RightPadding { get; set; }

    /// <summary>
    /// 获取或设置表格中单元格之间的间距（以磅为单位）。
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Word 在指定表格中排序单元格的方向。
    /// </summary>
    WdTableDirection TableDirection { get; set; }

    /// <summary>
    /// 获取或设置当当前文档保存为网页时，指定对象的标识标签。
    /// </summary>
    string ID { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, NeedConvert = true)]
    IWordStyle? Style { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    WdBuiltinStyle? StyleEnum { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    string? StyleName { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将标题行格式应用于所选表格的第一行。
    /// </summary>
    bool ApplyStyleHeadingRows { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将最后一行格式应用于指定表格的最后一行。
    /// </summary>
    bool ApplyStyleLastRow { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将第一列格式应用于指定表格的第一列。
    /// </summary>
    bool ApplyStyleFirstColumn { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将最后一列格式应用于指定表格的最后一列。
    /// </summary>
    bool ApplyStyleLastColumn { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示如果应用的预设表格样式为行提供样式条带，是否将样式条带应用于表格中的行。
    /// </summary>
    bool ApplyStyleRowBands { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示如果应用的预设表格样式为列提供样式条带，是否将样式条带应用于表格中的列。
    /// </summary>
    bool ApplyStyleColumnBands { get; set; }

    /// <summary>
    /// 获取或设置包含指定表格标题的字符串。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置包含指定表格描述的字符串。
    /// </summary>
    string Descr { get; set; }

    /// <summary>
    /// 返回表示表格中所有表格列的 Columns 集合。
    /// </summary>
    IWordColumns? Columns { get; }

    /// <summary>
    /// 返回表示表格中所有表格行的 Rows 集合。
    /// </summary>
    IWordRows Rows { get; }

    /// <summary>
    /// 获取或设置表示指定对象的所有边框的 Borders 集合。
    /// </summary>
    IWordBorders Borders { get; set; }

    /// <summary>
    /// 返回引用指定对象的阴影格式的 Shading 对象。
    /// </summary>
    IWordShading Shading { get; }

    /// <summary>
    /// 返回表示指定表格中所有表格的 Tables 集合。
    /// </summary>
    IWordTables Tables { get; }

    /// <summary>
    /// 选择指定的对象。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 按字母数字升序排序表格行。
    /// </summary>
    void SortAscending();

    /// <summary>
    /// 按字母数字降序排序表格行。
    /// </summary>
    void SortDescending();

    /// <summary>
    /// 将预定义的外观应用于表格。
    /// </summary>
    /// <param name="format">可选 Object。</param>
    /// <param name="applyBorders">可选 Object。如果为 True，则应用指定格式的边框属性。默认值为 True。</param>
    /// <param name="applyShading">可选 Object。如果为 True，则应用指定格式的阴影属性。默认值为 True。</param>
    /// <param name="applyFont">可选 Object。如果为 True，则应用指定格式的字体属性。默认值为 True。</param>
    /// <param name="applyColor">可选 Object。如果为 True，则应用指定格式的颜色属性。默认值为 True。</param>
    /// <param name="applyHeadingRows">可选 Object。如果为 True，则应用指定格式的标题行属性。默认值为 True。</param>
    /// <param name="applyLastRow">可选 Object。如果为 True，则应用指定格式的最后一行属性。默认值为 False。</param>
    /// <param name="applyFirstColumn">可选 Object。如果为 True，则应用指定格式的第一列属性。默认值为 True。</param>
    /// <param name="applyLastColumn">可选 Object。如果为 True，则应用指定格式的最后一列属性。默认值为 False。</param>
    /// <param name="autoFit">可选 Object。如果为 True，则在尽可能不更改单元格中文本换行方式的情况下减小表格列的宽度。默认值为 True。</param>
    void AutoFormat(object format = null, bool? applyBorders = null, bool? applyShading = null,
                    bool? applyFont = null, bool? applyColor = null, bool? applyHeadingRows = null,
                    bool? applyLastRow = null, bool? applyFirstColumn = null, bool? applyLastColumn = null,
                    bool? autoFit = null);

    /// <summary>
    /// 使用预定义表格格式的特征更新表格。
    /// </summary>
    void UpdateAutoFormat();

    /// <summary>
    /// 返回表示表格中单元格的 Cell 对象。
    /// </summary>
    /// <param name="row">必需 Integer。要返回的表格中的行号。可以是 1 到表格中行数之间的整数。</param>
    /// <param name="column">必需 Integer。要返回的表格中的单元格号。可以是 1 到表格中列数之间的整数。</param>
    /// <returns>Microsoft.Office.Interop.Word.Cell</returns>
    IWordCell? Cell(int row, int column);

    /// <summary>
    /// 在表格中指定行之前插入一个空段落，并返回一个 Table 对象，该对象包含指定行及其后的行。
    /// </summary>
    /// <param name="beforeRow">必需 Object。表格要拆分之前的行。可以是行号或 Row 对象。</param>
    /// <returns>Microsoft.Office.Interop.Word.Table</returns>
    IWordTable? Split(IWordRow beforeRow);

    /// <summary>
    /// 在表格中指定行之前插入一个空段落，并返回一个 Table 对象，该对象包含指定行及其后的行。
    /// </summary>
    /// <param name="beforeRow">必需 Object。表格要拆分之前的行。可以是行号或 Row 对象。</param>
    /// <returns>Microsoft.Office.Interop.Word.Table</returns>
    IWordTable? Split(int beforeRow);

    /// <summary>
    /// 将表格转换为文本，并返回表示分隔文本的 Range 对象。
    /// </summary>
    /// <param name="separator">可选 Object。分隔转换后列的字符（段落标记分隔转换后的行）。可以是以下任何 Microsoft.Office.Interop.Word.WdTableFieldSeparator 常量：wdSeparateByCommas、wdSeparateByDefaultListSeparator、wdSeparateByParagraphs、wdSeparateByTabs（默认）。</param>
    /// <param name="nestedTables">可选 Object。如果为 True，则嵌套表格将被转换为文本。如果 Separator 不是 wdSeparateByParagraphs，则忽略此参数。默认值为 True。</param>
    /// <returns>Microsoft.Office.Interop.Word.Range</returns>
    IWordRange? ConvertToText(WdTableFieldSeparator? separator = null, bool? nestedTables = null);

    /// <summary>
    /// 确定 Microsoft Word 在使用自动调整功能时如何调整表格大小。
    /// </summary>
    /// <param name="behavior">必需 WdAutoFitBehavior。当使用自动调整功能时 Word 如何调整指定表格的大小。可以是以下 WdAutoFitBehavior 常量之一：wdAutoFitContent、wdAutoFitWindow、wdAutoFitFixed。</param>
    void AutoFitBehavior(WdAutoFitBehavior behavior);

    /// <summary>
    /// 对指定表格进行排序。
    /// </summary>
    /// <param name="excludeHeader">可选 Object。如果为 True，则从排序操作中排除第一行或段落标题。默认值为 False。</param>
    /// <param name="fieldNumber">可选 Object。要排序的字段。Microsoft Word 按 FieldNumber、然后按 FieldNumber2、然后按 FieldNumber3 排序。</param>
    /// <param name="sortFieldType">可选 Object。FieldNumber、FieldNumber2 和 FieldNumber3 的相应排序类型。可以是以下 Microsoft.Office.Interop.Word.WdSortFieldType 常量之一：wdSortFieldAlphanumeric、wdSortFieldDate、wdSortFieldJapanJIS、wdSortFieldKoreaKS、wdSortFieldNumeric、wdSortFieldStroke、wdSortFieldSyllable。</param>
    /// <param name="sortOrder">可选 Object。对 FieldNumber、FieldNumber2 和 FieldNumber3 进行排序时使用的排序顺序。可以是以下 Microsoft.Office.Interop.Word.WdSortOrder 常量之一：wdSortOrderAscending（默认）、wdSortOrderDescending。</param>
    /// <param name="fieldNumber2">可选 Object。要排序的字段。Word 按 FieldNumber、然后按 FieldNumber2、然后按 FieldNumber3 排序。</param>
    /// <param name="sortFieldType2">可选 Object。FieldNumber、FieldNumber2 和 FieldNumber3 的相应排序类型。可以是以下 Microsoft.Office.Interop.Word.WdSortFieldType 常量之一：wdSortFieldAlphanumeric、wdSortFieldDate、wdSortFieldJapanJIS、wdSortFieldKoreaKS、wdSortFieldNumeric、wdSortFieldStroke、wdSortFieldSyllable。</param>
    /// <param name="sortOrder2">可选 Object。对 FieldNumber、FieldNumber2 和 FieldNumber3 进行排序时使用的排序顺序。可以是以下 Microsoft.Office.Interop.Word.WdSortOrder 常量之一：wdSortOrderAscending（默认）、wdSortOrderDescending。</param>
    /// <param name="fieldNumber3">可选 Object。要排序的字段。Microsoft Word 按 FieldNumber、然后按 FieldNumber2、然后按 FieldNumber3 排序。</param>
    /// <param name="sortFieldType3">可选 Object。FieldNumber、FieldNumber2 和 FieldNumber3 的相应排序类型。可以是以下 Microsoft.Office.Interop.Word.WdSortFieldType 常量之一：wdSortFieldAlphanumeric、wdSortFieldDate、wdSortFieldJapanJIS、wdSortFieldKoreaKS、wdSortFieldNumeric、wdSortFieldStroke、wdSortFieldSyllable。</param>
    /// <param name="sortOrder3">可选 Object。对 FieldNumber、FieldNumber2 和 FieldNumber3 进行排序时使用的排序顺序。可以是以下 Microsoft.Office.Interop.Word.WdSortOrder 常量之一：wdSortOrderAscending（默认）、wdSortOrderDescending。</param>
    /// <param name="caseSensitive">可选 Object。如果为 True，则进行区分大小写的排序。默认值为 False。</param>
    /// <param name="bidiSort">可选 Object。如果为 True，则基于从右到左语言规则进行排序。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="ignoreThe">可选 Object。如果为 True，则在排序从右到左语言文本时忽略阿拉伯字符 alef lam。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="ignoreKashida">可选 Object。如果为 True，则在排序从右到左语言文本时忽略 kashidas。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="ignoreDiacritics">可选 Object。如果为 True，则在排序从右到左语言文本时忽略双向控制字符。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="ignoreHe">可选 Object。如果为 True，则在排序从右到左语言文本时忽略希伯来字符 he。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="languageID">可选 Object。指定排序语言。可以是 Microsoft.Office.Interop.Word.WdLanguageID 常量之一。</param>
    void Sort(bool? excludeHeader = null, string? fieldNumber = null, WdSortFieldType? sortFieldType = null, WdSortFieldType? sortOrder = null,
            string? fieldNumber2 = null, WdSortFieldType? sortFieldType2 = null,
            WdSortFieldType? sortOrder2 = null, string? fieldNumber3 = null, WdSortFieldType? sortFieldType3 = null, WdSortFieldType? sortOrder3 = null,
            bool? caseSensitive = null, bool? bidiSort = null, bool? ignoreThe = null,
            bool? ignoreKashida = null, bool? ignoreDiacritics = null, bool? ignoreHe = null, WdLanguageID? languageID = null);

    /// <summary>
    /// 应用指定样式，但保留用户直接应用的任何格式。
    /// </summary>
    /// <param name="styleName">要应用的样式名称。</param>
    void ApplyStyleDirectFormatting(string styleName);
}