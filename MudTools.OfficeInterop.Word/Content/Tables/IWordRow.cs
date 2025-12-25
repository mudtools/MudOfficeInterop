//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格行的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordRow : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }


    /// <summary>
    /// 返回表示包含在指定对象中的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取或设置一个值，确定表格中的一行或多行文本是否可以跨页断点拆分。
    /// </summary>
    int AllowBreakAcrossPages { get; set; }

    /// <summary>
    /// 获取或设置一个 WdRowAlignment 常量，表示指定行的对齐方式。
    /// </summary>
    WdRowAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，确定指定的一行或多行是否格式化为表格标题。
    /// </summary>
    int HeadingFormat { get; set; }

    /// <summary>
    /// 获取或设置指定行中相邻列之间文本的距离（以磅为单位）。
    /// </summary>
    float SpaceBetweenColumns { get; set; }

    /// <summary>
    /// 获取或设置表格中指定行的高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置确定指定行高度的规则。
    /// </summary>
    WdRowHeightRule HeightRule { get; set; }

    /// <summary>
    /// 获取或设置指定表格行的左缩进值（以磅为单位）。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 确定指定列或行是否是表格中的最后一行。
    /// </summary>
    bool IsLast { get; }

    /// <summary>
    /// 确定指定列或行是否是表格中的第一行。
    /// </summary>
    bool IsFirst { get; }

    /// <summary>
    /// 返回表示集合中项目位置的整数。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 返回表示列、行、选择或范围中的表格单元格的 Cells 集合。
    /// </summary>
    IWordCells? Cells { get; }

    /// <summary>
    /// 返回表示指定对象的所有边框的 Borders 集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 返回引用指定对象的阴影格式的 Shading 对象。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 返回集合中的下一个对象。
    /// </summary>
    IWordRow? Next { get; }

    /// <summary>
    /// 返回集合中的上一个对象。
    /// </summary>
    IWordRow? Previous { get; }

    /// <summary>
    /// 获取指定行的嵌套级别。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置当当前文档保存为网页时，指定对象的标识标签。
    /// </summary>
    string ID { get; set; }

    /// <summary>
    /// 选择指定的对象。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 设置表格中一行或多行的缩进。
    /// </summary>
    /// <param name="leftIndent">必需 Single。指定行或多行的当前左边缘与所需左边缘之间的距离（以磅为单位）。</param>
    /// <param name="rulerStyle">必需 WdRulerStyle。控制当左缩进更改时 Microsoft Word 调整表格的方式。WdRulerStyle 可以是以下常量之一：wdAdjustNone - 调整行或多行的左边缘，通过向左或向右移动列来保留所有列的宽度。这是默认值。wdAdjustSameWidth - 调整第一列的左边缘，通过将指定行或多行中所有单元格的宽度设置为相同值来保留表格右边缘的位置。wdAdjustFirstColumn - 仅调整第一列的左边缘，保留其他列的位置和表格的右边缘。wdAdjustProportional - 调整第一列的左边缘，通过按比例调整指定行或多行中所有单元格的宽度来保留表格右边缘的位置。</param>
    void SetLeftIndent(float leftIndent, WdRulerStyle rulerStyle);

    /// <summary>
    /// 设置表格行的高度。
    /// </summary>
    /// <param name="rowHeight">必需 Single。行或多行的高度（以磅为单位）。</param>
    /// <param name="heightRule">必需 WdRowHeightRule。确定指定行高度的规则。WdRowHeightRule 可以是以下常量之一：wdRowHeightAtLeast、wdRowHeightExactly、wdRowHeightAuto。</param>
    void SetHeight(float rowHeight, WdRowHeightRule heightRule);

    /// <summary>
    /// 将表格转换为文本，并返回表示分隔文本的 Range 对象。
    /// </summary>
    /// <param name="separator">可选 Object。分隔转换后列的字符（段落标记分隔转换后的行）。可以是以下任何 Microsoft.Office.Interop.Word.WdTableFieldSeparator 常量：wdSeparateByCommas、wdSeparateByDefaultListSeparator、wdSeparateByParagraphs、wdSeparateByTabs（默认）。</param>
    /// <param name="nestedTables">可选 Object。如果为 True，则嵌套表格将被转换为文本。如果 Separator 不是 wdSeparateByParagraphs，则忽略此参数。默认值为 True。</param>
    /// <returns>Microsoft.Office.Interop.Word.Range</returns>
    IWordRange? ConvertToText(WdTableFieldSeparator? separator = null, bool? nestedTables = null);
}