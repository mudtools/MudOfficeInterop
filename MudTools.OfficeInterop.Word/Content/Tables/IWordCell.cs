//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格单元格的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCell : IOfficeObject<IWordCell>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取单元格行索引。
    /// </summary>
    int RowIndex { get; }

    /// <summary>
    /// 获取单元格列索引。
    /// </summary>
    int ColumnIndex { get; }

    /// <summary>
    /// 获取单元格范围。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取单元格所在的行。
    /// </summary>
    IWordRow? Row { get; }

    /// <summary>
    /// 获取单元格所在的列。
    /// </summary>
    IWordColumn? Column { get; }

    /// <summary>
    /// 获取或设置单元格宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置单元格高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置单元格垂直对齐方式。
    /// </summary>
    WdCellVerticalAlignment VerticalAlignment { get; set; }


    /// <summary>
    /// 获取单元格边框。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取单元格底纹。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取单元格首选宽度。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置单元格首选宽度类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 获取或设置单元格是否适应文本。
    /// </summary>
    bool FitText { get; set; }

    /// <summary>
    /// 获取单元格左边界位置。
    /// </summary>
    float LeftPadding { get; set; }

    /// <summary>
    /// 获取单元格右边界位置。
    /// </summary>
    float RightPadding { get; set; }

    /// <summary>
    /// 获取单元格上边界位置。
    /// </summary>
    float TopPadding { get; set; }

    /// <summary>
    /// 获取单元格下边界位置。
    /// </summary>
    float BottomPadding { get; set; }

    /// <summary>
    /// 获取单元格所在的表格。
    /// </summary>
    IWordTables? Tables { get; }

    /// <summary>
    /// 获取单元格的嵌套级别。嵌套级别表示表格在嵌套表格结构中的层级，主表格为级别1，嵌套在其中的表格为更高级别。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置单元格内的文本是否自动换行。
    /// </summary>
    bool WordWrap { get; set; }

    /// <summary>
    /// 获取或设置单元格的标识符。
    /// </summary>
    string ID { get; set; }

    /// <summary>
    /// 获取或设置单元格所在行的高度规则。
    /// </summary>
    WdRowHeightRule HeightRule { get; set; }

    /// <summary>
    /// 获取下一个相邻的单元格。
    /// </summary>
    IWordCell? Next { get; }

    /// <summary>
    /// 获取前一个相邻的单元格。
    /// </summary>
    IWordCell? Previous { get; }

    /// <summary>
    /// 选择单元格。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除单元格。
    /// </summary>
    void Delete(WdDeleteCells? shiftCells = null);

    /// <summary>
    /// 合并单元格。
    /// </summary>
    /// <param name="mergeTo">要合并到的单元格。</param>
    void Merge(IWordCell mergeTo);

    /// <summary>
    /// 拆分单元格。
    /// </summary>
    /// <param name="numRows">行数。</param>
    /// <param name="numColumns">列数。</param>
    void Split(int numRows, int numColumns);

    /// <summary>
    /// 在单元格中插入公式。
    /// </summary>
    /// <param name="formula">要插入的公式，如果为null则使用默认公式</param>
    /// <param name="numFormat">数字格式代码，用于格式化结果，如果为null则使用默认格式</param>
    void Formula(string? formula = null, string? numFormat = null);

    /// <summary>
    /// 设置单元格的宽度。
    /// </summary>
    /// <param name="columnWidth">列宽值（以磅为单位）</param>
    /// <param name="rulerStyle">标尺样式，指定如何测量宽度</param>
    void SetWidth(float columnWidth, WdRulerStyle rulerStyle);

    /// <summary>
    /// 设置单元格的高度。
    /// </summary>
    /// <param name="rowHeight">行高值（以磅为单位）</param>
    /// <param name="heightRule">行高规则，指定如何应用高度设置</param>
    void SetHeight(float rowHeight, WdRowHeightRule heightRule);

    /// <summary>
    /// 对单元格上方的数字进行自动求和。
    /// </summary>
    void AutoSum();
}
