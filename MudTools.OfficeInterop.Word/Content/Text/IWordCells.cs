//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格单元格集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordCells : IEnumerable<IWordCell?>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取单元格数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取单元格。
    /// </summary>
    IWordCell? this[int index] { get; }

    /// <summary>
    /// 获取或设置单元格的宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置单元格的高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置单元格高度的规则。
    /// </summary>
    WdRowHeightRule HeightRule { get; set; }

    /// <summary>
    /// 获取或设置单元格内容的垂直对齐方式。
    /// </summary>
    WdCellVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>
    /// 获取单元格的边框集合。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取单元格的底纹设置。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取单元格的嵌套级别。
    /// </summary>
    int NestingLevel { get; }

    /// <summary>
    /// 获取或设置单元格首选宽度。
    /// </summary>
    float PreferredWidth { get; set; }

    /// <summary>
    /// 获取或设置首选宽度的类型。
    /// </summary>
    WdPreferredWidthType PreferredWidthType { get; set; }

    /// <summary>
    /// 添加新的单元格。
    /// </summary>
    /// <param name="beforeCell">在指定单元格前添加。</param>
    /// <returns>新创建的单元格。</returns>
    IWordCell? Add(IWordCell? beforeCell = null);

    /// <summary>
    /// 删除指定索引的单元格。
    /// </summary>
    void Delete(WdDeleteCells? shiftCells);

    /// <summary>
    /// 设置单元格的宽度。
    /// </summary>
    /// <param name="columnWidth">列宽值（以磅为单位）。</param>
    /// <param name="rulerStyle">标尺样式，用于确定宽度设置的参考方式。</param>
    void SetWidth(float columnWidth, WdRulerStyle rulerStyle);

    /// <summary>
    /// 设置单元格的高度。
    /// </summary>
    /// <param name="rowHeight">行高值（以磅为单位）。</param>
    /// <param name="heightRule">行高规则，用于确定高度的计算方式。</param>
    void SetHeight(float rowHeight, WdRowHeightRule heightRule);

    /// <summary>
    /// 合并单元格范围。
    /// </summary>
    void Merge();

    /// <summary>
    /// 拆分所有单元格。
    /// </summary>
    /// <param name="numRows">行数。</param>
    /// <param name="numColumns">列数。</param>
    void Split(int? numRows = null, int? numColumns = null);

    /// <summary>
    /// 自动调整所有单元格大小。
    /// </summary>
    void AutoFit();
}