//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 表格行集合的封装接口。
/// </summary>
public interface IWordRows : IEnumerable<IWordRow>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取行数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取行。
    /// </summary>
    IWordRow this[int index] { get; }

    /// <summary>
    /// 获取第一行。
    /// </summary>
    IWordRow? First { get; }

    /// <summary>
    /// 获取最后一行。
    /// </summary>
    IWordRow? Last { get; }

    /// <summary>
    /// 获取或设置行的边框集合。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取或设置行的底纹格式。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置行的对齐方式。
    /// </summary>
    WdRowAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置行高规则，确定行高是固定值、最小值还是自动调整。
    /// </summary>
    WdRowHeightRule HeightRule { get; set; }

    /// <summary>
    /// 获取或设置表格的方向（从左到右或从右到左）。
    /// </summary>
    WdTableDirection TableDirection { get; set; }

    /// <summary>
    /// 获取或设置行相对于垂直位置的锚定点。
    /// </summary>
    WdRelativeVerticalPosition RelativeVerticalPosition { get; set; }

    /// <summary>
    /// 获取或设置行相对于水平位置的锚定点。
    /// </summary>
    WdRelativeHorizontalPosition RelativeHorizontalPosition { get; set; }

    /// <summary>
    /// 获取或设置行的高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置列之间的间距。
    /// </summary>
    float SpaceBetweenColumns { get; set; }

    /// <summary>
    /// 获取或设置标题格式。如果为 1，则该行将成为表格的标题行，会在跨页时重复显示；如果为 0，则不是标题行。
    /// </summary>
    int HeadingFormat { get; set; }

    /// <summary>
    /// 获取或设置是否允许行跨页断开显示。如果为 1，则允许行跨页断开；如果为 0，则不允许行跨页断开。
    /// </summary>
    int AllowBreakAcrossPages { get; set; }

    /// <summary>
    /// 获取或设置行的左缩进距离。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 设置行的左缩进距离和标尺样式。
    /// </summary>
    /// <param name="leftIndent">左缩进的距离值。</param>
    /// <param name="rulerStyle">应用于缩进的标尺样式。</param>
    void SetLeftIndent(float leftIndent, WdRulerStyle rulerStyle);

    /// <summary>
    /// 添加新的行。
    /// </summary>
    /// <param name="beforeRow">在指定行前添加。</param>
    /// <returns>新创建的行。</returns>
    IWordRow Add(object beforeRow = null);

    /// <summary>
    /// 删除指定索引的行。
    /// </summary>
    /// <param name="index">行索引。</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定范围的行。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="count">删除数量。</param>
    void DeleteRange(int startIndex, int count);

    /// <summary>
    /// 删除所有行。
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取所有行索引列表。
    /// </summary>
    /// <returns>行索引列表。</returns>
    List<int> GetIndexes();

    /// <summary>
    /// 在指定位置插入行。
    /// </summary>
    /// <param name="index">插入位置。</param>
    /// <param name="count">插入行数。</param>
    /// <returns>新插入的行列表。</returns>
    List<IWordRow> Insert(int index, int count = 1);

    /// <summary>
    /// 获取指定范围的行。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    /// <returns>行列表。</returns>
    List<IWordRow> GetRange(int startIndex, int endIndex);

    /// <summary>
    /// 选择所有行。
    /// </summary>
    void Select();

    /// <summary>
    /// 复制所有行。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切所有行。
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到所有行。
    /// </summary>
    void Paste();

    /// <summary>
    /// 清除所有行内容。
    /// </summary>
    void ClearContents();

    /// <summary>
    /// 自动调整所有行高。
    /// </summary>
    void AutoFit();

    /// <summary>
    /// 合并行范围。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    void Merge(int startIndex, int endIndex);

    /// <summary>
    /// 拆分所有行。
    /// </summary>
    /// <param name="numRows">每行拆分后的行数。</param>
    void Split(int numRows);

    /// <summary>
    /// 设置所有行边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条宽度。</param>
    /// <param name="color">颜色。</param>
    void SetBordersForAll(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic);

    /// <summary>
    /// 设置所有行底纹。
    /// </summary>
    /// <param name="pattern">图案。</param>
    /// <param name="foregroundColor">前景色。</param>
    /// <param name="backgroundColor">背景色。</param>
    void SetShadingForAll(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite);
}
