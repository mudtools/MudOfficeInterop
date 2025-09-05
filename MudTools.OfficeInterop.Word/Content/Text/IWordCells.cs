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
public interface IWordCells : IEnumerable<IWordCell>, IDisposable
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
    /// 获取单元格数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取单元格。
    /// </summary>
    IWordCell this[int index] { get; }

    /// <summary>
    /// 获取第一个单元格。
    /// </summary>
    IWordCell First { get; }

    /// <summary>
    /// 获取最后一个单元格。
    /// </summary>
    IWordCell Last { get; }


    /// <summary>
    /// 添加新的单元格。
    /// </summary>
    /// <param name="beforeCell">在指定单元格前添加。</param>
    /// <returns>新创建的单元格。</returns>
    IWordCell Add(object beforeCell = null);

    /// <summary>
    /// 删除指定索引的单元格。
    /// </summary>
    /// <param name="index">单元格索引。</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定范围的单元格。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="count">删除数量。</param>
    void DeleteRange(int startIndex, int count);

    /// <summary>
    /// 删除所有单元格。
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取所有单元格索引列表。
    /// </summary>
    /// <returns>单元格索引列表。</returns>
    List<int> GetIndexes();

    /// <summary>
    /// 合并单元格范围。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    void Merge(int startIndex, int endIndex);

    /// <summary>
    /// 拆分所有单元格。
    /// </summary>
    /// <param name="numRows">行数。</param>
    /// <param name="numColumns">列数。</param>
    void Split(int numRows, int numColumns);

    /// <summary>
    /// 获取指定范围的单元格。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    /// <returns>单元格列表。</returns>
    List<IWordCell> GetRange(int startIndex, int endIndex);

    /// <summary>
    /// 选择所有单元格。
    /// </summary>
    void Select();

    /// <summary>
    /// 复制所有单元格。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切所有单元格。
    /// </summary>
    void Cut();

    /// <summary>
    /// 粘贴内容到所有单元格。
    /// </summary>
    void Paste();

    /// <summary>
    /// 清除所有单元格内容。
    /// </summary>
    void ClearContents();

    /// <summary>
    /// 自动调整所有单元格大小。
    /// </summary>
    void AutoFit();

    /// <summary>
    /// 设置所有单元格边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条宽度。</param>
    /// <param name="color">颜色。</param>
    void SetBordersForAll(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic);

    /// <summary>
    /// 设置所有单元格底纹。
    /// </summary>
    /// <param name="pattern">图案。</param>
    /// <param name="foregroundColor">前景色。</param>
    /// <param name="backgroundColor">背景色。</param>
    void SetShadingForAll(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite);
}