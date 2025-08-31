namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Borders 的接口，用于操作边框集合。
/// </summary>
public interface IWordBorders : IEnumerable<IWordBorder>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取边框的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据边框类型获取边框。
    /// </summary>
    IWordBorder this[WdBorderType borderType] { get; }

    /// <summary>
    /// 获取或设置是否启用边框。
    /// </summary>
    bool Enable { get; set; }

    /// <summary>
    /// 应用边框样式到所有边框。
    /// </summary>
    /// <param name="lineStyle">线条样式。</param>
    /// <param name="lineWidth">线条粗细。</param>
    /// <param name="color">颜色。</param>
    void ApplyStyle(WdLineStyle lineStyle, WdLineWidth lineWidth, WdColor color);

    /// <summary>
    /// 获取指定类型的边框是否存在。
    /// </summary>
    /// <param name="borderType">边框类型。</param>
    /// <returns>是否存在。</returns>
    bool Contains(WdBorderType borderType);

    /// <summary>
    /// 获取所有边框类型的列表。
    /// </summary>
    /// <returns>边框类型列表。</returns>
    List<WdBorderType> GetBorderTypes();
}