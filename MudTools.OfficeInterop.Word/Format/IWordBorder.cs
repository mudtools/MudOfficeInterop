namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Border 的接口，用于操作单个边框样式。
/// </summary>
public interface IWordBorder : IDisposable
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
    /// 获取或设置边框是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置边框线条样式。
    /// </summary>
    WdLineStyle LineStyle { get; set; }

    /// <summary>
    /// 获取或设置边框线条粗细。
    /// </summary>
    WdLineWidth LineWidth { get; set; }

    /// <summary>
    /// 获取或设置边框颜色（RGB值）。
    /// </summary>
    WdColor Color { get; set; }

    /// <summary>
    /// 获取或设置边框颜色索引。
    /// </summary>
    WdColorIndex ColorIndex { get; set; }


    /// <summary>
    /// 获取或设置是否应用艺术边框样式。
    /// </summary>
    WdPageBorderArt ArtStyle { get; set; }

    /// <summary>
    /// 获取或设置艺术边框的长度。
    /// </summary>
    int ArtWidth { get; set; }

    /// <summary>
    /// 获取边框是否为内置样式。
    /// </summary>
    bool Inside { get; }
}
