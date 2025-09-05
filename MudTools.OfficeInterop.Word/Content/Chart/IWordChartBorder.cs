namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表边框格式的封装接口。
/// </summary>
public interface IWordChartBorder : IDisposable
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
    /// 获取或设置边框颜色。
    /// </summary>
    object Color { get; set; }

    /// <summary>
    /// 获取或设置边框颜色索引。
    /// </summary>
    XlColorIndex ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置边框线条样式。
    /// </summary>
    XlLineStyle LineStyle { get; set; }

    /// <summary>
    /// 获取或设置边框粗细。
    /// </summary>
    float Weight { get; set; }
}