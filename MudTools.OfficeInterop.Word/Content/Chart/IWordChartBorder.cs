using System.Drawing;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表边框格式的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartBorder : IDisposable
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
    /// 获取或设置边框颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

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
    [ComPropertyWrap(NeedConvert = true)]
    float Weight { get; set; }
}