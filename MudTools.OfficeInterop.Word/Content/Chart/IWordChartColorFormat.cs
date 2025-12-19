using System.Drawing;

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表颜色格式的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartColorFormat : IDisposable
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
    /// 获取或设置 RGB 颜色值。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color RGB { get; }

    /// <summary>
    /// 获取或设置图表元素的方案颜色。方案颜色是基于当前文档主题的一组预定义颜色之一。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color SchemeColor { get; set; }

    /// <summary>
    /// 获取颜色类型。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    MsoColorType Type { get; }
}