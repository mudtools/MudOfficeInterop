//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表绘图区域的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordPlotArea : IOfficeObject<IWordPlotArea>, IDisposable
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
    /// 获取或设置绘图区域名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置左边距。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置顶边距。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置宽度。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置高度。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置绘图区内部区域距离图表左边的距离。
    /// </summary>
    double InsideLeft { get; set; }

    /// <summary>
    /// 获取或设置绘图区内部区域距离图表顶部的距离。
    /// </summary>
    double InsideTop { get; set; }

    /// <summary>
    /// 获取或设置绘图区内部区域的宽度。
    /// </summary>
    double InsideWidth { get; set; }

    /// <summary>
    /// 获取或设置绘图区内部区域的高度。
    /// </summary>
    double InsideHeight { get; set; }

    /// <summary>
    /// 获取或设置绘图区的位置类型。
    /// </summary>
    XlChartElementPosition Position { get; set; }

    /// <summary>
    /// 获取或设置是否填充。
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取内部区域格式。
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取边框格式。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 选择绘图区域。
    /// </summary>
    void Select();

    /// <summary>
    /// 清除绘图区域格式。
    /// </summary>
    void ClearFormats();
}