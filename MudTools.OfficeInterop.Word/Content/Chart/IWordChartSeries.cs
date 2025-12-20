//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表系列的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord", ComClassName = "Series")]
public interface IWordChartSeries : IDisposable
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
    /// 获取或设置系列名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置系列值。
    /// </summary>
    object Values { get; set; }

    /// <summary>
    /// 获取或设置分类轴标签。
    /// </summary>
    object XValues { get; set; }

    /// <summary>
    /// 获取或设置气泡大小值。
    /// </summary>
    object BubbleSizes { get; set; }

    /// <summary>
    /// 获取或设置图表类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    XlChartType ChartType { get; set; }

    /// <summary>
    /// 获取或设置是否平滑线。
    /// </summary>
    bool Smooth { get; set; }

    /// <summary>
    /// 获取或设置标记大小。
    /// </summary>
    int MarkerSize { get; set; }

    /// <summary>
    /// 获取或设置标记背景色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerBackgroundColor { get; set; }

    /// <summary>
    /// 获取或设置标记前景色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerForegroundColor { get; set; }

    /// <summary>
    /// 获取或设置是否显示数据标签。
    /// </summary>
    bool HasDataLabels { get; set; }

    /// <summary>
    /// 获取或设置是否显示趋势线。
    /// </summary>
    bool HasErrorBars { get; set; }

    /// <summary>
    /// 获取系列格式。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 获取数据标签集合。
    /// </summary>
    [ReturnValueConvert]
    IWordDataLabels? DataLabels();

    /// <summary>
    /// 获取数据标签。
    /// </summary>
    [ReturnValueConvert]
    IWordDataLabel? DataLabels(int index);

    /// <summary>
    /// 获取数据标签。
    /// </summary>
    [ReturnValueConvert]
    IWordDataLabel? DataLabels(string name);

    /// <summary>
    /// 获取趋势线集合。
    /// </summary>
    [ReturnValueConvert]
    IWordTrendlines? Trendlines();

    /// <summary>
    /// 获取趋势线。
    /// </summary>
    [ReturnValueConvert]
    IWordTrendline? Trendlines(int index);

    /// <summary>
    /// 获取趋势线。
    /// </summary>
    [ReturnValueConvert]
    IWordTrendline? Trendlines(string name);

    /// <summary>
    /// 获取误差线对象。
    /// </summary>
    IWordErrorBars? ErrorBars { get; }

    /// <summary>
    /// 应用数据标签。
    /// </summary>
    /// <param name="type">标签类型。</param>
    /// <param name="legendKey">是否显示图例项标示。</param>
    /// <param name="autoText">是否自动文本。</param>
    /// <param name="hasLeaderLines">是否有引导线。</param>
    void ApplyDataLabels(XlDataLabelsType type, bool legendKey, bool autoText, bool hasLeaderLines);

    /// <summary>
    /// 选择系列。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除系列。
    /// </summary>
    void Delete();
}