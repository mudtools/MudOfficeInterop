//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 图表的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChart : IDisposable
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
    /// 获取或设置图表类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    XlChartType ChartType { get; set; }

    /// <summary>
    /// 获取或设置是否显示图例。
    /// </summary>
    bool HasLegend { get; set; }

    /// <summary>
    /// 获取或设置是否显示数据表。
    /// </summary>
    bool HasDataTable { get; set; }

    /// <summary>
    /// 获取或设置是否显示标题。
    /// </summary>
    bool HasTitle { get; set; }

    /// <summary>
    /// 获取图表区域。
    /// </summary>
    IWordChartArea? ChartArea { get; }

    /// <summary>
    /// 获取绘图区域。
    /// </summary>
    IWordPlotArea? PlotArea { get; }

    /// <summary>
    /// 获取图例。
    /// </summary>
    IWordLegend? Legend { get; }

    /// <summary>
    /// 获取图表标题。
    /// </summary>
    IWordChartTitle? ChartTitle { get; }

    /// <summary>
    /// 获取数据表。
    /// </summary>
    IWordDataTable? DataTable { get; }

    /// <summary>
    /// 获取图表数据。
    /// </summary>
    IWordChartData? ChartData { get; }

    /// <summary>
    /// 获取系列集合。
    /// </summary>
    [ReturnValueConvert]
    IWordSeriesCollection? SeriesCollection();

    /// <summary>
    /// 根据索引获取系列集合。
    /// </summary>
    /// <param name="index">要获取的系列集合的从零开始的索引。</param>
    /// <returns>指定索引处的系列集合。</returns>
    [ReturnValueConvert]
    IWordSeriesCollection? SeriesCollection(int index);

    /// <summary>
    /// 根据名称获取系列集合。
    /// </summary>
    /// <param name="name">要获取的系列集合的名称。</param>
    /// <returns>具有指定名称的系列集合。</returns>
    [ReturnValueConvert]
    IWordSeriesCollection? SeriesCollection(string name);

    /// <summary>
    /// 获取图表组集合。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    IWordChartGroups? ChartGroups { get; }

    /// <summary>
    /// 获取墙壁。
    /// </summary>
    IWordWalls? Walls { get; }

    /// <summary>
    /// 获取地板。
    /// </summary>
    IWordFloor? Floor { get; }

    /// <summary>
    /// 应用数据标签。
    /// </summary>
    /// <param name="type">标签类型。</param>
    void ApplyDataLabels(XlDataLabelsType type);

    /// <summary>
    /// 设置图表数据源。
    /// </summary>
    /// <param name="source">数据源。</param>
    /// <param name="plotBy">绘制方式。</param>
    void SetSourceData(string source, XlRowCol plotBy);

    /// <summary>
    /// 选择图表。
    /// </summary>
    void Select();

    /// <summary>
    /// 复制图表。
    /// </summary>
    void Copy();

    /// <summary>
    /// 删除图表。
    /// </summary>
    void Delete();

    /// <summary>
    /// 刷新图表。
    /// </summary>
    void Refresh();

    /// <summary>
    /// 导出图表。
    /// </summary>
    /// <param name="filename">文件名。</param>
    /// <param name="filterName">过滤器名称。</param>
    /// <param name="interactive">是否交互式。</param>
    void Export(string filename, string? filterName = null, bool? interactive = null);

    /// <summary>
    /// 设置图表元素。
    /// </summary>
    /// <param name="element">图表元素。</param>
    void SetElement([ComNamespace("MsCore")] MsoChartElementType element);
}