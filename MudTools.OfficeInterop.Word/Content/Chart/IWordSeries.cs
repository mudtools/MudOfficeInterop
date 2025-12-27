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
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordSeries : IOfficeObject<IWordSeries>, IDisposable
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
    /// 获取或设置坐标轴组。
    /// </summary>
    XlAxisGroup AxisGroup { get; set; }

    /// <summary>
    /// 获取或设置扇形图扇区偏离饼图中心的距离（以磅为单位）。
    /// </summary>
    int Explosion { get; set; }

    /// <summary>
    /// 获取或设置对象的公式，使用 A1 样式引用。
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置对象的公式，使用用户语言的宏语言本地化形式和 A1 样式引用。
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置对象的公式，使用 R1C1 样式引用。
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置对象的公式，使用用户语言的宏语言本地化形式和 R1C1 样式引用。
    /// </summary>
    string FormulaR1C1Local { get; set; }

    /// <summary>
    /// 获取或设置当数值为负数时是否反转图案或填充。
    /// </summary>
    bool InvertIfNegative { get; set; }

    /// <summary>
    /// 获取或设置图片在系列上的显示方式。
    /// </summary>
    XlChartPictureType PictureType { get; set; }

    /// <summary>
    /// 获取或设置图表图片单位大小（以磅为单位）。
    /// </summary>
    double PictureUnit { get; set; }


    /// <summary>
    /// 获取或设置数据标记的背景色索引。
    /// </summary>
    XlColorIndex MarkerBackgroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置数据标记的前景色索引。
    /// </summary>
    XlColorIndex MarkerForegroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置数据标记的样式。
    /// </summary>
    XlMarkerStyle MarkerStyle { get; set; }

    /// <summary>
    /// 获取系列的内部区域格式属性。
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取系列的填充格式属性。
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取或设置系列在图表中的绘制顺序。
    /// </summary>
    int PlotOrder { get; set; }


    /// <summary>
    /// 获取或设置是否显示数据标签。
    /// </summary>
    IWordChartBorder? Border { get; }

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
    [ValueConvert]
    IWordDataLabels? DataLabels();

    /// <summary>
    /// 获取数据标签。
    /// </summary>
    [ValueConvert]
    IWordDataLabel? DataLabels(int index);

    /// <summary>
    /// 获取数据标签。
    /// </summary>
    [ValueConvert]
    IWordDataLabel? DataLabels(string name);

    /// <summary>
    /// 获取趋势线集合。
    /// </summary>
    [ValueConvert]
    IWordTrendlines? Trendlines();

    /// <summary>
    /// 获取趋势线。
    /// </summary>
    [ValueConvert]
    IWordTrendline? Trendlines(int index);

    /// <summary>
    /// 获取趋势线。
    /// </summary>
    [ValueConvert]
    IWordTrendline? Trendlines(string name);

    /// <summary>
    /// 获取误差线对象。
    /// </summary>
    IWordErrorBars? ErrorBars { get; }

    /// <summary>
    /// 为图表系列添加误差线
    /// </summary>
    /// <param name="direction">误差线的方向，可以是X轴或Y轴方向</param>
    /// <param name="include">指定误差线包含哪些部分，如正值、负值或两者</param>
    /// <param name="type">误差线类型，如固定值、百分比、标准误差等</param>
    /// <param name="amount">正值误差量</param>
    /// <param name="minusValues">负值误差量</param>
    /// <returns>返回表示误差线的对象</returns>
    object? ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type, double? amount = null, double? minusValues = null);

    /// <summary>
    /// 为图表系列应用数据标签。
    /// </summary>
    /// <param name="type">指定要显示的数据标签类型，默认显示数值</param>
    /// <param name="legendKey">是否在数据标签旁边显示图例项标示</param>
    /// <param name="autoText">数据标签的自动文本</param>
    /// <param name="hasLeaderLines">数据标签是否有引导线</param>
    /// <param name="showSeriesName">是否显示系列名称</param>
    /// <param name="showCategoryName">是否显示分类名称</param>
    /// <param name="showValue">是否显示值</param>
    /// <param name="showPercentage">是否显示百分比</param>
    /// <param name="showBubbleSize">是否显示气泡大小</param>
    /// <param name="separator">数据标签中各项之间的分隔符</param>
    void ApplyDataLabels(XlDataLabelsType type = XlDataLabelsType.xlDataLabelsShowValue,
                        bool? legendKey = null, string? autoText = null, bool? hasLeaderLines = null,
                        bool? showSeriesName = null, bool? showCategoryName = null,
                        bool? showValue = null, bool? showPercentage = null,
                        bool? showBubbleSize = null, string? separator = null);

    /// <summary>
    /// 选择系列。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除系列。
    /// </summary>
    void Delete();

    /// <summary>
    /// 删除系列的格式设置并恢复默认格式。
    /// </summary>
    /// <returns>返回操作结果对象。</returns>
    object? ClearFormats();

    /// <summary>
    /// 复制系列对象到剪贴板。
    /// </summary>
    /// <returns>返回复制操作的结果对象。</returns>
    object? Copy();

    /// <summary>
    /// 将剪贴板中的内容粘贴到系列中。
    /// </summary>
    /// <returns>返回粘贴操作的结果对象。</returns>
    object? Paste();

    /// <summary>
    /// 获取系列中的特定单个数据点。
    /// </summary>
    /// <param name="Index">要获取的数据点索引，如果省略则返回所有数据点。</param>
    /// <returns>返回表示数据点或数据点集合的对象。</returns>
    [ValueConvert]
    IWordPoint? Points(int Index);

    /// <summary>
    /// 获取系列中的特定数据点集合。
    /// </summary>
    /// <returns>返回表示数据点或数据点集合的对象。</returns>
    [ValueConvert]
    IWordPoints? Points();
}