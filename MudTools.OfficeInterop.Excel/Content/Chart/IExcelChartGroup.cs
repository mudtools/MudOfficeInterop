
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 图表中一组同类型数据系列的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ChartGroup
/// 用于控制系列的间隙宽度、重叠、坐标轴、数据标签、趋势线等共享属性。
/// </summary>
public interface IExcelChartGroup : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Chart）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置系列之间的间隙宽度（0-500，100=默认）。
    /// 值越大，柱形/条形之间的空隙越大。
    /// </summary>
    int GapWidth { get; set; }

    /// <summary>
    /// 获取或设置同组系列之间的重叠程度（-100 到 100）。
    /// 正值表示重叠，负值表示分离。
    /// </summary>
    int Overlap { get; set; }

    /// <summary>
    /// 获取或设置是否显示高低点连线（Hi-Lo Lines）。
    /// </summary>
    bool HasHiLoLines { get; set; }

    /// <summary>
    /// 获取或设置是否显示垂线（Drop Lines）。
    /// </summary>
    bool HasDropLines { get; set; }

    /// <summary>
    /// 获取或设置是否显示上涨/下跌柱（Up/Down Bars）。
    /// </summary>
    bool HasUpDownBars { get; set; }

    /// <summary>
    /// 获取图表组的类型（柱形图、折线图、饼图等）。
    /// 使用 <see cref="XlChartType"/> 枚举。
    /// </summary>
    XlChartType Type { get; }

    /// <summary>
    /// 获取图表组的索引（从 1 开始）。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取图表组中的系列集合。
    /// 返回封装后的 <see cref="IExcelSeriesCollection"/> 接口。
    /// </summary>
    IExcelSeriesCollection? SeriesCollection(int? index = null);

    /// <summary>
    /// 获取图表组的系列连线对象，用于设置系列之间的连接线格式
    /// 返回封装后的 <see cref="IExcelSeriesLines"/> 接口
    /// </summary>
    IExcelSeriesLines? SeriesLines { get; }

    /// <summary>
    /// 获取雷达图坐标轴标签对象，用于设置雷达图坐标轴的标签格式
    /// 返回封装后的 <see cref="IExcelTickLabels"/> 接口
    /// </summary>
    IExcelTickLabels? RadarAxisLabels { get; }

    /// <summary>
    /// 获取或设置图表组的子类型（与主类型组合定义具体图表类型）
    /// </summary>
    int SubType { get; set; }

    /// <summary>
    /// 获取或设置气泡图中气泡的缩放比例（1-300，100=默认）
    /// 值越大，气泡显示得越大
    /// </summary>
    int BubbleScale { get; set; }

    /// <summary>
    /// 获取或设置是否在气泡图中显示负值气泡
    /// true = 显示负值气泡，false = 隐藏负值气泡
    /// </summary>
    bool ShowNegativeBubbles { get; set; }

    /// <summary>
    /// 获取或设置是否按类别使用不同颜色
    /// true = 每个类别使用不同颜色，false = 整个系列使用相同颜色
    /// </summary>
    bool VaryByCategories { get; set; }

    /// <summary>
    /// 获取或设置图表分割类型，用于复合图表中主次图表的分割方式
    /// 使用 <see cref="XlChartSplitType"/> 枚举定义分割方式
    /// </summary>
    XlChartSplitType SplitType { get; set; }

    /// <summary>
    /// 获取或设置图表大小所表示的含义（面积或宽度）
    /// 使用 <see cref="XlSizeRepresents"/> 枚举定义大小含义
    /// </summary>
    XlSizeRepresents SizeRepresents { get; set; }

    /// <summary>
    /// 获取或设置图表分割的临界值
    /// 根据 <see cref="SplitType"/> 的不同，表示位置、值或百分比
    /// </summary>
    double SplitValue { get; set; }

    /// <summary>
    /// 获取或设置第二绘图区的大小（5-200，100=默认）
    /// 用于复合饼图或复合条形图中的第二个图表区域
    /// </summary>
    int SecondPlotSize { get; set; }

    /// <summary>
    /// 获取或设置三维图表的着色样式
    /// true = 使用 3D 着色效果，false = 使用平面着色效果
    /// </summary>
    bool Has3DShading { get; set; }

    /// <summary>
    /// 获取图表组的上涨柱（UpBars）格式。
    /// 返回封装后的 <see cref="IExcelUpBars"/> 接口。
    /// </summary>
    IExcelUpBars? UpBars { get; }

    /// <summary>
    /// 获取图表组的下跌柱（DownBars）格式。
    /// 返回封装后的 <see cref="IExcelDownBars"/> 接口。
    /// </summary>
    IExcelDownBars? DownBars { get; }

    /// <summary>
    /// 获取图表组的高低线（HiLoLines）格式。
    /// 返回封装后的 <see cref="IExcelHiLoLines"/> 接口。
    /// </summary>
    IExcelHiLoLines? HiLoLines { get; }

    /// <summary>
    /// 获取图表组的垂线（DropLines）格式。
    /// 返回封装后的 <see cref="IExcelDropLines"/> 接口。
    /// </summary>
    IExcelDropLines? DropLines { get; }
}