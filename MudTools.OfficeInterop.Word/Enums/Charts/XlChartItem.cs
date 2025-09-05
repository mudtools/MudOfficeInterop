//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定图表中的不同项类型，用于在操作图表时标识特定的图表元素
/// </summary>
public enum XlChartItem
{
    /// <summary>
    /// 数据标签
    /// </summary>
    xlDataLabel = 0,
    /// <summary>
    /// 图表区
    /// </summary>
    xlChartArea = 2,
    /// <summary>
    /// 数据系列
    /// </summary>
    xlSeries = 3,
    /// <summary>
    /// 图表标题
    /// </summary>
    xlChartTitle = 4,
    /// <summary>
    /// 墙面
    /// </summary>
    xlWalls = 5,
    /// <summary>
    /// 角落
    /// </summary>
    xlCorners = 6,
    /// <summary>
    /// 数据表
    /// </summary>
    xlDataTable = 7,
    /// <summary>
    /// 趋势线
    /// </summary>
    xlTrendline = 8,
    /// <summary>
    /// 误差线
    /// </summary>
    xlErrorBars = 9,
    /// <summary>
    /// X轴误差线
    /// </summary>
    xlXErrorBars = 10,
    /// <summary>
    /// Y轴误差线
    /// </summary>
    xlYErrorBars = 11,
    /// <summary>
    /// 图例项
    /// </summary>
    xlLegendEntry = 12,
    /// <summary>
    /// 图例键
    /// </summary>
    xlLegendKey = 13,
    /// <summary>
    /// 形状
    /// </summary>
    xlShape = 14,
    /// <summary>
    /// 主网格线
    /// </summary>
    xlMajorGridlines = 15,
    /// <summary>
    /// 次网格线
    /// </summary>
    xlMinorGridlines = 16,
    /// <summary>
    /// 坐标轴标题
    /// </summary>
    xlAxisTitle = 17,
    /// <summary>
    /// 上涨柱线
    /// </summary>
    xlUpBars = 18,
    /// <summary>
    /// 绘图区
    /// </summary>
    xlPlotArea = 19,
    /// <summary>
    /// 下降柱线
    /// </summary>
    xlDownBars = 20,
    /// <summary>
    /// 坐标轴
    /// </summary>
    xlAxis = 21,
    /// <summary>
    /// 系列线
    /// </summary>
    xlSeriesLines = 22,
    /// <summary>
    /// 底板
    /// </summary>
    xlFloor = 23,
    /// <summary>
    /// 图例
    /// </summary>
    xlLegend = 24,
    /// <summary>
    /// 高低点连线
    /// </summary>
    xlHiLoLines = 25,
    /// <summary>
    /// 垂直线
    /// </summary>
    xlDropLines = 26,
    /// <summary>
    /// 雷达图坐标轴标签
    /// </summary>
    xlRadarAxisLabels = 27,
    /// <summary>
    /// 无
    /// </summary>
    xlNothing = 28,
    /// <summary>
    /// 引导线
    /// </summary>
    xlLeaderLines = 29,
    /// <summary>
    /// 显示单位标签
    /// </summary>
    xlDisplayUnitLabel = 30,
    /// <summary>
    /// 数据透视图字段按钮
    /// </summary>
    xlPivotChartFieldButton = 31,
    /// <summary>
    /// 数据透视图拖放区域
    /// </summary>
    xlPivotChartDropZone = 32
}