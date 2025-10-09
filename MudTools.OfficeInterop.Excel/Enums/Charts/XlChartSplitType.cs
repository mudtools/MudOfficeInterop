namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 指定图表数据系列的分割类型，主要用于饼图和其他需要分割显示的数据可视化图表
/// </summary>
public enum XlChartSplitType
{
    /// <summary>
    /// 按位置分割，根据数据点在序列中的位置进行分割
    /// </summary>
    xlSplitByPosition = 1,

    /// <summary>
    /// 按值分割，根据数据点的数值大小进行分割
    /// </summary>
    xlSplitByValue = 2,

    /// <summary>
    /// 按百分比值分割，根据数据点占总数的百分比进行分割
    /// </summary>
    xlSplitByPercentValue = 3,

    /// <summary>
    /// 按自定义分割，允许用户自定义分割方式
    /// </summary>
    xlSplitByCustomSplit = 4
}