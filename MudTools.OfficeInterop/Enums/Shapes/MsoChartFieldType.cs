namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图表中可用的字段类型枚举
/// </summary>
public enum MsoChartFieldType
{
    /// <summary>
    /// 气泡图的气泡大小字段
    /// </summary>
    msoChartFieldBubbleSize = 1,

    /// <summary>
    /// 图表分类名称字段
    /// </summary>
    msoChartFieldCategoryName,

    /// <summary>
    /// 百分比字段
    /// </summary>
    msoChartFieldPercentage,

    /// <summary>
    /// 系列名称字段
    /// </summary>
    msoChartFieldSeriesName,

    /// <summary>
    /// 数值字段
    /// </summary>
    msoChartFieldValue,

    /// <summary>
    /// 公式字段
    /// </summary>
    msoChartFieldFormula,

    /// <summary>
    /// 范围字段
    /// </summary>
    msoChartFieldRange
}