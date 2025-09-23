
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定Excel图表元素的位置设置
/// </summary>
public enum XlChartElementPosition
{
    /// <summary>
    /// 自动位置 - 图表元素由Excel自动定位
    /// </summary>
    xlChartElementPositionAutomatic = -4105,
    
    /// <summary>
    /// 自定义位置 - 图表元素使用用户定义的位置
    /// </summary>
    xlChartElementPositionCustom = -4114
}