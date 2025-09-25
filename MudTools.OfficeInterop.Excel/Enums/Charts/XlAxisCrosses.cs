namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定坐标轴与其他坐标轴相交的位置
/// </summary>
public enum XlAxisCrosses
{
    /// <summary>
    /// Excel自动设置坐标轴交点
    /// </summary>
    xlAxisCrossesAutomatic = -4105,
    
    /// <summary>
    /// 使用自定义值设置坐标轴交点（通过CrossesAt属性指定具体交点值）
    /// </summary>
    xlAxisCrossesCustom = -4114,
    
    /// <summary>
    /// 坐标轴在最大值处相交
    /// </summary>
    xlAxisCrossesMaximum = 2,
    
    /// <summary>
    /// 坐标轴在最小值处相交
    /// </summary>
    xlAxisCrossesMinimum = 4
}