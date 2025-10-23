namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定误差线的类型
/// </summary>
public enum XlErrorBarType
{

    /// <summary>
    /// 自定义误差线类型
    /// </summary>
    xlErrorBarTypeCustom = -4114,

    /// <summary>
    /// 固定值误差线类型
    /// </summary>
    xlErrorBarTypeFixedValue = 1,

    /// <summary>
    /// 百分比误差线类型
    /// </summary>
    xlErrorBarTypePercent = 2,

    /// <summary>
    /// 标准偏差误差线类型
    /// </summary>
    xlErrorBarTypeStDev = -4155,

    /// <summary>
    /// 标准误差线类型
    /// </summary>
    xlErrorBarTypeStError = 4
}