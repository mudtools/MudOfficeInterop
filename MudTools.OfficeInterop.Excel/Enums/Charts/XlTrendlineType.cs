namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定图表中趋势线的类型
/// </summary>
public enum XlTrendlineType
{

    /// <summary>
    /// 指数趋势线
    /// </summary>
    xlExponential = 5,

    /// <summary>
    /// 线性趋势线
    /// </summary>
    xlLinear = -4132,

    /// <summary>
    /// 对数趋势线
    /// </summary>
    xlLogarithmic = -4133,

    /// <summary>
    /// 移动平均趋势线
    /// </summary>
    xlMovingAvg = 6,

    /// <summary>
    /// 多项式趋势线
    /// </summary>
    xlPolynomial = 3,

    /// <summary>
    /// 幂函数趋势线
    /// </summary>
    xlPower = 4
}