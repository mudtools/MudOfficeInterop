namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定在图表中显示误差线的方式
/// </summary>
public enum XlErrorBarInclude
{

    /// <summary>
    /// 包括正负两个方向的误差线
    /// </summary>
    xlErrorBarIncludeBoth = 1,

    /// <summary>
    /// 只包括负值方向的误差线
    /// </summary>
    xlErrorBarIncludeMinusValues = 3,

    /// <summary>
    /// 不包括误差线
    /// </summary>
    xlErrorBarIncludeNone = -4142,

    /// <summary>
    /// 只包括正值方向的误差线
    /// </summary>
    xlErrorBarIncludePlusValues = 2
}