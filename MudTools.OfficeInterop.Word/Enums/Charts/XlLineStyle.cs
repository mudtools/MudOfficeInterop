namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定线条样式，用于定义Excel中单元格边框或其他线条的外观样式
/// </summary>
public enum XlLineStyle
{
    /// <summary>
    /// 实线样式
    /// </summary>
    xlContinuous = 1,

    /// <summary>
    /// 虚线样式
    /// </summary>
    xlDash = -4115,

    /// <summary>
    /// 点划线样式（一点一划线）
    /// </summary>
    xlDashDot = 4,

    /// <summary>
    /// 双点划线样式（两点一划线）
    /// </summary>
    xlDashDotDot = 5,

    /// <summary>
    /// 点线样式
    /// </summary>
    xlDot = -4118,

    /// <summary>
    /// 双线样式
    /// </summary>
    xlDouble = -4119,

    /// <summary>
    /// 倾斜虚线点划线样式
    /// </summary>
    xlSlantDashDot = 13,

    /// <summary>
    /// 无线条样式（无边框）
    /// </summary>
    xlLineStyleNone = -4142
}