namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定当文本垂直溢出图表对象边界时的处理方式
/// </summary>
public enum XlOartVerticalOverflow
{
    /// <summary>
    /// 允许文本溢出对象边界显示
    /// </summary>
    xlOartVerticalOverflowOverflow,
    /// <summary>
    /// 裁剪溢出的文本，不显示超出边界的部分
    /// </summary>
    xlOartVerticalOverflowClip,
    /// <summary>
    /// 用省略号(...)表示被截断的文本
    /// </summary>
    xlOartVerticalOverflowEllipsis
}