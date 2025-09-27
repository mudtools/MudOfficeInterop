namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定页面分隔符的范围类型
/// </summary>
public enum XlPageBreakExtent
{
    /// <summary>
    /// 完整的页面分隔符，应用于整页
    /// </summary>
    xlPageBreakFull = 1,

    /// <summary>
    /// 部分页面分隔符，仅应用于部分内容
    /// </summary>
    xlPageBreakPartial
}