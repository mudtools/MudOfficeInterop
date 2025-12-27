namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定在共享工作簿中显示的更改集
/// </summary>
public enum XlHighlightChangesTime
{
    /// <summary>
    /// 显示自上次用户保存以来所做的更改
    /// </summary>
    xlSinceMyLastSave = 1,

    /// <summary>
    /// 显示所有更改
    /// </summary>
    xlAllChanges,

    /// <summary>
    /// 仅显示尚未审阅的更改
    /// </summary>
    xlNotYetReviewed
}