namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定打印文档时要包含的页面类型
/// </summary>
public enum WdPrintOutPages
{

    /// <summary>
    /// 打印所有页面
    /// </summary>
    wdPrintAllPages,

    /// <summary>
    /// 仅打印奇数页
    /// </summary>
    wdPrintOddPagesOnly,

    /// <summary>
    /// 仅打印偶数页
    /// </summary>
    wdPrintEvenPagesOnly
}