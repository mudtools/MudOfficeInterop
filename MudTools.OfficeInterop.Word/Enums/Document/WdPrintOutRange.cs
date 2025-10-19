namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定打印文档时要打印的范围选项
/// </summary>
public enum WdPrintOutRange
{
    /// <summary>
    /// 打印整个文档
    /// </summary>
    wdPrintAllDocument,

    /// <summary>
    /// 打印当前选中的内容
    /// </summary>
    wdPrintSelection,

    /// <summary>
    /// 打印光标所在的当前页面
    /// </summary>
    wdPrintCurrentPage,

    /// <summary>
    /// 打印从某一页到另一页的范围
    /// </summary>
    wdPrintFromTo,

    /// <summary>
    /// 打印指定页面范围
    /// </summary>
    wdPrintRangeOfPages
}