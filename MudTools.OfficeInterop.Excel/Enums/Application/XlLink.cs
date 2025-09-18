namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定链接的类型
/// </summary>
public enum XlLink
{
    /// <summary>
    /// 到 Excel 工作表的链接
    /// </summary>
    xlExcelLinks = 1,

    /// <summary>
    /// 到 OLE 源的链接
    /// </summary>
    xlOLELinks = 2,

    /// <summary>
    /// 仅限 Macintosh 的发布者链接
    /// </summary>
    xlPublishers = 5,

    /// <summary>
    /// 仅限 Macintosh 的订阅者链接
    /// </summary>
    xlSubscribers = 6
}