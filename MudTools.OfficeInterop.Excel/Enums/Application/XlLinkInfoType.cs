namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定要返回信息的链接类型
/// </summary>
public enum XlLinkInfoType
{
    /// <summary>
    /// OLE 链接类型
    /// </summary>
    xlLinkInfoOLELinks = 2,
    
    /// <summary>
    /// 发布者链接类型（仅限 Macintosh）
    /// </summary>
    xlLinkInfoPublishers = 5,
    
    /// <summary>
    /// 订阅者链接类型（仅限 Macintosh）
    /// </summary>
    xlLinkInfoSubscribers = 6
}