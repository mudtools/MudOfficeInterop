namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定要从链接返回的信息类型
/// </summary>
public enum XlLinkInfo
{
    /// <summary>
    /// 链接的更新状态（1表示自动更新，2表示手动更新）
    /// </summary>
    xlUpdateState = 1,

    /// <summary>
    /// 链接的版本日期
    /// </summary>
    xlEditionDate = 2,

    /// <summary>
    /// 链接的状态信息
    /// </summary>
    xlLinkInfoStatus = 3
}