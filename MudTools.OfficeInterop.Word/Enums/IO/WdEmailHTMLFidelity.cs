
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定是否保留或移除显示不需要的 HTML 标签
/// </summary>
public enum WdEmailHTMLFidelity
{
    /// <summary>
    /// 移除所有不影响消息显示的 HTML 标签
    /// </summary>
    wdEmailHTMLFidelityLow = 1,

    /// <summary>
    /// 不支持
    /// </summary>
    wdEmailHTMLFidelityMedium,

    /// <summary>
    /// 保持 HTML 完整
    /// </summary>
    wdEmailHTMLFidelityHigh
}