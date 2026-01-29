namespace MudTools.OfficeInterop;

/// <summary>
/// 指定超链接的类型。
/// </summary>
public enum MsoHyperlinkType
{
    /// <summary>
    /// 应用于文本范围（Range 对象）的超链接。
    /// </summary>
    msoHyperlinkRange,

    /// <summary>
    /// 应用于形状（Shape 对象）的超链接。
    /// </summary>
    msoHyperlinkShape,

    /// <summary>
    /// 应用于内嵌形状的超链接（仅适用于 Microsoft Word）。
    /// </summary>
    msoHyperlinkInlineShape
}