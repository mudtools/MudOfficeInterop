

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定剪贴板数据格式的枚举
/// </summary>
public enum MsoClipboardFormat
{
    /// <summary>
    /// 混合格式（用于表示多种格式混合的情况）
    /// </summary>
    msoClipboardFormatMixed = -2,
    /// <summary>
    /// 本机格式（特定应用程序的原生数据格式）
    /// </summary>
    msoClipboardFormatNative = 1,
    /// <summary>
    /// HTML格式
    /// </summary>
    msoClipboardFormatHTML = 2,
    /// <summary>
    /// RTF格式（富文本格式）
    /// </summary>
    msoClipboardFormatRTF = 3,
    /// <summary>
    /// 纯文本格式
    /// </summary>
    msoClipboardFormatPlainText = 4
}