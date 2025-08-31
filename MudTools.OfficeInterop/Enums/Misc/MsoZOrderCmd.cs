namespace MudTools.OfficeInterop;

/// <summary>
/// 指定如何重新排列对象的堆叠顺序（Z顺序）的命令
/// </summary>
public enum MsoZOrderCmd
{
    /// <summary>
    /// 将对象移到堆叠顺序的最前面
    /// </summary>
    msoBringToFront,
    /// <summary>
    /// 将对象移到堆叠顺序的最后面
    /// </summary>
    msoSendToBack,
    /// <summary>
    /// 将对象向前移动一个位置
    /// </summary>
    msoBringForward,
    /// <summary>
    /// 将对象向后移动一个位置
    /// </summary>
    msoSendBackward,
    /// <summary>
    /// 将对象放在文本前面
    /// </summary>
    msoBringInFrontOfText,
    /// <summary>
    /// 将对象放在文本后面
    /// </summary>
    msoSendBehindText
}