namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中图片或形状的文字环绕方式
/// </summary>
public enum WdWrapType
{
    /// <summary>
    /// 四周型环绕
    /// </summary>
    wdWrapSquare = 0,

    /// <summary>
    /// 紧密型环绕
    /// </summary>
    wdWrapTight = 1,

    /// <summary>
    /// 穿越型环绕
    /// </summary>
    wdWrapThrough = 2,

    /// <summary>
    /// 无环绕（上下型环绕）
    /// </summary>
    wdWrapNone = 3,

    /// <summary>
    /// 上下型环绕
    /// </summary>
    wdWrapTopBottom = 4,

    /// <summary>
    /// 衬于文字下方
    /// </summary>
    wdWrapBehind = 5,

    /// <summary>
    /// 浮于文字上方
    /// </summary>
    wdWrapFront = 3,

    /// <summary>
    /// 嵌入型
    /// </summary>
    wdWrapInline = 7
}