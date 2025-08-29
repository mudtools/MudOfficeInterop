namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定段落的行距选项
/// </summary>
public enum WdLineSpacing
{
    /// <summary>
    /// 单倍行距
    /// </summary>
    wdLineSpaceSingle,

    /// <summary>
    /// 1.5倍行距
    /// </summary>
    wdLineSpace1pt5,

    /// <summary>
    /// 双倍行距
    /// </summary>
    wdLineSpaceDouble,

    /// <summary>
    /// 最小行距，至少指定的数值
    /// </summary>
    wdLineSpaceAtLeast,

    /// <summary>
    /// 固定行距，精确指定的数值
    /// </summary>
    wdLineSpaceExactly,

    /// <summary>
    /// 多倍行距，指定倍数
    /// </summary>
    wdLineSpaceMultiple
}