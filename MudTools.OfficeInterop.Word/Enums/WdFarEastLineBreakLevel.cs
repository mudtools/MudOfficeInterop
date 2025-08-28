namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定东亚语言文本的换行级别控制选项
/// </summary>
public enum WdFarEastLineBreakLevel
{
    /// <summary>
    /// 普通换行级别控制
    /// </summary>
    wdFarEastLineBreakLevelNormal,

    /// <summary>
    /// 严格换行级别控制
    /// </summary>
    wdFarEastLineBreakLevelStrict,

    /// <summary>
    /// 自定义换行级别控制
    /// </summary>
    wdFarEastLineBreakLevelCustom
}