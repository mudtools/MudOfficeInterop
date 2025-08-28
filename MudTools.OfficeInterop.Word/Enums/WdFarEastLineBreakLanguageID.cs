namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定远东语言的换行规则ID，用于控制不同语言环境下的文本换行行为
/// </summary>
public enum WdFarEastLineBreakLanguageID
{
    /// <summary>
    /// 日文换行规则，对应LCID 1041
    /// </summary>
    wdLineBreakJapanese = 1041,
    
    /// <summary>
    /// 韩文换行规则，对应LCID 1042
    /// </summary>
    wdLineBreakKorean = 1042,
    
    /// <summary>
    /// 简体中文换行规则，对应LCID 2052
    /// </summary>
    wdLineBreakSimplifiedChinese = 2052,
    
    /// <summary>
    /// 繁体中文换行规则，对应LCID 1028
    /// </summary>
    wdLineBreakTraditionalChinese = 1028
}