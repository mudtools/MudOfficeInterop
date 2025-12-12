
namespace MudTools.OfficeInterop;

/// <summary>
/// 指定文本大小写转换的类型
/// </summary>
public enum MsoTextChangeCase
{
    /// <summary>
    /// 句子大小写 - 只有第一个字母大写，专有名词除外
    /// </summary>
    msoCaseSentence = 1,

    /// <summary>
    /// 全小写 - 所有字母都小写
    /// </summary>
    msoCaseLower,

    /// <summary>
    /// 全大写 - 所有字母都大写
    /// </summary>
    msoCaseUpper,

    /// <summary>
    /// 首字母大写 - 每个单词的首字母大写
    /// </summary>
    msoCaseTitle,

    /// <summary>
    /// 切换大小写 - 在大写和小写之间切换
    /// </summary>
    msoCaseToggle
}