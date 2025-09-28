namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在Word文档中执行查找和替换操作时的替换选项
/// </summary>
public enum WdReplace
{
    /// <summary>
    /// 不执行替换操作，仅查找
    /// </summary>
    wdReplaceNone,
    /// <summary>
    /// 只替换第一个匹配项
    /// </summary>
    wdReplaceOne,
    /// <summary>
    /// 替换所有匹配项
    /// </summary>
    wdReplaceAll
}