namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定一行是文本行还是表格行
/// </summary>
public enum WdLineType
{
    /// <summary>
    /// 文档正文中的文本行
    /// </summary>
    wdTextLine,

    /// <summary>
    /// 表格行
    /// </summary>
    wdTableRow
}