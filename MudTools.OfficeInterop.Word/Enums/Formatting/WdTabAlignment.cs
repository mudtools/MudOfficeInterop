namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在文档中对齐制表符的方式
/// </summary>
public enum WdTabAlignment
{
    /// <summary>
    /// 左对齐制表符
    /// </summary>
    wdAlignTabLeft = 0,
    /// <summary>
    /// 居中对齐制表符
    /// </summary>
    wdAlignTabCenter = 1,
    /// <summary>
    /// 右对齐制表符
    /// </summary>
    wdAlignTabRight = 2,
    /// <summary>
    /// 小数点对齐制表符
    /// </summary>
    wdAlignTabDecimal = 3,
    /// <summary>
    /// 条形对齐制表符
    /// </summary>
    wdAlignTabBar = 4,
    /// <summary>
    /// 列表对齐制表符
    /// </summary>
    wdAlignTabList = 6
}