namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定表格中应用条件格式的区域
/// </summary>
public enum WdConditionCode
{
    /// <summary>
    /// 第一行
    /// </summary>
    wdFirstRow,
    /// <summary>
    /// 最后一行
    /// </summary>
    wdLastRow,
    /// <summary>
    /// 奇数行带纹
    /// </summary>
    wdOddRowBanding,
    /// <summary>
    /// 偶数行带纹
    /// </summary>
    wdEvenRowBanding,
    /// <summary>
    /// 第一列
    /// </summary>
    wdFirstColumn,
    /// <summary>
    /// 最后一列
    /// </summary>
    wdLastColumn,
    /// <summary>
    /// 奇数列带纹
    /// </summary>
    wdOddColumnBanding,
    /// <summary>
    /// 偶数列带纹
    /// </summary>
    wdEvenColumnBanding,
    /// <summary>
    /// 右上角单元格
    /// </summary>
    wdNECell,
    /// <summary>
    /// 左上角单元格
    /// </summary>
    wdNWCell,
    /// <summary>
    /// 右下角单元格
    /// </summary>
    wdSECell,
    /// <summary>
    /// 左下角单元格
    /// </summary>
    wdSWCell
}