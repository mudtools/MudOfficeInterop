namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定文档元素的边框类型
/// </summary>
public enum WdBorderType
{
    /// <summary>
    /// 顶部边框
    /// </summary>
    wdBorderTop = -1,

    /// <summary>
    /// 左侧边框
    /// </summary>
    wdBorderLeft = -2,

    /// <summary>
    /// 底部边框
    /// </summary>
    wdBorderBottom = -3,

    /// <summary>
    /// 右侧边框
    /// </summary>
    wdBorderRight = -4,

    /// <summary>
    /// 水平边框
    /// </summary>
    wdBorderHorizontal = -5,

    /// <summary>
    /// 垂直边框
    /// </summary>
    wdBorderVertical = -6,

    /// <summary>
    /// 向下对角线边框
    /// </summary>
    wdBorderDiagonalDown = -7,

    /// <summary>
    /// 向上对角线边框
    /// </summary>
    wdBorderDiagonalUp = -8
}