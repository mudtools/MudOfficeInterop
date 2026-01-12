namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定形状在文档中的位置
/// </summary>
public enum WdShapePosition
{
    /// <summary>
    /// 形状位于顶部
    /// </summary>
    wdShapeTop = -999999,
    /// <summary>
    /// 形状位于左侧
    /// </summary>
    wdShapeLeft,
    /// <summary>
    /// 形状位于底部
    /// </summary>
    wdShapeBottom,
    /// <summary>
    /// 形状位于右侧
    /// </summary>
    wdShapeRight,
    /// <summary>
    /// 形状居中
    /// </summary>
    wdShapeCenter,
    /// <summary>
    /// 形状位于内部
    /// </summary>
    wdShapeInside,
    /// <summary>
    /// 形状位于外部
    /// </summary>
    wdShapeOutside
}