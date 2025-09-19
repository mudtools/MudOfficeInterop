namespace MudTools.OfficeInterop;

/// <summary>
/// 指定形状节点的编辑类型
/// </summary>
public enum MsoEditingType
{
    /// <summary>
    /// 自动编辑类型，由应用程序自动决定节点的控制点位置
    /// </summary>
    msoEditingAuto,
    /// <summary>
    /// 角点编辑类型，节点的进入和离开线段形成尖角
    /// </summary>
    msoEditingCorner,
    /// <summary>
    /// 平滑编辑类型，节点的进入和离开线段形成平滑曲线
    /// </summary>
    msoEditingSmooth,
    /// <summary>
    /// 对称编辑类型，节点两侧的控制点保持对称，形成均匀的曲线
    /// </summary>
    msoEditingSymmetric
}