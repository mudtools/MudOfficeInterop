namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定图表类型，用于Microsoft Office Interop操作
/// </summary>
public enum MsoDiagramType
{
    /// <summary>
    /// 混合图表类型
    /// </summary>
    msoDiagramMixed = -2,

    /// <summary>
    /// 组织结构图类型
    /// </summary>
    msoDiagramOrgChart = 1,

    /// <summary>
    /// 循环关系图类型，显示连续循环步骤的流程图
    /// </summary>
    msoDiagramCycle = 2,

    /// <summary>
    /// 径向关系图类型，显示与核心元素的关系
    /// </summary>
    msoDiagramRadial = 3,

    /// <summary>
    /// 金字塔关系图类型，基于基础的关系图
    /// </summary>
    msoDiagramPyramid = 4,

    /// <summary>
    /// 维恩关系图类型，显示元素之间的重叠区域
    /// </summary>
    msoDiagramVenn = 5,

    /// <summary>
    /// 目标关系图类型，显示实现目标的步骤
    /// </summary>
    msoDiagramTarget = 6
}