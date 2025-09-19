
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示自由形状（Freeform）路径中的一个节点，包含位置、类型和控制点信息。
/// </summary>
public interface IExcelShapeNode : IDisposable
{
    /// <summary>
    /// 获取此节点所属的父对象（通常是 ShapeNodes 集合）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此节点所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取节点类型（角点或曲线点）。
    /// </summary>
    MsoEditingType EditingType { get; }

    /// <summary>
    /// 获取节点在工作表中的 X 坐标（单位：磅，points）。
    /// </summary>
    float PointsX { get; }

    /// <summary>
    /// 获取节点在工作表中的 Y 坐标（单位：磅，points）。
    /// </summary>
    float PointsY { get; }

    /// <summary>
    /// 获取第一个控制点相对于节点的 X 偏移量（仅对曲线节点有效）。
    /// </summary>
    float? ControlPoint1X { get; }

    /// <summary>
    /// 获取第一个控制点相对于节点的 Y 偏移量（仅对曲线节点有效）。
    /// </summary>
    float? ControlPoint1Y { get; }

    /// <summary>
    /// 获取第二个控制点相对于节点的 X 偏移量（仅对曲线节点有效）。
    /// </summary>
    float? ControlPoint2X { get; }

    /// <summary>
    /// 获取第二个控制点相对于节点的 Y 偏移量（仅对曲线节点有效）。
    /// </summary>
    float? ControlPoint2Y { get; }
}