//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

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