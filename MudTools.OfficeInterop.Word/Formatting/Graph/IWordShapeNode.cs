//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示用户自定义自由曲线中节点的几何形状和几何编辑属性。
/// 此接口提供对曲线节点信息的访问，包括位置、编辑类型和线段类型。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordShapeNode : IOfficeObject<IWordShapeNode, MsWord.ShapeNode>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取指定节点的编辑类型。
    /// 如果指定节点是顶点，则返回指示对该节点的更改如何影响连接到该节点的两个线段的值。
    /// 如果节点是曲线段的控制点，则返回相邻顶点的编辑类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEditingType EditingType { get; }

    /// <summary>
    /// 获取指定节点的位置坐标对。
    /// 每个坐标以磅为单位表示。
    /// </summary>
    object Points { get; }

    /// <summary>
    /// 获取一个值，指示与指定节点关联的线段是直线还是曲线。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoSegmentType SegmentType { get; }
}