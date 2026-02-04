//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示形状中的一个节点，该节点定义了自由形状的几何结构。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointShapeNode : IOfficeObject<IPowerPointShapeNode, MsPowerPoint.ShapeNode>, IDisposable
{
    /// <summary>
    /// 获取创建此形状节点的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取形状节点的父对象。
    /// </summary>
    /// <value>父对象，通常是形状节点集合。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取节点的编辑类型。
    /// </summary>
    /// <value>节点的编辑类型，指示节点是否可以编辑。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEditingType EditingType { get; }

    /// <summary>
    /// 获取节点的坐标点。
    /// </summary>
    /// <value>包含节点坐标的数组，格式取决于节点类型。</value>
    object? Points { get; }

    /// <summary>
    /// 获取节点的线段类型。
    /// </summary>
    /// <value>节点的线段类型，指示节点之间的连接方式。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoSegmentType SegmentType { get; }
}