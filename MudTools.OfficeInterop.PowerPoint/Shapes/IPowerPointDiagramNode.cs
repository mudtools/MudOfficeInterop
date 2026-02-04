//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示图示中的一个节点，提供对节点属性和操作的访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointDiagramNode : IOfficeObject<IPowerPointDiagramNode, MsPowerPoint.DiagramNode>, IDisposable
{
    /// <summary>
    /// 获取创建此图示节点的应用程序。
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
    /// 在当前节点附近添加一个新节点。
    /// </summary>
    /// <param name="pos">新节点相对于当前节点的位置。</param>
    /// <param name="nodeType">要添加的节点类型。</param>
    /// <returns>新添加的图示节点。</returns>
    IPowerPointDiagramNode? AddNode([ComNamespace("MsCore")] MsoRelativeNodePosition pos = MsoRelativeNodePosition.msoAfterNode, [ComNamespace("MsCore")] MsoDiagramNodeType nodeType = MsoDiagramNodeType.msoDiagramNode);

    /// <summary>
    /// 删除当前节点及其所有子节点。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将当前节点移动到目标节点附近。
    /// </summary>
    /// <param name="targetNode">目标节点，当前节点将移动到该节点附近。</param>
    /// <param name="pos">当前节点相对于目标节点的位置。</param>
    void MoveNode(IPowerPointDiagramNode targetNode, [ComNamespace("MsCore")] MsoRelativeNodePosition pos);

    /// <summary>
    /// 用当前节点替换目标节点。
    /// </summary>
    /// <param name="targetNode">要被替换的目标节点。</param>
    void ReplaceNode(IPowerPointDiagramNode targetNode);

    /// <summary>
    /// 交换当前节点与目标节点的位置。
    /// </summary>
    /// <param name="targetNode">要与当前节点交换的目标节点。</param>
    /// <param name="swapChildren">指示是否同时交换子节点。</param>
    void SwapNode(IPowerPointDiagramNode targetNode, bool swapChildren = true);

    /// <summary>
    /// 克隆当前节点到目标节点附近。
    /// </summary>
    /// <param name="copyChildren">指示是否同时克隆子节点。</param>
    /// <param name="targetNode">目标节点，克隆的节点将放置在该节点附近。</param>
    /// <param name="pos">克隆节点相对于目标节点的位置。</param>
    /// <returns>新克隆的节点。</returns>
    IPowerPointDiagramNode? CloneNode(bool copyChildren, IPowerPointDiagramNode targetNode, [ComNamespace("MsCore")] MsoRelativeNodePosition pos = MsoRelativeNodePosition.msoAfterNode);

    /// <summary>
    /// 将当前节点的所有子节点转移到接收节点。
    /// </summary>
    /// <param name="receivingNode">接收子节点的目标节点。</param>
    void TransferChildren(IPowerPointDiagramNode receivingNode);

    /// <summary>
    /// 获取当前节点的下一个同级节点。
    /// </summary>
    /// <returns>下一个同级节点，如果没有则返回 null。</returns>
    IPowerPointDiagramNode? NextNode();

    /// <summary>
    /// 获取当前节点的上一个同级节点。
    /// </summary>
    /// <returns>上一个同级节点，如果没有则返回 null。</returns>
    IPowerPointDiagramNode? PrevNode();

    /// <summary>
    /// 获取图示节点的父对象。
    /// </summary>
    /// <value>父对象，通常是图示或另一个节点。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取当前节点的子节点集合。
    /// </summary>
    /// <value>子节点的集合。</value>
    IPowerPointDiagramNodeChildren? Children { get; }

    /// <summary>
    /// 获取与当前节点关联的形状对象。
    /// </summary>
    /// <value>表示节点形状的对象。</value>
    IPowerPointShape? Shape { get; }

    /// <summary>
    /// 获取图示的根节点。
    /// </summary>
    /// <value>图示的根节点。</value>
    IPowerPointDiagramNode? Root { get; }

    /// <summary>
    /// 获取当前节点所属的图示。
    /// </summary>
    /// <value>包含当前节点的图示对象。</value>
    IPowerPointDiagram? Diagram { get; }

    /// <summary>
    /// 获取或设置节点的组织结构图布局类型。
    /// </summary>
    /// <value>节点的布局类型。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoOrgChartLayoutType Layout { get; set; }

    /// <summary>
    /// 获取与当前节点关联的文本形状。
    /// </summary>
    /// <value>包含节点文本的形状对象。</value>
    IPowerPointShape? TextShape { get; }
}