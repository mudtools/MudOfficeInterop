//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;


/// <summary>
/// 表示 Office 图表中的一个节点，提供对图表节点的各种操作方法
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeDiagramNode : IOfficeObject<IOfficeDiagramNode, MsCore.DiagramNode>, IDisposable
{

    /// <summary>
    /// 获取与此图表节点关联的形状对象
    /// </summary>
    IOfficeShape? Shape { get; }

    /// <summary>
    /// 获取图表节点的根节点
    /// </summary>
    IOfficeDiagramNode? Root { get; }

    /// <summary>
    /// 获取图表节点的子节点集合
    /// </summary>
    IOfficeDiagramNodeChildren? Children { get; }

    /// <summary>
    /// 获取此节点所属的图表对象
    /// </summary>
    IOfficeDiagram? Diagram { get; }

    /// <summary>
    /// 获取或设置组织结构图的布局类型
    /// </summary>
    MsoOrgChartLayoutType Layout { get; set; }

    /// <summary>
    /// 获取与此图表节点关联的文本形状对象
    /// </summary>
    IOfficeShape? TextShape { get; }

    /// <summary>
    /// 在当前节点的指定位置添加一个新节点
    /// </summary>
    /// <param name="Pos">新节点相对于当前节点的位置，默认为在当前节点之后</param>
    /// <param name="NodeType">要添加的节点类型，默认为主节点</param>
    /// <returns>新添加的图表节点</returns>
    IOfficeDiagramNode? AddNode(MsoRelativeNodePosition Pos = MsoRelativeNodePosition.msoAfterNode,
                        MsoDiagramNodeType NodeType = MsoDiagramNodeType.msoDiagramNode);

    /// <summary>
    /// 删除当前节点
    /// </summary>
    void Delete();

    /// <summary>
    /// 将当前节点移动到目标节点的指定位置
    /// </summary>
    /// <param name="TargetNode">目标节点</param>
    /// <param name="Pos">相对于目标节点的位置</param>
    void MoveNode(IOfficeDiagramNode TargetNode, MsoRelativeNodePosition Pos);

    /// <summary>
    /// 用当前节点替换目标节点
    /// </summary>
    /// <param name="TargetNode">要被替换的目标节点</param>
    void ReplaceNode(IOfficeDiagramNode TargetNode);

    /// <summary>
    /// 与目标节点交换位置
    /// </summary>
    /// <param name="TargetNode">要交换的目标节点</param>
    /// <param name="SwapChildren">是否同时交换子节点，默认为true</param>
    void SwapNode(IOfficeDiagramNode TargetNode, bool SwapChildren = true);

    /// <summary>
    /// 克隆当前节点到目标节点的指定位置
    /// </summary>
    /// <param name="CopyChildren">是否复制子节点</param>
    /// <param name="TargetNode">目标节点</param>
    /// <param name="Pos">相对于目标节点的位置，默认为在目标节点之后</param>
    /// <returns>克隆的新节点</returns>
    IOfficeDiagramNode? CloneNode(bool CopyChildren, IOfficeDiagramNode TargetNode, MsoRelativeNodePosition Pos = MsoRelativeNodePosition.msoAfterNode);

    /// <summary>
    /// 将当前节点的子节点转移到接收节点
    /// </summary>
    /// <param name="ReceivingNode">接收子节点的节点</param>
    void TransferChildren(IOfficeDiagramNode ReceivingNode);

    /// <summary>
    /// 获取下一个相邻节点
    /// </summary>
    /// <returns>下一个节点，如果没有则返回null</returns>
    IOfficeDiagramNode? NextNode();

    /// <summary>
    /// 获取上一个相邻节点
    /// </summary>
    /// <returns>上一个节点，如果没有则返回null</returns>
    IOfficeDiagramNode? PrevNode();
}