namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office图表节点子节点集合的接口，继承自IEnumerable&lt;IOfficeDiagramNode&gt;和IDisposable接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
[ItemIndex]
public interface IOfficeDiagramNodeChildren : IEnumerable<IOfficeDiagramNode>, IDisposable
{
    /// <summary>
    /// 获取集合中子节点的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取子节点
    /// </summary>
    /// <param name="index">子节点的索引</param>
    /// <returns>指定索引的子节点</returns>
    IOfficeDiagramNode this[int index] { get; }

    /// <summary>
    /// 获取第一个子节点
    /// </summary>
    IOfficeDiagramNode FirstChild { get; }

    /// <summary>
    /// 获取最后一个子节点
    /// </summary>
    IOfficeDiagramNode LastChild { get; }

    /// <summary>
    /// 向集合中添加一个新的节点
    /// </summary>
    /// <param name="Index">新节点插入的位置索引，默认为-1表示添加到末尾</param>
    /// <param name="NodeType">节点类型，默认为msoDiagramNode主节点类型</param>
    /// <returns>新添加的节点</returns>
    IOfficeDiagramNode AddNode(int Index = -1, MsoDiagramNodeType NodeType = MsoDiagramNodeType.msoDiagramNode);

    /// <summary>
    /// 选择集合中的所有节点
    /// </summary>
    void SelectAll();
}