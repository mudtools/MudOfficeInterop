namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 SmartArt 图表节点集合的接口
/// 继承自 IEnumerable<IOfficeDiagramNode> 和 IDisposable，支持遍历和资源释放
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
[ItemIndex]
public interface IOfficeDiagramNodes : IEnumerable<IOfficeDiagramNode>, IDisposable
{
    /// <summary>
    /// 选择集合中的所有图表节点
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 获取集合中图表节点的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的特定图表节点
    /// </summary>
    /// <param name="index">要获取的节点在集合中的索引位置</param>
    /// <returns>指定索引位置的图表节点</returns>
    IOfficeDiagramNode this[int index] { get; }
}