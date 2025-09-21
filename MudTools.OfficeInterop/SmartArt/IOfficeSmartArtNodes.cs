
namespace MudTools.OfficeInterop;

/// <summary>
/// SmartArtNodes 集合封装接口
/// </summary>
public interface IOfficeSmartArtNodes : IEnumerable<IOfficeSmartArtNode>, IDisposable
{
    /// <summary>
    /// 获取集合中节点总数
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取节点（索引从1开始，兼容COM习惯）
    /// </summary>
    /// <param name="index">节点索引（1起始）</param>
    /// <returns>对应节点或 null</returns>
    IOfficeSmartArtNode? this[int index] { get; }

    /// <summary>
    /// 在集合末尾添加一个新节点
    /// </summary>
    /// <param name="text">节点文本</param>
    /// <returns>新创建的节点</returns>
    IOfficeSmartArtNode? Add(string text);

    /// <summary>
    /// 清空所有节点
    /// </summary>
    void Clear();
}