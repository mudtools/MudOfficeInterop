//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 SmartArt 图表节点集合的接口
/// 继承自 IEnumerable[IOfficeDiagramNode] 和 IDisposable，支持遍历和资源释放
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeDiagramNodes : IEnumerable<IOfficeDiagramNode?>, IDisposable
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
    IOfficeDiagramNode? this[int index] { get; }

    /// <summary>
    /// 通过索引获取集合中的特定图表节点
    /// </summary>
    /// <param name="index">要获取的节点在集合中的索引位置</param>
    /// <returns>指定索引位置的图表节点</returns>
    IOfficeDiagramNode? this[string index] { get; }
}