//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// SmartArtNodes 集合封装接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
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
    /// <returns>新创建的节点</returns>
    IOfficeSmartArtNode? Add();
}