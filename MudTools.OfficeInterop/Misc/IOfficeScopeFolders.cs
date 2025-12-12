//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中搜索范围文件夹集合的接口封装。
/// 该接口提供对搜索范围文件夹集合的访问和管理。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeScopeFolders : IEnumerable<IOfficeScopeFolder>, IDisposable
{
    /// <summary>
    /// 获取搜索范围文件夹集合中项的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取搜索范围文件夹（索引从 1 开始）。
    /// </summary>
    /// <param name="index">搜索范围文件夹索引。</param>
    /// <returns>搜索范围文件夹对象。</returns>
    IOfficeScopeFolder this[int index] { get; }
}