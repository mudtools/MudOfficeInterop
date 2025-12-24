//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示在文件搜索操作中找到的文件的路径名集合。
/// <para>注：FoundFiles 对象是 FileSearch 对象的成员。</para>
/// <para>注：此接口基于对 Office FileSearch 对象模型的理解实现，因为 Word 特定的 FoundFiles SDK 文档有限。</para>
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeFoundFiles : IEnumerable<string>, IDisposable
{
    #region 基本属性 (Basic Properties)
    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取集合中的文件路径名数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过从 1 开始的索引号获取单个找到的文件的路径名。
    /// </summary>
    /// <param name="index">文件的索引号（从 1 开始）。</param>
    /// <returns>指定文件的完整路径名。</returns>
    string? this[int index] { get; }

    #endregion
}