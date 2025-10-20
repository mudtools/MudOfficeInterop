//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示文档中冲突的集合。
/// <para>注：使用 <see cref="AcceptAll"/> 方法接受所有冲突并将其合并到文档中，或使用 <see cref="RejectAll"/> 方法拒绝所有冲突更改。</para>
/// </summary>
public interface IWordConflicts : IEnumerable<IWordConflict>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的冲突数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取单个冲突。
    /// </summary>
    /// <param name="index">索引（从 1 开始）。</param>
    /// <returns>指定的冲突对象。</returns>
    IWordConflict? this[int index] { get; }

    /// <summary>
    /// 接受所有冲突更改，删除冲突，并将更改合并到文档的服务器副本中。
    /// </summary>
    void AcceptAll();

    /// <summary>
    /// 拒绝所有用户的更改，并保留文档的服务器副本。
    /// </summary>
    void RejectAll();
}