//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示最近访问的文件的集合。
/// <para>注：使用 RecentFiles 属性可返回 RecentFiles 集合。</para>
/// <para>注：使用 RecentFiles(index)（其中 index 是索引号）可返回单个 RecentFile 对象。索引序号代表该文件在“文件”菜单上的位置。</para>
/// </summary>
public interface IWordRecentFiles : IEnumerable<IWordRecentFile>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的最近使用文件数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引号获取单个最近使用文件。
    /// </summary>
    /// <param name="index">索引号（从 1 开始）。</param>
    /// <returns>指定的最近使用文件对象。</returns>
    IWordRecentFile this[int index] { get; }

    /// <summary>
    /// 将文件添加到最近使用的文件列表中。
    /// </summary>
    /// <param name="fileName">要添加的文件的名称。</param>
    /// <param name="readOnly">如果为 true，则以只读方式打开文件。</param>
    /// <returns>表示添加的最近使用文件的对象。</returns>
    IWordRecentFile Add(string fileName, object readOnly);
}