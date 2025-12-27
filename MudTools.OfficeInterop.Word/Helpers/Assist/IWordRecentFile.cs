//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示最近使用的文件。
/// <para>注：RecentFile 对象是 RecentFiles 集合的成员。RecentFiles 集合中的各项包含最近使用的所有文件。集合中的各项显示在“文件”菜单的底部。</para>
/// </summary>
public interface IWordRecentFile : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取一个值，该值表示集合中项的位置。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置指定对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取指定对象的磁盘或 Web 路径。
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示最近使用的文件是否以只读方式打开。
    /// </summary>
    bool ReadOnly { get; set; }

    /// <summary>
    /// 删除指定的最近使用文件条目。
    /// </summary>
    void Delete();

    /// <summary>
    /// 打开指定的最近使用文件。
    /// </summary>
    /// <returns>表示打开的文档的对象。</returns>
    IWordDocument Open();
}