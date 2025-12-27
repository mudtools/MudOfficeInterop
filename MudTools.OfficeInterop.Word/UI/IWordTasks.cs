//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示当前在系统上运行的所有应用程序任务的集合。
/// <para>注：使用 Tasks 属性可返回 Tasks 集合。</para>
/// <para>注：使用 Tasks(index) 可返回单个 Task 对象，其中 index 是应用程序名称或索引号。</para>
/// </summary>
public interface IWordTasks : IEnumerable<IWordTask>, IDisposable
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
    /// 获取集合中的任务数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过应用程序名称或索引号获取单个任务。
    /// </summary>
    /// <param name="index">应用程序名称（字符串）或索引号（整数）。</param>
    /// <returns>指定的任务对象。</returns>
    IWordTask this[object index] { get; }

    /// <summary>
    /// 确定指定的任务是否存在。
    /// </summary>
    /// <param name="name">要检查的任务名称。</param>
    /// <returns>如果任务存在则返回 true，否则返回 false。</returns>
    bool Exists(string name);

    /// <summary>
    /// 关闭所有打开的应用程序，退出 Microsoft Windows，注销当前用户。
    /// </summary>
    void ExitWindows();
}