//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// ProtectedViewWindows 接口及实现类
/// </summary>
public interface IWordProtectedViewWindows : IEnumerable<IWordProtectedViewWindow>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取集合中受保护的视图窗口的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 返回集合中指定的 <see cref="IWordProtectedViewWindow"/> 对象。
    /// </summary>
    /// <param name="index">要返回的单个对象。可以是代表序号位置的 Number 类型的值。</param>
    /// <returns>指定索引处的 <see cref="IWordProtectedViewWindow"/> 对象。</returns>
    IWordProtectedViewWindow? this[int index] { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 打开一个文档并在新的受保护的视图窗口中显示。
    /// </summary>
    /// <param name="FileName">要打开的文档的名称。</param>
    /// <param name="AddToRecentFiles">如果为 True，则将该文件添加到“文件”菜单上的最近使用的文件列表中。</param>
    /// <param name="PasswordDocument">打开文档所需的密码。</param>
    /// <param name="Visible">如果为 True，则在可见的受保护视图窗口中打开文档。</param>
    /// <param name="OpenAndRepair">如果为 True，则修复并打开文档。</param>
    /// <returns>返回新创建的 <see cref="IWordProtectedViewWindow"/> 对象。</returns>
    IWordProtectedViewWindow Open(string FileName, ref object AddToRecentFiles, ref object PasswordDocument, ref object Visible, ref object OpenAndRepair);
}