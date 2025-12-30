//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// 表示VBA编辑器中一组链接的窗口集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsVb"), ItemIndex]
public interface IVbeLinkedWindows : IEnumerable<IVbeWindow?>, IOfficeObject<IVbeLinkedWindows>, IDisposable
{
    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取集合中窗口的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取集合中的窗口
    /// </summary>
    /// <param name="index">窗口在集合中的索引位置，从0开始</param>
    /// <returns>指定索引位置的窗口对象，如果不存在则返回null</returns>
    IVbeWindow? this[int index] { get; }

    /// <summary>
    /// 根据名称获取集合中的窗口
    /// </summary>
    /// <param name="name">窗口的名称</param>
    /// <returns>指定名称的窗口对象，如果不存在则返回null</returns>
    IVbeWindow? this[string name] { get; }

    /// <summary>
    /// 从集合中移除指定的窗口
    /// </summary>
    /// <param name="Window">要从集合中移除的窗口对象</param>
    void Remove(IVbeWindow Window);

    /// <summary>
    /// 向集合中添加一个窗口
    /// </summary>
    /// <param name="Window">要添加到集合中的窗口对象</param>
    void Add(IVbeWindow Window);
}