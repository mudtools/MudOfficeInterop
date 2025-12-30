//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE VBProjects 集合对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.VBProjects 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsVb"), ItemIndex]
public interface IVbeWindows : IEnumerable<IVbeWindow?>, IOfficeObject<IVbeWindows>, IDisposable
{
    /// <summary>
    /// 获取 VB 项目的 VBE 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取集合中的窗口数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取窗口对象
    /// </summary>
    /// <param name="index">窗口在集合中的索引位置，从0开始</param>
    /// <returns>指定索引位置的IVbeWindow对象，如果索引无效则返回null</returns>
    IVbeWindow? this[int index] { get; }

    /// <summary>
    /// 根据名称获取窗口对象
    /// </summary>
    /// <param name="name">窗口的名称</param>
    /// <returns>指定名称的IVbeWindow对象，如果找不到则返回null</returns>
    IVbeWindow? this[string name] { get; }

    /// <summary>
    /// 创建一个工具窗口
    /// </summary>
    /// <param name="addInInst">附加组件实例，用于标识创建窗口的插件</param>
    /// <param name="progId">程序标识符，指定要嵌入的OLE对象的ProgID</param>
    /// <param name="caption">窗口的标题文本</param>
    /// <param name="guidPosition">窗口位置的GUID标识符</param>
    /// <param name="DocObj">可选参数，要嵌入的文档对象，默认为null</param>
    /// <returns>新创建的IVbeWindow对象，如果创建失败则返回null</returns>
    IVbeWindow? CreateToolWindow(IVbeAddIn addInInst, string progId, string caption, string guidPosition, object? DocObj = null);
}