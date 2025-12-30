//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;
/// <summary>
/// VBE Application 对象的二次封装接口
/// 提供对 Microsoft.Vbe.Interop.Application (通过 VBE 对象访问) 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsVb", ComClassName = "VBE")]
public interface IVbeApplication : IOfficeObject<IVbeApplication>, IDisposable
{
    /// <summary>
    /// 获取 VBE 的版本号。
    /// </summary>
    string Version { get; }

    /// <summary>
    /// 获取 VBE 中的 VBProjects 集合，包含所有 VB 项目。
    /// </summary>
    IVbeVBProjects? VBProjects { get; }

    /// <summary>
    /// 获取 VBE 中的命令栏集合。
    /// </summary>
    IOfficeCommandBars? CommandBars { get; }

    /// <summary>
    /// 获取 VBE 中的代码窗格集合。
    /// </summary>
    IVbeCodePanes? CodePanes { get; }

    /// <summary>
    /// 获取 VBE 中的窗口集合。
    /// </summary>
    IVbeWindows? Windows { get; }

    /// <summary>
    /// 获取或设置当前活动的 VB 项目。
    /// </summary>
    IVbeVBProject? ActiveVBProject { get; set; }

    /// <summary>
    /// 获取当前选定的 VB 组件。
    /// </summary>
    IVbeVBComponent? SelectedVBComponent { get; }

    /// <summary>
    /// 获取 VBE 的主窗口。
    /// </summary>
    IVbeWindow? MainWindow { get; }

    /// <summary>
    /// 获取当前活动的窗口。
    /// </summary>
    IVbeWindow? ActiveWindow { get; }

    /// <summary>
    /// 获取或设置当前活动的代码窗格。
    /// </summary>
    IVbeCodePane? ActiveCodePane { get; set; }

    /// <summary>
    /// 获取 VBE 中的加载项集合。
    /// </summary>
    IVbeAddins? Addins { get; }
}