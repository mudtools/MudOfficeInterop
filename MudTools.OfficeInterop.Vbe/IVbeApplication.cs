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
public interface IVbeApplication : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 VBE 应用程序的版本号
    /// 对应 Application.Version 属性 (通过 VBE)
    /// </summary>
    string Version { get; }
    #endregion

    #region 状态属性
    /// <summary>
    /// 获取 VBE 应用程序是否可见
    /// 对应 VBE.MainWindow.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    #endregion

    #region 核心对象集合和属性
    /// <summary>
    /// 获取 VB 项目集合
    /// 对应 VBE.VBProjects 属性
    /// </summary>
    IVbeVBProjects VBProjects { get; }

    /// <summary>
    /// 获取当前活动的 VB 项目
    /// 对应 VBE.ActiveVBProject 属性
    /// </summary>
    object ActiveVBProject { get; } // 使用 object 作为通用占位符

    /// <summary>
    /// 获取当前活动的 VB 组件
    /// 对应 VBE.SelectedVBComponent 属性
    /// </summary>
    IVbeVBComponent ActiveVBComponent { get; }

    /// <summary>
    /// 获取当前活动的代码窗格
    /// 对应 VBE.ActiveCodePane 属性
    /// </summary>
    object ActiveCodePane { get; } // 使用 object 作为通用占位符
    #endregion

    #region 环境和设置
    /// <summary>
    /// 获取或设置 VBE 主窗口的窗口状态（正常、最小化、最大化）
    /// 对应 VBE.MainWindow.WindowState 属性
    /// </summary>
    vbext_WindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置 VBE 主窗口的左边距
    /// 对应 VBE.MainWindow.Left 属性
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// 获取或设置 VBE 主窗口的顶边距
    /// 对应 VBE.MainWindow.Top 属性
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// 获取或设置 VBE 主窗口的宽度
    /// 对应 VBE.MainWindow.Width 属性
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取或设置 VBE 主窗口的高度
    /// 对应 VBE.MainWindow.Height 属性
    /// </summary>
    int Height { get; set; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 退出 VBE 应用程序 (通常通过关闭宿主应用程序实现)
    /// </summary>
    void Quit();

    /// <summary>
    /// 保存所有打开的 VB 项目
    /// </summary>
    void SaveAll();
    #endregion 
}