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
[ComObjectWrap(ComNamespace = "MsVb")]
public interface IVbeWindow : IOfficeObject<IVbeWindow>, IDisposable
{
    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    IVbeApplication VBE { get; }

    /// <summary>
    /// 获取包含所有窗口的 Windows 集合对象。
    /// </summary>
    IVbeWindows Collection { get; }

    /// <summary>
    /// 关闭窗口。
    /// </summary>
    void Close();

    /// <summary>
    /// 获取窗口的标题文本。
    /// </summary>
    string Caption { get; }

    /// <summary>
    /// 获取或设置一个值，指示窗口是否可见。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置窗口左边缘相对于屏幕左边缘的水平位置（以像素为单位）。
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// 获取或设置窗口上边缘相对于屏幕上边缘的垂直位置（以像素为单位）。
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// 获取或设置窗口的宽度（以像素为单位）。
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取或设置窗口的高度（以像素为单位）。
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// 获取或设置窗口的状态（最小化、最大化或正常）。
    /// </summary>
    vbext_WindowState WindowState { get; set; }

    /// <summary>
    /// 将焦点设置到该窗口。
    /// </summary>
    void SetFocus();

    /// <summary>
    /// 获取窗口的类型（如代码窗口、对象浏览器等）。
    /// </summary>
    vbext_WindowType Type { get; }

    /// <summary>
    /// 设置窗口的类型。
    /// </summary>
    /// <param name="kind">要设置的窗口类型。</param>
    void SetKind(vbext_WindowType kind);

    /// <summary>
    /// 获取窗口所属的链接窗口集合。
    /// </summary>
    IVbeLinkedWindows LinkedWindows { get; }

    /// <summary>
    /// 获取窗口的链接窗口框架。
    /// </summary>
    IVbeWindow LinkedWindowFrame { get; }

    /// <summary>
    /// 将窗口从其链接窗口集合中分离。
    /// </summary>
    void Detach();

    /// <summary>
    /// 将窗口附加到指定的窗口句柄。
    /// </summary>
    /// <param name="windowHandle">要附加的窗口句柄。</param>
    void Attach(int windowHandle);

    /// <summary>
    /// 获取窗口的句柄（HWND）。
    /// </summary>
    int HWnd { get; }
}