//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Vbe;

/// <summary>
/// 表示 VBA 编辑器中的代码窗格。
/// </summary>
[ComObjectWrap(ComNamespace = "MsVb")]
public interface IVbeCodePane : IOfficeObject<IVbeCodePane>, IDisposable
{
    /// <summary>
    /// 获取表示 VBA 编辑器环境的 VBE 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 获取包含此代码窗格的 CodePanes 集合。
    /// </summary>
    IVbeCodePanes? Collection { get; }

    /// <summary>
    /// 获取包含此代码窗格的窗口对象。
    /// </summary>
    IVbeWindow? Window { get; }

    /// <summary>
    /// 获取代码窗格中的当前选区。
    /// </summary>
    /// <param name="startLine">输出参数：选区起始行号。</param>
    /// <param name="startColumn">输出参数：选区起始列号。</param>
    /// <param name="endLine">输出参数：选区结束行号。</param>
    /// <param name="endColumn">输出参数：选区结束列号。</param>
    void GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);

    /// <summary>
    /// 设置代码窗格中的选区。
    /// </summary>
    /// <param name="startLine">选区起始行号。</param>
    /// <param name="startColumn">选区起始列号。</param>
    /// <param name="endLine">选区结束行号。</param>
    /// <param name="endColumn">选区结束列号。</param>
    void SetSelection(int startLine, int startColumn, int endLine, int endColumn);

    /// <summary>
    /// 获取或设置代码窗格中可见的第一行行号。
    /// </summary>
    int TopLine { get; set; }

    /// <summary>
    /// 获取代码窗格中可见的行数。
    /// </summary>
    int CountOfVisibleLines { get; }

    /// <summary>
    /// 获取与此代码窗格关联的代码模块。
    /// </summary>
    IVbeCodeModule? CodeModule { get; }

    /// <summary>
    /// 显示代码窗格（使其成为活动窗格）。
    /// </summary>
    void Show();

    /// <summary>
    /// 获取代码窗格的视图模式。
    /// </summary>
    vbext_CodePaneview CodePaneView { get; }
}