//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// ProtectedViewWindow 接口及实现类
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordProtectedViewWindow : IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表指定对象的父对象的对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取受保护的视图窗口的窗口标题。
    /// </summary>
    string Caption { get; }

    /// <summary>
    /// 获取受保护的视图窗口的高度（以磅为单位）。
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// 获取或设置受保护的视图窗口的水平位置（以磅为单位）。
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// 获取受保护的视图窗口的垂直位置（以磅为单位）。
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// 获取或设置受保护的视图窗口的宽度（以磅为单位）。
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取受保护的视图窗口的窗口状态。
    /// </summary>
    WdWindowState WindowState { get; set; }

    /// <summary>
    /// 获取受保护的视图窗口的显示状态。
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取受保护的视图窗口所显示的文档对象。
    /// </summary>
    IWordDocument? Document { get; }

    /// <summary>
    /// 获取受保护的视图窗口的唯一标识符。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值代表是否激活受保护的视图窗口。
    /// </summary>
    bool Active { get; }

    /// <summary>
    /// 获取受保护视图窗口中打开的文档的源文件名（不包含路径）。
    /// </summary>
    string SourceName { get; }

    /// <summary>
    /// 获取受保护视图窗口中打开的文档的源文件完整路径。
    /// </summary>
    string SourcePath { get; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 关闭受保护的视图窗口。
    /// </summary>
    void Close();

    /// <summary>
    /// 编辑受保护的视图窗口中的文档。
    /// </summary>
    /// <param name="passwordTemplate">打开模板时所需的密码。</param>
    /// <param name="writePasswordDocument">打开文档时所需的写入密码。</param>
    /// <param name="writePasswordTemplate">打开模板时所需的写入密码。</param>
    /// <returns>返回编辑中的文档对象。</returns>
    IWordDocument? Edit(string? passwordTemplate = null, string? writePasswordDocument = null, string? writePasswordTemplate = null);

    /// <summary>
    /// 激活受保护的视图窗口。
    /// </summary>
    void Activate();

    void ToggleRibbon();
}