//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// 表示 PowerPoint 的幻灯片放映窗口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSlideShowWindow : IDisposable
{
    /// <summary>
    /// 获取创建此幻灯片放映窗口的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此幻灯片放映窗口的父对象。
    /// </summary>
    /// <value>表示此幻灯片放映窗口父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取幻灯片放映视图。
    /// </summary>
    /// <value>表示幻灯片放映视图的 <see cref="IPowerPointSlideShowView"/> 对象。</value>
    IPowerPointSlideShowView? View { get; }

    /// <summary>
    /// 获取幻灯片放映窗口关联的演示文稿。
    /// </summary>
    /// <value>表示演示文稿的 <see cref="IPowerPointPresentation"/> 对象。</value>
    IPowerPointPresentation? Presentation { get; }

    /// <summary>
    /// 获取一个值，指示幻灯片放映是否以全屏模式显示。
    /// </summary>
    /// <value>指示是否为全屏模式的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool IsFullScreen { get; }

    /// <summary>
    /// 获取或设置幻灯片放映窗口的左边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示左边缘位置的浮点数。</value>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映窗口的上边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示上边缘位置的浮点数。</value>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映窗口的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示宽度的浮点数。</value>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映窗口的高度（以磅为单位）。
    /// </summary>
    /// <value>表示高度的浮点数。</value>
    float Height { get; set; }

    /// <summary>
    /// 获取幻灯片放映窗口的窗口句柄。
    /// </summary>
    /// <value>表示窗口句柄的整数值。</value>
    int HWND { get; }

    /// <summary>
    /// 获取一个值，指示幻灯片放映窗口是否为活动窗口。
    /// </summary>
    /// <value>指示是否为活动窗口的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Active { get; }

    /// <summary>
    /// 激活幻灯片放映窗口。
    /// </summary>
    void Activate();

    /// <summary>
    /// 获取幻灯片导航对象。
    /// </summary>
    /// <value>表示幻灯片导航的 <see cref="IPowerPointSlideNavigation"/> 对象。</value>
    IPowerPointSlideNavigation? SlideNavigation { get; }
}