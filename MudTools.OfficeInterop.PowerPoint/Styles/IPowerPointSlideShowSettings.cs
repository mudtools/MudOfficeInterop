//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 演示文稿的幻灯片放映设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSlideShowSettings : IDisposable
{
    /// <summary>
    /// 获取创建此幻灯片放映设置的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此幻灯片放映设置的父对象。
    /// </summary>
    /// <value>表示此幻灯片放映设置父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取幻灯片放映中指针的颜色设置。
    /// </summary>
    /// <value>表示指针颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? PointerColor { get; }

    /// <summary>
    /// 获取演示文稿中的命名幻灯片放映集合。
    /// </summary>
    /// <value>表示命名幻灯片放映集合的 <see cref="IPowerPointNamedSlideShows"/> 对象。</value>
    IPowerPointNamedSlideShows? NamedSlideShows { get; }

    /// <summary>
    /// 获取或设置幻灯片放映的起始幻灯片编号。
    /// </summary>
    /// <value>表示起始幻灯片编号的整数值。</value>
    int StartingSlide { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的结束幻灯片编号。
    /// </summary>
    /// <value>表示结束幻灯片编号的整数值。</value>
    int EndingSlide { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的推进模式。
    /// </summary>
    /// <value>表示推进模式的 <see cref="PpSlideShowAdvanceMode"/> 枚举值。</value>
    PpSlideShowAdvanceMode AdvanceMode { get; set; }

    /// <summary>
    /// 运行幻灯片放映。
    /// </summary>
    /// <returns>表示幻灯片放映窗口的 <see cref="IPowerPointSlideShowWindow"/> 对象。</returns>
    IPowerPointSlideShowWindow? Run();

    /// <summary>
    /// 获取或设置一个值，指示幻灯片放映是否循环播放直到手动停止。
    /// </summary>
    /// <value>指示是否循环播放的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LoopUntilStopped { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的显示类型。
    /// </summary>
    /// <value>表示显示类型的 <see cref="PpSlideShowType"/> 枚举值。</value>
    PpSlideShowType ShowType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示幻灯片放映时是否播放旁白。
    /// </summary>
    /// <value>指示是否播放旁白的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowWithNarration { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示幻灯片放映时是否播放动画。
    /// </summary>
    /// <value>指示是否播放动画的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowWithAnimation { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的名称。
    /// </summary>
    /// <value>表示幻灯片放映名称的字符串。</value>
    string? SlideShowName { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的范围类型。
    /// </summary>
    /// <value>表示范围类型的 <see cref="PpSlideShowRangeType"/> 枚举值。</value>
    PpSlideShowRangeType RangeType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示幻灯片放映时是否显示滚动条。
    /// </summary>
    /// <value>指示是否显示滚动条的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowScrollbar { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示幻灯片放映时是否显示演示者视图。
    /// </summary>
    /// <value>指示是否显示演示者视图的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowPresenterView { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示幻灯片放映时是否显示媒体控件。
    /// </summary>
    /// <value>指示是否显示媒体控件的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowMediaControls { get; set; }
}