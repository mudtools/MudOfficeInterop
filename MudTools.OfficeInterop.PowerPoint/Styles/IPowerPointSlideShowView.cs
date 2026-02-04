//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// 表示 PowerPoint 幻灯片放映的视图。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSlideShowView : IOfficeObject<IPowerPointSlideShowView, MsPowerPoint.SlideShowView>, IDisposable
{
    /// <summary>
    /// 获取创建此幻灯片放映视图的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此幻灯片放映视图的父对象。
    /// </summary>
    /// <value>表示此幻灯片放映视图父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取幻灯片放映的缩放比例。
    /// </summary>
    /// <value>表示缩放比例的整数值。</value>
    int Zoom { get; }

    /// <summary>
    /// 获取当前显示的幻灯片。
    /// </summary>
    /// <value>表示当前幻灯片的 <see cref="IPowerPointSlide"/> 对象。</value>
    IPowerPointSlide? Slide { get; }

    /// <summary>
    /// 获取或设置幻灯片放映中指针的类型。
    /// </summary>
    /// <value>表示指针类型的 <see cref="PpSlideShowPointerType"/> 枚举值。</value>
    PpSlideShowPointerType PointerType { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的状态。
    /// </summary>
    /// <value>表示幻灯片放映状态的 <see cref="PpSlideShowState"/> 枚举值。</value>
    PpSlideShowState State { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用快捷键。
    /// </summary>
    /// <value>指示是否启用快捷键的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool AcceleratorsEnabled { get; set; }

    /// <summary>
    /// 获取幻灯片放映已运行的累计时间（以秒为单位）。
    /// </summary>
    /// <value>表示累计时间的浮点数。</value>
    float PresentationElapsedTime { get; }

    /// <summary>
    /// 获取或设置当前幻灯片已显示的时间（以秒为单位）。
    /// </summary>
    /// <value>表示幻灯片显示时间的浮点数。</value>
    float SlideElapsedTime { get; set; }

    /// <summary>
    /// 获取最后查看的幻灯片。
    /// </summary>
    /// <value>表示最后查看幻灯片的 <see cref="IPowerPointSlide"/> 对象。</value>
    IPowerPointSlide? LastSlideViewed { get; }

    /// <summary>
    /// 获取幻灯片放映的推进模式。
    /// </summary>
    /// <value>表示推进模式的 <see cref="PpSlideShowAdvanceMode"/> 枚举值。</value>
    PpSlideShowAdvanceMode AdvanceMode { get; }

    /// <summary>
    /// 获取幻灯片放映中指针的颜色设置。
    /// </summary>
    /// <value>表示指针颜色的 <see cref="IPowerPointColorFormat"/> 对象。</value>
    IPowerPointColorFormat? PointerColor { get; }

    /// <summary>
    /// 获取一个值，指示是否为命名幻灯片放映。
    /// </summary>
    /// <value>指示是否为命名幻灯片放映的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool IsNamedShow { get; }

    /// <summary>
    /// 获取幻灯片放映的名称。
    /// </summary>
    /// <value>表示幻灯片放映名称的字符串。</value>
    string? SlideShowName { get; }

    /// <summary>
    /// 在幻灯片上绘制一条线。
    /// </summary>
    /// <param name="beginX">起始点的 X 坐标（以磅为单位）。</param>
    /// <param name="beginY">起始点的 Y 坐标（以磅为单位）。</param>
    /// <param name="endX">结束点的 X 坐标（以磅为单位）。</param>
    /// <param name="endY">结束点的 Y 坐标（以磅为单位）。</param>
    void DrawLine(float beginX, float beginY, float endX, float endY);

    /// <summary>
    /// 擦除幻灯片上的所有绘图。
    /// </summary>
    void EraseDrawing();

    /// <summary>
    /// 转到第一张幻灯片。
    /// </summary>
    void First();

    /// <summary>
    /// 转到最后一张幻灯片。
    /// </summary>
    void Last();

    /// <summary>
    /// 转到下一张幻灯片。
    /// </summary>
    void Next();

    /// <summary>
    /// 转到上一张幻灯片。
    /// </summary>
    void Previous();

    /// <summary>
    /// 转到指定索引的幻灯片。
    /// </summary>
    /// <param name="index">要转到的幻灯片的索引（从1开始）。</param>
    /// <param name="resetSlide">指示是否重置幻灯片时间的布尔值。</param>
    void GotoSlide(int index, [ConvertTriState] bool resetSlide = true);

    /// <summary>
    /// 转到指定名称的命名幻灯片放映。
    /// </summary>
    /// <param name="slideShowName">命名幻灯片放映的名称。</param>
    void GotoNamedShow(string slideShowName);

    /// <summary>
    /// 结束命名幻灯片放映。
    /// </summary>
    void EndNamedShow();

    /// <summary>
    /// 重置当前幻灯片的显示时间。
    /// </summary>
    void ResetSlideTime();

    /// <summary>
    /// 退出幻灯片放映。
    /// </summary>
    void Exit();

    /// <summary>
    /// 获取幻灯片放映中的当前显示位置。
    /// </summary>
    /// <value>表示当前显示位置的整数值。</value>
    int CurrentShowPosition { get; }

    /// <summary>
    /// 转到指定的点击位置。
    /// </summary>
    /// <param name="index">点击位置的索引。</param>
    void GotoClick(int index);

    /// <summary>
    /// 获取当前点击的索引。
    /// </summary>
    /// <returns>当前点击的索引。</returns>
    int? GetClickIndex();

    /// <summary>
    /// 获取当前幻灯片的总点击次数。
    /// </summary>
    /// <returns>总点击次数。</returns>
    int? GetClickCount();

    /// <summary>
    /// 检查第一个动画是否为自动播放。
    /// </summary>
    /// <returns>如果第一个动画是自动播放，则返回 true；否则返回 false。</returns>
    bool? FirstAnimationIsAutomatic();

    /// <summary>
    /// 获取指定形状的播放器。
    /// </summary>
    /// <param name="shapeId">形状的标识符。</param>
    /// <returns>指定形状的 <see cref="IPowerPointPlayer"/> 对象。</returns>
    IPowerPointPlayer? Player(object shapeId);

    /// <summary>
    /// 获取一个值，指示媒体控件是否可见。
    /// </summary>
    /// <value>指示媒体控件是否可见的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool MediaControlsVisible { get; }

    /// <summary>
    /// 获取媒体控件的左边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示左边缘位置的浮点数。</value>
    float MediaControlsLeft { get; }

    /// <summary>
    /// 获取媒体控件的上边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示上边缘位置的浮点数。</value>
    float MediaControlsTop { get; }

    /// <summary>
    /// 获取媒体控件的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示宽度的浮点数。</value>
    float MediaControlsWidth { get; }

    /// <summary>
    /// 获取媒体控件的高度（以磅为单位）。
    /// </summary>
    /// <value>表示高度的浮点数。</value>
    float MediaControlsHeight { get; }
}