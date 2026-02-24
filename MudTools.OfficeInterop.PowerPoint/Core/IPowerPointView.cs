//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 文档窗口的视图。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointView : IOfficeObject<IPowerPointView, MsPowerPoint.View>, IDisposable
{
    /// <summary>
    /// 获取创建此视图的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此视图的父对象。
    /// </summary>
    /// <value>表示此视图父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取视图的类型。
    /// </summary>
    /// <value>表示视图类型的 <see cref="PpViewType"/> 枚举值。</value>
    PpViewType Type { get; }

    /// <summary>
    /// 获取或设置视图的缩放比例。
    /// </summary>
    /// <value>表示缩放比例的整数值。</value>
    int Zoom { get; set; }

    /// <summary>
    /// 粘贴剪贴板内容到视图中。
    /// </summary>
    void Paste();

    /// <summary>
    /// 获取或设置视图中当前显示的幻灯片。
    /// </summary>
    /// <value>表示当前幻灯片的对象。</value>
    object? Slide { get; set; }

    /// <summary>
    /// 转到指定索引的幻灯片。
    /// </summary>
    /// <param name="index">要转到的幻灯片的索引（从1开始）。</param>
    void GotoSlide(int index);

    /// <summary>
    /// 获取或设置一个值，指示是否显示幻灯片缩略图。
    /// </summary>
    /// <value>指示是否显示幻灯片缩略图的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool DisplaySlideMiniature { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否缩放以适应窗口。
    /// </summary>
    /// <value>指示是否缩放以适应窗口的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ZoomToFit { get; set; }

    /// <summary>
    /// 以特殊格式粘贴剪贴板内容到视图中。
    /// </summary>
    /// <param name="dataType">粘贴的数据类型。</param>
    /// <param name="displayAsIcon">指示是否显示为图标的布尔值。</param>
    /// <param name="iconFileName">图标文件的名称。</param>
    /// <param name="iconIndex">图标的索引。</param>
    /// <param name="iconLabel">图标的标签。</param>
    /// <param name="link">指示是否链接到文件的布尔值。</param>
    void PasteSpecial(PpPasteDataType dataType = PpPasteDataType.ppPasteDefault, [ConvertTriState] bool displayAsIcon = false, string iconFileName = "", int iconIndex = 0, string iconLabel = "", [ConvertTriState] bool link = false);

    /// <summary>
    /// 获取视图的打印选项。
    /// </summary>
    /// <value>表示打印选项的 <see cref="IPowerPointPrintOptions"/> 对象。</value>
    IPowerPointPrintOptions? PrintOptions { get; }

    /// <summary>
    /// 打印视图中的内容。
    /// </summary>
    /// <param name="from">打印的起始页。值为-1表示从第一页开始。</param>
    /// <param name="to">打印的结束页。值为-1表示打印到最后一页。</param>
    /// <param name="printToFile">要打印到的文件名。</param>
    /// <param name="copies">打印份数。</param>
    /// <param name="collate">指示是否逐份打印的布尔值。</param>
    void PrintOut(int from = -1, int to = -1, string printToFile = "", int copies = 0, [ConvertTriState] bool collate = false);

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