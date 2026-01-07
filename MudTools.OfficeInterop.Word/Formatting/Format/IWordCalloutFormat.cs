
namespace MudTools.OfficeInterop.Word;


[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCalloutFormat : IOfficeObject<IWordCalloutFormat, MsWord.CalloutFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用垂直分隔条将标注文本与标注线分开。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Accent { get; set; }

    /// <summary>
    /// 获取或设置标注线的角度类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutAngleType Angle { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否自动设置标注线的长度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoLength { get; }

    /// <summary>
    /// 获取或设置一个值，指示指定标注中的文本是否被边框包围。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Border { get; set; }

    /// <summary>
    /// 对于具有显式设置的垂直距离的标注，获取从文本框边缘到标注线附着点的垂直距离（以磅为单位）。
    /// </summary>
    float Drop { get; }

    /// <summary>
    /// 获取标注线附着到标注文本框的位置类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutDropType DropType { get; }

    /// <summary>
    /// 获取或设置标注线末端与文本框边界之间的水平距离（以磅为单位）。
    /// </summary>
    float Gap { get; set; }

    /// <summary>
    /// 当指定标注的 AutoLength 属性设置为 False 时，获取标注线第一段（附着到文本框的段）的长度（以磅为单位）。
    /// </summary>
    float Length { get; }

    /// <summary>
    /// 获取或设置标注的类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoCalloutType Type { get; set; }

    /// <summary>
    /// 指定当标注移动时，自动缩放标注线的第一段（附着到文本框的段）。
    /// </summary>
    void AutomaticLength();

    /// <summary>
    /// 设置从文本框边缘到标注线附着点的垂直距离（以磅为单位）。
    /// </summary>
    /// <param name="drop">垂直距离，以磅为单位。</param>
    void CustomDrop(float drop);

    /// <summary>
    /// 指定无论标注如何移动，标注线的第一段（附着到文本框的段）都保持固定长度。
    /// </summary>
    /// <param name="length">标注线第一段的长度，以磅为单位。</param>
    void CustomLength(float length);

    /// <summary>
    /// 指定标注线是附着在文本框的顶部、底部、中心，还是附着在距离文本框顶部或底部指定距离的位置。
    /// </summary>
    /// <param name="dropType">标注线相对于文本框边界的起始位置。</param>
    void PresetDrop([ComNamespace("MsCore")] MsoCalloutDropType dropType);
}