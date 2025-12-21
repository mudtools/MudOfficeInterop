using System.Drawing;

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// 表示 Word 图表字体格式的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordChartFont : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置字体名称。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float Size { get; set; }

    /// <summary>
    /// 获取或设置是否粗体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置下划线类型。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlUnderlineStyle Underline { get; set; }

    /// <summary>
    /// 获取或设置字体颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置字体颜色索引。
    /// </summary>
    XlColorIndex ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置图表字体的背景色样式。
    /// </summary>
    XlBackground Background { get; set; }

    /// <summary>
    /// 获取或设置字体样式（如常规、粗体、斜体等）。
    /// </summary>
    object FontStyle { get; set; }


    /// <summary>
    /// 获取或设置是否使用空心字体效果。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool OutlineFont { get; set; }

    /// <summary>
    /// 获取或设置是否使用阴影字体效果。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置是否使用删除线字体效果。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool StrikeThrough { get; set; }

    /// <summary>
    /// 获取或设置是否使用下标字体效果。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置是否使用上标字体效果。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Superscript { get; set; }
}