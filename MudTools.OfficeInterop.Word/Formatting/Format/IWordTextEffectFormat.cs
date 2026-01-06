//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示与艺术字对象关联的文本效果格式。
/// <para>注：使用 Shape 对象的 TextEffect 属性可返回 TextEffectFormat 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTextEffectFormat : IOfficeObject<IWordTextEffectFormat, MsWord.TextEffectFormat>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置艺术字的对齐方式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextEffectAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置艺术字字体是否为粗体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontBold { get; set; }

    /// <summary>
    /// 获取或设置艺术字字体是否为斜体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontItalic { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体名称。
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体大小（以磅为单位）。
    /// </summary>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示艺术字字符对是否经过字距调整。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool KernedPairs { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将艺术字中的所有字符调整为相同高度。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool NormalizedHeight { get; set; }

    /// <summary>
    /// 获取或设置艺术字的预设形状。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffectShape PresetShape { get; set; }

    /// <summary>
    /// 获取或设置艺术字的预设文本效果。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffect PresetTextEffect { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示艺术字中的字符是否旋转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotatedChars { get; set; }

    /// <summary>
    /// 获取或设置艺术字中的文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置艺术字中文本的字符间距比例。
    /// </summary>
    float Tracking { get; set; }

    /// <summary>
    /// 在艺术字中的文本排列方式（水平或垂直）之间切换。
    /// </summary>
    void ToggleVerticalText();
}