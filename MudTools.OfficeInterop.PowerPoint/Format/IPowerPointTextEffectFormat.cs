//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示艺术字文本效果格式，用于控制艺术字对象的文本外观和布局。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTextEffectFormat : IOfficeObject<IPowerPointTextEffectFormat, MsPowerPoint.TextEffectFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此对象的应用程序对象。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此对象的创建者标识符。
    /// </summary>
    /// <value>创建者的整数标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此文本效果格式的父对象。
    /// </summary>
    /// <value>父对象。</value>
    object Parent { get; }

    /// <summary>
    /// 切换文本的垂直方向。
    /// </summary>
    void ToggleVerticalText();

    /// <summary>
    /// 获取或设置文本效果的对齐方式。
    /// </summary>
    /// <value>文本对齐方式的枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextEffectAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文本是否使用粗体。
    /// </summary>
    /// <value>如果文本为粗体则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontBold { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文本是否使用斜体。
    /// </summary>
    /// <value>如果文本为斜体则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontItalic { get; set; }

    /// <summary>
    /// 获取或设置文本使用的字体名称。
    /// </summary>
    /// <value>字体名称字符串。</value>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置文本的字体大小。
    /// </summary>
    /// <value>字体大小值。</value>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否对字符对进行字距调整。
    /// </summary>
    /// <value>如果启用了字距调整则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool KernedPairs { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文本高度是否已标准化。
    /// </summary>
    /// <value>如果文本高度已标准化则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool NormalizedHeight { get; set; }

    /// <summary>
    /// 获取或设置预设的文本效果形状。
    /// </summary>
    /// <value>预设形状的枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffectShape PresetShape { get; set; }

    /// <summary>
    /// 获取或设置预设的文本效果。
    /// </summary>
    /// <value>预设文本效果的枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffect PresetTextEffect { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示字符是否已旋转。
    /// </summary>
    /// <value>如果字符已旋转则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotatedChars { get; set; }

    /// <summary>
    /// 获取或设置艺术字对象中的文本内容。
    /// </summary>
    /// <value>文本内容字符串。</value>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置字符间距。
    /// </summary>
    /// <value>字符间距值。</value>
    float Tracking { get; set; }
}