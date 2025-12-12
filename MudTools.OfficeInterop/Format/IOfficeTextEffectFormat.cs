//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中艺术字格式的接口封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeTextEffectFormat : IDisposable
{
    /// <summary>
    /// 获取或设置艺术字的文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体名称。
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体大小。
    /// </summary>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体粗细。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontBold { get; set; }

    /// <summary>
    /// 获取或设置艺术字的字体倾斜。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool FontItalic { get; set; }

    /// <summary>
    /// 获取或设置艺术字的对齐方式。
    /// </summary>
    MsoTextEffectAlignment Alignment { get; set; }

    /// <summary>
    /// 获取或设置艺术字的样式。
    /// </summary>
    MsoPresetTextEffect PresetTextEffect { get; set; }

    /// <summary>
    /// 获取或设置艺术字的形状类型。
    /// </summary>
    MsoPresetTextEffectShape PresetShape { get; set; }

    /// <summary>
    /// 获取或设置艺术字旋转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool RotatedChars { get; set; }

    /// <summary>
    /// 获取或设置艺术字的水平缩放。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool NormalizedHeight { get; set; }

    /// <summary>
    /// 获取或设置艺术字的间距。
    /// </summary>
    float Tracking { get; set; }

    /// <summary>
    /// 切换艺术字文本的垂直显示状态。
    /// </summary>
    void ToggleVerticalText();
}