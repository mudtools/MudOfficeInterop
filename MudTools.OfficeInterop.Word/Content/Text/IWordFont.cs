//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Font 的接口，用于操作 Word 文档中文字的字体样式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFont : IDisposable
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
    /// 获取或设置字体名称（如 "Arial", "Times New Roman" 等）。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小（单位为磅）。
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置是否为粗体。
    /// </summary>
    [IgnoreGenerator]
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否为斜体。
    /// </summary>
    [IgnoreGenerator]
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置是否带下划线。
    /// </summary>
    WdUnderline Underline { get; set; }

    /// <summary>
    /// 获取或设置字体颜色（RGB 值）。
    /// </summary>
    WdColor Color { get; set; }

    /// <summary>
    /// 获取或设置上标（如数学中的平方）。
    /// </summary>
    [IgnoreGenerator]
    bool Superscript { get; set; }

    /// <summary>
    /// 获取或设置下标（如化学式 H₂O）。
    /// </summary>
    [IgnoreGenerator]
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置字符间距（单位为磅）。
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置字符缩放比例（百分比，如 100 表示正常大小）。
    /// </summary>
    int Scaling { get; set; }

    /// <summary>
    /// 获取或设置字符位置偏移（正值为上移，负值为下移）。
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// 获取或设置数字间距类型。
    /// </summary>
    WdNumberSpacing NumberSpacing { get; set; }

    /// <summary>
    /// 获取或设置下划线颜色（RGB 值）。
    /// </summary>
    WdColor UnderlineColor { get; set; }

    /// <summary>
    /// 获取或设置东亚语言字体名称。
    /// </summary>
    string NameFarEast { get; set; }

    /// <summary>
    /// 获取或设置是否为轮廓字体。
    /// </summary>
    [IgnoreGenerator]
    bool Outline { get; set; }

    /// <summary>
    /// 获取或设置是否具有阴影效果。
    /// </summary>
    [IgnoreGenerator]
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置是否禁用字符间距网格。
    /// </summary>
    bool DisableCharacterSpaceGrid { get; set; }

    /// <summary>
    /// 获取或设置字体是否隐藏。
    /// </summary>
    [IgnoreGenerator]
    bool Hidden { get; set; }

    /// <summary>
    /// 获取或设置字体颜色索引。
    /// </summary>
    WdColorIndex ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置样式集。
    /// </summary>
    WdStylisticSet StylisticSet { get; set; }

    /// <summary>
    /// 获取字体的副本。
    /// </summary>
    IWordFont? Duplicate { get; }

    /// <summary>
    /// 获取字体边框格式。
    /// </summary>
    IWordBorders? Borders { get; }

    /// <summary>
    /// 获取字体填充格式。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取字体发光效果格式。
    /// </summary>
    IWordGlowFormat? Glow { get; }

    /// <summary>
    /// 获取字体线条格式。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取字体倒影效果格式。
    /// </summary>
    IWordReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取字体底纹格式。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取字体颜色格式。
    /// </summary>
    IWordColorFormat? TextColor { get; }

    /// <summary>
    /// 获取字体阴影格式。
    /// </summary>
    IWordShadowFormat? TextShadow { get; }

    /// <summary>
    /// 获取字体三维格式。
    /// </summary>
    IWordThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 增大字体大小。
    /// </summary>
    void Grow();

    /// <summary>
    /// 减小字体大小。
    /// </summary>
    void Shrink();

    /// <summary>
    /// 将当前字体设置为模板默认字体。
    /// </summary>
    void SetAsTemplateDefault();

    /// <summary>
    /// 重置字体格式为默认设置。
    /// </summary>
    void Reset();
}