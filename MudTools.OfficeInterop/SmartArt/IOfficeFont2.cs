//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 封装 Microsoft.Office.Core.Font2 对象的接口，用于操作 Office 应用程序中的字体格式设置
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeFont2 : IDisposable
{
    /// <summary>
    /// 获取字体对象的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置 ASCII 字符的字体名称
    /// </summary>
    string NameAscii { get; set; }

    /// <summary>
    /// 获取或设置复杂脚本字符的字体名称
    /// </summary>
    string NameComplexScript { get; set; }

    /// <summary>
    /// 获取或设置其他字符的字体名称
    /// </summary>
    string NameOther { get; set; }

    /// <summary>
    /// 获取或设置远东字符的字体名称
    /// </summary>
    string NameFarEast { get; set; }

    /// <summary>
    /// 获取或设置字体是否为粗体
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置字体是否为斜体
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置文本删除线样式
    /// </summary>
    MsoTextStrike Strike { get; set; }

    /// <summary>
    /// 获取或设置文本大写格式
    /// </summary>
    MsoTextCaps Caps { get; set; }

    /// <summary>
    /// 获取或设置数字是否自动旋转
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutorotateNumbers { get; set; }

    /// <summary>
    /// 获取或设置基线下偏移量
    /// </summary>
    float BaselineOffset { get; set; }

    /// <summary>
    /// 获取或设置字符间距
    /// </summary>
    float Kerning { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    float Size { get; set; }

    /// <summary>
    /// 获取或设置字符间距
    /// </summary>
    float Spacing { get; set; }

    /// <summary>
    /// 获取或设置下划线样式
    /// </summary>
    MsoTextUnderlineType UnderlineStyle { get; set; }

    /// <summary>
    /// 获取或设置是否全部大写
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Allcaps { get; set; }

    /// <summary>
    /// 获取或设置是否双删除线
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool DoubleStrikeThrough { get; set; }

    /// <summary>
    /// 获取或设置字符间距是否相等
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Equalize { get; set; }

    /// <summary>
    /// 获取或设置是否小型大写字母
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Smallcaps { get; set; }

    /// <summary>
    /// 获取或设置是否有删除线
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool StrikeThrough { get; set; }

    /// <summary>
    /// 获取或设置是否下标
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置是否上标
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Superscript { get; set; }

    /// <summary>
    /// 获取字体是否可嵌入
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Embeddable { get; }

    /// <summary>
    /// 获取字体是否已嵌入
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Embedded { get; }

    /// <summary>
    /// 获取或设置艺术字格式
    /// </summary>
    MsoPresetTextEffect WordArtformat { get; set; }

    /// <summary>
    /// 获取或设置柔化边缘格式
    /// </summary>
    MsoSoftEdgeType SoftEdgeFormat { get; set; }

    /// <summary>
    /// 获取字体的填充格式
    /// </summary>
    IOfficeFillFormat Fill { get; }

    /// <summary>
    /// 获取字体的辉光格式
    /// </summary>
    IOfficeGlowFormat Glow { get; }

    /// <summary>
    /// 获取字体的倒影格式
    /// </summary>
    IOfficeReflectionFormat Reflection { get; }

    /// <summary>
    /// 获取字体的线条格式
    /// </summary>
    IOfficeLineFormat Line { get; }

    /// <summary>
    /// 获取字体的阴影格式
    /// </summary>
    IOfficeShadowFormat Shadow { get; }

    /// <summary>
    /// 获取字体的高亮颜色格式
    /// </summary>
    IOfficeColorFormat Highlight { get; }

    /// <summary>
    /// 获取字体的下划线颜色格式
    /// </summary>
    IOfficeColorFormat UnderlineColor { get; }
}