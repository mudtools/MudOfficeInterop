//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定文档主题颜色索引的枚举类型，用于标识Word文档中使用的不同主题颜色
/// </summary>
public enum WdThemeColorIndex
{
    /// <summary>
    /// 表示非主题颜色，值为-1
    /// </summary>
    wdNotThemeColor = -1,

    /// <summary>
    /// 主题深色1
    /// </summary>
    wdThemeColorMainDark1,

    /// <summary>
    /// 主题浅色1
    /// </summary>
    wdThemeColorMainLight1,

    /// <summary>
    /// 主题深色2
    /// </summary>
    wdThemeColorMainDark2,

    /// <summary>
    /// 主题浅色2
    /// </summary>
    wdThemeColorMainLight2,

    /// <summary>
    /// 强调颜色1
    /// </summary>
    wdThemeColorAccent1,

    /// <summary>
    /// 强调颜色2
    /// </summary>
    wdThemeColorAccent2,

    /// <summary>
    /// 强调颜色3
    /// </summary>
    wdThemeColorAccent3,

    /// <summary>
    /// 强调颜色4
    /// </summary>
    wdThemeColorAccent4,

    /// <summary>
    /// 强调颜色5
    /// </summary>
    wdThemeColorAccent5,

    /// <summary>
    /// 强调颜色6
    /// </summary>
    wdThemeColorAccent6,

    /// <summary>
    /// 超链接颜色
    /// </summary>
    wdThemeColorHyperlink,

    /// <summary>
    /// 已访问的超链接颜色
    /// </summary>
    wdThemeColorHyperlinkFollowed,

    /// <summary>
    /// 背景颜色1
    /// </summary>
    wdThemeColorBackground1,

    /// <summary>
    /// 文本颜色1
    /// </summary>
    wdThemeColorText1,

    /// <summary>
    /// 背景颜色2
    /// </summary>
    wdThemeColorBackground2,

    /// <summary>
    /// 文本颜色2
    /// </summary>
    wdThemeColorText2
}