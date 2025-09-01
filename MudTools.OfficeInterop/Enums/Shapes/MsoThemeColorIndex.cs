//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 主题颜色索引的枚举类型
/// </summary>
public enum MsoThemeColorIndex
{
    /// <summary>
/// 混合主题颜色
/// </summary>
    msoThemeColorMixed = -2,
    /// <summary>
/// 非主题颜色
/// </summary>
    msoNotThemeColor = 0,
    /// <summary>
/// 主题深色 1
/// </summary>
    msoThemeColorDark1 = 1,
    /// <summary>
/// 主题浅色 1
/// </summary>
    msoThemeColorLight1 = 2,
    /// <summary>
/// 主题深色 2
/// </summary>
    msoThemeColorDark2 = 3,
    /// <summary>
/// 主题浅色 2
/// </summary>
    msoThemeColorLight2 = 4,
    /// <summary>
/// 强调色 1
/// </summary>
    msoThemeColorAccent1 = 5,
    /// <summary>
/// 强调色 2
/// </summary>
    msoThemeColorAccent2 = 6,
    /// <summary>
/// 强调色 3
/// </summary>
    msoThemeColorAccent3 = 7,
    /// <summary>
/// 强调色 4
/// </summary>
    msoThemeColorAccent4 = 8,
    /// <summary>
/// 强调色 5
/// </summary>
    msoThemeColorAccent5 = 9,
    /// <summary>
/// 强调色 6
/// </summary>
    msoThemeColorAccent6 = 10,
    /// <summary>
/// 超链接颜色
/// </summary>
    msoThemeColorHyperlink = 11,
    /// <summary>
/// 已访问的超链接颜色
/// </summary>
    msoThemeColorFollowedHyperlink = 12,
    /// <summary>
/// 文本颜色 1
/// </summary>
    msoThemeColorText1 = 13,
    /// <summary>
/// 背景颜色 1
/// </summary>
    msoThemeColorBackground1 = 14,
    /// <summary>
/// 文本颜色 2
/// </summary>
    msoThemeColorText2 = 15,
    /// <summary>
/// 背景颜色 2
/// </summary>
    msoThemeColorBackground2 = 16
}