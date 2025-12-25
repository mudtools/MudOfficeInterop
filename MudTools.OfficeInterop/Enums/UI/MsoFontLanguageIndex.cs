//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft.Office.Core.ThemeFonts 集合中包含的三种语言字体之一
/// </summary>
public enum MsoFontLanguageIndex
{
    /// <summary>
    /// 表示拉丁字体
    /// </summary>
    msoThemeLatin = 1,

    /// <summary>
    /// 表示复杂脚本语言的字体。复杂脚本语言集合支持阿拉伯语、格鲁吉亚语、希伯来语、印度语、泰语和越南语字母
    /// </summary>
    msoThemeComplexScript,

    /// <summary>
    /// 表示东亚字体。东亚语言包括简体中文、繁体中文、日语和韩语
    /// </summary>
    msoThemeEastAsian
}