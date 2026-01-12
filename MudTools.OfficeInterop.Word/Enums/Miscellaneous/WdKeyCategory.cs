//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定键类别，用于键盘快捷键的分类
/// </summary>
public enum WdKeyCategory
{
    /// <summary>
    /// 空键类别（无效值）
    /// </summary>
    wdKeyCategoryNil = -1,
    /// <summary>
    /// 禁用的键类别
    /// </summary>
    wdKeyCategoryDisable,
    /// <summary>
    /// 命令键类别
    /// </summary>
    wdKeyCategoryCommand,
    /// <summary>
    /// 宏键类别
    /// </summary>
    wdKeyCategoryMacro,
    /// <summary>
    /// 字体键类别
    /// </summary>
    wdKeyCategoryFont,
    /// <summary>
    /// 自动图文集键类别
    /// </summary>
    wdKeyCategoryAutoText,
    /// <summary>
    /// 样式键类别
    /// </summary>
    wdKeyCategoryStyle,
    /// <summary>
    /// 符号键类别
    /// </summary>
    wdKeyCategorySymbol,
    /// <summary>
    /// 前缀键类别
    /// </summary>
    wdKeyCategoryPrefix
}