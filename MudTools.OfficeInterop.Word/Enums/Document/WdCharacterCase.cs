//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定字符的大小写格式类型
/// </summary>
public enum WdCharacterCase
{
    /// <summary>
/// 切换到下一个大小写状态
/// </summary>
    wdNextCase = -1,
    /// <summary>
/// 小写
/// </summary>
    wdLowerCase = 0,
    /// <summary>
/// 大写
/// </summary>
    wdUpperCase = 1,
    /// <summary>
/// 每个单词首字母大写
/// </summary>
    wdTitleWord = 2,
    /// <summary>
/// 句子首字母大写
/// </summary>
    wdTitleSentence = 4,
    /// <summary>
/// 切换大小写
/// </summary>
    wdToggleCase = 5,
    /// <summary>
/// 半角字符
/// </summary>
    wdHalfWidth = 6,
    /// <summary>
/// 全角字符
/// </summary>
    wdFullWidth = 7,
    /// <summary>
/// 片假名
/// </summary>
    wdKatakana = 8,
    /// <summary>
/// 平假名
/// </summary>
    wdHiragana = 9
}