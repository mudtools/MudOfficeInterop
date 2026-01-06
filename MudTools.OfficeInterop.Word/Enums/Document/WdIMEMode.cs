//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定日语输入法编辑器 (IME) 的默认启动模式
/// </summary>
public enum WdIMEMode
{
    /// <summary>
    /// 不更改 IME 模式
    /// </summary>
    wdIMEModeNoControl = 0,

    /// <summary>
    /// 激活 IME
    /// </summary>
    wdIMEModeOn = 1,

    /// <summary>
    /// 禁用 IME 并激活拉丁文本输入
    /// </summary>
    wdIMEModeOff = 2,

    /// <summary>
    /// 在全角平假名模式下激活 IME
    /// </summary>
    wdIMEModeHiragana = 4,

    /// <summary>
    /// 在全角片假名模式下激活 IME
    /// </summary>
    wdIMEModeKatakana = 5,

    /// <summary>
    /// 在半角片假名模式下激活 IME
    /// </summary>
    wdIMEModeKatakanaHalf = 6,

    /// <summary>
    /// 在全角拉丁模式下激活 IME
    /// </summary>
    wdIMEModeAlphaFull = 7,

    /// <summary>
    /// 在半角拉丁模式下激活 IME
    /// </summary>
    wdIMEModeAlpha = 8,

    /// <summary>
    /// 在全角韩文模式下激活 IME
    /// </summary>
    wdIMEModeHangulFull = 9,

    /// <summary>
    /// 在半角韩文模式下激活 IME
    /// </summary>
    wdIMEModeHangul = 10
}