//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Microsoft Office Word 如何紧密环绕文本框的文本
/// </summary>
public enum WdTextboxTightWrap
{
    /// <summary>
    /// 不将文本紧密环绕在文本框内容周围
    /// </summary>
    wdTightNone,

    /// <summary>
    /// 将文本在所有行上紧密环绕文本框，紧贴文本框内容
    /// </summary>
    wdTightAll,

    /// <summary>
    /// 仅在首行和末行紧密环绕文本
    /// </summary>
    wdTightFirstAndLastLines,

    /// <summary>
    /// 仅在首行紧密环绕文本
    /// </summary>
    wdTightFirstLineOnly,

    /// <summary>
    /// 仅在末行紧密环绕文本
    /// </summary>
    wdTightLastLineOnly
}