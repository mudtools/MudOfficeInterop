//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定在文本文件中用作文本识别符的字符类型
/// </summary>
public enum XlTextQualifier
{
    /// <summary>
    /// 指定双引号 (") 作为文本识别符
    /// </summary>
    xlTextQualifierDoubleQuote = 1,

    /// <summary>
    /// 指定没有文本识别符
    /// </summary>
    xlTextQualifierNone = -4142,

    /// <summary>
    /// 指定单引号 (') 作为文本识别符
    /// </summary>
    xlTextQualifierSingleQuote = 2
}