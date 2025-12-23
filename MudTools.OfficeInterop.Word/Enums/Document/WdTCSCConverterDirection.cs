//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定如何在繁体中文和简体中文之间进行转换
/// </summary>
public enum WdTCSCConverterDirection
{
    /// <summary>
    /// 从简体中文转换为繁体中文
    /// </summary>
    wdTCSCConverterDirectionSCTC,

    /// <summary>
    /// 从繁体中文转换为简体中文
    /// </summary>
    wdTCSCConverterDirectionTCSC,

    /// <summary>
    /// 根据指定范围的检测语言自动选择适当的方向进行转换
    /// </summary>
    wdTCSCConverterDirectionAuto
}