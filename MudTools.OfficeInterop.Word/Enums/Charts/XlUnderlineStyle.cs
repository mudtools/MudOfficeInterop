//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定应用于字体的下划线类型
/// </summary>
public enum XlUnderlineStyle
{
    /// <summary>
    /// 双粗下划线
    /// </summary>
    xlUnderlineStyleDouble = -4119,

    /// <summary>
    /// 两条细下划线紧密排列
    /// </summary>
    xlUnderlineStyleDoubleAccounting = 5,

    /// <summary>
    /// 无下划线
    /// </summary>
    xlUnderlineStyleNone = -4142,

    /// <summary>
    /// 单下划线
    /// </summary>
    xlUnderlineStyleSingle = 2,

    /// <summary>
    /// 不支持
    /// </summary>
    xlUnderlineStyleSingleAccounting = 4
}