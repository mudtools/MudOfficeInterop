//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Word 2010 打开文档时使用的兼容性模式
/// </summary>
public enum WdCompatibilityMode
{
    /// <summary>
    /// Word 2010 进入与 Word 2003 最兼容的模式。在此模式下，Word 2010 的新功能将被禁用。
    /// </summary>
    wdWord2003 = 11,

    /// <summary>
    /// Word 2010 进入与 Word 2007 最兼容的模式。在此模式下，Word 2010 的新功能将被禁用。
    /// </summary>
    wdWord2007 = 12,

    /// <summary>
    /// 默认值。启用所有 Word 2010 功能。
    /// </summary>
    wdWord2010 = 14,

    /// <summary>
    /// Word 2013 兼容模式
    /// </summary>
    wdWord2013 = 15,

    /// <summary>
    /// 与最新版本 Word 等效的兼容性模式
    /// </summary>
    wdCurrent = 65535
}