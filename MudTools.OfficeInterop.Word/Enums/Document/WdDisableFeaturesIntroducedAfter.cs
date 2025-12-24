//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定要禁用该版本之后引入的所有功能的 Microsoft Word 版本。仅适用于通过 DisableFeaturesIntroducedAfter 属性设置的文档，或通过 DisableFeaturesIntroducedAfterbyDefault 属性应用于所有文档。
/// </summary>
public enum WdDisableFeaturesIntroducedAfter
{
    /// <summary>
    /// 指定 Windows 95 版 Word，版本 7.0 和 7.0a
    /// </summary>
    wd70,

    /// <summary>
    /// 指定 Windows 95 版 Word，版本 7.0 和 7.0a，亚洲版
    /// </summary>
    wd70FE,

    /// <summary>
    /// 指定 Windows 版 Word 97。默认值。
    /// </summary>
    wd80
}