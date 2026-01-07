//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示将文档保存为网页时用于特定字符集的默认字体。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeWebPageFont : IOfficeObject<IOfficeWebPageFont, MsCore.WebPageFont>, IDisposable
{
    /// <summary>
    /// 获取表示对象容器应用程序的Application对象。
    /// </summary>
    object? Application { get; }

    /// <summary>
    /// 获取一个32位整数，指示创建指定对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置宿主应用程序中的比例字体设置。
    /// </summary>
    /// <value>比例字体的名称。</value>
    string ProportionalFont { get; set; }

    /// <summary>
    /// 获取或设置宿主应用程序中的比例字体大小设置（以磅为单位）。
    /// </summary>
    /// <value>比例字体的大小。</value>
    float ProportionalFontSize { get; set; }

    /// <summary>
    /// 获取或设置宿主应用程序中的等宽字体设置。
    /// </summary>
    /// <value>等宽字体的名称。</value>
    string FixedWidthFont { get; set; }

    /// <summary>
    /// 获取或设置宿主应用程序中的等宽字体大小设置（以磅为单位）。
    /// </summary>
    /// <value>等宽字体的大小。</value>
    float FixedWidthFontSize { get; set; }
}