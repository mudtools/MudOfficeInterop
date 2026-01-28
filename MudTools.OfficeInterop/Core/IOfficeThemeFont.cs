//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft Office 2007 主题颜色方案中的颜色。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeThemeFont : IOfficeObject<IOfficeThemeFont, MsCore.ThemeFont>, IDisposable
{
    /// <summary>
    /// 获取 Microsoft.Office.Core.ThemeFont 对象的父对象。只读。
    /// </summary>
    /// <returns>Object</returns>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置 Microsoft Office 主题字体方案中字体的名称。可读写。
    /// </summary>
    /// <returns>String</returns>
    string Name { get; set; }
}