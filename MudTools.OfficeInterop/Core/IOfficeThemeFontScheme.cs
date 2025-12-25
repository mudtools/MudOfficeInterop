//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft Office 2007 主题的字体方案。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeThemeFontScheme : IDisposable
{
    /// <summary>
    /// 获取 Microsoft.Office.Core.ThemeFontScheme 对象的父对象。只读。
    /// </summary>
    /// <returns>Object</returns>
    object Parent { get; }

    /// <summary>
    /// 获取文档"标题"的字体设置。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.ThemeFonts</returns>
    IOfficeThemeFonts? MajorFont { get; }

    /// <summary>
    /// 获取文档"正文"的字体设置。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.ThemeFonts</returns>
    IOfficeThemeFonts? MinorFont { get; }

    /// <summary>
    /// 从文件加载 Microsoft Office 主题的字体方案。
    /// </summary>
    /// <param name="fileName">字体方案文件的名称。</param>
    void Load(string fileName);

    /// <summary>
    /// 将 Microsoft Office 主题的字体方案保存到文件。
    /// </summary>
    /// <param name="fileName">文件的名称。</param>
    void Save(string fileName);
}