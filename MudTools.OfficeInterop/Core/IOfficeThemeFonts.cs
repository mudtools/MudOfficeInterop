//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office主题字体集合的接口，提供对Office文档中主题字体的访问
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore"), ItemIndex]
public interface IOfficeThemeFonts : IOfficeObject<IOfficeThemeFonts>, IEnumerable<IOfficeThemeFont?>, IDisposable
{
    /// <summary>
    /// 获取 Microsoft.Office.Core.ThemeFonts 对象的父对象。只读。
    /// </summary>
    /// <returns>Object</returns>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中主题字体的总数
    /// </summary>
    /// <returns>主题字体的数量</returns>
    int Count { get; }

    /// <summary>
    /// 获取指定索引处的主题字体
    /// </summary>
    /// <param name="index">要获取的字体的从零开始的索引</param>
    /// <returns>指定索引处的IOfficeThemeFont对象，如果索引无效则返回null</returns>
    IOfficeThemeFont? this[MsoFontLanguageIndex index] { get; }
}