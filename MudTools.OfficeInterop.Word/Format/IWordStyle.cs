//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Style 的接口，用于操作文档样式。
/// </summary>
public interface IWordStyle : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    bool InUse { get; }

    /// <summary>
    /// 获取样式的本地化名称。
    /// </summary>
    string NameLocal { get; }

    /// <summary>
    /// 获取或设置样式的类型。
    /// </summary>
    WdStyleType Type { get; }

    /// <summary>
    /// 获取或设置样式的下一个段落样式名称。
    /// </summary>
    string NextParagraphStyle { get; set; }

    /// <summary>
    /// 获取或设置是否自动更新样式。
    /// </summary>
    bool AutomaticallyUpdate { get; set; }

    /// <summary>
    /// 获取或设置是否为快捷样式。
    /// </summary>
    bool QuickStyle { get; set; }

    /// <summary>
    /// 获取或设置是否可见。
    /// </summary>
    bool Visibility { get; set; }

    /// <summary>
    /// 获取样式的字体格式封装对象。
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取样式的段落格式封装对象。
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }

    /// <summary>
    /// 获取样式的编号格式封装对象。
    /// </summary>
    IWordListTemplate ListTemplate { get; }

    /// <summary>
    /// 删除此样式。
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制样式到另一个名称。
    /// </summary>
    /// <param name="newName">新样式名称。</param>
    /// <returns>复制的新样式。</returns>
    IWordStyle Copy(string newName);

    /// <summary>
    /// 应用样式到指定范围。
    /// </summary>
    /// <param name="range">要应用样式的范围。</param>
    void ApplyTo(IWordRange range);

    /// <summary>
    /// 检查样式是否为内置样式。
    /// </summary>
    bool IsBuiltIn { get; }
}