//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word AutoCorrectEntry 对象的封装接口。
/// 用于定义“键入时自动替换”行为，例如将 'teh' 替换为 'the'。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordAutoCorrectEntry : IOfficeObject<IWordAutoCorrectEntry, MsWord.AutoCorrectEntry>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 获取自动更正条目的名称（即触发词，如 "teh"）。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置自动更正条目的替换值（即替换后的内容，如 "the"）。
    /// 注意：设置此属性会修改 Word 中的实际条目。
    /// </summary>
    string Value { get; set; }

    /// <summary>
    /// 删除当前自动更正条目。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将当前自动更正条目应用到指定的文本范围。
    /// </summary>
    /// <param name="range">要应用自动更正的文本范围。</param>
    void Apply(IWordRange range);
}