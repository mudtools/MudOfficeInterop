//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Fields 的接口，用于操作Word域集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordFields : IEnumerable<IWordField?>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取域集合中的域数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取域（从1开始）。
    /// </summary>
    IWordField? this[int index] { get; }

    /// <summary>
    /// 获取域集合的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，确定 Fields 集合中的所有字段是否被锁定。
    /// </summary>
    int Locked { get; set; }

    /// <summary>
    /// 在字段代码和字段结果之间切换字段的显示。
    /// </summary>
    void ToggleShowCodes();

    /// <summary>
    /// 更新字段对象的结果。
    /// </summary>
    /// <returns>更新操作的返回值。</returns>
    int? Update();

    /// <summary>
    /// 用它们的最新结果替换 Fields 集合中的所有字段。
    /// </summary>
    void Unlink();

    /// <summary>
    /// 将对 INCLUDETEXT 字段结果所做的更改保存回源文档。
    /// </summary>
    void UpdateSource();

    /// <summary>
    /// 将 Field 对象添加到 Fields 集合。
    /// </summary>
    /// <param name="range">必需 Range 对象。要添加字段的范围。如果范围未折叠，字段将替换该范围。</param>
    /// <param name="type">可选 Object。可以是任何 WdFieldType 常量。有关有效常量的列表，请查阅对象浏览器。默认值为 wdFieldEmpty。</param>
    /// <param name="text">可选 Object。字段需要的附加文本。例如，如果要为字段指定开关，可以在此处添加。</param>
    /// <param name="preserveFormatting">可选 Object。如果为 True，则在对字段更新期间保留应用于字段的格式。</param>
    /// <returns>Microsoft.Office.Interop.Word.Field</returns>
    IWordField? Add(IWordRange range, WdFieldType? type = null, string? text = null, bool? preserveFormatting = null);
}