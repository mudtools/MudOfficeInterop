//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示单个题注标签。
/// 此接口提供对题注标签属性的访问，包括标签名称、编号样式、章节编号设置等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCaptionLabel : IOfficeObject<IWordCaptionLabel, MsWord.CaptionLabel>, IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 <see cref="IWordApplication"/> 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取指定对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取一个值，指示指定对象是否是 Microsoft Word 中的内置样式或题注标签之一。
    /// </summary>
    bool BuiltIn { get; }

    /// <summary>
    /// 如果 CaptionLabel 对象的 BuiltIn 属性为 True，则获取表示指定题注标签类型的值。
    /// </summary>
    WdCaptionLabelID ID { get; }

    /// <summary>
    /// 获取或设置一个值，指示页码或题注标签是否包含章节编号。
    /// </summary>
    bool IncludeChapterNumber { get; set; }

    /// <summary>
    /// 获取或设置 CaptionLabel 对象的编号样式。
    /// </summary>
    WdCaptionNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置在题注标签包含章节编号时标记新章节的标题样式。
    /// </summary>
    int ChapterStyleLevel { get; set; }

    /// <summary>
    /// 获取或设置章节编号和序列号之间的字符。
    /// </summary>
    WdSeparatorType Separator { get; set; }

    /// <summary>
    /// 获取或设置题注标签文本的位置。
    /// </summary>
    WdCaptionPosition Position { get; set; }

    /// <summary>
    /// 删除指定的题注标签。
    /// </summary>
    void Delete();
}