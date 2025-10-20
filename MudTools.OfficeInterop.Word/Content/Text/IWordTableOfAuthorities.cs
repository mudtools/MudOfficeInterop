
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的一个引文目录 (Table of Authorities, TOA) 的二次封装接口。
/// 此接口提供了对引文目录内容、格式和操作的访问，同时管理底层 COM 对象的生命周期。
/// </summary>
public interface IWordTableOfAuthorities : IDisposable
{
    /// <summary>
    /// 获取此引文目录所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此引文目录的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此引文目录在文档中所占据的范围 (<see cref="IWordRange"/>)。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示引文目录是否在每个条目后显示 "passim"（表示该条目在多页上出现）。
    /// 此属性对应于 TOA 域的 \p 开关 [[7]]。
    /// </summary>
    bool Passim { get; set; }

    /// <summary>
    /// 获取或设置分隔引文目录中各项与其页码的字符（最多五个）。
    /// 此属性对应于 TOA 域的 \e 开关 [[24]]。
    /// </summary>
    string? EntrySeparator { get; set; }

    /// <summary>
    /// 获取或设置分隔引文目录中页码范围的字符（最多五个）。
    /// 例如，"12-15" 中的 "-" [[10]]。
    /// </summary>
    string? PageRangeSeparator { get; set; }

    /// <summary>
    /// 获取或设置一个书签的名称。如果设置了此属性，引文目录将仅包含该书签范围内的引文。
    /// </summary>
    string? Bookmark { get; set; }

    /// <summary>
    /// 获取或设置要包含在引文目录中的条目类别。
    /// 有效值为 1 到 16，对应于“引文目录”对话框中的类别列表 [[22]]。
    /// </summary>
    int Category { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否保留引文目录中条目的原始格式。
    /// </summary>
    bool KeepEntryFormatting { get; set; }

    /// <summary>
    /// 获取或设置在引文目录条目文本和页码之间使用的分隔符。
    /// </summary>
    string Separator { get; set; }

    /// <summary>
    /// 获取或设置要包含在引文目录中的序列名称。
    /// </summary>
    string IncludeSequenceName { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在引文目录中包含类别标题。
    /// </summary>
    bool IncludeCategoryHeader { get; set; }

    /// <summary>
    /// 获取或设置用于分隔页码的字符。
    /// </summary>
    string PageNumberSeparator { get; set; }

    /// <summary>
    /// 获取或设置引文目录中使用的制表符前导符类型。
    /// </summary>
    WdTabLeader TabLeader { get; set; }

    /// <summary>
    /// 更新引文目录中的所有条目，包括页码和条目文本。
    /// </summary>
    void Update();

    /// <summary>
    /// 从文档中删除此引文目录。
    /// </summary>
    void Delete();
}