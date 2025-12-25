//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的一个引文目录 (Table of Authorities, TOA) 的二次封装接口。
/// 此接口提供了对引文目录内容、格式和操作的访问，同时管理底层 COM 对象的生命周期。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTableOfAuthorities : IDisposable
{
    /// <summary>
    /// 获取此引文目录所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此引文目录的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示如果同一引文有五个或更多页引用，是否用"Passim"替换。
    /// </summary>
    bool Passim { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否将引文目录条目中的格式应用于指定引文目录中的条目。
    /// </summary>
    bool KeepEntryFormatting { get; set; }

    /// <summary>
    /// 获取或设置要包含在引文目录中的条目类别。
    /// </summary>
    int Category { get; set; }

    /// <summary>
    /// 获取或设置要从中收集引文目录条目的书签名称。
    /// </summary>
    string Bookmark { get; set; }

    /// <summary>
    /// 获取或设置序列号和页码之间的字符（最多五个字符）。
    /// </summary>
    string Separator { get; set; }

    /// <summary>
    /// 获取或设置引文目录的序列（SEQ）字段标识符。
    /// </summary>
    string IncludeSequenceName { get; set; }

    /// <summary>
    /// 获取或设置引文目录条目及其页码之间的字符（最多五个字符）。
    /// </summary>
    string EntrySeparator { get; set; }

    /// <summary>
    /// 获取或设置引文目录中页面范围之间的字符（最多五个字符）。
    /// </summary>
    string PageRangeSeparator { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示引文目录中是否显示一组条目的类别名称。
    /// </summary>
    bool IncludeCategoryHeader { get; set; }

    /// <summary>
    /// 获取或设置引文目录中各个页引用之间的字符（最多五个字符）。
    /// </summary>
    string PageNumberSeparator { get; set; }

    /// <summary>
    /// 获取或设置引文目录中条目及其页码之间的制表符前导符。
    /// </summary>
    WdTabLeader TabLeader { get; set; }

    /// <summary>
    /// 返回表示包含在指定对象中的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 更新指定引文目录中显示的条目。
    /// </summary>
    void Update();
}