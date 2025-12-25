//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中的一个目录 (Table of Contents, TOC) 的二次封装接口。
/// 此接口提供了对目录内容、样式和操作的访问，同时管理底层 COM 对象的生命周期。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTableOfContents : IDisposable
{
    /// <summary>
    /// 获取此目录所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此目录的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }
    /// <summary>
    /// 获取或设置一个值，指示是否使用内置标题样式创建目录。
    /// </summary>
    bool UseHeadingStyles { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用目录条目（TC）字段创建目录。
    /// </summary>
    bool UseFields { get; set; }

    /// <summary>
    /// 获取或设置目录的起始标题级别。
    /// </summary>
    int UpperHeadingLevel { get; set; }

    /// <summary>
    /// 获取或设置目录的结束标题级别。
    /// </summary>
    int LowerHeadingLevel { get; set; }

    /// <summary>
    /// 获取或设置用于从 TOC 字段构建目录的一个字母标识符。
    /// </summary>
    string TableID { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在目录中页码是否右对齐。
    /// </summary>
    bool RightAlignPageNumbers { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示目录中是否包含页码。
    /// </summary>
    bool IncludePageNumbers { get; set; }

    /// <summary>
    /// 获取或设置目录中条目及其页码之间的字符。
    /// </summary>
    WdTabLeader TabLeader { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在发布到 Web 时目录条目是否应格式化为超链接。
    /// </summary>
    bool UseHyperlinks { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在发布到 Web 时是否隐藏目录中的页码。
    /// </summary>
    bool HidePageNumbersInWeb { get; set; }

    /// <summary>
    /// 返回表示用于编译目录的附加样式（Heading 1 – Heading 9 样式之外的样式）的 HeadingStyles 对象。
    /// </summary>
    IWordHeadingStyles? HeadingStyles { get; }

    /// <summary>
    /// 返回表示包含在指定对象中的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 更新指定目录中项目的页码。
    /// </summary>
    void UpdatePageNumbers();

    /// <summary>
    /// 更新指定目录中显示的条目。
    /// </summary>
    void Update();
}