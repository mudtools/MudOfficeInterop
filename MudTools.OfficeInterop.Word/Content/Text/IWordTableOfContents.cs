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
public interface IWordTableOfContents : IDisposable
{
    /// <summary>
    /// 获取此目录所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此目录的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此目录在文档中所占据的范围 (<see cref="IWordRange"/>)。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取目录使用的标题样式集合。
    /// 标题样式集合包含除内置"标题 1-9"外，用于编译目录的其他样式。
    /// </summary>
    IWordHeadingStyles? HeadingStyles { get; }

    /// <summary>
    /// 获取或设置目录的唯一标识符。
    /// </summary>
    string TableID { get; set; }

    /// <summary>
    /// 获取或设置目录中包含的最低（最详细）标题级别。
    /// </summary>
    int LowerHeadingLevel { get; set; }

    /// <summary>
    /// 获取或设置目录中包含的最高（最一般）标题级别。
    /// </summary>
    int UpperHeadingLevel { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示目录是否基于字段。
    /// </summary>
    bool UseFields { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示目录是否使用标题样式来确定要包含的段落。
    /// </summary>
    bool UseHeadingStyles { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示页码是否在目录中右对齐。
    /// </summary>
    bool RightAlignPageNumbers { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示在 Web 视图中是否隐藏页码。
    /// </summary>
    bool HidePageNumbersInWeb { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示目录是否使用超链接。
    /// </summary>
    bool UseHyperlinks { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示目录是否包含页码。
    /// </summary>
    bool IncludePageNumbers { get; set; }

    /// <summary>
    /// 获取或设置目录的右对齐页码分隔符。
    /// </summary>
    WdTabLeader TabLeader { get; set; }

    /// <summary>
    /// 更新目录中的所有条目，包括页码和标题文本。
    /// </summary>
    void Update();

    /// <summary>
    /// 仅更新目录中的页码。
    /// </summary>
    void UpdatePageNumbers();

    /// <summary>
    /// 从文档中删除此目录。
    /// </summary>
    void Delete();
}