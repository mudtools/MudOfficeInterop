//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示一个书目源（引文）。
/// <para>注：Source 对象是 Sources 集合的成员。</para>
/// <para>注：使用 Sources(index) 可返回单个 Source 对象，其中 index 是源的索引号（从 1 开始）或 Tag 值（字符串）。</para>
/// <para>注：此接口基于对 Word 2007 及更高版本中引文和书目功能的理解实现，因为官方 SDK 文档信息有限。</para>
/// </summary>
public interface IWordSource : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #endregion

    #region 书目源属性 (Bibliography Source Properties)

    /// <summary>
    /// 获取或设置书目源的唯一标识标签。
    /// </summary>
    string Tag { get; }

    /// <summary>
    /// 获取书目源的数据，通常以 XML 字符串形式表示。
    /// </summary>
    string XML { get; }

    /// <summary>
    /// 获取或设置书目源的类型（例如，“书籍”、“期刊文章”等）。
    /// </summary>
    bool Cited { get; }
    #endregion

    #region 书目源方法 (Bibliography Source Methods)

    /// <summary>
    /// 删除指定的书目源。
    /// </summary>
    void Delete();
    #endregion
}