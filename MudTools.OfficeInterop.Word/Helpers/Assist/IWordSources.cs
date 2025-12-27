//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示文档中所有书目源（引文）的对象集合。
/// <para>注：使用 Application.Bibliography.Sources 属性可返回 Sources 集合。</para>
/// <para>注：使用 Sources(index)（其中 index 是源的索引号或 Tag 值）可返回单个 Source 对象。</para>
/// <para>注：此接口基于对 Word 2007 及更高版本中引文和书目功能的理解实现，因为官方 SDK 文档信息有限。</para>
/// </summary>
public interface IWordSources : IEnumerable<IWordSource>, IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取集合中的书目源数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引号（从 1 开始）或 Tag 值（字符串）获取单个书目源。
    /// </summary>
    /// <param name="index">索引号（整数）或 Tag 值（字符串）。</param>
    /// <returns>指定的书目源对象。</returns>
    IWordSource this[int index] { get; }

    #endregion

    #region 书目源集合方法 (Bibliography Sources Collection Methods)

    /// <summary>
    /// 将新的书目源添加到集合中。
    /// </summary>
    /// <param name="data">表示新书目源数据的 XML 字符串。</param>
    /// <returns>表示添加的书目源的对象。</returns>
    void Add(string data);
    #endregion
}