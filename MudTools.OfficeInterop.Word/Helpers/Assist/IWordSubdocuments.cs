//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档中所有子文档的集合。
/// </summary>
public interface IWordSubdocuments : IEnumerable<IWordSubdocument>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的子文档数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取子文档。
    /// </summary>
    /// <param name="index">从 1 开始的索引。</param>
    IWordSubdocument this[int index] { get; }

    /// <summary>
    /// 添加一个新的子文档。
    /// </summary>
    /// <param name="name">子文档的文件名。</param>
    /// <param name="confirmConversions">是否确认文件转换。</param>
    /// <param name="readOnly">是否以只读方式打开。</param>
    /// <param name="passwordDocument">文档密码。</param>
    /// <param name="passwordTemplate">模板密码。</param>
    /// <param name="revert">是否恢复原始版本。</param>
    /// <param name="writePasswordDocument">写入密码。</param>
    /// <param name="writePasswordTemplate">写入模板密码。</param>
    /// <returns>新创建的子文档。</returns>
    IWordSubdocument AddFromFile(string name, bool confirmConversions, bool readOnly, object passwordDocument,
                        string passwordTemplate, bool revert, string writePasswordDocument, string writePasswordTemplate);

    IWordSubdocument AddFromRange(IWordRange range);

    /// <summary>
    /// 删除所有子文档。
    /// </summary>
    void Delete();

    /// <summary>
    /// 展开所有子文档。
    /// </summary>
    void Select();
}
