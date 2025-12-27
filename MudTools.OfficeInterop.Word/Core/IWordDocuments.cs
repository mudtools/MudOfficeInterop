//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// Word 文档集合接口
/// </summary>
public interface IWordDocuments : IDisposable, IEnumerable<IWordDocument>
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取文档数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }


    /// <summary>
    /// 根据索引获取文档
    /// </summary>
    IWordDocument this[int index] { get; }

    /// <summary>
    /// 根据名称获取文档
    /// </summary>
    IWordDocument this[string name] { get; }

    /// <summary>
    /// 创建新文档
    /// </summary>
    /// <param name="template">模板路径（可选）</param>
    /// <returns>新文档对象</returns>
    IWordDocument Add(string? template = null);

    /// <summary>
    /// 打开现有文档
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="readOnly">是否只读打开</param>
    /// <param name="password">密码（可选）</param>
    /// <returns>文档对象</returns>
    IWordDocument Open(string fileName, bool readOnly = false, string? password = null);

    /// <summary>
    /// 打开一个现有文档。
    /// </summary>
    /// <param name="fileName">要打开的文档的文件名。</param>
    /// <param name="confirmConversions">如果为 true，则在文件不是 Word 格式时显示“转换文件”对话框。</param>
    /// <param name="readOnly">如果为 true，则以只读方式打开文档。</param>
    /// <param name="addToRecentFiles">如果为 true，则将文件添加到最近使用的文件列表中。</param>
    /// <param name="passwordDocument">打开文档所需的密码。</param>
    /// <param name="passwordTemplate">打开模板所需的密码。</param>
    /// <param name="revert">如果为 true，则将文档恢复到上次保存的版本。</param>
    /// <param name="writePasswordDocument">保存对文档所做的更改所需的密码。</param>
    /// <param name="writePasswordTemplate">保存对模板所做的更改所需的密码。</param>
    /// <param name="format">文档的格式。</param>
    /// <param name="encoding">文档的编码。</param>
    /// <param name="visible">如果为 true，则打开文档时使其可见。</param>
    /// <returns>打开的文档对象。</returns>
    IWordDocument? OpenDocument(string fileName, bool confirmConversions = true, bool readOnly = false, bool addToRecentFiles = true,
                                     string passwordDocument = "", string passwordTemplate = "", bool revert = true, string writePasswordDocument = "",
                                     string writePasswordTemplate = "", WdOpenFormat format = WdOpenFormat.wdOpenFormatAuto,
                                     MsoEncoding encoding = MsoEncoding.msoEncodingSimplifiedChineseAutoDetect, bool visible = true);

    /// <summary>
    /// 获取活动文档
    /// </summary>
    /// <returns>活动文档对象</returns>
    IWordDocument GetActiveDocument();

    /// <summary>
    /// 关闭所有文档
    /// </summary>
    void Close(WdSaveOptions saveChanges = WdSaveOptions.wdSaveChanges,
                     WdOriginalFormat originalFormat = WdOriginalFormat.wdWordDocument,
                     bool? routeDocument = null);

    /// <summary>
    /// 关闭所有文档
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    void Close(bool saveChanges = true);

    /// <summary>
    /// 保存所有文档
    /// </summary>
    void SaveAll();
}