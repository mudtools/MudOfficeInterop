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
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordDocuments : IDisposable, IOfficeObject<IWordDocuments, MsWord.Documents>, IEnumerable<IWordDocument?>
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
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
    /// 通过索引获取指定的文档。
    /// </summary>
    /// <param name="index">文档的序号位置或表示文档名称的字符串。</param>
    /// <returns>指定索引处的文档对象。</returns>
    IWordDocument? this[int index] { get; }

    /// <summary>
    /// 通过索引获取指定的文档。
    /// </summary>
    /// <param name="name">文档的序号位置或表示文档名称的字符串。</param>
    /// <returns>指定索引处的文档对象。</returns>
    IWordDocument? this[string name] { get; }


    /// <summary>
    /// 关闭指定的文档或文档集合。
    /// </summary>
    /// <param name="saveChanges">指定文档的保存操作。可以是以下 WdSaveOptions 常量之一：
    /// wdDoNotSaveChanges（不保存更改）、wdPromptToSaveChanges（提示保存更改）、wdSaveChanges（保存更改）。</param>
    /// <param name="originalFormat">指定文档的保存格式。可以是以下 WdOriginalFormat 常量之一：
    /// wdOriginalDocumentFormat（原始文档格式）、wdPromptUser（提示用户）、wdWordDocument（Word 文档格式）。</param>
    /// <param name="routeDocument">True 将文档路由到下一个收件人。如果文档没有附加路由单，则忽略此参数。</param>
    void Close(WdSaveOptions? saveChanges = null, WdOriginalFormat? originalFormat = null, bool? routeDocument = null);

    /// <summary>
    /// 保存 Microsoft.Office.Interop.Word._Application.Documents 集合中的所有文档。
    /// </summary>
    /// <param name="noPrompt">True 表示 Microsoft Word 自动保存所有文档；
    /// False 表示 Word 提示用户保存自上次保存以来已更改的每个文档。</param>
    /// <param name="originalFormat">指定文档的保存方式。可以是任何 WdOriginalFormat 常量。</param>
    void Save(bool? noPrompt = null, WdOriginalFormat? originalFormat = null);

    /// <summary>
    /// 返回表示添加到已打开文档集合中的新空白文档的 Document 对象。
    /// </summary>
    /// <param name="template">用于新文档的模板名称。如果省略此参数，则使用 Normal 模板。</param>
    /// <param name="newTemplate">True 将文档作为模板打开。默认值为 False。</param>
    /// <param name="documentType">可以是以下 WdNewDocumentType 常量之一：
    /// wdNewBlankDocument（新空白文档）、wdNewEmailMessage（新电子邮件）、wdNewFrameset（新框架集）、wdNewWebPage（新网页）。
    /// 默认常量为 wdNewBlankDocument。</param>
    /// <param name="visible">True 在可见窗口中打开文档。如果此值为 False，Microsoft Word 打开文档但将文档窗口的 Visible 属性设置为 False。
    /// 默认值为 True。</param>
    /// <returns>新创建的文档对象。</returns>
    IWordDocument? Add(string? template = null, bool? newTemplate = null, WdNewDocumentType? documentType = null, bool? visible = null);

    /// <summary>
    /// 将服务器上的指定文档签出到本地计算机进行编辑。
    /// </summary>
    /// <param name="fileName">要签出的文件名。</param>
    void CheckOut(string fileName);

    /// <summary>
    /// 确定 Microsoft Word 是否可以从服务器签出指定的文档。可读/写 Boolean。
    /// </summary>
    /// <param name="fileName">文档的服务器路径和名称。</param>
    /// <returns>如果可以签出文档，则为 True；否则为 False。</returns>
    bool? CanCheckOut(string fileName);

    /// <summary>
    /// 打开指定的文档并将其添加到 Documents 集合。
    /// </summary>
    /// <param name="fileName">文档的名称（接受路径）。</param>
    /// <param name="confirmConversions">如果文件不是 Microsoft Word 格式，则为 True 显示"转换文件"对话框。</param>
    /// <param name="readOnly">True 以只读方式打开文档。
    /// 注意：此参数不会覆盖已保存文档上的"建议只读"设置。
    /// 例如，如果文档已保存并启用了"建议只读"设置，将 ReadOnly 参数设置为 False 不会导致文件以读/写方式打开。</param>
    /// <param name="addToRecentFiles">True 将文件名添加到"文件"菜单底部的最近使用文件列表中。</param>
    /// <param name="passwordDocument">打开文档的密码。</param>
    /// <param name="passwordTemplate">打开模板的密码。</param>
    /// <param name="revert">控制如果 FileName 是已打开文档的名称时的行为。
    /// True 放弃对已打开文档的任何未保存更改并重新打开文件。
    /// False 激活已打开的文档。</param>
    /// <param name="writePasswordDocument">保存文档更改的密码。</param>
    /// <param name="writePasswordTemplate">保存模板更改的密码。</param>
    /// <param name="format">用于打开文档的文件转换器。可以是 WdOpenFormat 常量。
    /// 要指定外部文件格式，请将 FileConverter.OpenFormat 属性应用于 FileConverter 对象以确定用于此参数的值。</param>
    /// <param name="encoding">查看保存文档时 Microsoft Word 使用的文档编码（代码页或字符集）。
    /// 可以是任何有效的 MsoEncoding 常量。默认值为系统代码页。</param>
    /// <param name="visible">True 在可见窗口中打开文档。默认值为 True。</param>
    /// <param name="openAndRepair">True 修复文档以防止文档损坏。</param>
    /// <param name="documentDirection">可以是 WdDocumentDirection 常量。</param>
    /// <param name="noEncodingDialog">如果无法识别文本编码，则为 True 跳过显示 Word 显示的编码对话框。默认值为 False。</param>
    /// <param name="xmlTransform">指定要使用的转换。</param>
    /// <returns>打开的文档对象。</returns>
    IWordDocument? Open(string? fileName, bool? confirmConversions = null, bool? readOnly = null,
                        bool? addToRecentFiles = null, string? passwordDocument = null, string? passwordTemplate = null,
                        bool? revert = null, string? writePasswordDocument = null, string? writePasswordTemplate = null,
                        WdOpenFormat? format = null, [ComNamespace("MsCore")] MsoEncoding? encoding = null, bool? visible = null, bool? openAndRepair = null,
                        WdDocumentDirection? documentDirection = null, bool? noEncodingDialog = null, object? xmlTransform = null);

    /// <summary>
    /// 打开指定的文档并将其添加到 Documents 集合。
    /// </summary>
    /// <param name="fileName">文档的名称（接受路径）。</param>
    /// <param name="confirmConversions">如果文件不是 Microsoft Word 格式，则为 True 显示"转换文件"对话框。</param>
    /// <param name="readOnly">True 以只读方式打开文档。此参数不会覆盖已保存文档上的"建议只读"设置。
    /// 例如，如果文档已保存并启用了"建议只读"设置，将 ReadOnly 参数设置为 False 不会导致文件以读/写方式打开。</param>
    /// <param name="addToRecentFiles">True 将文件名添加到"文件"菜单底部的最近使用文件列表中。</param>
    /// <param name="passwordDocument">打开文档的密码。</param>
    /// <param name="passwordTemplate">打开模板的密码。</param>
    /// <param name="revert">控制如果 FileName 是已打开文档的名称时的行为。
    /// True 放弃对已打开文档的任何未保存更改并重新打开文件。
    /// False 激活已打开的文档。</param>
    /// <param name="writePasswordDocument">保存文档更改的密码。</param>
    /// <param name="writePasswordTemplate">保存模板更改的密码。</param>
    /// <param name="format">用于打开文档的文件转换器。可以是 WdOpenFormat 常量之一。
    /// 默认值为 WdOpenFormat.wdOpenFormatAuto。</param>
    /// <param name="encoding">查看保存文档时 Microsoft Word 使用的文档编码（代码页或字符集）。
    /// 可以是任何有效的 MsoEncoding 枚举值。默认值为系统代码页。</param>
    /// <param name="visible">True 在可见窗口中打开文档。默认值为 True。</param>
    /// <param name="openAndRepair">True 修复文档以防止文档损坏。</param>
    /// <param name="documentDirection">指示文档中文本的水平流向。可以是任何有效的 WdDocumentDirection 常量。
    /// 默认值为 WdDocumentDirection.wdLeftToRight。</param>
    /// <param name="noEncodingDialog">如果无法识别文本编码，则为 True 跳过显示 Word 显示的编码对话框。默认值为 False。</param>
    /// <param name="xmlTransform">指定要使用的转换。</param>
    /// <returns>打开的文档对象。</returns>
    IWordDocument? OpenNoRepairDialog(string fileName, bool? confirmConversions = null, bool? readOnly = null,
                                      bool? addToRecentFiles = null, string? passwordDocument = null,
                                      string? passwordTemplate = null, bool? revert = null, string? writePasswordDocument = null,
                                      string? writePasswordTemplate = null, WdOpenFormat? format = null, [ComNamespace("MsCore")] MsoEncoding? encoding = null,
                                      bool? visible = null, bool? openAndRepair = null, WdDocumentDirection? documentDirection = null,
                                      bool? noEncodingDialog = null, object? xmlTransform = null);

    /// <summary>
    /// 返回一个对象，该对象表示 Microsoft Office Word 发布到前三个参数描述的账户的新博客文档。
    /// </summary>
    /// <param name="providerID">提供者在 Word 中注册自身时使用的唯一值 GUID。</param>
    /// <param name="postURL">用于向博客添加帖子的 URL。</param>
    /// <param name="blogName">将在 Word 中使用的博客的显示名称。</param>
    /// <param name="postID">用于填充使用 AddBlogDocument 方法创建的文档的现有帖子的 ID。</param>
    /// <returns>新创建的博客文档对象。</returns>
    IWordDocument? AddBlogDocument(string providerID, string postURL, string blogName, string postID = "");
}