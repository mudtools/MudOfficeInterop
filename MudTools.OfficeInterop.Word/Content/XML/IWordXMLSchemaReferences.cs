//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示附加到文档的唯一命名空间的 Microsoft.Office.Interop.Word.XMLSchemaReference 对象集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordXMLSchemaReferences : IEnumerable<IWordXMLSchemaReference?>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中建议的数量
    /// </summary>
    /// <returns>集合中建议的总数</returns>
    int Count { get; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="index">指示序号位置的 Object 或表示单个对象名称的字符串。</param>
    /// <returns>Microsoft.Office.Interop.Word.XMLSchemaReference</returns>
    IWordXMLSchemaReference? this[int index] { get; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="name">指示序号位置的 Object 或表示单个对象名称的字符串。</param>
    /// <returns>Microsoft.Office.Interop.Word.XMLSchemaReference</returns>
    IWordXMLSchemaReference? this[string name] { get; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Word 是否会在用户输入时验证文档中的 XML。
    /// </summary>
    bool AutomaticValidation { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Word 是否在保存文档时验证文档中的 XML。
    /// </summary>
    bool AllowSaveAsXMLWithoutValidation { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示当 Microsoft Word 验证文档中的 XML 时是否隐藏架构违规。
    /// </summary>
    bool HideValidationErrors { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示 Microsoft Word 是否对具有元素兄弟的文本节点执行验证，并指定当 Microsoft.Office.Interop.Word._Document.XMLSaveDataOnly 属性为 True 时是否将这些文本节点保存在 XML 中。
    /// </summary>
    bool IgnoreMixedContent { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否为文档中的 XML 元素显示自动占位符文本。True 显示占位符文本，False 隐藏占位符文本。
    /// </summary>
    bool ShowPlaceholderText { get; set; }

    /// <summary>
    /// 根据附加的 XML 架构验证单个 XML 元素或整个文档。
    /// </summary>
    void Validate();

    /// <summary>
    /// 返回表示应用于文档的架构的 XMLSchemaReference。
    /// </summary>
    /// <param name="namespaceURI">可选 String。架构中定义的架构名称。Namespace 参数区分大小写，必须完全按照其在架构中出现的拼写。如果在附加到文档的任何架构中找不到指定的命名空间，将显示错误。</param>
    /// <param name="alias">可选 String。架构在"模板和加载项"对话框的"架构"选项卡中显示的名称。</param>
    /// <param name="fileName">可选 String。架构的路径和文件名。可以是本地文件路径、网络路径或 Internet 地址。</param>
    /// <param name="installForAllUsers">可选 Boolean。如果登录到计算机的所有用户都可以访问和使用新架构，则为 True。默认值为 False。</param>
    /// <returns>Microsoft.Office.Interop.Word.XMLSchemaReference</returns>
    IWordXMLSchemaReference? Add(string? namespaceURI = null, string? alias = null, string? fileName = null, bool installForAllUsers = false);
}