//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft.Office.Core._CustomXMLParts 集合中的单个自定义 XML 部分。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeCustomXMLPart : IDisposable
{
    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取文档中绑定数据区域的根元素。如果区域为空，该属性返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? DocumentElement { get; }

    /// <summary>
    /// 获取包含分配给当前 Microsoft.Office.Core._CustomXMLPart 对象的 GUID 的字符串。只读。
    /// </summary>
    /// <returns>String</returns>
    string Id { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core._CustomXMLPart 对象命名空间的唯一地址标识符。只读。
    /// </summary>
    /// <returns>String</returns>
    string NamespaceURI { get; }

    /// <summary>
    /// 获取或设置表示附加到文档中绑定数据区域的模式集合的 CustomXMLSchemaCollection 对象。可读写。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLSchemaCollection</returns>
    IOfficeCustomXMLSchemaCollection? SchemaCollection { get; set; }

    /// <summary>
    /// 获取用于当前 Microsoft.Office.Core._CustomXMLPart 对象的命名空间前缀映射集合。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLPrefixMappings</returns>
    IOfficeCustomXMLPrefixMappings? NamespaceManager { get; }

    /// <summary>
    /// 获取当前 Microsoft.Office.Core._CustomXMLPart 对象的 XML 表示形式。只读。
    /// </summary>
    /// <returns>String</returns>
    string XML { get; }

    /// <summary>
    /// 获取提供对任何 XML 验证错误的访问的 CustomXMLValidationErrors 对象。如果没有验证错误，此属性返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLValidationErrors</returns>
    IOfficeCustomXMLValidationErrors? Errors { get; }

    /// <summary>
    /// 获取一个值，指示 Microsoft.Office.Core._CustomXMLPart 是否是内置的。只读。
    /// </summary>
    /// <returns>Boolean</returns>
    bool BuiltIn { get; }

    /// <summary>
    /// 将节点添加到 XML 树中。
    /// </summary>
    /// <param name="parent">表示应在其下添加此节点的节点。如果添加属性，该参数表示应添加属性的元素。</param>
    /// <param name="name">表示要添加的节点的基本名称。</param>
    /// <param name="namespaceURI">表示要追加元素的命名空间。对于追加类型为 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 的节点，此参数是必需的，否则将被忽略。</param>
    /// <param name="nextSibling">表示应成为新节点的下一个兄弟节点的节点。如果未指定，节点将添加到父节点的子节点末尾。对于类型为 msoCustomXMLNodeAttribute 的添加，此参数将被忽略。如果节点不是父节点的子节点，将显示错误。</param>
    /// <param name="nodeType">指定要追加的节点类型。如果未指定参数，则假定为 msoCustomXMLNodeElement 类型。</param>
    /// <param name="nodeValue">用于为允许文本的节点设置追加节点的值。如果节点不允许文本，此参数将被忽略。</param>
    void AddNode(IOfficeCustomXMLNode parent, string name = "", string namespaceURI = "", IOfficeCustomXMLNode nextSibling = null, MsoCustomXMLNodeType nodeType = MsoCustomXMLNodeType.msoCustomXMLNodeElement, string nodeValue = "");

    /// <summary>
    /// 从数据存储（IXMLDataStore 接口）中删除当前 Microsoft.Office.Core._CustomXMLPart。
    /// </summary>
    void Delete();

    /// <summary>
    /// 允许模板作者从现有文件填充 Microsoft.Office.Core._CustomXMLPart。如果加载成功，则返回 True。
    /// </summary>
    /// <param name="filePath">指向用户计算机或网络上包含要加载的 XML 的文件。</param>
    /// <returns>Boolean</returns>
    bool? Load(string filePath);

    /// <summary>
    /// 允许模板作者从 XML 字符串填充 Microsoft.Office.Core._CustomXMLPart 对象。如果加载成功，则返回 True。
    /// </summary>
    /// <param name="xml">包含要加载的 XML。</param>
    /// <returns>Boolean</returns>
    bool? LoadXML(string xml);

    /// <summary>
    /// 从自定义 XML 部分中选择节点集合。
    /// </summary>
    /// <param name="xpath">包含 XPath 表达式。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLNodes</returns>
    IOfficeCustomXMLNodes? SelectNodes(string xpath);

    /// <summary>
    /// 在自定义 XML 部分中选择与 XPath 表达式匹配的单个节点。
    /// </summary>
    /// <param name="xpath">包含 XPath 表达式。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLNodes</returns>
    IOfficeCustomXMLNode? SelectSingleNode(string xpath);
}