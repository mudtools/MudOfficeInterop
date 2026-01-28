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
public interface IOfficeCustomXMLNode : IOfficeObject<IOfficeCustomXMLNode, MsCore.CustomXMLNode>, IDisposable
{
    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示当前节点中当前元素属性的 CustomXMLNode 集合。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNodes</returns>
    IOfficeCustomXMLNodes? Attributes { get; }

    /// <summary>
    /// 获取不带命名空间前缀的节点基本名称（如果存在）在文档对象模型（DOM）中。只读。
    /// </summary>
    /// <returns>String</returns>
    string BaseName { get; }

    /// <summary>
    /// 获取包含当前节点所有子元素的 CustomXMLNodes 集合。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNodes</returns>
    IOfficeCustomXMLNodes? ChildNodes { get; }

    /// <summary>
    /// 获取对应于当前节点第一个子元素的 CustomXMLNode 对象。如果节点没有子元素（或者不是 msoCustomXMLNodeElement 类型），返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? FirstChild { get; }

    /// <summary>
    /// 获取对应于当前节点最后一个子元素的 CustomXMLNode 对象。如果节点没有子元素（或者不是 msoCustomXMLNodeElement 类型），返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? LastChild { get; }

    /// <summary>
    /// 获取 Microsoft.Office.Core.CustomXMLNode 对象命名空间的唯一地址标识符。只读。
    /// </summary>
    /// <returns>String</returns>
    string NamespaceURI { get; }

    /// <summary>
    /// 获取当前节点的下一个兄弟节点（元素、注释或处理指令）。如果节点是其级别的最后一个兄弟节点，该属性返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? NextSibling { get; }

    /// <summary>
    /// 获取当前节点的类型。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.MsoCustomXMLNodeType</returns>
    MsoCustomXMLNodeType NodeType { get; }

    /// <summary>
    /// 获取或设置当前节点的值。可读写。
    /// </summary>
    /// <returns>String</returns>
    string NodeValue { get; set; }

    /// <summary>
    /// 获取表示与此节点关联的 Microsoft Office Excel 工作簿、Microsoft Office PowerPoint 演示文稿或 Microsoft Office Word 文档的对象。只读。
    /// </summary>
    /// <returns>Object</returns>
    object OwnerDocument { get; }

    /// <summary>
    /// 获取表示与此节点关联的部分的对象。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core._CustomXMLPart</returns>
    IOfficeCustomXMLPart? OwnerPart { get; }

    /// <summary>
    /// 获取当前节点的前一个兄弟节点（元素、注释或处理指令）。如果当前节点是其级别的第一个兄弟节点，该属性返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? PreviousSibling { get; }

    /// <summary>
    /// 获取当前节点的父元素节点。如果当前节点在根级别，该属性返回 Nothing。只读。
    /// </summary>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? ParentNode { get; }

    /// <summary>
    /// 获取或设置当前节点的文本。可读写。
    /// </summary>
    /// <returns>String</returns>
    string Text { get; set; }

    /// <summary>
    /// 获取当前节点的规范化 XPath 字符串。如果节点不再存在于文档对象模型（DOM）中，该属性返回错误消息。只读。
    /// </summary>
    /// <returns>String</returns>
    string XPath { get; }

    /// <summary>
    /// 获取当前节点及其子节点（如果存在）的 XML 表示形式。只读。
    /// </summary>
    /// <returns>String</returns>
    string XML { get; }

    /// <summary>
    /// 将单个节点作为最后一个子节点追加到树中上下文元素节点下。
    /// </summary>
    /// <param name="name">表示要追加元素的基本名称。</param>
    /// <param name="namespaceURI">表示要追加元素的命名空间。对于追加类型为 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 的节点，此参数是必需的，否则将被忽略。</param>
    /// <param name="nodeType">指定要追加的节点类型。如果未指定参数，则假定为 msoCustomXMLNodeElement 类型。</param>
    /// <param name="nodeValue">用于为允许文本的节点设置追加节点的值。如果节点不允许文本，此参数将被忽略。</param>
    void AppendChildNode(string name = "", string namespaceURI = "", MsoCustomXMLNodeType nodeType = MsoCustomXMLNodeType.msoCustomXMLNodeElement, string nodeValue = "");

    /// <summary>
    /// 将子树作为最后一个子节点添加到树中上下文元素节点下。
    /// </summary>
    /// <param name="xml">表示要添加的子树。</param>
    void AppendChildSubtree(string xml);

    /// <summary>
    /// 从树中删除当前节点（包括其所有子节点，如果存在）。
    /// </summary>
    void Delete();

    /// <summary>
    /// 如果当前元素节点有子元素节点，则返回 True。
    /// </summary>
    /// <returns>Boolean</returns>
    bool? HasChildNodes();

    /// <summary>
    /// 在树中刚好在上下文节点之前插入新节点。
    /// </summary>
    /// <param name="name">表示要添加节点的基本名称。</param>
    /// <param name="namespaceURI">表示要添加元素的命名空间。对于添加类型为 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 的节点，此参数是必需的，否则将被忽略。</param>
    /// <param name="nodeType">指定要添加节点的类型。如果未指定参数，则假定为 msoCustomXMLNodeElement 类型的节点。</param>
    /// <param name="nodeValue">用于为允许文本的节点设置要添加节点的值。如果节点不允许文本，此参数将被忽略。</param>
    /// <param name="nextSibling">表示上下文节点。</param>
    void InsertNodeBefore(string name = "", string namespaceURI = "", MsoCustomXMLNodeType nodeType = MsoCustomXMLNodeType.msoCustomXMLNodeElement, string nodeValue = "", IOfficeCustomXMLNode nextSibling = null);

    /// <summary>
    /// 在刚好在上下文节点之前的位置插入指定的子树。
    /// </summary>
    /// <param name="xml">表示要添加的子树。</param>
    /// <param name="nextSibling">指定上下文节点。</param>
    void InsertSubtreeBefore(string xml, IOfficeCustomXMLNode nextSibling = null);

    /// <summary>
    /// 从树中移除指定的子节点。
    /// </summary>
    /// <param name="child">表示上下文节点的子节点。</param>
    void RemoveChild(IOfficeCustomXMLNode child);

    /// <summary>
    /// 从主树中移除指定的子节点（及其子树），并在相同位置用不同的节点替换它。
    /// </summary>
    /// <param name="oldNode">表示要替换的子节点。</param>
    /// <param name="name">表示要添加元素的基本名称。</param>
    /// <param name="namespaceURI">表示要添加元素的命名空间。对于添加类型为 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 的节点，此参数是必需的，否则将被忽略。</param>
    /// <param name="nodeType">指定要添加节点的类型。如果未指定参数，则假定为 msoCustomXMLNodeElement 类型。</param>
    /// <param name="nodeValue">用于为允许文本的节点设置要添加节点的值。如果节点不允许文本，此参数将被忽略。</param>
    void ReplaceChildNode(IOfficeCustomXMLNode oldNode, string name = "", string namespaceURI = "", MsoCustomXMLNodeType nodeType = MsoCustomXMLNodeType.msoCustomXMLNodeElement, string nodeValue = "");

    /// <summary>
    /// 从主树中移除指定的节点（及其子树），并在相同位置用不同的子树替换它。
    /// </summary>
    /// <param name="xml">表示要添加的子树。</param>
    /// <param name="oldNode">表示要替换的子节点。</param>
    void ReplaceChildSubtree(string xml, IOfficeCustomXMLNode oldNode);

    /// <summary>
    /// 选择匹配 XPath 表达式的节点集合。此方法与 Microsoft.Office.Core._CustomXMLPart.SelectNodes(System.String) 方法的不同之处在于，XPath 表达式将以'expression'节点作为上下文节点开始计算。
    /// </summary>
    /// <param name="xpath">包含 XPath 表达式。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLNodes</returns>
    IOfficeCustomXMLNodes? SelectNodes(string xpath);

    /// <summary>
    /// 从匹配 XPath 表达式的集合中选择单个节点。此方法与 Microsoft.Office.Core._CustomXMLPart.SelectSingleNode(System.String) 方法的不同之处在于，XPath 表达式将以'expression'节点作为上下文节点开始计算。
    /// </summary>
    /// <param name="xpath">包含 XPath 表达式。</param>
    /// <returns>Microsoft.Office.Core.CustomXMLNode</returns>
    IOfficeCustomXMLNode? SelectSingleNode(string xpath);
}