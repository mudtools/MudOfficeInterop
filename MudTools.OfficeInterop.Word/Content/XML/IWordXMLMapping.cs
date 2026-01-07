//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 ContentControl 对象上的 XML 映射，用于在自定义 XML 和内容控件之间建立关联。
/// XML 映射是内容控件中的文本与文档的自定义 XML 数据存储中的 XML 元素之间的链接。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordXMLMapping : IOfficeObject<IWordXMLMapping, MsWord.XMLMapping>, IDisposable
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
    /// 获取一个布尔值，该值表示文档中的内容控件是否映射到文档 XML 数据存储中的 XML 节点。只读。
    /// </summary>
    bool IsMapped { get; }

    /// <summary>
    /// 获取表示文档中的内容控件所映射到的自定义 XML 部分的 CustomXMLPart 对象。
    /// </summary>
    IOfficeCustomXMLPart? CustomXMLPart { get; }

    /// <summary>
    /// 获取表示数据存储中自定义 XML 节点的 CustomXMLNode 对象，文档中的内容控件映射到该节点。
    /// </summary>
    IOfficeCustomXMLNode? CustomXMLNode { get; }

    /// <summary>
    /// 允许创建或更改内容控件上的 XML 映射。
    /// 如果 Microsoft Word 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，则返回 True。
    /// </summary>
    /// <param name="xPath">指定表示要映射内容控件的 XML 节点的 XPath 字符串。无效的 XPath 字符串会导致运行时错误。</param>
    /// <param name="prefixMapping">指定查询 XPath 参数中提供的表达式时要使用的前缀映射。如果省略，Word 将使用当前文档中指定自定义 XML 部分的前缀映射集。</param>
    /// <param name="source">指定要映射内容控件的目标自定义 XML 数据。如果省略此参数，XPath 将针对当前文档中的所有自定义 XML 进行计算，并且映射将建立在第一个 CustomXMLPart 中，其中 XPath 解析为 XML 节点。</param>
    /// <returns>如果映射成功，则为 True；否则为 False。</returns>
    bool? SetMapping(string xPath, string prefixMapping = "", IOfficeCustomXMLPart source = null);

    /// <summary>
    /// 从父内容控件中删除 XML 映射。
    /// </summary>
    void Delete();

    /// <summary>
    /// 允许创建或更改内容控件上的 XML 数据映射。
    /// 如果 Microsoft Word 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，则返回 True。
    /// </summary>
    /// <param name="node">指定要映射当前内容控件的 XML 节点。</param>
    /// <returns>如果映射成功，则为 True；否则为 False。</returns>
    bool? SetMappingByNode(IOfficeCustomXMLNode node);

    /// <summary>
    /// 获取表示 XML 映射的 XPath 的字符串，该 XPath 计算结果为当前映射的 XML 节点。只读。
    /// </summary>
    string XPath { get; }

    /// <summary>
    /// 获取表示用于计算当前 XML 映射 XPath 的前缀映射的字符串。只读。
    /// </summary>
    string PrefixMappings { get; }
}